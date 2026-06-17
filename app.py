import io
import re
from copy import copy
from pathlib import Path
from urllib.parse import (
    urljoin,
    urlparse,
    parse_qs,
    urlencode,
    urlunparse,
    quote,
    unquote,
)
from io import StringIO

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.worksheet.datavalidation import DataValidation


# ============================================================
# SETTINGS
# ============================================================

BASE_DIR = Path(__file__).parent
TEMPLATE_XLSM_PATH = BASE_DIR / "IDR_Template.xlsm"
TEMPLATE_XLSX_PATH = BASE_DIR / "IDR_Template.xlsx"

if TEMPLATE_XLSM_PATH.exists():
    TEMPLATE_PATH = TEMPLATE_XLSM_PATH
else:
    TEMPLATE_PATH = TEMPLATE_XLSX_PATH

BASE_URL = "https://webapps1.dot.illinois.gov"
IDOT_HOME_URL = "https://webapps1.dot.illinois.gov/WCTB/LBHome"

HIDDEN_PAY_ITEMS_SHEET_NAME = "IDOT_Pay_Items"
JOB_INFO_SHEET_NAME = "IDOT_Job_Info"
OLD_MATERIALS_SHEET_NAME = "Materials Data"


SEARCH_MAX_PAGES_PER_LETTING = 25

CELL_MAP = {
    "date": "C6",
    "contractor_label": "B8",
    "contractor": "C8",
    "work_description": "C9",
    "weather": "C10",
    "remarks": "C25",

    "county": "L2",
    "section": "L3",
    "route": "L4",
    "district": "L5",
    "contract_number": "L6",
    "job_number": "L7",
    "project": "L8",

    "start_row": 13,
    "end_row": 18,
    "item_code_col": "B",
    "item_description_col": "D",
    "location_col": "F",

    # H is the quantity the user manually fills out.
    # I is the unit that the program/formula fills out.
    "quantity_col": "H",
    "unit_col": "I",
}

CONTRACTOR_LABEL_TEXT = "Contractor or Sub."

# ============================================================
# DESCRIPTION TEXT FIT SETTINGS
# ============================================================

# Font-only resizing for wrapped description cells. This does not change row height.
DESCRIPTION_TEXT_FIT_RULES = [
    (35, 10),
    (70, 9),
    (105, 8),
    (140, 7),
    (9999, 6),
]

UNIT_WORDS = [
    "CU YD",
    "CUYD",
    "SQ YD",
    "SQYD",
    "SQ FT",
    "SQFT",
    "FOOT",
    "EACH",
    "L SUM",
    "LSUM",
    "CAL DA",
    "CAL MO",
    "POUND",
    "HOUR",
    "TON",
    "GALLON",
    "ACRE",
    "UNIT",
    "SQ M",
    "METER",
    "LITER",
    "M GAL",
    "L FOOT",
]


# ============================================================
# BASIC WEB HELPERS
# ============================================================

def get_headers():
    return {
        "User-Agent": "Mozilla/5.0 IDR Generator",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }


def make_session():
    session = requests.Session()
    session.headers.update(get_headers())
    return session


def get_html(session, url):
    response = session.get(url, timeout=30)
    response.raise_for_status()
    return response.text


def get_response(session, url):
    response = session.get(url, timeout=30)
    response.raise_for_status()
    return response


def absolute_url(href, base=BASE_URL + "/WCTB/"):
    return urljoin(base, href)


def set_query_param(url, key, value):
    parsed = urlparse(url)
    query = parse_qs(parsed.query)
    query[key] = [str(value)]

    return urlunparse((
        parsed.scheme,
        parsed.netloc,
        parsed.path,
        parsed.params,
        urlencode(query, doseq=True),
        parsed.fragment,
    ))


def clean_line(text):
    text = str(text)

    if text.lower() in ["nan", "none"]:
        return ""

    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def strip_bullet(text):
    text = clean_line(text)
    text = text.lstrip("*").strip()
    return text




def decode_content(content):
    for encoding in ["utf-8", "utf-16", "cp1252", "latin1"]:
        try:
            return content.decode(encoding)
        except Exception:
            continue

    return content.decode("latin1", errors="ignore")


# ============================================================
# PARSE CONTRACT DETAIL PAGE
# ============================================================

def parse_region_project_from_values(values, metadata):
    cleaned_values = [clean_line(x) for x in values if clean_line(x)]

    if len(cleaned_values) >= 1:
        metadata["region"] = cleaned_values[0]

    if len(cleaned_values) >= 2:
        metadata["district"] = cleaned_values[1]

    if len(cleaned_values) >= 3:
        metadata["dbe_percent"] = cleaned_values[2]

    if len(cleaned_values) >= 4:
        metadata["vbp_percent"] = cleaned_values[3]

    if len(cleaned_values) >= 5:
        metadata["federal_project"] = cleaned_values[4]


def parse_region_project_from_lines(lines, metadata):
    for i, line in enumerate(lines):
        line_lower = line.lower()

        if (
            "region" in line_lower
            and "district" in line_lower
            and "federal project" in line_lower
        ):
            if i + 1 < len(lines):
                values = lines[i + 1].split()
                parse_region_project_from_values(values, metadata)
            return


def parse_region_project_from_flat_text(flat_text, metadata):
    pattern = re.compile(
        r"Region\s+District\s+DBE\s*%\s+VBP\s*%\s+Federal\s+Project\s*#\s+"
        r"(?P<region>\S+)\s+"
        r"(?P<district>\S+)\s+"
        r"(?P<dbe>\S+)\s+"
        r"(?P<vbp>\S+)\s+"
        r"(?P<federal_project>\S+)",
        re.IGNORECASE,
    )

    match = pattern.search(flat_text)

    if not match:
        return

    metadata["region"] = clean_line(match.group("region"))
    metadata["district"] = clean_line(match.group("district"))
    metadata["dbe_percent"] = clean_line(match.group("dbe"))
    metadata["vbp_percent"] = clean_line(match.group("vbp"))
    metadata["federal_project"] = clean_line(match.group("federal_project"))


def parse_county_route_from_values(values, metadata):
    cleaned_values = []

    for value in values:
        value = strip_bullet(value)

        if not value:
            continue

        value_lower = value.lower()

        if "county" in value_lower and "key route" in value_lower:
            continue

        if value_lower.startswith("contract specifics"):
            continue

        cleaned_values.append(value)

    if len(cleaned_values) >= 1:
        metadata["county"] = cleaned_values[0]

    if len(cleaned_values) >= 2:
        metadata["key_route"] = cleaned_values[1]

    if len(cleaned_values) >= 3:
        metadata["marked_route"] = cleaned_values[2]

    if len(cleaned_values) >= 4:
        metadata["website_section"] = cleaned_values[3]

    if len(cleaned_values) >= 5:
        metadata["state_job"] = cleaned_values[4]

    if len(cleaned_values) >= 6:
        metadata["pps"] = " / ".join(cleaned_values[5:])


def parse_county_route_from_lines(lines, metadata):
    """
    Most reliable parser for this IDOT page:
    BeautifulSoup.stripped_strings usually returns this section as separate lines:

    County(s) Key Route(s) Marked Route(s) Section(s) State Job #(s) PPS #(s)
    Cook
    FAP 350
    IL 50 (CICERO AVE)
    FAP 0350 22 RS
    C-91-308-22
    1-80995-0000

    This handles sections that are NOT shaped like 2019-161-W.
    Example: FAP 0350 22 RS
    """
    for i, line in enumerate(lines):
        line_lower = line.lower()

        if (
            "county" in line_lower
            and "key route" in line_lower
            and "marked route" in line_lower
            and "state job" in line_lower
        ):
            values = []

            for next_line in lines[i + 1:]:
                next_line = strip_bullet(next_line)

                if not next_line:
                    continue

                if next_line.lower().startswith("contract specifics"):
                    break

                values.append(next_line)

            parse_county_route_from_values(values, metadata)
            return


def parse_county_route_from_flat_text(flat_text, metadata):
    pattern = re.compile(
        r"County\s*\(s\)\s+"
        r"Key\s+Route\s*\(s\)\s+"
        r"Marked\s+Route\s*\(s\)\s+"
        r"Section\s*\(s\)\s+"
        r"State\s+Job\s*#\s*\(s\)\s+"
        r"PPS\s*#\s*\(s\)\s+"
        r"(?P<block>.*?)\s+Contract\s+Specifics",
        re.IGNORECASE | re.DOTALL,
    )

    match = pattern.search(flat_text)

    if not match:
        return

    block = clean_line(match.group("block"))

    bullet_values = re.findall(r"\*\s*([^*]+?)(?=\s*\*|$)", block)

    if bullet_values:
        parse_county_route_from_values(bullet_values, metadata)
        return

    # Fallback when the values are smashed into one flat string.
    state_job_match = re.search(r"\bC-\d{2}-\d{3}-\d{2}\b", block)

    if not state_job_match:
        return

    before_state_job = block[:state_job_match.start()].strip()
    after_state_job = block[state_job_match.end():].strip()

    metadata["state_job"] = state_job_match.group(0)

    # Split known route styles from the beginning.
    key_route_match = re.search(
        r"\b(?:FAP|FAU|FAS|FAI|SBI|CH|TR|IL|US|I)\s*[A-Z0-9.-]+\b",
        before_state_job,
        re.IGNORECASE,
    )

    if not key_route_match:
        return

    metadata["county"] = clean_line(before_state_job[:key_route_match.start()])
    metadata["key_route"] = clean_line(key_route_match.group(0))

    after_key_route = before_state_job[key_route_match.end():].strip()

    # Find common IDOT section formats near the end of the remaining text.
    section_patterns = [
        # Example: 2019-161-W
        r"\b\d{4}-[A-Z0-9-]+\b",

        # Example: FAP 0350 22 RS
        r"\b(?:FAP|FAU|FAS|FAI|SBI)\s+\d{3,4}\s+[A-Z0-9]+\s+[A-Z0-9]+\b",

        # Example: 22 RS, 2022 RS, 123 RS-1
        r"\b\d{2,4}\s+[A-Z0-9-]+\b",
    ]

    best_section_match = None

    for pattern_text in section_patterns:
        matches = list(re.finditer(pattern_text, after_key_route, re.IGNORECASE))

        if matches:
            best_section_match = matches[-1]
            break

    if best_section_match:
        metadata["marked_route"] = clean_line(after_key_route[:best_section_match.start()])
        metadata["website_section"] = clean_line(best_section_match.group(0))
    else:
        metadata["marked_route"] = clean_line(after_key_route)

    if after_state_job:
        metadata["pps"] = clean_line(after_state_job)


def parse_metadata_from_contract_page(html, contract_url):
    soup = BeautifulSoup(html, "html.parser")

    lines = [clean_line(x) for x in soup.stripped_strings if clean_line(x)]
    flat_text = clean_line(soup.get_text(" "))

    metadata = {
        "contract_url": contract_url,
        "item_contract": "",
        "letting_date": "",
        "region": "",
        "district": "",
        "dbe_percent": "",
        "vbp_percent": "",
        "federal_project": "",
        "county": "",
        "key_route": "",
        "marked_route": "",
        "website_section": "",
        "state_job": "",
        "pps": "",
        "working_days": "",
    }

    letting_match = re.search(
        r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December)"
        r"\s+\d{1,2},\s+\d{4}\s+Letting\s+\d{1,2}:\d{2}\s+(?:AM|PM)",
        flat_text,
        re.IGNORECASE,
    )

    if letting_match:
        metadata["letting_date"] = clean_line(letting_match.group(0))

    contract_match = re.search(r"\b\d{3}-[A-Z0-9]{5}\b", flat_text)

    if contract_match:
        metadata["item_contract"] = clean_line(contract_match.group(0))

    # Try line-based first because it preserves the table values separately.
    parse_region_project_from_lines(lines, metadata)
    parse_county_route_from_lines(lines, metadata)

    # Then use flat-text fallback if anything important is missing.
    if not metadata["region"] or not metadata["district"] or not metadata["federal_project"]:
        parse_region_project_from_flat_text(flat_text, metadata)

    if (
        not metadata["county"]
        or not metadata["key_route"]
        or not metadata["marked_route"]
        or not metadata["website_section"]
        or not metadata["state_job"]
    ):
        parse_county_route_from_flat_text(flat_text, metadata)

    working_days_match = re.search(
        r"(\d+)\s+Working\s+Days",
        flat_text,
        re.IGNORECASE,
    )

    if working_days_match:
        metadata["working_days"] = working_days_match.group(1)

    return metadata


# ============================================================
# FIND IDOT CONTRACT DETAIL PAGE
# ============================================================

def normalize_contract_input(value):
    value = clean_line(value).upper()
    value = value.replace(" ", "")
    return value


def extract_contract_label(text):
    text = clean_line(text).upper()

    match = re.search(r"\b\d{3}-[A-Z0-9]{5}\b", text)

    if match:
        return match.group(0)

    return ""


def contract_matches(user_job_number, contract_label):
    user_job_number = normalize_contract_input(user_job_number)
    contract_label = normalize_contract_input(contract_label)

    if not user_job_number or not contract_label:
        return False

    if user_job_number == contract_label:
        return True

    contract_suffix = contract_label.split("-")[-1]

    if user_job_number == contract_suffix:
        return True

    return False


def get_page_signature(contract_links):
    labels = []

    for item in contract_links[:15]:
        labels.append(item.get("label", "") + "|" + item.get("url", ""))

    return "||".join(labels)


def extract_archive_dates_from_home(html):
    soup = BeautifulSoup(html, "html.parser")
    text = clean_line(soup.get_text(" "))

    start = text.find("Transportation Bulletin Archives")

    if start != -1:
        text = text[start:]

    end = text.find("Prior Lettings")

    if end != -1:
        text = text[:end]

    date_pattern = re.compile(
        r"\b("
        r"January|February|March|April|May|June|July|August|September|October|November|December"
        r")\s+\d{1,2},\s+\d{4}\b"
    )

    dates = []

    for match in date_pattern.finditer(text):
        date_text = match.group(0)

        if date_text not in dates:
            dates.append(date_text)

    return dates


def extract_letting_links_from_any_html(html):
    links = []

    raw_matches = re.findall(
        r"https?://(?:webapps1|webapps)\.dot\.illinois\.gov/WCTB/LbLettingDetail/Index/[0-9a-fA-F-]+",
        html,
    )

    for url in raw_matches:
        links.append({
            "text": "Letting",
            "url": url.replace("webapps.dot.illinois.gov", "webapps1.dot.illinois.gov"),
            "source": "raw-html",
        })

    relative_matches = re.findall(
        r"/WCTB/LbLettingDetail/Index/[0-9a-fA-F-]+",
        html,
    )

    for href in relative_matches:
        links.append({
            "text": "Letting",
            "url": absolute_url(href),
            "source": "raw-html-relative",
        })

    soup = BeautifulSoup(html, "html.parser")

    for tag in soup.find_all(True):
        visible_text = clean_line(tag.get_text(" "))

        for attr_name, attr_value in tag.attrs.items():
            if isinstance(attr_value, list):
                attr_value = " ".join(attr_value)

            attr_value = str(attr_value)

            if "LbLettingDetail/Index" in attr_value or "lblettingdetail/index" in attr_value.lower():
                matches = re.findall(
                    r"(?:https?://(?:webapps1|webapps)\.dot\.illinois\.gov)?/?WCTB/LbLettingDetail/Index/[0-9a-fA-F-]+",
                    attr_value,
                )

                for href in matches:
                    if href.startswith("http"):
                        url = href
                    elif href.startswith("/"):
                        url = BASE_URL + href
                    else:
                        url = BASE_URL + "/" + href

                    url = url.replace("webapps.dot.illinois.gov", "webapps1.dot.illinois.gov")

                    links.append({
                        "text": visible_text or "Letting",
                        "url": url,
                        "source": f"tag-attr-{attr_name}",
                    })

    unique_links = []
    seen = set()

    for item in links:
        url = item["url"]

        if url in seen:
            continue

        seen.add(url)
        unique_links.append(item)

    return unique_links


def extract_contract_urls_from_search_text(text):
    urls = []

    urls.extend(
        re.findall(
            r"https://(?:webapps1|webapps)\.dot\.illinois\.gov/WCTB/LbContractDetail/Index/[^\"'<> \n\r]+",
            text,
            flags=re.IGNORECASE,
        )
    )

    urls.extend(
        re.findall(
            r"<link>(https://(?:webapps1|webapps)\.dot\.illinois\.gov/WCTB/LbContractDetail/Index/.*?)</link>",
            text,
            flags=re.IGNORECASE,
        )
    )

    cleaned_urls = []

    for url in urls:
        url = url.replace("&amp;", "&")
        url = url.replace("webapps.dot.illinois.gov", "webapps1.dot.illinois.gov")
        url = url.split("&form=")[0]
        url = url.split("&ved=")[0]
        url = url.strip()

        if "LbContractDetail/Index" not in url:
            continue

        if url not in cleaned_urls:
            cleaned_urls.append(url)

    return cleaned_urls


def extract_letting_urls_from_search_text(text):
    urls = []

    urls.extend(
        re.findall(
            r"https://(?:webapps1|webapps)\.dot\.illinois\.gov/WCTB/LbLettingDetail/Index/[0-9a-fA-F-]+",
            text,
            flags=re.IGNORECASE,
        )
    )

    urls.extend(
        re.findall(
            r"<link>(https://(?:webapps1|webapps)\.dot\.illinois\.gov/WCTB/LbLettingDetail/Index/[0-9a-fA-F-]+)</link>",
            text,
            flags=re.IGNORECASE,
        )
    )

    cleaned_urls = []

    for url in urls:
        url = url.replace("&amp;", "&")
        url = url.replace("webapps.dot.illinois.gov", "webapps1.dot.illinois.gov")
        url = url.split("&form=")[0]
        url = url.split("&ved=")[0]
        url = url.strip()

        if "LbLettingDetail/Index" not in url:
            continue

        if url not in cleaned_urls:
            cleaned_urls.append(url)

    return cleaned_urls


def bing_rss_search(session, query):
    search_url = "https://www.bing.com/search?q=" + quote(query) + "&format=rss"

    response = session.get(
        search_url,
        timeout=30,
        headers={
            "User-Agent": "Mozilla/5.0 IDR Generator",
            "Accept": "application/rss+xml,application/xml,text/xml,*/*",
        },
    )

    response.raise_for_status()
    return response.text, search_url


def extract_duckduckgo_result_urls(html):
    soup = BeautifulSoup(html, "html.parser")
    urls = []

    for a in soup.find_all("a", href=True):
        href = a["href"]

        if "uddg=" in href:
            parsed = urlparse(href)
            query = parse_qs(parsed.query)
            uddg = query.get("uddg", [""])[0]

            if uddg:
                url = unquote(uddg)

                if "webapps1.dot.illinois.gov/WCTB/LbContractDetail/Index" in url:
                    urls.append(url)

                if "webapps.dot.illinois.gov/WCTB/LbContractDetail/Index" in url:
                    urls.append(url)

        elif "LbContractDetail/Index" in href:
            urls.append(href)

    cleaned_urls = []

    for url in urls:
        url = url.replace("&amp;", "&")
        url = url.replace("webapps.dot.illinois.gov", "webapps1.dot.illinois.gov")
        url = url.split("&rut=")[0]
        url = url.strip()

        if url not in cleaned_urls:
            cleaned_urls.append(url)

    return cleaned_urls


def resolve_archive_date_to_letting_url(session, date_text):
    queries = [
        f'site:webapps1.dot.illinois.gov/WCTB/LbLettingDetail/Index "{date_text} Letting"',
        f'site:webapps.dot.illinois.gov/WCTB/LbLettingDetail/Index "{date_text} Letting"',
        f'"{date_text} Letting 12:00 PM" "LbLettingDetail"',
    ]

    for query in queries:
        try:
            text, _ = bing_rss_search(session, query)
            letting_urls = extract_letting_urls_from_search_text(text)

            for url in letting_urls:
                try:
                    html = get_html(session, url)
                    page_text = clean_line(BeautifulSoup(html, "html.parser").get_text(" "))

                    if date_text.lower() in page_text.lower():
                        return url

                except Exception:
                    continue

        except Exception:
            continue

    return ""


def get_current_letting_link(session, html):
    links = extract_letting_links_from_any_html(html)

    if links:
        links[0]["text"] = "Current Notice of Letting"
        links[0]["source"] = "current"
        return links[0]

    return None


def get_all_archive_letting_links_newest_first(session):
    home_html = get_html(session, IDOT_HOME_URL)

    archive_dates = extract_archive_dates_from_home(home_html)
    direct_links = extract_letting_links_from_any_html(home_html)

    archive_links = []
    seen_urls = set()

    for date_text in archive_dates:
        found_url = ""

        for link in direct_links:
            if date_text.lower() in link.get("text", "").lower():
                found_url = link["url"]
                break

        if not found_url:
            found_url = resolve_archive_date_to_letting_url(session, date_text)

        if found_url and found_url not in seen_urls:
            seen_urls.add(found_url)

            archive_links.append({
                "text": date_text,
                "url": found_url,
                "source": "archive-date-list",
            })

    current = get_current_letting_link(session, home_html)

    final_links = []

    if current and current["url"] not in seen_urls:
        final_links.append(current)
        seen_urls.add(current["url"])

    for link in archive_links:
        if link["url"] not in seen_urls:
            final_links.append(link)
            seen_urls.add(link["url"])

    for link in direct_links:
        if link["url"] not in seen_urls:
            final_links.append(link)
            seen_urls.add(link["url"])

    return final_links


def get_contract_links_from_letting_page(html):
    soup = BeautifulSoup(html, "html.parser")

    contract_links = []

    for a in soup.find_all("a", href=True):
        href = a["href"]

        if "lbcontractdetail" not in href.lower():
            continue

        link_text = clean_line(a.get_text(" "))
        label = extract_contract_label(link_text)

        parent_row = a.find_parent("tr")
        row_text = ""

        if parent_row is not None:
            row_text = clean_line(parent_row.get_text(" "))

            if not label:
                label = extract_contract_label(row_text)

        contract_links.append({
            "label": label,
            "url": absolute_url(href),
            "text": link_text,
            "row_text": row_text,
        })

    return contract_links


def letting_page_matches_contract(session, letting, user_job_number):
    seen_page_signatures = set()

    for page_num in range(1, SEARCH_MAX_PAGES_PER_LETTING + 1):
        if page_num == 1:
            page_url = letting["url"]
        else:
            page_url = set_query_param(letting["url"], "page", page_num)

        try:
            html = get_html(session, page_url)
        except Exception:
            break

        contract_links = get_contract_links_from_letting_page(html)

        if not contract_links:
            break

        page_signature = get_page_signature(contract_links)

        if page_signature in seen_page_signatures:
            break

        seen_page_signatures.add(page_signature)

        for contract in contract_links:
            label = contract.get("label", "")

            if contract_matches(user_job_number, label):
                return {
                    "label": label,
                    "url": contract["url"],
                    "letting": letting["text"],
                    "letting_url": letting["url"],
                    "page": page_num,
                    "source": letting.get("source", ""),
                }

    return None


def find_contract_detail_url_from_public_search(session, user_job_number):
    user_job_number = normalize_contract_input(user_job_number)

    search_queries = [
        f'site:webapps1.dot.illinois.gov/WCTB/LbContractDetail "{user_job_number}"',
        f'site:webapps.dot.illinois.gov/WCTB/LbContractDetail "{user_job_number}"',
        f'"{user_job_number}" "LbContractDetail"',
        f'"{user_job_number}" "webapps1.dot.illinois.gov/WCTB"',
    ]

    candidate_urls = []

    for query in search_queries:
        try:
            text, _ = bing_rss_search(session, query)
            candidate_urls.extend(extract_contract_urls_from_search_text(text))
        except Exception:
            pass

        duck_url = "https://duckduckgo.com/html/?q=" + quote(query)

        try:
            response = session.get(
                duck_url,
                timeout=30,
                headers={
                    "User-Agent": "Mozilla/5.0 IDR Generator",
                    "Accept": "text/html,*/*",
                },
            )
            response.raise_for_status()

            candidate_urls.extend(extract_duckduckgo_result_urls(response.text))
            candidate_urls.extend(extract_contract_urls_from_search_text(response.text))

        except Exception:
            pass

    cleaned_urls = []

    for url in candidate_urls:
        url = url.replace("&amp;", "&")
        url = url.replace("webapps.dot.illinois.gov", "webapps1.dot.illinois.gov")
        url = url.strip()

        if "LbContractDetail/Index" in url and url not in cleaned_urls:
            cleaned_urls.append(url)

    for url in cleaned_urls:
        try:
            contract_html = get_html(session, url)
            metadata = parse_metadata_from_contract_page(contract_html, url)
            label = metadata.get("item_contract", "")

            if contract_matches(user_job_number, label):
                return {
                    "label": label,
                    "url": url,
                    "letting": metadata.get("letting_date", "Found by public contract search"),
                    "letting_url": "",
                    "page": "",
                    "source": "public-contract-search",
                }

        except Exception:
            continue

    return None


def find_contract_detail_url(session, job_number):
    original_input = job_number.strip()

    if not original_input:
        raise ValueError("Enter a job number first.")

    if original_input.lower().startswith("http") and "lbcontractdetail" in original_input.lower():
        html = get_html(session, original_input)
        metadata = parse_metadata_from_contract_page(html, original_input)

        if not metadata.get("item_contract"):
            raise ValueError(
                "The direct URL opened, but the contract number could not be parsed from that page."
            )

        return {
            "label": metadata.get("item_contract", ""),
            "url": original_input,
            "letting": metadata.get("letting_date", "Direct URL"),
            "letting_url": "",
            "page": "",
            "source": "direct-url",
        }

    user_job_number = normalize_contract_input(original_input)

    letting_links = get_all_archive_letting_links_newest_first(session)
    checked_lettings = 0

    for letting in letting_links:
        checked_lettings += 1

        result = letting_page_matches_contract(
            session=session,
            letting=letting,
            user_job_number=user_job_number,
        )

        if result is not None:
            return result

    result = find_contract_detail_url_from_public_search(session, user_job_number)

    if result is not None:
        return result

    raise ValueError(
        f"Could not find contract '{user_job_number}'. "
        f"I checked {checked_lettings} current/archive letting page(s), newest to oldest. "
        "Try the full item-contract number like 001-62K33, or paste the direct contract detail URL."
    )


# ============================================================
# PAY ITEM PARSING
# ============================================================

def unit_regex():
    escaped_units = sorted([re.escape(u) for u in UNIT_WORDS], key=len, reverse=True)
    return "(" + "|".join(escaped_units) + ")"


def normalize_unit(unit):
    unit = clean_line(unit).upper()

    replacements = {
        "CUYD": "CU YD",
        "SQYD": "SQ YD",
        "SQFT": "SQ FT",
        "LSUM": "L SUM",
    }

    return replacements.get(unit, unit)


def normalize_pay_item_code(value):
    return clean_line(value).upper()


def get_pay_item_report_url(contract_url, html):
    soup = BeautifulSoup(html, "html.parser")

    for a in soup.find_all("a", href=True):
        text = clean_line(a.get_text(" ")).lower()
        href = a["href"]
        href_lower = href.lower()

        if "pay item report" in text or "getpayitemexcelfile" in href_lower:
            return urljoin(contract_url, href)

    return ""


def normalize_pay_item_df(df):
    rename = {}

    for col in df.columns:
        low = str(col).lower().strip()

        if ("pay item" in low and "#" in low) or low in ["pay item", "item #", "item number", "item"]:
            rename[col] = "item_code"
        elif "uom" in low:
            rename[col] = "unit"
        elif low in ["unit", "units"]:
            rename[col] = "unit"
        elif "description" in low:
            rename[col] = "item_description"
        elif "quantity" in low or low == "qty":
            rename[col] = "quantity"
        elif "unit price" in low or "price" in low:
            rename[col] = "unit_price"

    df = df.rename(columns=rename)

    needed = [
        "item_code",
        "unit",
        "item_description",
        "quantity",
        "unit_price",
    ]

    for col in needed:
        if col not in df.columns:
            df[col] = ""

    df = df[needed]

    df["item_code"] = df["item_code"].astype(str).str.strip().str.upper()
    df["unit"] = df["unit"].astype(str).str.strip().str.upper().apply(normalize_unit)
    df["item_description"] = df["item_description"].astype(str).str.strip()
    df["quantity"] = df["quantity"].astype(str).str.strip()
    df["unit_price"] = df["unit_price"].astype(str).str.strip()

    for col in needed:
        df[col] = df[col].replace("nan", "")

    df = df[df["item_code"].str.match(r"^[A-Z]?\d{6,8}[A-Z]?$", na=False)]

    df = df.drop_duplicates(subset=["item_code"], keep="first")
    df = df.reset_index(drop=True)

    return df


def parse_pay_items_from_html_tables(text):
    try:
        tables = pd.read_html(StringIO(text))
    except Exception:
        return pd.DataFrame()

    for table in tables:
        table = table.copy()

        if isinstance(table.columns, pd.MultiIndex):
            table.columns = [
                " ".join([clean_line(x) for x in col if clean_line(x)])
                for col in table.columns
            ]
        else:
            table.columns = [clean_line(c) for c in table.columns]

        normalized = normalize_pay_item_df(table)

        if not normalized.empty:
            return normalized

    return pd.DataFrame()


def parse_pay_items_from_row_tokens(text):
    soup = BeautifulSoup(text, "html.parser")
    lines = [clean_line(x) for x in soup.get_text("\n").splitlines() if clean_line(x)]

    rows = []
    current = []

    item_start = re.compile(
        rf"^(?P<item_code>[A-Z]?\d{{6,8}}[A-Z]?)\s+(?P<unit>{unit_regex()})\b",
        re.IGNORECASE,
    )

    for line in lines:
        if line.lower().startswith("no more pay items"):
            break

        if item_start.search(line):
            if current:
                rows.append(" ".join(current))
                current = []

            current.append(line)
        else:
            if current:
                current.append(line)

    if current:
        rows.append(" ".join(current))

    return parse_pay_item_row_strings(rows)


def parse_pay_item_row_strings(row_strings):
    parsed_rows = []

    row_pattern = re.compile(
        rf"^(?P<item_code>[A-Z]?\d{{6,8}}[A-Z]?)\s+"
        rf"(?P<unit>{unit_regex()})\s+"
        rf"(?P<body>.+)$",
        re.IGNORECASE,
    )

    for row in row_strings:
        row = clean_line(row)

        match = row_pattern.match(row)

        if not match:
            continue

        item_code = normalize_pay_item_code(match.group("item_code"))
        unit = normalize_unit(match.group("unit"))
        body = clean_line(match.group("body"))

        body = re.sub(
            r"\b(Base items|Specialty items|Non-bid items)\b",
            " ",
            body,
            flags=re.IGNORECASE,
        )

        body = clean_line(body)

        number_matches = list(
            re.finditer(
                r"\$?\d{1,3}(?:,\d{3})*(?:\.\d+)?|\$?\d+(?:\.\d+)?",
                body,
            )
        )

        if not number_matches:
            continue

        unit_price = ""
        quantity = ""

        last_number = number_matches[-1].group(0)

        if last_number.startswith("$"):
            unit_price = last_number

            if len(number_matches) < 2:
                continue

            quantity = number_matches[-2].group(0)
            desc_end = number_matches[-2].start()
        else:
            quantity = last_number
            desc_end = number_matches[-1].start()

        description = clean_line(body[:desc_end])

        if not description:
            continue

        parsed_rows.append({
            "item_code": item_code,
            "unit": unit,
            "item_description": description,
            "quantity": quantity,
            "unit_price": unit_price,
        })

    if not parsed_rows:
        return pd.DataFrame()

    return normalize_pay_item_df(pd.DataFrame(parsed_rows))


def parse_pay_items_from_flat_text(text):
    soup = BeautifulSoup(text, "html.parser")
    visible_text = soup.get_text(" ")

    visible_text = visible_text.replace("\u00a0", " ")
    visible_text = re.sub(r"\s+", " ", visible_text).strip()

    lower_text = visible_text.lower()

    starts = [
        lower_text.find("pay item #"),
        lower_text.find("pay item"),
        lower_text.find("item #"),
    ]

    starts = [x for x in starts if x != -1]

    if starts:
        visible_text = visible_text[min(starts):]

    end = visible_text.lower().find("no more pay items")

    if end != -1:
        visible_text = visible_text[:end]

    item_start_pattern = re.compile(
        rf"\b(?P<item_code>[A-Z]?\d{{6,8}}[A-Z]?)\s+(?P<unit>{unit_regex()})\b",
        re.IGNORECASE,
    )

    matches = list(item_start_pattern.finditer(visible_text))

    if not matches:
        return pd.DataFrame()

    row_strings = []

    for index, match in enumerate(matches):
        start = match.start()

        if index + 1 < len(matches):
            end = matches[index + 1].start()
        else:
            end = len(visible_text)

        row_strings.append(clean_line(visible_text[start:end]))

    return parse_pay_item_row_strings(row_strings)


def parse_pay_items_from_tab_or_csv_text(text):
    for sep in ["\t", ",", ";", "|"]:
        try:
            df = pd.read_csv(StringIO(text), sep=sep, engine="python")
        except Exception:
            continue

        normalized = normalize_pay_item_df(df)

        if not normalized.empty:
            return normalized

    return pd.DataFrame()


def find_table_inside_raw_excel(raw_df):
    for row_index in range(len(raw_df)):
        row_values = list(raw_df.iloc[row_index].values)
        joined = " ".join([clean_line(x).lower() for x in row_values])

        has_item = "pay item" in joined or "item #" in joined or "item number" in joined
        has_uom = "uom" in joined or "unit" in joined
        has_desc = "description" in joined
        has_qty = "quantity" in joined or "qty" in joined

        if has_item and has_uom and has_desc and has_qty:
            headers = [clean_line(x) for x in row_values]
            data = raw_df.iloc[row_index + 1:].copy()
            data.columns = headers
            return data

    return pd.DataFrame()


def parse_pay_items_from_excel_bytes(content):
    for engine in [None, "openpyxl", "xlrd"]:
        try:
            file_data = io.BytesIO(content)

            if engine is None:
                sheets = pd.read_excel(file_data, sheet_name=None, header=None)
            else:
                sheets = pd.read_excel(file_data, sheet_name=None, header=None, engine=engine)

            for sheet_name, raw_df in sheets.items():
                df = find_table_inside_raw_excel(raw_df)
                normalized = normalize_pay_item_df(df)

                if not normalized.empty:
                    return normalized

        except Exception:
            continue

    return pd.DataFrame()


def parse_pay_items_from_any_text(text):
    parsers = [
        parse_pay_items_from_html_tables,
        parse_pay_items_from_row_tokens,
        parse_pay_items_from_flat_text,
        parse_pay_items_from_tab_or_csv_text,
    ]

    for parser in parsers:
        try:
            pay_items = parser(text)

            if not pay_items.empty:
                return pay_items

        except Exception:
            continue

    return pd.DataFrame()


def parse_pay_items_from_pay_item_report(session, contract_url, html):
    pay_item_url = get_pay_item_report_url(contract_url, html)

    if not pay_item_url:
        return pd.DataFrame()

    response = get_response(session, pay_item_url)
    content = response.content

    if not content:
        return pd.DataFrame()

    text = decode_content(content)
    pay_items = parse_pay_items_from_any_text(text)

    if not pay_items.empty:
        return pay_items

    return parse_pay_items_from_excel_bytes(content)


def parse_pay_items_from_contract_page_text(html):
    return parse_pay_items_from_any_text(html)


def fetch_idot_job(job_number):
    session = make_session()

    match = find_contract_detail_url(session, job_number)

    html = get_html(session, match["url"])
    metadata = parse_metadata_from_contract_page(html, match["url"])

    if not metadata.get("item_contract"):
        metadata["item_contract"] = match.get("label", "")

    pay_items = parse_pay_items_from_pay_item_report(session, match["url"], html)

    if pay_items.empty:
        pay_items = parse_pay_items_from_contract_page_text(html)

    if pay_items.empty:
        raise ValueError(
            "Found the contract page, but could not extract pay items from the Pay Item Report or page text."
        )

    return metadata, pay_items, match

# ============================================================
# PDF HELPERS - NO EXCEL / NO MACROS
# ============================================================

from datetime import date as DateClass

try:
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from reportlab.pdfgen import canvas
except Exception:
    letter = None
    landscape = None
    colors = None
    inch = 72
    canvas = None

PDF_WEATHER_OPTIONS = [
    "Sunny",
    "Cloudy",
    "Light Rain",
    "Normal Rain",
    "Heavy Rain",
    "Snow",
]

PDF_ROW_COUNT = 6


def get_today_default():
    return DateClass.today()


def format_report_date(value):
    if hasattr(value, "strftime"):
        return value.strftime("%m/%d/%Y")
    return clean_line(value)


def format_report_day(value):
    if hasattr(value, "strftime"):
        return value.strftime("%A")
    return ""


def format_pdf_filename(contract_number):
    contract_number = clean_line(contract_number) or "IDOT"
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", contract_number).strip("_")
    return f"{safe}_IDR.pdf"


def dataframe_records_by_code(pay_items):
    records = {}
    if pay_items is None or pay_items.empty:
        return records
    for _, row in pay_items.iterrows():
        code = normalize_pay_item_code(row.get("item_code", ""))
        if code and code not in records:
            records[code] = {
                "item_code": code,
                "item_description": clean_line(row.get("item_description", "")),
                "unit": normalize_unit(row.get("unit", "")),
                "plan_quantity": clean_line(row.get("quantity", "")),
                "unit_price": clean_line(row.get("unit_price", "")),
                "is_custom": False,
            }
    return records


def dataframe_records_by_description(pay_items):
    records = {}
    if pay_items is None or pay_items.empty:
        return records
    for _, row in pay_items.iterrows():
        desc = clean_line(row.get("item_description", ""))
        if desc and desc not in records:
            code = normalize_pay_item_code(row.get("item_code", ""))
            records[desc] = {
                "item_code": code,
                "item_description": desc,
                "unit": normalize_unit(row.get("unit", "")),
                "plan_quantity": clean_line(row.get("quantity", "")),
                "unit_price": clean_line(row.get("unit_price", "")),
                "is_custom": False,
            }
    return records


def resolve_line_item(pay_items, lookup_mode, selected_code, selected_description, custom_code, custom_description, custom_unit, location, quantity):
    by_code = dataframe_records_by_code(pay_items)
    by_desc = dataframe_records_by_description(pay_items)

    if lookup_mode == "Custom Item":
        return {
            "item_code": normalize_pay_item_code(custom_code),
            "item_description": clean_line(custom_description),
            "unit": normalize_unit(custom_unit),
            "location": clean_line(location),
            "quantity": clean_line(quantity),
            "plan_quantity": "",
            "unit_price": "",
            "is_custom": True,
        }

    if lookup_mode == "Item Code":
        item = by_code.get(normalize_pay_item_code(selected_code), {}).copy()
    else:
        item = by_desc.get(clean_line(selected_description), {}).copy()

    if not item:
        item = {
            "item_code": normalize_pay_item_code(selected_code),
            "item_description": clean_line(selected_description),
            "unit": "",
            "plan_quantity": "",
            "unit_price": "",
            "is_custom": False,
        }

    item["location"] = clean_line(location)
    item["quantity"] = clean_line(quantity)
    item["unit"] = normalize_unit(item.get("unit", ""))
    return item


def build_idr_rows_form(pay_items):
    code_options = [""] + sorted(dataframe_records_by_code(pay_items).keys())
    description_options = [""] + sorted(dataframe_records_by_description(pay_items).keys())

    rows = []

    st.subheader("IDR Pay Item Rows")
    st.caption("Each row can be filled by item code, by item description, or as a custom pay item. The website autofills the matching code, description, and unit before creating the PDF.")

    for idx in range(PDF_ROW_COUNT):
        row_num = idx + 1
        with st.expander(f"Pay Item Row {row_num}", expanded=(idx == 0)):
            lookup_mode = st.radio(
                "Fill this row by",
                ["Item Code", "Item Description", "Custom Item"],
                horizontal=True,
                key=f"lookup_mode_{idx}",
            )

            selected_code = ""
            selected_description = ""
            custom_code = ""
            custom_description = ""
            custom_unit = ""

            if lookup_mode == "Item Code":
                selected_code = st.selectbox(
                    "Item Code",
                    code_options,
                    key=f"selected_code_{idx}",
                )
                selected_item = dataframe_records_by_code(pay_items).get(selected_code, {})

                c1, c2 = st.columns([3, 1])
                with c1:
                    st.text_input(
                        "Autofilled Item Description",
                        value=selected_item.get("item_description", ""),
                        disabled=True,
                        key=f"autofill_desc_from_code_{idx}",
                    )
                with c2:
                    st.text_input(
                        "Autofilled Unit",
                        value=selected_item.get("unit", ""),
                        disabled=True,
                        key=f"autofill_unit_from_code_{idx}",
                    )

            elif lookup_mode == "Item Description":
                selected_description = st.selectbox(
                    "Item Description",
                    description_options,
                    key=f"selected_desc_{idx}",
                )
                selected_item = dataframe_records_by_description(pay_items).get(selected_description, {})

                c1, c2 = st.columns([1, 1])
                with c1:
                    st.text_input(
                        "Autofilled Item Code",
                        value=selected_item.get("item_code", ""),
                        disabled=True,
                        key=f"autofill_code_from_desc_{idx}",
                    )
                with c2:
                    st.text_input(
                        "Autofilled Unit",
                        value=selected_item.get("unit", ""),
                        disabled=True,
                        key=f"autofill_unit_from_desc_{idx}",
                    )

            else:
                c1, c2, c3 = st.columns([1, 3, 1])
                with c1:
                    custom_code = st.text_input("Custom Item Code", key=f"custom_code_{idx}")
                with c2:
                    custom_description = st.text_input("Custom Item Description", key=f"custom_desc_{idx}")
                with c3:
                    custom_unit = st.text_input("Custom Unit", key=f"custom_unit_{idx}")

            c1, c2 = st.columns([2, 1])
            with c1:
                location = st.text_input("Location", key=f"location_{idx}")
            with c2:
                quantity = st.text_input("Quantity Used", key=f"quantity_{idx}")

            rows.append(resolve_line_item(
                pay_items=pay_items,
                lookup_mode=lookup_mode,
                selected_code=selected_code,
                selected_description=selected_description,
                custom_code=custom_code,
                custom_description=custom_description,
                custom_unit=custom_unit,
                location=location,
                quantity=quantity,
            ))

    return rows


def wrap_text_to_lines(text, max_chars, max_lines):
    text = clean_line(text)
    if not text:
        return [""]

    words = text.split()
    lines = []
    current = ""

    for word in words:
        test = word if not current else current + " " + word
        if len(test) <= max_chars:
            current = test
        else:
            if current:
                lines.append(current)
            current = word

        if len(lines) >= max_lines:
            break

    if current and len(lines) < max_lines:
        lines.append(current)

    if len(lines) == max_lines and len(" ".join(words)) > len(" ".join(lines)):
        lines[-1] = lines[-1][:max(0, max_chars - 3)].rstrip() + "..."

    return lines or [""]


def draw_label_value(c, label, value, x, y, width, height=18, field_name=None, editable=True, font_size=8):
    c.setFont("Helvetica-Bold", 7)
    c.drawString(x, y + height + 3, label)
    c.rect(x, y, width, height, stroke=1, fill=0)

    value = clean_line(value)
    if editable and field_name:
        c.acroForm.textfield(
            name=field_name,
            value=value,
            x=x + 2,
            y=y + 2,
            width=width - 4,
            height=height - 4,
            borderWidth=0,
            fontName="Helvetica",
            fontSize=font_size,
            textColor=colors.black,
            fillColor=None,
            forceBorder=False,
        )
    else:
        c.setFont("Helvetica", font_size)
        c.drawString(x + 2, y + 5, value[:80])


def draw_multiline_field(c, name, value, x, y, width, height, font_size=7):
    c.rect(x, y, width, height, stroke=1, fill=0)
    value = clean_line(value)

    # ReportLab textfield supports basic AcroForm fields. Many PDF viewers allow editing these fields.
    try:
        c.acroForm.textfield(
            name=name,
            value=value,
            x=x + 2,
            y=y + 2,
            width=width - 4,
            height=height - 4,
            borderWidth=0,
            fontName="Helvetica",
            fontSize=font_size,
            textColor=colors.black,
            fillColor=None,
            forceBorder=False,
            fieldFlags="multiline",
        )
    except TypeError:
        # Older reportlab versions may not accept fieldFlags as a string.
        c.setFont("Helvetica", font_size)
        lines = wrap_text_to_lines(value, 55, 4)
        text_y = y + height - 10
        for line in lines:
            c.drawString(x + 2, text_y, line)
            text_y -= font_size + 2


def make_fillable_idr_pdf(metadata, idr_info, rows):
    if canvas is None:
        raise RuntimeError("The reportlab package is required. Add reportlab to requirements.txt.")

    output = io.BytesIO()
    c = canvas.Canvas(output, pagesize=landscape(letter))
    page_width, page_height = landscape(letter)

    margin = 28
    top = page_height - margin

    report_date = idr_info.get("date", get_today_default())
    report_date_text = format_report_date(report_date)
    report_day_text = format_report_day(report_date)

    # Title
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(page_width / 2, top - 5, "IDOT INSPECTOR DAILY REPORT")
    c.setFont("Helvetica", 8)
    c.drawCentredString(page_width / 2, top - 18, "Generated from the Streamlit IDR website - macro-free PDF")

    # Header date blocks. Top-left and top-right update from selected date.
    draw_label_value(c, "Date", report_date_text, margin, top - 42, 90, field_name="top_left_date")
    draw_label_value(c, "Day", report_day_text, margin + 98, top - 42, 90, field_name="top_left_day")
    draw_label_value(c, "Report Date", report_date_text, page_width - margin - 190, top - 42, 90, field_name="top_right_date")
    draw_label_value(c, "Day", report_day_text, page_width - margin - 92, top - 42, 92, field_name="top_right_day")

    # Job info blocks.
    y = top - 82
    left_x = margin
    mid_x = margin + 260
    right_x = margin + 520

    draw_label_value(c, "Contractor or Sub.", idr_info.get("contractor", ""), left_x, y, 230, field_name="contractor")
    draw_label_value(c, "Weather", idr_info.get("weather", ""), mid_x, y, 120, field_name="weather")
    draw_label_value(c, "Contract No.", metadata.get("item_contract", ""), right_x, y, 110, field_name="contract_no")
    draw_label_value(c, "Job No.", metadata.get("state_job", ""), right_x + 118, y, 110, field_name="job_no")

    y -= 40
    draw_label_value(c, "County", metadata.get("county", ""), left_x, y, 150, field_name="county")
    draw_label_value(c, "Section", metadata.get("key_route", ""), left_x + 160, y, 170, field_name="section")
    draw_label_value(c, "Route", metadata.get("marked_route", ""), left_x + 340, y, 180, field_name="route")
    draw_label_value(c, "District", metadata.get("district", ""), left_x + 530, y, 70, field_name="district")
    draw_label_value(c, "Project", metadata.get("federal_project", ""), left_x + 610, y, 120, field_name="project")

    y -= 47
    c.setFont("Helvetica-Bold", 7)
    c.drawString(left_x, y + 24, "Work Description")
    draw_multiline_field(c, "work_description", idr_info.get("work_description", ""), left_x, y - 5, page_width - (2 * margin), 28, font_size=7)

    # Pay item table.
    table_top = y - 25
    row_h = 45
    headers = [
        ("Item Code", 82),
        ("Item Description", 255),
        ("Location", 150),
        ("Quantity", 70),
        ("Unit", 62),
        ("Custom?", 54),
    ]

    x = margin
    c.setFillColor(colors.lightgrey)
    c.rect(x, table_top - 17, sum(w for _, w in headers), 17, stroke=1, fill=1)
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 7)
    cx = x
    for label, w in headers:
        c.drawString(cx + 3, table_top - 12, label)
        c.rect(cx, table_top - 17, w, 17, stroke=1, fill=0)
        cx += w

    start_y = table_top - 17
    for i in range(PDF_ROW_COUNT):
        row = rows[i] if i < len(rows) else {}
        row_y = start_y - ((i + 1) * row_h)
        cx = x

        code = clean_line(row.get("item_code", ""))
        desc = clean_line(row.get("item_description", ""))
        loc = clean_line(row.get("location", ""))
        qty = clean_line(row.get("quantity", ""))
        unit = clean_line(row.get("unit", ""))
        custom = "Yes" if row.get("is_custom") else ""

        values = [code, desc, loc, qty, unit, custom]
        field_names = [
            f"row_{i+1}_item_code",
            f"row_{i+1}_item_description",
            f"row_{i+1}_location",
            f"row_{i+1}_quantity",
            f"row_{i+1}_unit",
            f"row_{i+1}_custom",
        ]

        for (label, w), value, field_name in zip(headers, values, field_names):
            c.rect(cx, row_y, w, row_h, stroke=1, fill=0)
            if label == "Item Description":
                draw_multiline_field(c, field_name, value, cx, row_y, w, row_h, font_size=6)
            else:
                try:
                    c.acroForm.textfield(
                        name=field_name,
                        value=value,
                        x=cx + 2,
                        y=row_y + 14,
                        width=w - 4,
                        height=14,
                        borderWidth=0,
                        fontName="Helvetica",
                        fontSize=7,
                        textColor=colors.black,
                        fillColor=None,
                        forceBorder=False,
                    )
                except Exception:
                    c.setFont("Helvetica", 7)
                    c.drawString(cx + 3, row_y + row_h - 12, value[:20])
            cx += w

    # Remarks area.
    remarks_y = margin + 34
    c.setFont("Helvetica-Bold", 7)
    c.drawString(margin, remarks_y + 42, "Remarks / Notes")
    draw_multiline_field(c, "remarks", idr_info.get("remarks", ""), margin, remarks_y, page_width - (2 * margin), 40, font_size=7)

    # Signature fields.
    sig_y = margin
    draw_label_value(c, "Inspector Name", idr_info.get("inspector", ""), margin, sig_y, 200, field_name="inspector")
    draw_label_value(c, "Signature", "", margin + 220, sig_y, 220, field_name="signature")
    draw_label_value(c, "Date", report_date_text, margin + 460, sig_y, 100, field_name="signature_date")

    c.showPage()
    c.save()
    output.seek(0)
    return output


def make_pay_items_excel(metadata, pay_items):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pay_items.to_excel(writer, index=False, sheet_name="Pay Items")

        info_df = pd.DataFrame(
            [{"field": k, "value": v} for k, v in metadata.items()]
        )

        info_df.to_excel(writer, index=False, sheet_name="Job Info")

        workbook = writer.book

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, 1000, 20)

            header_format = workbook.add_format({
                "bold": True,
                "border": 1,
                "text_wrap": True,
            })

            if sheet_name == "Pay Items":
                cols = pay_items.columns
            else:
                cols = info_df.columns

            for col_num, col in enumerate(cols):
                ws.write(0, col_num, col, header_format)
                ws.set_column(col_num, col_num, 24)

            if sheet_name == "Pay Items":
                ws.set_column(2, 2, 48)

    output.seek(0)
    return output


# ============================================================
# STREAMLIT APP
# ============================================================

st.set_page_config(
    page_title="IDOT Job IDR PDF Generator",
    page_icon="📄",
    layout="wide",
)

st.title("IDOT Job IDR PDF Generator")

st.write(
    "Enter an IDOT job/contract number or paste the direct IDOT contract URL. "
    "The app pulls the job information and pay items, then creates a macro-free editable PDF."
)

if canvas is None:
    st.error("Missing dependency: reportlab. Add reportlab to requirements.txt and redeploy the app.")
    st.stop()

with st.sidebar:
    st.header("Job Lookup")

    job_number = st.text_input(
        "IDOT Job / Contract Number or Contract Detail URL",
        placeholder="Example: 62K33, 001-62K33, or paste the IDOT contract URL",
    )

    st.caption("No Excel macros are used. The website creates a fillable PDF output.")


if "metadata" not in st.session_state:
    st.session_state.metadata = None

if "pay_items" not in st.session_state:
    st.session_state.pay_items = pd.DataFrame()

if "match" not in st.session_state:
    st.session_state.match = None

if st.button("Find IDOT Job"):
    try:
        with st.spinner("Searching IDOT Transportation Bulletin archives newest to oldest..."):
            metadata, pay_items, match = fetch_idot_job(job_number)

        st.session_state.metadata = metadata
        st.session_state.pay_items = pay_items
        st.session_state.match = match

        st.success(
            f"Found {metadata.get('item_contract', job_number)} with {len(pay_items)} pay items."
        )

        if match.get("letting_url"):
            st.info(
                f"Found on letting/archive page: {match.get('letting', '')}, "
                f"page {match.get('page', '')}"
            )
        else:
            st.info(f"Found using: {match.get('letting', '')}")

    except Exception as e:
        st.error(str(e))

metadata = st.session_state.metadata
pay_items = st.session_state.pay_items

if metadata is not None:
    st.subheader("Job Information")

    required_fields = {
        "County": metadata.get("county", ""),
        "Section": metadata.get("key_route", ""),
        "Route": metadata.get("marked_route", ""),
        "District": metadata.get("district", ""),
        "Contract No.": metadata.get("item_contract", ""),
        "Job No.": metadata.get("state_job", ""),
        "Project": metadata.get("federal_project", ""),
    }

    missing_fields = [name for name, value in required_fields.items() if not value]

    if missing_fields:
        st.warning(
            "Some job fields did not parse correctly: "
            + ", ".join(missing_fields)
            + ". You can still generate the PDF, then edit those fields inside the PDF if needed."
        )

        with st.expander("Parser debug info"):
            st.write(metadata)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.write(f"**County:** {metadata.get('county', '')}")
        st.write(f"**Section:** {metadata.get('key_route', '')}")
        st.write(f"**Route:** {metadata.get('marked_route', '')}")

    with col2:
        st.write(f"**District:** {metadata.get('district', '')}")
        st.write(f"**Contract No.:** {metadata.get('item_contract', '')}")
        st.write(f"**Job No.:** {metadata.get('state_job', '')}")

    with col3:
        st.write(f"**Project:** {metadata.get('federal_project', '')}")
        st.write(f"**Working Days:** {metadata.get('working_days', '')}")

if metadata is not None and not pay_items.empty:
    st.subheader("IDR Header Fields")

    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        idr_date = st.date_input("IDR Date", value=get_today_default())
    with c2:
        contractor = st.text_input("Contractor or Subcontractor")
    with c3:
        weather = st.selectbox("Weather", [""] + PDF_WEATHER_OPTIONS)

    work_description = st.text_area("Work Description", height=70)
    remarks = st.text_area("Remarks / Notes", height=80)
    inspector = st.text_input("Inspector Name")

    st.divider()

    rows = build_idr_rows_form(pay_items)

    st.divider()

    col_a, col_b = st.columns(2)

    with col_a:
        pay_items_file = make_pay_items_excel(metadata, pay_items)

        st.download_button(
            label="Download IDOT Pay Item Table",
            data=pay_items_file,
            file_name=f"{metadata.get('item_contract', 'idot_job')}_pay_items.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with col_b:
        idr_info = {
            "date": idr_date,
            "contractor": contractor,
            "weather": weather,
            "work_description": work_description,
            "remarks": remarks,
            "inspector": inspector,
        }

        try:
            pdf_file = make_fillable_idr_pdf(metadata, idr_info, rows)
            st.download_button(
                label="Download Editable IDR PDF",
                data=pdf_file,
                file_name=format_pdf_filename(metadata.get("item_contract", "IDOT")),
                mime="application/pdf",
            )
        except Exception as e:
            st.error(f"Could not generate PDF: {e}")

    with st.expander("Preview selected/resolved IDR rows"):
        st.dataframe(pd.DataFrame(rows), use_container_width=True)
