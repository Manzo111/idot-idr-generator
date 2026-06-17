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
TEMPLATE_XLSM_PATH = BASE_DIR / "IDR_template.xlsm"
TEMPLATE_XLSX_PATH = BASE_DIR / "IDR_template.xlsx"

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
# EXCEL HELPERS
# ============================================================

def get_effective_cell_address(ws, cell):
    """
    If a cell is part of a merged range, return the top-left cell of that merged range.
    Otherwise, return the original cell.

    This is important because writing blank text into any cell inside a merged range
    actually writes to the top-left cell, which can accidentally erase labels.
    """
    for merged_range in ws.merged_cells.ranges:
        if cell in merged_range:
            return merged_range.coord.split(":")[0]

    return cell


def write_contractor_section(ws):
    """
    Keeps the Contractor/Subcontractor label on the Excel sheet.

    The bug was caused by this situation:
    - the template has a merged contractor area,
    - safe_set(ws, contractor_cell, "") writes to the top-left of that merged area,
    - that top-left cell is also where the label lives,
    - so the label gets erased.

    This function checks whether the label cell and contractor entry cell resolve
    to the same merged-cell anchor. If they do, it only writes the label and does
    not blank the contractor cell.
    """
    label_cell = CELL_MAP["contractor_label"]
    contractor_cell = CELL_MAP["contractor"]

    label_anchor = get_effective_cell_address(ws, label_cell)
    contractor_anchor = get_effective_cell_address(ws, contractor_cell)

    safe_set(ws, label_cell, CONTRACTOR_LABEL_TEXT)

    # Only clear the contractor entry area if it is actually separate from the label.
    # If they share one merged range, clearing contractor would erase the label.
    if contractor_anchor != label_anchor:
        safe_set(ws, contractor_cell, "")


def safe_set(ws, cell, value):
    try:
        for merged_range in ws.merged_cells.ranges:
            if cell in merged_range:
                top_left_cell = merged_range.coord.split(":")[0]
                ws[top_left_cell] = value
                return

        ws[cell] = value

    except Exception as e:
        raise ValueError(f"Could not write value '{value}' to cell {cell}: {e}")


def unlock_all_cells(ws):
    for row in ws.iter_rows():
        for cell in row:
            cell.protection = copy(cell.protection)
            cell.protection = cell.protection.copy(locked=False)


def lock_autofilled_cells(ws):
    locked_cells = [
        CELL_MAP["county"],
        CELL_MAP["section"],
        CELL_MAP["route"],
        CELL_MAP["district"],
        CELL_MAP["contract_number"],
        CELL_MAP["job_number"],
        CELL_MAP["project"],
    ]

    for cell_addr in locked_cells:
        for merged_range in ws.merged_cells.ranges:
            if cell_addr in merged_range:
                top_left_cell = merged_range.coord.split(":")[0]
                ws[top_left_cell].protection = ws[top_left_cell].protection.copy(locked=True)
                break
        else:
            ws[cell_addr].protection = ws[cell_addr].protection.copy(locked=True)

    for row in range(CELL_MAP["start_row"], CELL_MAP["end_row"] + 1):
        # Item code, item description, and quantity stay editable.
        # Unit is locked because it is auto-filled.
        unit_cell = ws[f"{CELL_MAP['unit_col']}{row}"]
        unit_cell.protection = unit_cell.protection.copy(locked=True)




def get_merged_anchor_cell(ws, cell_address):
    """
    Return the real cell Excel uses for formatting/writing.

    If a visible box is merged, Excel only stores the value/style on the
    top-left cell of the merged range. Formatting a non-anchor cell inside
    the merged range will not reliably show up in the generated workbook.
    """
    for merged_range in ws.merged_cells.ranges:
        if cell_address in merged_range:
            return merged_range.coord.split(":")[0]

    return cell_address







def get_description_font_size(description):
    """Return a readable font size based on the final wrapped description text."""
    description_length = len(clean_line(description))

    for max_length, font_size in DESCRIPTION_TEXT_FIT_RULES:
        if description_length <= max_length:
            return font_size

    return 6


def make_font_with_size(original_font, size):
    """
    Build a real openpyxl Font object with the same main style settings,
    but with a different font size.

    This is more reliable than mutating a copied style object because openpyxl
    styles are immutable/proxied internally.
    """
    return Font(
        name=original_font.name,
        sz=size,
        b=original_font.b,
        i=original_font.i,
        vertAlign=original_font.vertAlign,
        underline=original_font.underline,
        strike=original_font.strike,
        color=copy(original_font.color),
        scheme=original_font.scheme,
        family=original_font.family,
        charset=original_font.charset,
        outline=original_font.outline,
        shadow=original_font.shadow,
        condense=original_font.condense,
        extend=original_font.extend,
    )


def get_text_for_cell(ws, cell_address):
    """
    Get the value from the actual visible cell.

    If the visible cell is merged, Excel stores the value in the top-left
    merged anchor. This function returns that anchor value.
    """
    anchor_address = get_merged_anchor_cell(ws, cell_address)
    value = ws[anchor_address].value

    if value is None:
        return ""

    return str(value)


def format_item_description_cells(ws):
    """Apply wrapped text and font-only resizing to item description cells."""
    for row in range(CELL_MAP["start_row"], CELL_MAP["end_row"] + 1):
        visible_cell_address = f"{CELL_MAP['item_description_col']}{row}"
        anchor_cell_address = get_merged_anchor_cell(ws, visible_cell_address)
        anchor_cell = ws[anchor_cell_address]

        description_text = get_text_for_cell(ws, visible_cell_address)
        font_size = get_description_font_size(description_text)

        current_alignment = copy(anchor_cell.alignment)

        anchor_cell.alignment = Alignment(
            horizontal=current_alignment.horizontal or "left",
            vertical="top",
            text_rotation=current_alignment.text_rotation,
            wrap_text=True,
            shrink_to_fit=False,
            indent=current_alignment.indent,
            relativeIndent=current_alignment.relativeIndent,
            justifyLastLine=current_alignment.justifyLastLine,
            readingOrder=current_alignment.readingOrder,
        )

        anchor_cell.font = make_font_with_size(anchor_cell.font, font_size)

        for merged_range in ws.merged_cells.ranges:
            if visible_cell_address in merged_range:
                for cell_row in ws.iter_rows(
                    min_row=merged_range.min_row,
                    max_row=merged_range.max_row,
                    min_col=merged_range.min_col,
                    max_col=merged_range.max_col,
                ):
                    for range_cell in cell_row:
                        try:
                            range_cell.alignment = copy(anchor_cell.alignment)
                            range_cell.font = copy(anchor_cell.font)
                        except Exception:
                            pass
                break

        ws.row_dimensions[row].hidden = False
        ws.row_dimensions[row].collapsed = False


def create_or_replace_sheet(wb, sheet_name):
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]

    return wb.create_sheet(sheet_name)


def delete_sheet_if_exists(wb, sheet_name):
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]


def write_hidden_pay_items_sheet(wb, pay_items):
    ws = create_or_replace_sheet(wb, HIDDEN_PAY_ITEMS_SHEET_NAME)

    headers = [
        "item_code",
        "unit",
        "item_description",
        "quantity",
        "unit_price",
    ]

    for col_num, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col_num, value=header)

    for row_num, (_, row) in enumerate(pay_items.iterrows(), start=2):
        ws.cell(row=row_num, column=1, value=row.get("item_code", ""))
        ws.cell(row=row_num, column=2, value=row.get("unit", ""))
        ws.cell(row=row_num, column=3, value=row.get("item_description", ""))
        ws.cell(row=row_num, column=4, value=row.get("quantity", ""))
        ws.cell(row=row_num, column=5, value=row.get("unit_price", ""))

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 55
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 18

    ws.sheet_state = "hidden"

    return ws


def write_job_info_sheet(wb, metadata):
    ws = create_or_replace_sheet(wb, JOB_INFO_SHEET_NAME)

    rows = [
        ("contract_url", metadata.get("contract_url", "")),
        ("item_contract", metadata.get("item_contract", "")),
        ("letting_date", metadata.get("letting_date", "")),
        ("region", metadata.get("region", "")),
        ("district", metadata.get("district", "")),
        ("dbe_percent", metadata.get("dbe_percent", "")),
        ("vbp_percent", metadata.get("vbp_percent", "")),
        ("federal_project", metadata.get("federal_project", "")),
        ("county", metadata.get("county", "")),
        ("key_route", metadata.get("key_route", "")),
        ("marked_route", metadata.get("marked_route", "")),
        ("website_section", metadata.get("website_section", "")),
        ("state_job", metadata.get("state_job", "")),
        ("pps", metadata.get("pps", "")),
        ("working_days", metadata.get("working_days", "")),
    ]

    for i, (key, value) in enumerate(rows, start=1):
        ws.cell(row=i, column=1, value=key)
        ws.cell(row=i, column=2, value=value)

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 90
    ws.sheet_state = "hidden"

    return ws


def apply_item_code_dropdown(ws, pay_items_count):
    formula_range = f"='{HIDDEN_PAY_ITEMS_SHEET_NAME}'!$A$2:$A${pay_items_count + 1}"

    dv = DataValidation(
        type="list",
        formula1=formula_range,
        allow_blank=True,
    )

    ws.add_data_validation(dv)

    for row in range(CELL_MAP["start_row"], CELL_MAP["end_row"] + 1):
        dv.add(ws[f"{CELL_MAP['item_code_col']}{row}"])



def apply_description_dropdown(ws, pay_items_count):
    formula_range = f"='{HIDDEN_PAY_ITEMS_SHEET_NAME}'!$C$2:$C${pay_items_count + 1}"

    dv = DataValidation(
        type="list",
        formula1=formula_range,
        allow_blank=True,
    )

    ws.add_data_validation(dv)

    for row in range(CELL_MAP["start_row"], CELL_MAP["end_row"] + 1):
        dv.add(ws[f"{CELL_MAP['item_description_col']}{row}"])

def add_lookup_formulas_for_blank_rows(ws, pay_items_count):
    lookup_range = f"'{HIDDEN_PAY_ITEMS_SHEET_NAME}'!$A$2:$E${pay_items_count + 1}"

    for row in range(CELL_MAP["start_row"], CELL_MAP["end_row"] + 1):
        code_cell = f"{CELL_MAP['item_code_col']}{row}"
        desc_cell = f"{CELL_MAP['item_description_col']}{row}"
        unit_cell = f"{CELL_MAP['unit_col']}{row}"

        # Unit is formula-driven only.
        # It checks item description first because when the user selects an item,
        # the description cell changes immediately.
        # Then it falls back to item code.
        ws[unit_cell] = (
            f'=IF({desc_cell}<>"",'
            f'IFERROR(INDEX(\'{HIDDEN_PAY_ITEMS_SHEET_NAME}\'!$B$2:$B${pay_items_count + 1},'
            f'MATCH({desc_cell},\'{HIDDEN_PAY_ITEMS_SHEET_NAME}\'!$C$2:$C${pay_items_count + 1},0)),""),'
            f'IF({code_cell}<>"",'
            f'IFERROR(VLOOKUP({code_cell},{lookup_range},2,FALSE),""),'
            f'""))'
        )

def apply_quantity_conditional_formatting(ws, pay_items_count):
    start_row = CELL_MAP["start_row"]
    end_row = CELL_MAP["end_row"]

    code_col = CELL_MAP["item_code_col"]
    desc_col = CELL_MAP["item_description_col"]
    quantity_col = CELL_MAP["quantity_col"]

    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    lookup_range = f"'{HIDDEN_PAY_ITEMS_SHEET_NAME}'!$A$2:$E${pay_items_count + 1}"

    for row in range(start_row, end_row + 1):
        code_cell = f"${code_col}${row}"
        desc_cell = f"${desc_col}${row}"
        qty_cell = f"${quantity_col}${row}"
        qty_range = f"{quantity_col}{row}"

        # Look up the max/plan quantity from item description first,
        # then fall back to item code.
        max_qty_formula = (
            f'IF({desc_cell}<>"",'
            f'INDEX(\'{HIDDEN_PAY_ITEMS_SHEET_NAME}\'!$D$2:$D${pay_items_count + 1},'
            f'MATCH({desc_cell},\'{HIDDEN_PAY_ITEMS_SHEET_NAME}\'!$C$2:$C${pay_items_count + 1},0)),'
            f'IF({code_cell}<>"",'
            f'VLOOKUP({code_cell},{lookup_range},4,FALSE),'
            f'0))'
        )

        # Convert both the entered quantity and max quantity to numbers.
        # This fixes cases where IDOT quantities have commas or are stored as text.
        entered_qty_number = f'VALUE(SUBSTITUTE({qty_cell},",",""))'
        max_qty_number = f'VALUE(SUBSTITUTE({max_qty_formula},",",""))'

        has_item_formula = f'OR({code_cell}<>"",{desc_cell}<>"")'

        # Red: actual quantity is greater than the allowed/plan quantity.
        ws.conditional_formatting.add(
            qty_range,
            FormulaRule(
                formula=[
                    f'AND({qty_cell}<>"",{has_item_formula},IFERROR({entered_qty_number}>{max_qty_number},FALSE))'
                ],
                fill=red_fill,
                stopIfTrue=True,
            )
        )

        # Yellow: actual quantity is close, from 90% to 100% of allowed/plan quantity.
        ws.conditional_formatting.add(
            qty_range,
            FormulaRule(
                formula=[
                    f'AND({qty_cell}<>"",{has_item_formula},IFERROR({entered_qty_number}>={max_qty_number}*0.9,FALSE),IFERROR({entered_qty_number}<={max_qty_number},FALSE))'
                ],
                fill=yellow_fill,
                stopIfTrue=True,
            )
        )

        # Green: actual quantity is below 90% of allowed/plan quantity.
        ws.conditional_formatting.add(
            qty_range,
            FormulaRule(
                formula=[
                    f'AND({qty_cell}<>"",{has_item_formula},IFERROR({entered_qty_number}<{max_qty_number}*0.9,FALSE))'
                ],
                fill=green_fill,
                stopIfTrue=True,
            )
        )

def clear_old_idr_values(ws):
    clear_cells = [
        CELL_MAP["date"],
        CELL_MAP["work_description"],
        CELL_MAP["weather"],
        CELL_MAP["remarks"],

        CELL_MAP["county"],
        CELL_MAP["section"],
        CELL_MAP["route"],
        CELL_MAP["district"],
        CELL_MAP["contract_number"],
        CELL_MAP["job_number"],
        CELL_MAP["project"],
    ]

    for cell in clear_cells:
        safe_set(ws, cell, None)

    for row in range(CELL_MAP["start_row"], CELL_MAP["end_row"] + 1):
        safe_set(ws, f"{CELL_MAP['item_code_col']}{row}", None)
        safe_set(ws, f"{CELL_MAP['item_description_col']}{row}", None)
        safe_set(ws, f"{CELL_MAP['location_col']}{row}", None)
        safe_set(ws, f"{CELL_MAP['quantity_col']}{row}", None)
        safe_set(ws, f"{CELL_MAP['unit_col']}{row}", None)


def fill_idr_template(metadata, pay_items, selected_rows):
    wb = load_workbook(TEMPLATE_PATH, keep_vba=(TEMPLATE_PATH.suffix.lower() == ".xlsm"))

    ws = wb["IDR_Form"] if "IDR_Form" in wb.sheetnames else wb.active

    clear_old_idr_values(ws)

    safe_set(ws, CELL_MAP["date"], "")

    # Keep the Contractor/Subcontractor label on the Excel sheet.
    # Do not clear the contractor cell if it shares a merged range with the label.
    write_contractor_section(ws)

    safe_set(ws, CELL_MAP["work_description"], "")
    safe_set(ws, CELL_MAP["weather"], "")
    safe_set(ws, CELL_MAP["remarks"], "")

    safe_set(ws, CELL_MAP["county"], metadata.get("county", ""))
    safe_set(ws, CELL_MAP["section"], metadata.get("key_route", ""))
    safe_set(ws, CELL_MAP["route"], metadata.get("marked_route", ""))
    safe_set(ws, CELL_MAP["district"], metadata.get("district", ""))
    safe_set(ws, CELL_MAP["contract_number"], metadata.get("item_contract", ""))
    safe_set(ws, CELL_MAP["job_number"], metadata.get("state_job", ""))
    safe_set(ws, CELL_MAP["project"], metadata.get("federal_project", ""))

    delete_sheet_if_exists(wb, OLD_MATERIALS_SHEET_NAME)

    write_job_info_sheet(wb, metadata)
    write_hidden_pay_items_sheet(wb, pay_items)

    max_rows = CELL_MAP["end_row"] - CELL_MAP["start_row"] + 1
    selected_rows = selected_rows.head(max_rows)

    for i, (_, row_data) in enumerate(selected_rows.iterrows()):
        excel_row = CELL_MAP["start_row"] + i

        # User can select/change either item code or item description in Excel.
        # The xlsm macro in the template handles the two-way sync.
        safe_set(ws, f"{CELL_MAP['item_code_col']}{excel_row}", row_data.get("item_code", ""))
        safe_set(ws, f"{CELL_MAP['item_description_col']}{excel_row}", row_data.get("item_description", ""))

        # Unit is auto-filled by formula only.
        safe_set(ws, f"{CELL_MAP['unit_col']}{excel_row}", "")

        # Quantity is intentionally blank. User fills actual used quantity.
        safe_set(ws, f"{CELL_MAP['quantity_col']}{excel_row}", "")

    apply_item_code_dropdown(ws, len(pay_items))
    apply_description_dropdown(ws, len(pay_items))
    add_lookup_formulas_for_blank_rows(ws, len(pay_items))
    apply_quantity_conditional_formatting(ws, len(pay_items))

    format_item_description_cells(ws)

    unlock_all_cells(ws)
    lock_autofilled_cells(ws)

    ws.protection.sheet = True
    ws.protection.enable()

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    output = io.BytesIO()
    wb.save(output)
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
    page_title="IDOT Job IDR Generator",
    page_icon="📄",
    layout="wide",
)

st.title("IDOT Job IDR Generator")

st.write(
    "Enter an IDOT job/contract number or paste the direct IDOT contract URL. "
    "The app pulls the job information and pay items, then creates a job-specific IDR draft."
)

with st.sidebar:
    st.header("Job Lookup")

    job_number = st.text_input(
        "IDOT Job / Contract Number or Contract Detail URL",
        placeholder="Example: 62K33, 001-62K33, or paste the IDOT contract URL",
    )


if "metadata" not in st.session_state:
    st.session_state.metadata = None

if "pay_items" not in st.session_state:
    st.session_state.pay_items = pd.DataFrame()

if "match" not in st.session_state:
    st.session_state.match = None

if st.button("Find IDOT Job"):
    try:
        if not TEMPLATE_PATH.exists():
            st.error(
                f"Template file not found.\n\n"
                f"Put IDR_template.xlsm or IDR_template.xlsx in the same folder as this app file.\n\n"
                f"Expected location:\n{TEMPLATE_PATH}"
            )
        else:
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
            + ". Do not generate the IDR until these fields show correctly."
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

if not pay_items.empty:
    st.subheader("Pay Items Pulled from IDOT")

    st.write(
        "Select the material rows you want to place on the visible IDR draft. "
        "Only the first 6 selected rows fit on this printed IDR form."
    )

    pay_items_display = pay_items.copy()
    pay_items_display.insert(0, "use_on_idr", False)

    edited_items = st.data_editor(
        pay_items_display,
        num_rows="fixed",
        use_container_width=True,
        column_config={
            "use_on_idr": st.column_config.CheckboxColumn("Use on IDR"),
            "item_code": st.column_config.TextColumn("Item Code"),
            "unit": st.column_config.TextColumn("Unit / UOM"),
            "item_description": st.column_config.TextColumn("Description"),
            "quantity": st.column_config.TextColumn("Plan Quantity"),
            "unit_price": st.column_config.TextColumn("Unit Price"),
        },
    )

    selected_rows = edited_items[edited_items["use_on_idr"] == True].drop(
        columns=["use_on_idr"]
    )

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
        if st.button("Generate IDR Draft Excel"):
            try:
                if not TEMPLATE_PATH.exists():
                    st.error(
                        f"Template file not found.\n\n"
                        f"Put IDR_template.xlsm or IDR_template.xlsx in the same folder as this app file.\n\n"
                        f"Expected location:\n{TEMPLATE_PATH}"
                    )
                else:
                    output = fill_idr_template(
                        metadata=metadata,
                        pay_items=pay_items,
                        selected_rows=selected_rows,
                    )

                    st.download_button(
                        label="Download IDR Draft",
                        data=output,
                        file_name=f"{metadata.get('item_contract', 'IDOT')}_IDR_Draft.xlsm" if TEMPLATE_PATH.suffix.lower() == ".xlsm" else f"{metadata.get('item_contract', 'IDOT')}_IDR_Draft.xlsx",
                        mime="application/vnd.ms-excel.sheet.macroEnabled.12" if TEMPLATE_PATH.suffix.lower() == ".xlsm" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

            except Exception as e:
                st.error(f"Could not generate IDR: {e}")
