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
TEMPLATE_CANDIDATES = [
    BASE_DIR / "IDR_Template.xlsx",
    BASE_DIR / "IDR_template.xlsx",
]
TEMPLATE_PATH = next((path for path in TEMPLATE_CANDIDATES if path.exists()), TEMPLATE_CANDIDATES[0])

BASE_URL = "https://webapps1.dot.illinois.gov"
IDOT_HOME_URL = "https://webapps1.dot.illinois.gov/WCTB/LBHome"

HIDDEN_PAY_ITEMS_SHEET_NAME = "IDOT_Pay_Items"
JOB_INFO_SHEET_NAME = "IDOT_Job_Info"
OLD_MATERIALS_SHEET_NAME = "Materials Data"


SEARCH_MAX_PAGES_PER_LETTING = 25

CELL_MAP = {
    "date": "C6",
    "contractor_label": "B8",
    "contractor": "D8",
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
# EXACT EXCEL-STYLE PDF HELPERS
# ============================================================

import shutil
import subprocess
import tempfile
from datetime import date as DateClass

PDF_ROW_COUNT = 6

WEATHER_OPTIONS = [
    "Sunny",
    "Cloudy",
    "Light Rain",
    "Normal Rain",
    "Heavy Rain",
    "Snow",
]


def get_today_default():
    return DateClass.today()


def format_report_date(value):
    if hasattr(value, "strftime"):
        return value.strftime("%-m/%-d/%Y") if hasattr(value, "strftime") else str(value)
    return clean_line(value)


def format_report_date_safe(value):
    if hasattr(value, "strftime"):
        return value.strftime("%m/%d/%Y")
    return clean_line(value)


def format_pdf_filename(contract_number):
    contract_number = clean_line(contract_number) or "IDOT"
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", contract_number).strip("_")
    return f"{safe}_IDR.pdf"


def format_xlsx_filename(contract_number):
    contract_number = clean_line(contract_number) or "IDOT"
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", contract_number).strip("_")
    return f"{safe}_IDR_filled.xlsx"


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


def parse_number(value):
    value = clean_line(value)
    if not value:
        return None
    value = value.replace(",", "").replace("$", "")
    value = re.sub(r"[^0-9.\-]", "", value)
    if not value or value in [".", "-", "-."]:
        return None
    try:
        return float(value)
    except Exception:
        return None


def get_quantity_status(quantity, plan_quantity):
    entered = parse_number(quantity)
    permitted = parse_number(plan_quantity)
    if entered is None or permitted is None or permitted <= 0:
        return {"status": "", "ratio": None, "color": ""}
    ratio = entered / permitted
    if entered > permitted:
        return {"status": f"OVER ({ratio:.0%})", "ratio": ratio, "color": "#ffc7ce"}
    if ratio >= 0.75:
        return {"status": f"Close ({ratio:.0%})", "ratio": ratio, "color": "#ffeb9c"}
    return {"status": f"OK ({ratio:.0%})", "ratio": ratio, "color": "#c6efce"}


def get_pay_item_options(pay_items):
    if pay_items is None or pay_items.empty:
        return [""], [""]
    codes = [""]
    descriptions = [""]
    for _, row in pay_items.iterrows():
        code = normalize_pay_item_code(row.get("item_code", ""))
        desc = clean_line(row.get("item_description", ""))
        if code and code not in codes:
            codes.append(code)
        if desc and desc not in descriptions:
            descriptions.append(desc)
    codes.append("Custom / Manual")
    descriptions.append("Custom / Manual")
    return codes, descriptions


def get_pay_item_by_code(pay_items, code):
    return dataframe_records_by_code(pay_items).get(normalize_pay_item_code(code))


def get_pay_item_by_description(pay_items, description):
    return dataframe_records_by_description(pay_items).get(clean_line(description))


def row_key(row_index, field):
    return f"idr_row_{row_index}_{field}"


def header_key(field):
    return f"idr_header_{field}"


def clear_idr_row_state():
    for row_index in range(PDF_ROW_COUNT):
        for field in [
            "item_code", "item_description", "custom_code", "custom_description",
            "location", "quantity", "unit", "custom_unit", "plan_quantity",
            "unit_price", "is_custom",
        ]:
            st.session_state.pop(row_key(row_index, field), None)


def ensure_row_defaults(row_index):
    defaults = {
        "item_code": "",
        "item_description": "",
        "custom_code": "",
        "custom_description": "",
        "location": "",
        "quantity": "",
        "unit": "",
        "custom_unit": "",
        "plan_quantity": "",
        "unit_price": "",
        "is_custom": False,
    }
    for field, value in defaults.items():
        st.session_state.setdefault(row_key(row_index, field), value)


def set_row_from_official_item(row_index, item):
    st.session_state[row_key(row_index, "item_code")] = clean_line(item.get("item_code", ""))
    st.session_state[row_key(row_index, "item_description")] = clean_line(item.get("item_description", ""))
    st.session_state[row_key(row_index, "unit")] = normalize_unit(item.get("unit", ""))
    st.session_state[row_key(row_index, "plan_quantity")] = clean_line(item.get("plan_quantity", ""))
    st.session_state[row_key(row_index, "unit_price")] = clean_line(item.get("unit_price", ""))
    st.session_state[row_key(row_index, "is_custom")] = False


def on_item_code_change(row_index, pay_items):
    code = st.session_state.get(row_key(row_index, "item_code"), "")
    if code == "Custom / Manual":
        st.session_state[row_key(row_index, "item_description")] = "Custom / Manual"
        st.session_state[row_key(row_index, "is_custom")] = True
        st.session_state[row_key(row_index, "plan_quantity")] = ""
        st.session_state[row_key(row_index, "unit_price")] = ""
        return
    if not code:
        st.session_state[row_key(row_index, "item_description")] = ""
        st.session_state[row_key(row_index, "unit")] = ""
        st.session_state[row_key(row_index, "plan_quantity")] = ""
        st.session_state[row_key(row_index, "unit_price")] = ""
        st.session_state[row_key(row_index, "is_custom")] = False
        return
    item = get_pay_item_by_code(pay_items, code)
    if item:
        set_row_from_official_item(row_index, item)


def on_item_description_change(row_index, pay_items):
    description = st.session_state.get(row_key(row_index, "item_description"), "")
    if description == "Custom / Manual":
        st.session_state[row_key(row_index, "item_code")] = "Custom / Manual"
        st.session_state[row_key(row_index, "is_custom")] = True
        st.session_state[row_key(row_index, "plan_quantity")] = ""
        st.session_state[row_key(row_index, "unit_price")] = ""
        return
    if not description:
        st.session_state[row_key(row_index, "item_code")] = ""
        st.session_state[row_key(row_index, "unit")] = ""
        st.session_state[row_key(row_index, "plan_quantity")] = ""
        st.session_state[row_key(row_index, "unit_price")] = ""
        st.session_state[row_key(row_index, "is_custom")] = False
        return
    item = get_pay_item_by_description(pay_items, description)
    if item:
        set_row_from_official_item(row_index, item)


def on_custom_row_change(row_index):
    if st.session_state.get(row_key(row_index, "is_custom"), False):
        st.session_state[row_key(row_index, "unit")] = normalize_unit(
            st.session_state.get(row_key(row_index, "custom_unit"), "")
        )


def get_row_for_output(row_index):
    is_custom = bool(st.session_state.get(row_key(row_index, "is_custom"), False))
    if is_custom:
        return {
            "item_code": normalize_pay_item_code(st.session_state.get(row_key(row_index, "custom_code"), "")),
            "item_description": clean_line(st.session_state.get(row_key(row_index, "custom_description"), "")),
            "location": clean_line(st.session_state.get(row_key(row_index, "location"), "")),
            "quantity": clean_line(st.session_state.get(row_key(row_index, "quantity"), "")),
            "unit": normalize_unit(st.session_state.get(row_key(row_index, "custom_unit"), "")),
            "plan_quantity": "",
            "unit_price": "",
            "is_custom": True,
        }
    return {
        "item_code": clean_line(st.session_state.get(row_key(row_index, "item_code"), "")),
        "item_description": clean_line(st.session_state.get(row_key(row_index, "item_description"), "")),
        "location": clean_line(st.session_state.get(row_key(row_index, "location"), "")),
        "quantity": clean_line(st.session_state.get(row_key(row_index, "quantity"), "")),
        "unit": normalize_unit(st.session_state.get(row_key(row_index, "unit"), "")),
        "plan_quantity": clean_line(st.session_state.get(row_key(row_index, "plan_quantity"), "")),
        "unit_price": clean_line(st.session_state.get(row_key(row_index, "unit_price"), "")),
        "is_custom": False,
    }


def quantity_status_badge_html(quantity, plan_quantity):
    status = get_quantity_status(quantity, plan_quantity)
    if not status["status"]:
        return "<div class='qty-badge qty-empty'>-</div>"
    return f"<div class='qty-badge' style='background:{status['color']};'>{status['status']}</div>"


def build_idr_header_form():
    st.subheader("IDR Header / Top of Form")
    st.caption(
        "Fill these boxes exactly like the top and signature sections of the Excel IDR. "
        "The labels shown here explain where each value prints on the final PDF."
    )

    with st.expander("Show field guide", expanded=False):
        st.markdown(
            """
            - **Date** → prints in the top-left date box and also fills the date boxes beside Inspected/Measured/Calculated/Checked.
            - **Contractor or Sub.** → prints on the Contractor/Subcontractor line.
            - **Weather** → prints on the Weather line.
            - **Inspected by / Measured by / Calculated by / Checked by** → prints in the signature/initial boxes on the right side of the form.
            - **This is** → checks either Estimated Progress Measurement or Final Field Measurement.
            - **Estimated item no. / Final item no.** → prints inside the parentheses beside the selected measurement checkbox.
            - **Remarks** → prints in the expanded Remarks box under the measurement section.
            """
        )

    st.markdown("**Top form fields**")
    row1 = st.columns([1.0, 2.0, 1.4, 1.1, 1.1, 1.1, 1.1])
    with row1[0]:
        idr_date = st.date_input("Date", value=get_today_default(), key=header_key("date"))
    with row1[1]:
        contractor = st.text_input("Contractor or Sub.", key=header_key("contractor"))
    with row1[2]:
        weather = st.selectbox("Weather", [""] + WEATHER_OPTIONS, key=header_key("weather"))
    with row1[3]:
        inspected_by = st.text_input("Inspected by", key=header_key("inspected_by"))
    with row1[4]:
        measured_by = st.text_input("Measured by", key=header_key("measured_by"))
    with row1[5]:
        calculated_by = st.text_input("Calculated by", key=header_key("calculated_by"))
    with row1[6]:
        checked_by = st.text_input("Checked by", key=header_key("checked_by"))

    st.markdown("**Measurement and remarks fields**")
    row2 = st.columns([1.0, 1.2, 1.2, 4.0])
    with row2[0]:
        measurement_type = st.selectbox(
            "This is",
            ["", "Estimated progress measurement", "Final field measurement"],
            key=header_key("measurement_type"),
        )
    with row2[1]:
        estimated_item_no = st.text_input("Estimated item no.", key=header_key("estimated_item_no"))
    with row2[2]:
        final_item_no = st.text_input("Final item no.", key=header_key("final_item_no"))
    with row2[3]:
        remarks = st.text_area("Remarks", height=70, key=header_key("remarks"))

    return {
        "date": idr_date,
        "contractor": contractor,
        "weather": weather,
        "inspected_by": inspected_by,
        "measured_by": measured_by,
        "calculated_by": calculated_by,
        "checked_by": checked_by,
        "measurement_type": measurement_type,
        "estimated_item_no": estimated_item_no,
        "final_item_no": final_item_no,
        "remarks": remarks,
    }


def build_idr_rows_form(pay_items):
    st.subheader("IDR Pay Item Table")
    st.caption(
        "Use this like the Excel item table. The column headers explain each box. "
        "Item Code and Item Description are searchable dropdowns. Selecting either one fills the other and the unit. "
        "Quantity status is website-only and will not print on the PDF."
    )
    st.info(
        "Table guide: Item Code # = IDOT pay item number, Fund = fund code if needed, "
        "Item = pay item description, Location = where work was performed, "
        "Quantity = amount used today, Unit = auto-filled unit, Status = quantity warning for the website only."
    )

    code_options, description_options = get_pay_item_options(pay_items)

    st.markdown(
        """
        <style>
        .idr-table-header {font-weight: 700; font-size: 0.82rem; padding: 0.25rem 0; border-bottom: 1px solid #d0d0d0;}
        .qty-badge {min-height: 36px; border: 1px solid #999; border-radius: 4px; display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 0.78rem; color: #111; margin-top: 1.70rem;}
        .qty-empty {background: #f4f4f4; color: #777; font-weight: 400;}
        </style>
        """,
        unsafe_allow_html=True,
    )

    header_cols = st.columns([1.1, 0.7, 3.1, 2.2, 1.05, 0.9, 0.95])
    headers = ["Item Code #", "Fund", "Item", "Location", "Quantity", "Unit", "Status"]
    for col, header in zip(header_cols, headers):
        col.markdown(f"<div class='idr-table-header'>{header}</div>", unsafe_allow_html=True)

    rows = []
    for row_index in range(PDF_ROW_COUNT):
        ensure_row_defaults(row_index)
        row_cols = st.columns([1.1, 0.7, 3.1, 2.2, 1.05, 0.9, 0.95])

        current_code = st.session_state.get(row_key(row_index, "item_code"), "")
        current_desc = st.session_state.get(row_key(row_index, "item_description"), "")
        if current_code not in code_options:
            st.session_state[row_key(row_index, "item_code")] = ""
        if current_desc not in description_options:
            st.session_state[row_key(row_index, "item_description")] = ""

        with row_cols[0]:
            st.selectbox(
                f"Row {row_index + 1} Item Code",
                code_options,
                key=row_key(row_index, "item_code"),
                on_change=on_item_code_change,
                args=(row_index, pay_items),
                label_visibility="collapsed",
            )
        with row_cols[1]:
            st.text_input(f"Row {row_index + 1} Fund", key=row_key(row_index, "fund_code"), label_visibility="collapsed")
        with row_cols[2]:
            st.selectbox(
                f"Row {row_index + 1} Item Description",
                description_options,
                key=row_key(row_index, "item_description"),
                on_change=on_item_description_change,
                args=(row_index, pay_items),
                label_visibility="collapsed",
            )

        is_custom = bool(st.session_state.get(row_key(row_index, "is_custom"), False))
        if is_custom:
            with row_cols[0]:
                st.text_input(
                    f"Row {row_index + 1} Custom Code",
                    key=row_key(row_index, "custom_code"),
                    placeholder="Custom ID",
                    label_visibility="collapsed",
                )
            with row_cols[2]:
                st.text_input(
                    f"Row {row_index + 1} Custom Description",
                    key=row_key(row_index, "custom_description"),
                    placeholder="Custom item description",
                    label_visibility="collapsed",
                )

        with row_cols[3]:
            st.text_input(f"Row {row_index + 1} Location", key=row_key(row_index, "location"), label_visibility="collapsed")
        with row_cols[4]:
            st.text_input(f"Row {row_index + 1} Quantity", key=row_key(row_index, "quantity"), label_visibility="collapsed")
        with row_cols[5]:
            if is_custom:
                st.text_input(
                    f"Row {row_index + 1} Unit",
                    key=row_key(row_index, "custom_unit"),
                    placeholder="Unit",
                    label_visibility="collapsed",
                    on_change=on_custom_row_change,
                    args=(row_index,),
                )
            else:
                st.text_input(f"Row {row_index + 1} Unit", key=row_key(row_index, "unit"), label_visibility="collapsed", disabled=True)

        row = get_row_for_output(row_index)
        row["fund_code"] = clean_line(st.session_state.get(row_key(row_index, "fund_code"), ""))
        rows.append(row)
        with row_cols[6]:
            st.markdown(quantity_status_badge_html(row.get("quantity", ""), row.get("plan_quantity", "")), unsafe_allow_html=True)

    return rows




def safe_set(ws, cell, value):
    """Write to a cell safely, including merged cells."""
    for merged_range in ws.merged_cells.ranges:
        if cell in merged_range:
            top_left_cell = merged_range.coord.split(":")[0]
            ws[top_left_cell] = value
            return
    ws[cell] = value


def get_merged_anchor_cell(ws, cell_address):
    for merged_range in ws.merged_cells.ranges:
        if cell_address in merged_range:
            return merged_range.coord.split(":")[0]
    return cell_address


def get_description_font_size(description):
    description_length = len(clean_line(description))
    for max_length, font_size in DESCRIPTION_TEXT_FIT_RULES:
        if description_length <= max_length:
            return font_size
    return 6


def make_font_with_size(original_font, size):
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
    anchor_address = get_merged_anchor_cell(ws, cell_address)
    value = ws[anchor_address].value
    return "" if value is None else str(value)


def format_item_description_cells(ws):
    for row in range(13, 19):
        visible_cell_address = f"D{row}"
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

def format_quantity_cells(ws):
    """
    Give the quantity/unit cells a little breathing room so the left edge
    of the number does not get clipped in the exported PDF.

    The template stores the visible quantity box around H13:H18. LibreOffice
    PDF export can render left-aligned text too tight against the border,
    so we keep the same Excel layout but center the value vertically and
    add a small indent.
    """
    for row in range(13, 19):
        visible_cell_address = f"H{row}"
        anchor_cell_address = get_merged_anchor_cell(ws, visible_cell_address)
        anchor_cell = ws[anchor_cell_address]
        current_alignment = copy(anchor_cell.alignment)

        anchor_cell.alignment = Alignment(
            horizontal="left",
            vertical=current_alignment.vertical or "center",
            text_rotation=current_alignment.text_rotation,
            wrap_text=current_alignment.wrap_text,
            shrink_to_fit=False,
            indent=1,
            relativeIndent=current_alignment.relativeIndent,
            justifyLastLine=current_alignment.justifyLastLine,
            readingOrder=current_alignment.readingOrder,
        )

        # Apply the same alignment across a merged quantity range if H is merged.
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
                        except Exception:
                            pass
                break


def unmerge_range_keep_style(ws, range_coord):
    target = None
    for merged_range in list(ws.merged_cells.ranges):
        if str(merged_range) == range_coord:
            target = merged_range
            break
    if target is None:
        return
    anchor = ws.cell(target.min_row, target.min_col)
    saved = {
        "font": copy(anchor.font),
        "fill": copy(anchor.fill),
        "border": copy(anchor.border),
        "alignment": copy(anchor.alignment),
        "number_format": anchor.number_format,
        "protection": copy(anchor.protection),
    }
    ws.unmerge_cells(range_coord)
    for row in range(target.min_row, target.max_row + 1):
        for col in range(target.min_col, target.max_col + 1):
            cell = ws.cell(row, col)
            cell.font = copy(saved["font"])
            cell.fill = copy(saved["fill"])
            cell.border = copy(saved["border"])
            cell.alignment = copy(saved["alignment"])
            cell.number_format = saved["number_format"]
            cell.protection = copy(saved["protection"])


def prepare_exact_print_layout(wb, ws):
    wb.active = wb.sheetnames.index(ws.title)
    for sheet in wb.worksheets:
        if sheet.title != ws.title:
            sheet.sheet_state = "hidden"

    ws.print_area = "A2:N30"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25
    ws.page_margins.header = 0
    ws.page_margins.footer = 0


def copy_cell_style(source_cell, target_cell):
    """Copy the visible Excel style from one cell to another."""
    target_cell.font = copy(source_cell.font)
    target_cell.fill = copy(source_cell.fill)
    target_cell.border = copy(source_cell.border)
    target_cell.alignment = copy(source_cell.alignment)
    target_cell.number_format = source_cell.number_format
    target_cell.protection = copy(source_cell.protection)


def rebuild_bottom_section_layout(ws):
    """
    Move the measurement checkbox section directly under the pay item table,
    expand the remarks area, and place the footer text where requested.

    New bottom layout:
    - Rows 19-20: "This is" measurement checkboxes
    - Rows 21-29: larger remarks box
    - A30: printed date
    - M30: BC 628 revision text
    """
    # Save styles before unmerging/clearing the old template area.
    label_style_source = ws["B21"]
    checkbox_style_source = ws["C21"]
    measurement_text_style_source = ws["D21"]
    close_paren_style_source = ws["H21"]
    remarks_label_style_source = ws["B25"]
    remarks_box_style_source = ws["C25"]
    footer_left_style_source = ws["B30"]
    footer_right_style_source = ws["N30"]

    # Remove old and possible regenerated merges in the bottom area.
    for range_coord in [
        "D21:G21", "D23:G23", "C25:N29",
        "D19:G19", "D20:G20", "C21:N29",
        "A30:D30", "M30:N30",
    ]:
        try:
            unmerge_range_keep_style(ws, range_coord)
        except Exception:
            pass

    # Clear old bottom values. Keep row heights/columns from the original template.
    for row in range(19, 30):
        for col in range(1, 15):
            ws.cell(row=row, column=col).value = None

    for col in range(1, 15):
        ws.cell(row=30, column=col).value = None

    # Rebuild the moved measurement section.
    safe_set(ws, "B19", "This is:")
    safe_set(ws, "C19", "☐")
    safe_set(ws, "D19", "an estimated progress measurement (item no.:")
    safe_set(ws, "H19", ")")
    safe_set(ws, "C20", "☐")
    safe_set(ws, "D20", "a final field measurement (item no.:")
    safe_set(ws, "H20", ")")

    # Rebuild the larger remarks section.
    safe_set(ws, "B21", "Remarks:")
    safe_set(ws, "C21", "")

    # Footer values are filled later, but the display cells are moved here.
    safe_set(ws, "A30", "")
    safe_set(ws, "M30", "BC 628 (Rev. 8/04)")

    # Recreate merged regions for long text.
    try:
        ws.merge_cells("D19:G19")
    except Exception:
        pass
    try:
        ws.merge_cells("D20:G20")
    except Exception:
        pass
    try:
        ws.merge_cells("C21:N29")
    except Exception:
        pass
    try:
        ws.merge_cells("A30:D30")
    except Exception:
        pass
    try:
        ws.merge_cells("M30:N30")
    except Exception:
        pass

    # Apply styles to the moved section.
    for cell_addr in ["B19"]:
        copy_cell_style(label_style_source, ws[cell_addr])

    for cell_addr in ["C19", "C20"]:
        copy_cell_style(checkbox_style_source, ws[cell_addr])

    for cell_addr in ["D19", "D20"]:
        copy_cell_style(measurement_text_style_source, ws[cell_addr])
        ws[cell_addr].alignment = Alignment(
            horizontal=ws[cell_addr].alignment.horizontal or "left",
            vertical=ws[cell_addr].alignment.vertical or "center",
            wrap_text=False,
        )

    for cell_addr in ["H19", "H20"]:
        copy_cell_style(close_paren_style_source, ws[cell_addr])

    copy_cell_style(remarks_label_style_source, ws["B21"])

    # Style the expanded remarks merged box.
    remarks_anchor = ws["C21"]
    copy_cell_style(remarks_box_style_source, remarks_anchor)
    remarks_anchor.alignment = Alignment(
        horizontal="left",
        vertical="top",
        wrap_text=True,
        shrink_to_fit=False,
    )

    for row in range(21, 30):
        for col in range(3, 15):
            try:
                copy_cell_style(remarks_anchor, ws.cell(row=row, column=col))
            except Exception:
                pass

    # Footer styling. A30/M30 are merged anchors.
    copy_cell_style(footer_left_style_source, ws["A30"])
    ws["A30"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
    copy_cell_style(footer_right_style_source, ws["M30"])
    ws["M30"].alignment = Alignment(horizontal="right", vertical="center", wrap_text=False)



def clear_exact_idr_values(ws):
    # Header values only - do not clear labels.
    for cell in [
        "C6", "D8", "C10", "G6", "H6", "G7", "H7", "G8", "H8", "G9", "H9",
        "L2", "L3", "L4", "L5", "L6", "L7", "L8",
        "C21", "H19", "H20", "A30", "M30",
    ]:
        safe_set(ws, cell, "")

    # Restore checkboxes and closing parentheses in the moved measurement section.
    safe_set(ws, "C19", "☐")
    safe_set(ws, "C20", "☐")
    safe_set(ws, "H19", ")")
    safe_set(ws, "H20", ")")
    safe_set(ws, "M30", "BC 628 (Rev. 8/04)")

    for row in range(13, 19):
        for col in ["B", "C", "D", "F", "H", "I", "J", "M"]:
            safe_set(ws, f"{col}{row}", "")


def fill_exact_idr_workbook(metadata, idr_info, rows):
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"Template file not found. Put IDR_Template.xlsx in the same folder as this app file. Expected: {TEMPLATE_PATH}"
        )

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb["IDR_Form"] if "IDR_Form" in wb.sheetnames else wb.active

    # Split the initial/date line cells so the right-side dates can be filled cleanly.
    for merged in ["G6:H6", "G7:H7", "G8:H8", "G9:H9"]:
        unmerge_range_keep_style(ws, merged)

    rebuild_bottom_section_layout(ws)
    clear_exact_idr_values(ws)

    report_date = format_report_date(idr_info.get("date"))
    safe_set(ws, "C6", report_date)
    safe_set(ws, "A30", f"Printed {report_date}")
    safe_set(ws, "M30", "BC 628 (Rev. 8/04)")

    # Header/manual fields.
    safe_set(ws, "D8", idr_info.get("contractor", ""))
    safe_set(ws, "C10", idr_info.get("weather", ""))
    safe_set(ws, "G6", idr_info.get("inspected_by", ""))
    safe_set(ws, "H6", report_date)
    safe_set(ws, "G7", idr_info.get("measured_by", ""))
    safe_set(ws, "H7", report_date)
    safe_set(ws, "G8", idr_info.get("calculated_by", ""))
    safe_set(ws, "H8", report_date)
    safe_set(ws, "G9", idr_info.get("checked_by", ""))
    safe_set(ws, "H9", report_date)

    # Job metadata from IDOT.
    safe_set(ws, "L2", metadata.get("county", ""))
    safe_set(ws, "L3", metadata.get("key_route", ""))
    safe_set(ws, "L4", metadata.get("marked_route", ""))
    safe_set(ws, "L5", metadata.get("district", ""))
    safe_set(ws, "L6", metadata.get("item_contract", ""))
    safe_set(ws, "L7", metadata.get("state_job", ""))
    safe_set(ws, "L8", metadata.get("federal_project", ""))

    measurement_type = clean_line(idr_info.get("measurement_type", ""))
    if measurement_type == "Estimated progress measurement":
        safe_set(ws, "C19", "☒")
        safe_set(ws, "H19", (clean_line(idr_info.get("estimated_item_no", "")) + ")") if clean_line(idr_info.get("estimated_item_no", "")) else ")")
    elif measurement_type == "Final field measurement":
        safe_set(ws, "C20", "☒")
        safe_set(ws, "H20", (clean_line(idr_info.get("final_item_no", "")) + ")") if clean_line(idr_info.get("final_item_no", "")) else ")")

    safe_set(ws, "C21", idr_info.get("remarks", ""))

    for i in range(PDF_ROW_COUNT):
        excel_row = 13 + i
        row = rows[i] if i < len(rows) else {}
        code = clean_line(row.get("item_code", ""))
        if code == "Custom / Manual":
            code = ""
        desc = clean_line(row.get("item_description", ""))
        if desc == "Custom / Manual":
            desc = ""
        qty = clean_line(row.get("quantity", ""))
        unit = normalize_unit(row.get("unit", ""))
        qty_unit = clean_line(f"{qty} {unit}") if qty or unit else ""

        safe_set(ws, f"B{excel_row}", code)
        safe_set(ws, f"C{excel_row}", clean_line(row.get("fund_code", "")))
        safe_set(ws, f"D{excel_row}", desc)
        safe_set(ws, f"F{excel_row}", clean_line(row.get("location", "")))
        safe_set(ws, f"H{excel_row}", qty_unit)

    try:
        format_item_description_cells(ws)
    except Exception:
        pass

    try:
        format_quantity_cells(ws)
    except Exception:
        pass

    prepare_exact_print_layout(wb, ws)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def find_libreoffice_executable():
    return shutil.which("libreoffice") or shutil.which("soffice")


def convert_xlsx_bytes_to_pdf(xlsx_bytes):
    executable = find_libreoffice_executable()
    if not executable:
        raise RuntimeError(
            "LibreOffice is required for exact Excel-to-PDF output. Add an apt packages file with libreoffice installed."
        )

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        xlsx_path = tmpdir_path / "filled_idr.xlsx"
        pdf_path = tmpdir_path / "filled_idr.pdf"
        xlsx_path.write_bytes(xlsx_bytes.getvalue())

        result = subprocess.run(
            [
                executable,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(tmpdir_path),
                str(xlsx_path),
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=90,
        )

        if result.returncode != 0 or not pdf_path.exists():
            raise RuntimeError(
                "LibreOffice could not convert the filled template to PDF.\n\n"
                f"stdout: {result.stdout}\n\nstderr: {result.stderr}"
            )

        return io.BytesIO(pdf_path.read_bytes())


def make_exact_idr_pdf(metadata, idr_info, rows):
    xlsx_output = fill_exact_idr_workbook(metadata, idr_info, rows)
    return convert_xlsx_bytes_to_pdf(xlsx_output)


def make_pay_items_excel(metadata, pay_items):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pay_items.to_excel(writer, index=False, sheet_name="Pay Items")
        info_df = pd.DataFrame([{"field": k, "value": v} for k, v in metadata.items()])
        info_df.to_excel(writer, index=False, sheet_name="Job Info")
        workbook = writer.book
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, 1000, 20)
            header_format = workbook.add_format({"bold": True, "border": 1, "text_wrap": True})
            cols = pay_items.columns if sheet_name == "Pay Items" else info_df.columns
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
    "The website fills the IDR form and exports a PDF that matches the Excel template layout."
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
        with st.spinner("Searching IDOT Transportation Bulletin archives newest to oldest..."):
            metadata, pay_items, match = fetch_idot_job(job_number)
        st.session_state.metadata = metadata
        st.session_state.pay_items = pay_items
        st.session_state.match = match
        clear_idr_row_state()
        st.success(f"Found {metadata.get('item_contract', job_number)} with {len(pay_items)} pay items.")
        if match.get("letting_url"):
            st.info(f"Found on letting/archive page: {match.get('letting', '')}, page {match.get('page', '')}")
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
        st.warning("Some job fields did not parse correctly: " + ", ".join(missing_fields))
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
    idr_info = build_idr_header_form()
    st.divider()
    rows = build_idr_rows_form(pay_items)
    st.divider()

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        pay_items_file = make_pay_items_excel(metadata, pay_items)
        st.download_button(
            label="Download IDOT Pay Item Table",
            data=pay_items_file,
            file_name=f"{metadata.get('item_contract', 'idot_job')}_pay_items.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col_b:
        try:
            filled_xlsx = fill_exact_idr_workbook(metadata, idr_info, rows)
            st.download_button(
                label="Download Filled IDR Excel Backup",
                data=filled_xlsx,
                file_name=format_xlsx_filename(metadata.get("item_contract", "IDOT")),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Could not prepare filled Excel backup: {e}")
    with col_c:
        try:
            pdf_file = make_exact_idr_pdf(metadata, idr_info, rows)
            st.download_button(
                label="Download Exact IDR PDF",
                data=pdf_file,
                file_name=format_pdf_filename(metadata.get("item_contract", "IDOT")),
                mime="application/pdf",
            )
        except Exception as e:
            st.error(f"Could not generate exact PDF: {e}")
            if not find_libreoffice_executable():
                st.info("On Streamlit Cloud, add a packages.txt file containing: libreoffice")
