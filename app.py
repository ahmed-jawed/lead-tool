"""
Lead Enrichment Tool
Upload scraped lead sheets (CSV / Excel) -> the tool visits each business website,
finds the email + owner name, works out the business type and working region,
removes repeated businesses, and gives you a clean table + CSV download.
"""

import json
import re
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import unquote, urljoin, urlparse

import pandas as pd
import requests
import streamlit as st
import urllib3
from bs4 import BeautifulSoup

urllib3.disable_warnings()

# =====================================================================
# SETTINGS
# =====================================================================
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "en",
}

LINK_KEYWORDS = ("contact", "about", "team", "our-story", "who-we-are", "meet", "owner")
FALLBACK_PATHS = ("/contact", "/contact-us", "/about", "/about-us", "/our-team", "/team")

SOCIAL_HOSTS = (
    "facebook.com", "instagram.com", "linkedin.com", "google.com", "goo.gl",
    "youtube.com", "tiktok.com", "twitter.com", "x.com", "hipages.com.au",
    "oneflare.com.au", "yellowpages.com.au", "truelocal.com.au", "yelp.com",
)

FREE_MAIL = {
    "gmail.com", "googlemail.com", "outlook.com", "hotmail.com", "live.com", "msn.com",
    "yahoo.com", "ymail.com", "icloud.com", "me.com", "proton.me", "protonmail.com",
    "aol.com", "gmx.com", "outlook.com.au", "hotmail.com.au", "yahoo.com.au",
    "live.com.au", "bigpond.com", "bigpond.net.au", "optusnet.com.au", "tpg.com.au",
    "iinet.net.au", "westnet.com.au", "internode.on.net", "hotmail.co.uk",
    "yahoo.co.uk", "outlook.co.uk", "btinternet.com", "sky.com", "rogers.com",
    "shaw.ca", "sympatico.ca", "bell.net", "telus.net",
}

# =====================================================================
# EMAIL + NAME EXTRACTION
# =====================================================================
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
JUNK_EMAIL_PARTS = (
    "sentry", "wixpress", "example.", "domain.com", "yourdomain", "email.com",
    "godaddy", "schema.org", "wordpress.", "yoursite", "test@", "you@", "name@",
    "user@", "@google.", "@gstatic.", "@facebook.", "@instagram.", "@wix.",
    "@squarespace.", "@cloudflare.", "@hubspot.", "@mailchimp.", "@shopify.",
    "@w3.org", "@apple.", "@microsoft.", "@yelp.", "@2x",
    ".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".css", ".js",
)

ROLE = (r"(?i:owner(?:\s*(?:/|&|and)\s*operator)?|co-?founder|founder|"
        r"managing\s+director|director|proprietor|president|ceo)")
TOKEN = r"(?:O['’][A-Z][a-z]+|[A-Z][a-z]+(?:[A-Z][a-z]+)?(?:[\-][A-Z][a-z]+)?)"
NAME = r"(" + TOKEN + r"(?:\s" + TOKEN + r"){1,3})"
NAME_PATTERNS = [
    re.compile(ROLE + r"\s*[:\-–—|]\s*" + NAME),
    re.compile(NAME + r"\s*[,\-–—|(]\s*" + ROLE),
    re.compile(r"(?i:founded|owned|started|established|run|operated|led)\s+(?i:and\s+operated\s+)?(?i:by)\s+" + NAME),
    re.compile(r"(?i:owner\s+is|founder\s+is)\s+" + NAME),
]
NAME_STOPWORDS = {
    "meet", "hi", "hello", "about", "contact", "our", "the", "team", "by", "from",
    "your", "we", "us", "call", "email", "mr", "mrs", "ms", "dr", "licensed",
    "family", "owned", "operated", "local", "founded", "managing", "director",
    "owner", "founder", "ceo", "president", "and", "of", "at", "with", "led",
}
NAME_BLACKLIST = {
    "roofing", "roof", "roofs", "services", "service", "company", "contractors",
    "contractor", "construction", "home", "homes", "repair", "repairs", "restoration",
    "solutions", "group", "limited", "ltd", "inc", "llc", "pty", "pty.", "builders",
    "building", "maintenance", "installation", "replacement", "commercial",
    "residential", "free", "quote", "estimate", "get", "read", "more", "years",
    "experience", "quality", "best", "top", "north", "south", "east", "west",
    "new", "old", "united", "kingdom", "states", "australia", "canada", "qatar",
    "street", "road", "avenue", "terms", "privacy", "policy", "all", "rights",
    "reserved", "copyright", "gutters", "guttering", "metal", "tile", "tiles",
    "slate", "flat", "pitched", "emergency", "specialists", "specialist",
    "experts", "expert", "professional", "trusted", "premier", "reliable",
    "professionals", "who", "care", "bio", "shield", "plumbing", "plumbers",
    "gutter", "guard", "clean", "coat", "painting", "electrical", "air",
    "conditioning", "australian", "sydney", "melbourne", "brisbane", "perth",
    "pro", "central", "wide", "skyline",
}

# =====================================================================
# BUSINESS TYPE
# =====================================================================
TYPE_KEYWORDS = {
    "Roofing": ["roof", "roofing", "roofer", "roofers", "re-roof", "reroof", "colorbond",
                "roof tile", "roof restoration", "roof repair", "metal roof", "slate roof"],
    "Plumbing": ["plumb", "plumber", "plumbers", "plumbing", "hot water", "blocked drain",
                 "drainage", "gas fitting", "gasfitting"],
    "Electrical": ["electric", "electrician", "electricians", "electrical", "switchboard", "wiring"],
    "Gutters & Downpipes": ["gutter", "guttering", "downpipe", "fascia", "leaf guard", "gutter guard"],
    "Air Conditioning & HVAC": ["air conditioning", "aircon", "hvac", "heating and cooling",
                                "refrigeration", "ducted"],
    "Solar": ["solar", "photovoltaic", "battery storage"],
    "Building & Construction": ["builder", "builders", "building", "construction", "renovation",
                                "renovations", "carpentry", "carpenter", "extension", "remodel"],
    "Painting": ["painter", "painters", "painting"],
    "Cleaning": ["cleaning", "cleaner", "cleaners", "pressure wash", "pressure cleaning"],
    "Landscaping & Fencing": ["landscap", "garden", "fencing", "fence", "paving", "concreting", "turf"],
    "Pest Control": ["pest", "termite"],
    "Flooring & Tiling": ["tiler", "tilers", "tiling", "flooring"],
    "Waterproofing": ["waterproof"],
    "Insulation": ["insulation"],
}
TYPE_PATTERNS = {
    t: [re.compile(r"\b" + re.escape(k)) for k in kws] for t, kws in TYPE_KEYWORDS.items()
}


def classify_type(name_text, meta_text, body_text):
    """Score each business type: name hits count most, then title/headings, then body text."""
    name_text, meta_text, body_text = name_text.lower(), meta_text.lower(), body_text.lower()
    scores = {}
    for t, pats in TYPE_PATTERNS.items():
        s = 0
        for p in pats:
            if p.search(name_text):
                s += 6
            if p.search(meta_text):
                s += 3
            s += min(len(p.findall(body_text)), 6) * 0.5
        scores[t] = s
    best = max(scores, key=scores.get)
    return best if scores[best] >= 3 else "Unknown"


# =====================================================================
# REGION
# =====================================================================
# (canonical region, country, [aliases], needs_country_hint)
REGIONS = [
    ("Sydney, NSW", "AU", ["sydney", "parramatta", "penrith", "blacktown", "campbelltown", "hornsby",
                           "the hills district", "hills district", "northern beaches", "north shore",
                           "inner west", "sutherland shire", "st george", "eastern suburbs",
                           "bankstown", "castle hill", "baulkham hills", "chatswood", "cronulla",
                           "western sydney", "south west sydney", "north west sydney"], False),
    ("Sydney, NSW", "AU", ["liverpool"], True),
    ("Newcastle, NSW", "AU", ["newcastle", "lake macquarie", "maitland", "hunter valley"], True),
    ("Central Coast, NSW", "AU", ["central coast", "gosford", "wyong"], False),
    ("Wollongong, NSW", "AU", ["wollongong", "illawarra", "shellharbour"], False),
    ("Blue Mountains, NSW", "AU", ["blue mountains", "katoomba"], False),
    ("Melbourne, VIC", "AU", ["melbourne", "geelong", "dandenong", "frankston", "mornington peninsula"], False),
    ("Brisbane, QLD", "AU", ["brisbane", "ipswich", "logan", "moreton bay", "redcliffe"], False),
    ("Gold Coast, QLD", "AU", ["gold coast", "tweed heads"], False),
    ("Sunshine Coast, QLD", "AU", ["sunshine coast", "maroochydore", "caloundra"], False),
    ("Perth, WA", "AU", ["perth", "fremantle", "joondalup", "mandurah"], True),
    ("Adelaide, SA", "AU", ["adelaide"], False),
    ("Canberra, ACT", "AU", ["canberra", "queanbeyan"], False),
    ("Hobart, TAS", "AU", ["hobart", "launceston"], False),
    ("Darwin, NT", "AU", ["darwin"], False),
    ("London, UK", "UK", ["london"], False),
    ("Manchester, UK", "UK", ["manchester"], False),
    ("Birmingham, UK", "UK", ["birmingham"], False),
    ("Leeds, UK", "UK", ["leeds"], False),
    ("Liverpool, UK", "UK", ["liverpool"], True),
    ("Bristol, UK", "UK", ["bristol"], False),
    ("Sheffield, UK", "UK", ["sheffield"], False),
    ("Glasgow, UK", "UK", ["glasgow"], False),
    ("Edinburgh, UK", "UK", ["edinburgh"], False),
    ("Cardiff, UK", "UK", ["cardiff"], False),
    ("Newcastle, UK", "UK", ["newcastle upon tyne", "newcastle"], True),
    ("Nottingham, UK", "UK", ["nottingham"], False),
    ("Leicester, UK", "UK", ["leicester"], False),
    ("Southampton, UK", "UK", ["southampton"], False),
    ("Brighton, UK", "UK", ["brighton"], False),
    ("Toronto, ON", "CA", ["toronto", "mississauga", "brampton", "scarborough"], False),
    ("Vancouver, BC", "CA", ["vancouver", "surrey", "burnaby", "richmond bc"], False),
    ("Calgary, AB", "CA", ["calgary"], False),
    ("Edmonton, AB", "CA", ["edmonton"], False),
    ("Ottawa, ON", "CA", ["ottawa"], False),
    ("Montreal, QC", "CA", ["montreal", "montréal"], False),
    ("Winnipeg, MB", "CA", ["winnipeg"], False),
    ("Hamilton, ON", "CA", ["hamilton"], True),
    ("Halifax, NS", "CA", ["halifax"], False),
    ("Doha, Qatar", "QA", ["doha", "lusail", "al wakrah", "al rayyan"], False),
]
REGION_PATTERNS = [
    (canon, country, needs, [re.compile(r"\b" + re.escape(a) + r"\b") for a in aliases])
    for canon, country, aliases, needs in REGIONS
]


def country_from_host(host):
    host = host.lower()
    if host.endswith(".com.au") or host.endswith(".au") or host.endswith(".net.au"):
        return "AU"
    if host.endswith(".co.uk") or host.endswith(".uk"):
        return "UK"
    if host.endswith(".ca"):
        return "CA"
    if host.endswith(".qa"):
        return "QA"
    return None


def detect_region(weighted_texts, country):
    """weighted_texts: list of (text, weight). Returns (main_region, other_regions)."""
    counts = Counter()
    for text, weight in weighted_texts:
        t = text.lower()
        if not t:
            continue
        for canon, ctry, needs, pats in REGION_PATTERNS:
            if country and ctry != country:
                continue
            if needs and not country:
                continue
            n = sum(len(p.findall(t)) for p in pats)
            if n:
                counts[canon] += n * weight
    if not counts:
        return "", ""
    ranked = [r for r, _ in counts.most_common()]
    return ranked[0], ", ".join(ranked[1:4])


# =====================================================================
# FETCHING
# =====================================================================
def fetch(url, timeout):
    """Returns (html, final_url, error_reason)."""
    err = "Could not open site"
    for verify in (True, False):
        try:
            r = requests.get(url, headers=HEADERS, timeout=timeout, verify=verify)
            if r.status_code in (401, 403, 429):
                return None, None, f"Blocked by site ({r.status_code})"
            if r.status_code >= 400:
                return None, None, f"Page error ({r.status_code})"
            if "html" not in r.headers.get("Content-Type", "").lower():
                return None, None, "Not a web page"
            return r.text, r.url, ""
        except requests.exceptions.SSLError:
            err = "SSL problem"
            continue
        except requests.exceptions.Timeout:
            return None, None, "Timed out"
        except requests.exceptions.ConnectionError:
            return None, None, "Site not reachable"
        except Exception:
            return None, None, "Could not open site"
    return None, None, err


def url_variants(url):
    p = urlparse(url)
    host = p.netloc
    alt = host[4:] if host.startswith("www.") else "www." + host
    variants = [url, p._replace(netloc=alt).geturl()]
    if p.scheme == "https":
        variants.append(p._replace(scheme="http").geturl())
    return variants


def normalize_url(raw):
    raw = str(raw or "").strip()
    if not raw or raw.lower() in ("nan", "none"):
        return None
    if not raw.startswith(("http://", "https://")):
        raw = "https://" + raw
    return raw


def host_of(url):
    return urlparse(url).netloc.lower().replace("www.", "")


# =====================================================================
# PARSING HELPERS
# =====================================================================
def decode_cf_email(hexstr):
    try:
        key = int(hexstr[:2], 16)
        return "".join(chr(int(hexstr[i:i + 2], 16) ^ key) for i in range(2, len(hexstr), 2))
    except Exception:
        return ""


def clean_email(e):
    e = e.strip().strip(".,;:()<>[]\"'").lower()
    if not EMAIL_RE.fullmatch(e):
        return None
    if any(j in e for j in JUNK_EMAIL_PARTS):
        return None
    return e


def extract_emails(soup):
    found = []
    for a in soup.select('a[href^="mailto:"]'):
        mail = unquote(a["href"][7:].split("?")[0])
        found += [m.strip() for m in mail.split(",")]
    for el in soup.select("[data-cfemail]"):
        found.append(decode_cf_email(el["data-cfemail"]))
    found += EMAIL_RE.findall(soup.get_text(" ", strip=True))
    out = []
    for e in found:
        e = clean_email(e)
        if e and e not in out:
            out.append(e)
    return out


def tidy_name(raw):
    tokens = raw.split()
    while tokens and tokens[0].lower() in NAME_STOPWORDS:
        tokens.pop(0)
    while tokens and tokens[-1].lower() in NAME_STOPWORDS:
        tokens.pop()
    if not (2 <= len(tokens) <= 3):
        return None
    for t in tokens:
        low = t.lower().strip(".")
        if low in NAME_BLACKLIST or low in NAME_STOPWORDS:
            return None
        if not re.fullmatch(TOKEN, t):
            return None
    return " ".join(tokens)


def jsonld_blocks(soup):
    blocks = []
    for s in soup.find_all("script", type="application/ld+json"):
        try:
            blocks.append(json.loads(s.string or ""))
        except Exception:
            continue
    return blocks


def names_from_jsonld(blocks):
    results = []

    def walk(node, parent_key=""):
        if isinstance(node, list):
            for n in node:
                walk(n, parent_key)
        elif isinstance(node, dict):
            t = node.get("@type", "")
            t = " ".join(t) if isinstance(t, list) else str(t)
            if "Person" in t and node.get("name"):
                title = str(node.get("jobTitle", "")).lower()
                if parent_key in ("founder", "founders", "owner") or any(
                        w in title for w in ("owner", "founder", "director", "proprietor")):
                    results.append(str(node["name"]))
            for k, v in node.items():
                walk(v, k.lower())

    for b in blocks:
        walk(b)
    return results


def places_from_jsonld(blocks):
    """Collect address / area-served text from structured data (very reliable when present)."""
    out = []

    def walk(node):
        if isinstance(node, list):
            for n in node:
                walk(n)
        elif isinstance(node, dict):
            for key in ("addressLocality", "addressRegion"):
                if isinstance(node.get(key), str):
                    out.append(node[key])
            area = node.get("areaServed")
            if isinstance(area, str):
                out.append(area)
            elif isinstance(area, (list, dict)):
                walk(area)
            if node.get("@type") in ("City", "State", "AdministrativeArea", "Place") and node.get("name"):
                out.append(str(node["name"]))
            for v in node.values():
                if isinstance(v, (dict, list)):
                    walk(v)

    for b in blocks:
        walk(b)
    return " . ".join(out)


def find_owner_name(soup, blocks):
    for raw in names_from_jsonld(blocks):
        name = tidy_name(raw)
        if name:
            return name
    text = soup.get_text(" ", strip=True)
    for pat in NAME_PATTERNS:
        for m in pat.finditer(text):
            name = tidy_name(m.group(1))
            if name:
                return name
    return None


def find_links(soup, base_url):
    host = host_of(base_url)
    links = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if href.startswith(("mailto:", "tel:", "#", "javascript:")):
            continue
        full = urljoin(base_url, href).split("#")[0]
        p = urlparse(full)
        if p.netloc.lower().replace("www.", "") != host:
            continue
        label = (a.get_text(" ", strip=True) + " " + p.path).lower()
        if any(k in label for k in LINK_KEYWORDS) and full not in links:
            links.append(full)
    if not links:
        links = [urljoin(base_url, p) for p in FALLBACK_PATHS]
    return links


def page_meta_text(soup):
    parts = []
    if soup.title and soup.title.string:
        parts.append(soup.title.string)
    for m in soup.find_all("meta"):
        if (m.get("name") or m.get("property") or "").lower() in (
                "description", "og:title", "og:description", "og:site_name"):
            parts.append(m.get("content", ""))
    for h in soup.find_all(["h1", "h2"])[:8]:
        parts.append(h.get_text(" ", strip=True))
    return " . ".join(parts)


def related_domain(email_domain, site_host):
    return (email_domain == site_host or email_domain.endswith("." + site_host)
            or site_host.endswith("." + email_domain))


def pick_email(emails, site_host):
    """Prefer an address on the business's own domain, then a free-mail address.
    Addresses on other company domains (web developers, agencies) are skipped."""
    own = [e for e in emails if related_domain(e[0].split("@")[1], site_host)]
    free = [e for e in emails if e[0].split("@")[1] in FREE_MAIL]
    skipped = [e[0] for e in emails if e not in own and e not in free]
    best = (own or free or [(None, None)])[0]
    return best[0], best[1], skipped


# =====================================================================
# ONE BUSINESS
# =====================================================================
def process_site(name_hint, raw_url, category_hint, address_hint, timeout, max_pages):
    out = {"Owner Full Name": "", "Owner First Name": "", "Owner Last Name": "",
           "Email": "", "Business Type": "", "Working Region": "", "Other Regions": "",
           "Name Source Page": "", "Email Source Page": "", "Skipped Emails": "",
           "Status": ""}

    name_text = f"{name_hint} {category_hint}"
    url = normalize_url(raw_url)
    host = host_of(url) if url else ""
    country = country_from_host(host) if host else None

    meta_text, body_text, place_text = "", "", ""
    emails, pages = [], []

    if not url:
        out["Status"] = "No website"
    elif any(host == h or host.endswith("." + h) for h in SOCIAL_HOSTS):
        out["Status"] = "Social/listing page, not a business website"
    else:
        html, final_url, first_err = None, None, ""
        for variant in url_variants(url):
            html, final_url, err = fetch(variant, timeout)
            if html:
                break
            first_err = first_err or err
        if not html:
            out["Status"] = f"Could not open site ({first_err})"
        else:
            host = host_of(final_url)
            country = country_from_host(host) or country
            pages = [(final_url, BeautifulSoup(html, "html.parser"), html)]
            seen = {final_url}
            for link in find_links(pages[0][1], final_url):
                if len(pages) > max_pages:
                    break
                if link in seen:
                    continue
                seen.add(link)
                h, u, _ = fetch(link, timeout)
                if h:
                    pages.append((u, BeautifulSoup(h, "html.parser"), h))

            for page_url, soup, _raw in pages:
                blocks = jsonld_blocks(soup)
                place_text += " . " + places_from_jsonld(blocks)
                meta_text += " . " + page_meta_text(soup)
                if len(body_text) < 12000:
                    body_text += " " + soup.get_text(" ", strip=True)[:6000]
                for e in extract_emails(soup):
                    if e not in [x[0] for x in emails]:
                        emails.append((e, page_url))
                if not out["Owner Full Name"]:
                    nm = find_owner_name(soup, blocks)
                    if nm:
                        out["Owner Full Name"] = nm
                        out["Owner First Name"] = nm.split()[0]
                        out["Owner Last Name"] = nm.split()[-1]
                        out["Name Source Page"] = page_url

            if not emails:  # fallback: emails hidden in the page code
                for page_url, _s, raw in pages:
                    for e in EMAIL_RE.findall(unquote(raw)):
                        e = clean_email(e)
                        if e and e not in [x[0] for x in emails]:
                            emails.append((e, page_url))

            best, src, skipped = pick_email(emails, host)
            out["Email"] = best or ""
            out["Email Source Page"] = src or ""
            out["Skipped Emails"] = "; ".join(skipped)
            if best:
                out["Status"] = "OK"
            elif out["Owner Full Name"]:
                out["Status"] = "Name only (no usable email on site)"
            elif skipped:
                out["Status"] = "Only third-party emails found (skipped)"
            else:
                out["Status"] = "Nothing found"

    out["Business Type"] = classify_type(name_text, meta_text, body_text)
    region, others = detect_region(
        [(address_hint, 20), (place_text, 10), (meta_text, 3), (body_text, 1)], country)
    out["Working Region"], out["Other Regions"] = region, others
    return out


# =====================================================================
# SHEET HANDLING
# =====================================================================
def read_any(file):
    name = file.name.lower()
    if name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(file, dtype=str)
    else:
        try:
            df = pd.read_csv(file, dtype=str, encoding="utf-8-sig")
        except UnicodeDecodeError:
            file.seek(0)
            df = pd.read_csv(file, dtype=str, encoding="latin-1")
    df = df.fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def guess_col(cols, keywords, exclude=()):
    for kw in keywords:
        for c in cols:
            low = c.lower()
            if kw in low and not any(x in low for x in exclude):
                return c
    return None


def dedupe_key(name, raw_url):
    url = normalize_url(raw_url)
    if url:
        host = host_of(url)
        if any(host == h or host.endswith("." + h) for h in SOCIAL_HOSTS):
            return "url:" + url.lower().rstrip("/")
        return "site:" + host
    n = re.sub(r"[^a-z0-9]", "", str(name).lower())
    return "name:" + n if n else None


def safe_cell(v):
    v = "" if v is None else str(v)
    return "'" + v if v[:1] in ("=", "@") else v


# =====================================================================
# APP
# =====================================================================
st.set_page_config(page_title="Lead Enrichment Tool", page_icon="📇", layout="wide")
st.title("📇 Lead Enrichment Tool")
st.caption("Upload your scraped sheets. The tool finds emails and owner names, works out the "
           "business type and region, removes repeats, and gives you a ready-to-use CSV.")

with st.sidebar:
    st.header("Settings")
    workers = st.slider("Sites processed at once", 2, 16, 8)
    max_pages = st.slider("Extra pages checked per site", 1, 6, 4,
                          help="Contact, About and Team pages.")
    timeout = st.slider("Seconds to wait per page", 5, 30, 12)
    st.info("Tip: run 100 to 300 businesses at a time and keep this tab open until it finishes.")

files = st.file_uploader("Upload your CSV or Excel sheets (you can add several at once)",
                         type=["csv", "xlsx", "xls"], accept_multiple_files=True)

if files:
    frames, read_errors = [], []
    for f in files:
        try:
            frames.append(read_any(f))
        except Exception as e:
            read_errors.append(f"{f.name}: {e}")
    for msg in read_errors:
        st.error(f"Could not read {msg}")

    if frames:
        raw = pd.concat(frames, ignore_index=True).fillna("")
        cols = list(raw.columns)
        st.success(f"Loaded {len(raw)} rows from {len(frames)} file(s).")

        st.subheader("1. Check the columns")
        st.caption("The tool guesses these. Change any that are wrong.")
        none_opt = "(none)"
        opts = [none_opt] + cols

        def idx(guess):
            return opts.index(guess) if guess in opts else 0

        g_name = guess_col(cols, ["business name", "company", "business", "title", "name"],
                           exclude=("full", "first", "last"))
        g_url = guess_col(cols, ["website url", "webiste", "website", "url", "web", "site"])
        g_cat = guess_col(cols, ["category", "industry", "type"])
        g_addr = guess_col(cols, ["address", "location", "suburb", "city", "region"])

        c1, c2, c3, c4 = st.columns(4)
        col_name = c1.selectbox("Business name", opts, index=idx(g_name))
        col_url = c2.selectbox("Website URL", opts, index=idx(g_url))
        col_cat = c3.selectbox("Category (optional)", opts, index=idx(g_cat))
        col_addr = c4.selectbox("Address / location (optional)", opts, index=idx(g_addr))

        if col_name == none_opt or col_url == none_opt:
            st.warning("Please choose the Business name and Website URL columns.")
        else:
            work = pd.DataFrame({
                "Business Name": raw[col_name].astype(str).str.strip(),
                "Website URL": raw[col_url].astype(str).str.strip(),
                "_cat": raw[col_cat].astype(str) if col_cat != none_opt else "",
                "_addr": raw[col_addr].astype(str) if col_addr != none_opt else "",
            })
            work = work[(work["Business Name"] != "") | (work["Website URL"] != "")]
            keys = [dedupe_key(n, u) for n, u in zip(work["Business Name"], work["Website URL"])]
            work["_key"] = keys
            before = len(work)
            work = work[work["_key"].isna() | ~work["_key"].duplicated(keep="first")]
            removed_pre = before - len(work)
            work = work.reset_index(drop=True)

            st.write(f"**{len(work)} unique businesses** to process "
                     f"({removed_pre} repeated rows removed).")

            st.subheader("2. Process")
            if st.button("▶ Start processing", type="primary"):
                total = len(work)
                results = {}
                progress = st.progress(0.0, text="Starting...")
                with ThreadPoolExecutor(max_workers=workers) as ex:
                    futs = {
                        ex.submit(process_site, r["Business Name"], r["Website URL"],
                                  r["_cat"], r["_addr"], timeout, max_pages): i
                        for i, r in work.iterrows()
                    }
                    for n, fut in enumerate(as_completed(futs), start=1):
                        i = futs[fut]
                        try:
                            results[i] = fut.result()
                        except Exception as e:  # never let one bad site stop the run
                            results[i] = {"Status": f"Error: {e}"}
                        progress.progress(n / total, text=f"{n} of {total} businesses done")
                progress.empty()

                out_cols = ["Owner Full Name", "Owner First Name", "Owner Last Name", "Email",
                            "Business Type", "Working Region", "Other Regions",
                            "Name Source Page", "Email Source Page", "Skipped Emails", "Status"]
                for c in out_cols:
                    work[c] = [results.get(i, {}).get(c, "") for i in range(len(work))]

                before2 = len(work)
                has_email = work["Email"] != ""
                dup_email = has_email & work["Email"].duplicated(keep="first")
                work = work[~dup_email].reset_index(drop=True)
                removed_post = before2 - len(work)

                st.session_state["results"] = work.drop(columns=["_cat", "_addr", "_key"])
                st.session_state["stats"] = {
                    "uploaded": before, "removed_pre": removed_pre, "removed_post": removed_post,
                }
                st.rerun()

if "results" in st.session_state:
    df = st.session_state["results"]
    stats = st.session_state["stats"]

    st.divider()
    st.subheader("3. Your results")

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Rows uploaded", stats["uploaded"])
    m2.metric("Repeats removed", stats["removed_pre"] + stats["removed_post"])
    m3.metric("Businesses kept", len(df))
    m4.metric("With email", int((df["Email"] != "").sum()))
    m5.metric("With owner name", int((df["Owner Full Name"] != "").sum()))

    f1, f2, f3 = st.columns([2, 2, 1])
    types = sorted(df["Business Type"].unique())
    regions = sorted([r for r in df["Working Region"].unique() if r])
    pick_types = f1.multiselect("Business type", types, default=types)
    pick_regions = f2.multiselect("Working region (empty = all)", regions)
    only_email = f3.checkbox("Only with email")
    show_details = st.checkbox("Show source pages and status (to verify results)")

    view = df[df["Business Type"].isin(pick_types)]
    if pick_regions:
        view = view[view["Working Region"].isin(pick_regions)]
    if only_email:
        view = view[view["Email"] != ""]

    main_cols = ["Business Name", "Business Type", "Working Region", "Owner Full Name",
                 "Owner First Name", "Owner Last Name", "Email", "Website URL"]
    detail_cols = ["Other Regions", "Name Source Page", "Email Source Page",
                   "Skipped Emails", "Status"]
    show_cols = main_cols + (detail_cols if show_details else [])
    view = view[show_cols].reset_index(drop=True)

    st.dataframe(view, hide_index=True,
                 column_config={"Website URL": st.column_config.LinkColumn("Website URL")})
    st.caption(f"Showing {len(view)} of {len(df)} businesses.")

    csv_df = view.apply(lambda col: col.map(safe_cell))
    st.download_button("⬇ Download CSV", data=csv_df.to_csv(index=False).encode("utf-8-sig"),
                       file_name="leads_ready.csv", mime="text/csv", type="primary")
    st.caption("Before sending, verify the emails (MillionVerifier or ZeroBounce) and "
               "spot-check names against the source pages.")

    if st.button("Clear results and start over"):
        del st.session_state["results"]
        del st.session_state["stats"]
        st.rerun()
