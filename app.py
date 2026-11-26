"""
Lead Enrichment Tool (premium edition)
Upload scraped lead sheets -> finds emails + owner names, business type, working region,
scores each homepage for SEO, removes repeats and leads with no website/email,
and exports to CSV, Excel, Google Sheets, PDF, JSON or copy-paste.
"""

import io
import json
import re
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from urllib.parse import unquote, urljoin, urlparse
from xml.sax.saxutils import escape as xml_escape

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
    """Returns (html, final_url, error_reason, seconds_taken)."""
    err = "Could not open site"
    for verify in (True, False):
        try:
            r = requests.get(url, headers=HEADERS, timeout=timeout, verify=verify)
            took = r.elapsed.total_seconds() if getattr(r, "elapsed", None) else 0.0
            if r.status_code in (401, 403, 429):
                return None, None, f"Blocked by site ({r.status_code})", took
            if r.status_code >= 400:
                return None, None, f"Page error ({r.status_code})", took
            if "html" not in r.headers.get("Content-Type", "").lower():
                return None, None, "Not a web page", took
            return r.text, r.url, "", took
        except requests.exceptions.SSLError:
            err = "SSL problem"
            continue
        except requests.exceptions.Timeout:
            return None, None, "Timed out", 0.0
        except requests.exceptions.ConnectionError:
            return None, None, "Site not reachable", 0.0
        except Exception:
            return None, None, "Could not open site", 0.0
    return None, None, err, 0.0


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
# SEO SCORE (homepage, 100 points)
# =====================================================================
def seo_audit(soup, final_url, elapsed, region):
    """Returns (score 0-100, list of issues). Based on the homepage only."""
    score, issues = 0, []

    if final_url.startswith("https://"):
        score += 10
    else:
        issues.append("Site is not on HTTPS")

    title = soup.title.string.strip() if soup.title and soup.title.string else ""
    if title:
        if 25 <= len(title) <= 65:
            score += 12
        else:
            score += 6
            issues.append(f"Page title length is off ({len(title)} characters)")
    else:
        issues.append("Missing page title")

    md = soup.find("meta", attrs={"name": re.compile(r"^description$", re.I)})
    desc = (md.get("content") or "").strip() if md else ""
    if desc:
        if 100 <= len(desc) <= 170:
            score += 12
        else:
            score += 6
            issues.append("Meta description length is off")
    else:
        issues.append("Missing meta description")

    h1s = soup.find_all("h1")
    if len(h1s) == 1:
        score += 10
    elif len(h1s) > 1:
        score += 5
        issues.append("More than one H1 heading")
    else:
        issues.append("No H1 heading")

    if soup.find("meta", attrs={"name": re.compile(r"^viewport$", re.I)}):
        score += 10
    else:
        issues.append("Not mobile-optimised (no viewport tag)")

    imgs = soup.find_all("img")
    if imgs:
        ratio = sum(1 for i in imgs if (i.get("alt") or "").strip()) / len(imgs)
        if ratio >= 0.8:
            score += 8
        elif ratio >= 0.5:
            score += 4
            issues.append("Many images have no alt text")
        else:
            issues.append("Most images have no alt text")
    else:
        score += 8

    if soup.find("link", rel="canonical"):
        score += 4
    else:
        issues.append("No canonical tag")

    if soup.find("script", type="application/ld+json"):
        score += 8
    else:
        issues.append("No structured data (schema)")

    if soup.find("meta", attrs={"property": "og:title"}):
        score += 4
    else:
        issues.append("No social sharing tags")

    words = len(" ".join(t.get_text(" ", strip=True)
                         for t in soup.find_all(["p", "h1", "h2", "h3", "h4", "li"])).split())
    if words >= 300:
        score += 8
    elif words >= 150:
        score += 4
        issues.append("Thin homepage content")
    else:
        issues.append("Very little text on the homepage")

    host = host_of(final_url)
    internal = 0
    for a in soup.find_all("a", href=True):
        h = a["href"].strip()
        if h.startswith("/") or (host and host in h):
            internal += 1
    if internal >= 5:
        score += 4
    else:
        issues.append("Few internal links")

    if elapsed and elapsed < 2.5:
        score += 6
    elif elapsed and elapsed < 5:
        score += 3
    elif elapsed:
        issues.append("Slow server response")
    else:
        score += 3

    city = region.split(",")[0].strip().lower() if region else ""
    head_text = (title + " " + " ".join(h.get_text(" ", strip=True) for h in h1s)).lower()
    if city and city in head_text:
        score += 4
    else:
        issues.append("Location not in title or main heading")

    return score, issues


def seo_rating(score):
    if score is None or score == "":
        return ""
    return "Strong" if score >= 75 else "Average" if score >= 50 else "Weak"


def pagespeed(url, key):
    """Optional Google PageSpeed (Lighthouse). Returns (performance, seo) or (None, None)."""
    params = [("url", url), ("strategy", "mobile"), ("category", "performance"), ("category", "seo")]
    if key:
        params.append(("key", key))
    try:
        r = requests.get("https://www.googleapis.com/pagespeedonline/v5/runPagespeed",
                         params=params, timeout=70)
        if r.status_code != 200:
            return None, None
        cats = r.json()["lighthouseResult"]["categories"]
        return round(cats["performance"]["score"] * 100), round(cats["seo"]["score"] * 100)
    except Exception:
        return None, None


# =====================================================================
# EMAIL DOMAIN CHECK
# =====================================================================
_MAIL_CACHE = {}


def mail_domain_status(domain):
    """Checks that the email's domain can receive mail (it cannot confirm the mailbox itself)."""
    if domain in _MAIL_CACHE:
        return _MAIL_CACHE[domain]
    if domain in FREE_MAIL:
        return "Mail domain OK"
    status = "Unchecked"
    try:
        import dns.resolver
        try:
            dns.resolver.resolve(domain, "MX", lifetime=6)
            status = "Mail domain OK"
        except dns.resolver.NoAnswer:
            try:
                dns.resolver.resolve(domain, "A", lifetime=6)
                status = "Mail domain OK"
            except Exception:
                status = "No mail server found"
        except dns.resolver.NXDOMAIN:
            status = "Domain not found"
    except Exception:
        status = "Unchecked"
    _MAIL_CACHE[domain] = status
    return status


# =====================================================================
# ONE BUSINESS
# =====================================================================
def process_site(name_hint, raw_url, category_hint, address_hint, timeout, max_pages,
                 psi_on=False, psi_key=""):
    out = {"Owner Full Name": "", "Owner First Name": "", "Owner Last Name": "",
           "Email": "", "Email Check": "", "Business Type": "", "Working Region": "",
           "Other Regions": "", "SEO Score": None, "SEO Rating": "", "SEO Issues": "",
           "PageSpeed Performance": None, "PageSpeed SEO": None,
           "Name Source Page": "", "Email Source Page": "", "Skipped Emails": "",
           "Status": ""}

    name_text = f"{name_hint} {category_hint}"
    url = normalize_url(raw_url)
    host = host_of(url) if url else ""
    country = country_from_host(host) if host else None

    meta_text, body_text, place_text = "", "", ""
    emails, pages = [], []
    home = None  # (soup, final_url, elapsed)

    if not url:
        out["Status"] = "No website"
    elif any(host == h or host.endswith("." + h) for h in SOCIAL_HOSTS):
        out["Status"] = "Social/listing page, not a business website"
    else:
        html, final_url, first_err, took = None, None, "", 0.0
        for variant in url_variants(url):
            html, final_url, err, took = fetch(variant, timeout)
            if html:
                break
            first_err = first_err or err
        if not html:
            out["Status"] = f"Could not open site ({first_err})"
        else:
            host = host_of(final_url)
            country = country_from_host(host) or country
            home_soup = BeautifulSoup(html, "html.parser")
            home = (home_soup, final_url, took)
            pages = [(final_url, home_soup, html)]
            seen = {final_url}
            for link in find_links(home_soup, final_url):
                if len(pages) > max_pages:
                    break
                if link in seen:
                    continue
                seen.add(link)
                h, u, _, _ = fetch(link, timeout)
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
                out["Email Check"] = mail_domain_status(best.split("@")[1])
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

    if home:
        score, issues = seo_audit(home[0], home[1], home[2], region)
        out["SEO Score"] = score
        out["SEO Rating"] = seo_rating(score)
        out["SEO Issues"] = "; ".join(issues[:4])
        if psi_on and out["Email"]:
            out["PageSpeed Performance"], out["PageSpeed SEO"] = pagespeed(home[1], psi_key)
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
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    v = str(v)
    return "'" + v if v[:1] in ("=", "@") else v


# =====================================================================
# EXPORTS
# =====================================================================
NAVY, CREAM, GOLD = "#0B1F3A", "#F7F2E8", "#B8924A"


def to_xlsx(df):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Leads")
        ws = writer.sheets["Leads"]
        for cell in ws[1]:
            cell.font = Font(bold=True, color="F7F2E8")
            cell.fill = PatternFill("solid", fgColor="0B1F3A")
            cell.alignment = Alignment(vertical="center")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for i, col in enumerate(df.columns, start=1):
            longest = max([len(str(col))] + [len(str(v)) for v in df[col].head(300)])
            ws.column_dimensions[get_column_letter(i)].width = min(max(12, longest + 2), 55)
    return buf.getvalue()


def to_pdf(df):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    wanted = [("Business Name", 150), ("Business Type", 90), ("Working Region", 90),
              ("Owner Full Name", 100), ("Email", 150), ("Website URL", 150), ("SEO Score", 40)]
    cols = [(c, w) for c, w in wanted if c in df.columns]
    navy, cream, gold = colors.HexColor(NAVY), colors.HexColor(CREAM), colors.HexColor(GOLD)

    cell = ParagraphStyle("cell", fontName="Helvetica", fontSize=7.5, leading=9.5, textColor=navy)
    head = ParagraphStyle("head", fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=cream)
    title = ParagraphStyle("title", fontName="Helvetica-Bold", fontSize=18, leading=22, textColor=navy)
    sub = ParagraphStyle("sub", fontName="Helvetica", fontSize=9, leading=12, textColor=gold)

    data = [[Paragraph(xml_escape(c), head) for c, _ in cols]]
    for _, row in df.iterrows():
        data.append([Paragraph(xml_escape(safe_cell(row[c])), cell) for c, _ in cols])

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=24, rightMargin=24,
                            topMargin=28, bottomMargin=28, title="Lead list")
    table = Table(data, colWidths=[w for _, w in cols], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), navy),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F3ECDD")]),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9CFBA")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    doc.build([
        Paragraph("Lead list", title),
        Paragraph(f"{len(df)} businesses · prepared {datetime.now():%d %B %Y}", sub),
        Spacer(1, 12), table,
    ])
    return buf.getvalue()


def sheets_configured():
    try:
        return "gcp_service_account" in st.secrets
    except Exception:
        return False


def push_to_sheets(df, sheet_url):
    import gspread
    from google.oauth2.service_account import Credentials
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"])
    gc = gspread.authorize(creds)
    sh = gc.open_by_url(sheet_url)
    title = f"Leads {datetime.now():%d %b %H:%M}"
    rows = [df.columns.tolist()] + [[safe_cell(v) for v in r] for r in df.itertuples(index=False)]
    ws = sh.add_worksheet(title=title, rows=len(rows) + 5, cols=len(df.columns) + 1)
    ws.update(range_name="A1", values=rows, value_input_option="RAW")
    return f"{sh.url}#gid={ws.id}", title


# =====================================================================
# DESIGN
# =====================================================================
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Inter:wght@400;500;600&display=swap');

:root { --navy:#0B1F3A; --navy2:#14304F; --cream:#F7F2E8; --cream2:#EFE7D6;
        --gold:#B8924A; --line:#D9CFBA; --muted:#5B6B80; }

html, body, .stApp, [class*="st-"] { font-family: 'Inter', -apple-system, sans-serif; }
.stApp { background: var(--cream); color: var(--navy); }
.block-container { max-width: 1180px; padding-top: 1.6rem; padding-bottom: 4rem; }
header[data-testid="stHeader"] { background: transparent; }
footer { visibility: hidden; }

h1, h2, h3, h4 { font-family: 'Fraunces', Georgia, serif !important; color: var(--navy);
                 letter-spacing: -0.01em; font-weight: 500; }

.hero { background: linear-gradient(135deg, #0B1F3A 0%, #14304F 100%);
        border-radius: 22px; padding: 46px 52px; margin-bottom: 34px;
        box-shadow: 0 18px 40px rgba(11,31,58,.18); position: relative; overflow: hidden; }
.hero:after { content:""; position:absolute; right:-80px; top:-80px; width:260px; height:260px;
              border-radius:50%; border:1px solid rgba(184,146,74,.45); }
.hero .eyebrow { color: var(--gold); font-size: 12px; letter-spacing: .22em; font-weight: 600;
                 text-transform: uppercase; margin-bottom: 14px; }
.hero h1 { font-family:'Fraunces',serif !important; color: var(--cream) !important; font-size: 46px;
           line-height: 1.08; margin: 0 0 14px 0; font-weight: 500; }
.hero p { color: #D9D2C0; font-size: 16px; max-width: 640px; line-height: 1.6; margin: 0; }

.sec { display:flex; align-items:baseline; gap:16px; margin: 42px 0 14px 0;
       border-bottom: 1px solid var(--line); padding-bottom: 12px; }
.sec .num { font-family:'Fraunces',serif; color: var(--gold); font-size: 15px; letter-spacing: .12em; }
.sec .ttl { font-family:'Fraunces',serif; font-size: 27px; color: var(--navy); font-weight: 500; }
.sec .sub { color: var(--muted); font-size: 14px; margin-left: auto; }

[data-testid="stFileUploader"] section { background: #FBF8F1; border: 1.5px dashed var(--gold);
        border-radius: 16px; padding: 26px; }
[data-testid="stFileUploader"] section:hover { background: #fff; }

.stButton > button, .stDownloadButton > button, .stLinkButton > a {
        background: var(--navy) !important; color: var(--cream) !important; border: 1px solid var(--navy) !important;
        border-radius: 12px !important; padding: .72rem 1.5rem !important; font-weight: 600 !important;
        letter-spacing: .01em; transition: all .18s ease; box-shadow: 0 6px 16px rgba(11,31,58,.16); }
.stButton > button:hover, .stDownloadButton > button:hover, .stLinkButton > a:hover {
        background: var(--gold) !important; border-color: var(--gold) !important; color: var(--navy) !important;
        transform: translateY(-1px); }
.stButton > button p, .stDownloadButton > button p, .stLinkButton > a p { color: inherit !important; }

[data-testid="stMetric"] { background: #FBF8F1; border: 1px solid var(--line); border-radius: 16px;
        padding: 18px 20px; }
[data-testid="stMetricLabel"] p { color: var(--muted); font-size: 12px; letter-spacing: .08em;
        text-transform: uppercase; font-weight: 600; }
[data-testid="stMetricValue"] { font-family: 'Fraunces', serif; color: var(--navy); font-weight: 500; }

[data-testid="stExpander"] { background: #FBF8F1; border: 1px solid var(--line); border-radius: 14px; }
[data-baseweb="select"] > div, .stTextInput input { background: #FBF8F1 !important;
        border-color: var(--line) !important; border-radius: 10px !important; }
[data-testid="stDataFrame"] { border: 1px solid var(--line); border-radius: 14px; overflow: hidden; }
.stAlert { border-radius: 12px; }
.note { color: var(--muted); font-size: 13px; line-height: 1.6; }
</style>
"""


def section(num, title, sub=""):
    st.markdown(
        f'<div class="sec"><span class="num">{num}</span><span class="ttl">{title}</span>'
        f'<span class="sub">{sub}</span></div>', unsafe_allow_html=True)


# =====================================================================
# APP
# =====================================================================
st.set_page_config(page_title="Lead Intelligence", page_icon="📇", layout="wide")
st.markdown(CSS, unsafe_allow_html=True)
st.markdown(
    '<div class="hero"><div class="eyebrow">Lead Intelligence</div>'
    '<h1>Turn a scraped list into<br>ready-to-contact leads.</h1>'
    '<p>Upload your sheets. We find the email and owner, identify the business type and region, '
    'score each homepage for SEO, remove repeats, and hand you a clean list in the format you need.</p>'
    '</div>', unsafe_allow_html=True)

section("01", "Upload", "CSV or Excel · several files at once")
files = st.file_uploader("Upload your sheets", type=["csv", "xlsx", "xls"],
                         accept_multiple_files=True, label_visibility="collapsed")

with st.expander("Advanced settings"):
    a1, a2, a3 = st.columns(3)
    workers = a1.slider("Sites processed at once", 2, 16, 8)
    max_pages = a2.slider("Extra pages checked per site", 1, 6, 4,
                          help="Contact, About and Team pages.")
    timeout = a3.slider("Seconds to wait per page", 5, 30, 12)
    drop_unreachable = st.checkbox(
        "Remove leads with no website or no email", value=True,
        help="Leads are never removed for a good SEO score or a good-looking site.")
    psi_on = st.checkbox(
        "Also run Google PageSpeed on each lead (slower)", value=False,
        help="Adds Google's mobile performance and SEO scores. Roughly 10 extra minutes per 200 leads.")
    psi_key = st.text_input("Google PageSpeed API key (optional)", type="password",
                            help="Free from Google Cloud. Without a key, Google limits how many checks you can run.")

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

        section("02", "Check the columns", f"{len(raw)} rows loaded from {len(frames)} file(s)")
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
        col_addr = c4.selectbox("Address (optional)", opts, index=idx(g_addr))

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
            work["_key"] = [dedupe_key(n, u) for n, u in zip(work["Business Name"], work["Website URL"])]
            before = len(work)
            work = work[work["_key"].isna() | ~work["_key"].duplicated(keep="first")]
            removed_pre = before - len(work)
            work = work.reset_index(drop=True)

            section("03", "Process", f"{len(work)} unique businesses · {removed_pre} repeats removed")
            if st.button("Start processing"):
                total = len(work)
                results = {}
                progress = st.progress(0.0, text="Starting...")
                with ThreadPoolExecutor(max_workers=workers) as ex:
                    futs = {
                        ex.submit(process_site, r["Business Name"], r["Website URL"], r["_cat"],
                                  r["_addr"], timeout, max_pages, psi_on, psi_key): i
                        for i, r in work.iterrows()
                    }
                    for n, fut in enumerate(as_completed(futs), start=1):
                        i = futs[fut]
                        try:
                            results[i] = fut.result()
                        except Exception as e:  # one bad site never stops the run
                            results[i] = {"Status": f"Error: {e}"}
                        progress.progress(n / total, text=f"{n} of {total} businesses done")
                progress.empty()

                out_cols = ["Owner Full Name", "Owner First Name", "Owner Last Name", "Email",
                            "Email Check", "Business Type", "Working Region", "Other Regions",
                            "SEO Score", "SEO Rating", "SEO Issues", "PageSpeed Performance",
                            "PageSpeed SEO", "Name Source Page", "Email Source Page",
                            "Skipped Emails", "Status"]
                for c in out_cols:
                    work[c] = [results.get(i, {}).get(c, None) for i in range(len(work))]
                for c in out_cols:
                    if c not in ("SEO Score", "PageSpeed Performance", "PageSpeed SEO"):
                        work[c] = work[c].fillna("")
                for c in ("SEO Score", "PageSpeed Performance", "PageSpeed SEO"):
                    work[c] = pd.to_numeric(work[c], errors="coerce")

                before2 = len(work)
                dup_email = (work["Email"] != "") & work["Email"].duplicated(keep="first")
                work = work[~dup_email].reset_index(drop=True)
                removed_post = before2 - len(work)

                has_site = work["Website URL"].str.strip() != ""
                has_email = work["Email"] != ""
                if drop_unreachable:
                    keep = has_site & has_email
                else:
                    keep = pd.Series(True, index=work.index)
                kept = work[keep].drop(columns=["_cat", "_addr", "_key"]).reset_index(drop=True)
                dropped = work[~keep].drop(columns=["_cat", "_addr", "_key"]).reset_index(drop=True)

                st.session_state["results"] = kept
                st.session_state["removed"] = dropped
                st.session_state["stats"] = {
                    "uploaded": before, "dupes": removed_pre + removed_post, "noemail": len(dropped),
                }
                st.rerun()

if "results" in st.session_state:
    df = st.session_state["results"]
    dropped = st.session_state["removed"]
    stats = st.session_state["stats"]

    section("04", "Your leads", "Filter, review and export")

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Rows uploaded", stats["uploaded"])
    m2.metric("Repeats removed", stats["dupes"])
    m3.metric("No email / website", stats["noemail"])
    m4.metric("Leads ready", len(df))
    m5.metric("With owner name", int((df["Owner Full Name"] != "").sum()))
    st.write("")

    f1, f2, f3 = st.columns(3)
    types = sorted(df["Business Type"].unique())
    regions = sorted([r for r in df["Working Region"].unique() if r])
    ratings = [r for r in ["Weak", "Average", "Strong"] if r in set(df["SEO Rating"])]
    pick_types = f1.multiselect("Business type", types, default=types)
    pick_regions = f2.multiselect("Working region (empty = all)", regions)
    pick_ratings = f3.multiselect("SEO rating (empty = all)", ratings)
    show_details = st.checkbox("Show source pages and checks (to verify results)")

    view = df[df["Business Type"].isin(pick_types)]
    if pick_regions:
        view = view[view["Working Region"].isin(pick_regions)]
    if pick_ratings:
        view = view[view["SEO Rating"].isin(pick_ratings)]

    main_cols = ["Business Name", "Business Type", "Working Region", "Owner Full Name",
                 "Owner First Name", "Owner Last Name", "Email", "Website URL",
                 "SEO Score", "SEO Rating", "SEO Issues"]
    detail_cols = ["Email Check", "Other Regions", "Name Source Page", "Email Source Page",
                   "Skipped Emails", "Status"]
    if df["PageSpeed Performance"].notna().any():
        detail_cols = ["PageSpeed Performance", "PageSpeed SEO"] + detail_cols
    show_cols = main_cols + (detail_cols if show_details else [])
    view = view[show_cols].reset_index(drop=True)

    st.dataframe(
        view, hide_index=True,
        column_config={
            "Website URL": st.column_config.LinkColumn(
                "Website URL", display_text=r"https?://(?:www\.)?([^/]+)"),
            "SEO Score": st.column_config.ProgressColumn(
                "SEO Score", min_value=0, max_value=100, format="%d"),
        })
    st.markdown(
        f'<div class="note">Showing {len(view)} of {len(df)} leads. SEO score is a 100-point check of '
        'each homepage (title, description, headings, mobile tag, schema, speed and more). '
        'It is a guide, not a ranking. No lead is removed for a good score.</div>',
        unsafe_allow_html=True)

    if len(dropped):
        st.write("")
        with st.expander(f"Removed leads ({len(dropped)}): no usable email or website"):
            blocked = int(dropped["Status"].str.contains("Blocked").sum())
            if blocked:
                st.info(f"{blocked} sites blocked automatic access. They may still have an email, "
                        "so open those by hand if you want them.")
            st.dataframe(dropped[["Business Name", "Website URL", "Status"]], hide_index=True)

    section("05", "Export", "Choose a format")
    fmt = st.selectbox("Export format",
                       ["CSV", "Excel (.xlsx)", "Google Sheets", "PDF", "JSON",
                        "Copy & paste (tab-separated)"], label_visibility="collapsed")
    stamp = f"{datetime.now():%Y-%m-%d}"
    export_df = view.copy()
    if fmt == "CSV":
        csv_df = export_df.apply(lambda col: col.map(safe_cell))
        st.download_button("Download CSV", csv_df.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"leads_{stamp}.csv", mime="text/csv")
    elif fmt.startswith("Excel"):
        st.download_button("Download Excel file", to_xlsx(export_df.apply(lambda c: c.map(safe_cell))),
                           file_name=f"leads_{stamp}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    elif fmt == "PDF":
        st.download_button("Download PDF", to_pdf(export_df), file_name=f"leads_{stamp}.pdf",
                           mime="application/pdf")
        st.markdown('<div class="note">The PDF shows the main columns only. Use CSV or Excel for the full data.</div>',
                    unsafe_allow_html=True)
    elif fmt == "JSON":
        st.download_button("Download JSON",
                           export_df.where(export_df.notna(), None).to_json(orient="records", indent=2, force_ascii=False),
                           file_name=f"leads_{stamp}.json", mime="application/json")
    elif fmt == "Google Sheets":
        if sheets_configured():
            sheet_url = st.text_input("Paste the link of your Google Sheet",
                                      placeholder="https://docs.google.com/spreadsheets/d/...")
            if st.button("Send to Google Sheets"):
                if not sheet_url.strip():
                    st.warning("Paste your Google Sheet link first.")
                else:
                    try:
                        link, tab = push_to_sheets(export_df, sheet_url.strip())
                        st.success(f"Sent. A new tab called “{tab}” was added.")
                        st.markdown(f"[Open the sheet]({link})")
                    except Exception as e:
                        st.error(f"Could not send to Google Sheets: {e}")
            st.markdown('<div class="note">Or use the manual route below.</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="note">Manual route: 1) download the file, 2) open a blank Google Sheet, '
            '3) choose File, Import, Upload and select the file.</div>', unsafe_allow_html=True)
        g1, g2 = st.columns([1, 1])
        with g1:
            st.download_button("Download file for Google Sheets",
                               to_xlsx(export_df.apply(lambda c: c.map(safe_cell))),
                               file_name=f"leads_{stamp}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with g2:
            st.link_button("Open a blank Google Sheet", "https://sheets.new")
    else:
        tsv = export_df.apply(lambda col: col.map(safe_cell)).to_csv(sep="\t", index=False)
        st.code(tsv, language=None)
        st.markdown('<div class="note">Use the copy icon at the top right of the box, then paste into '
                    'Google Sheets or Excel.</div>', unsafe_allow_html=True)

    st.write("")
    if st.button("Clear results and start over"):
        for k in ("results", "removed", "stats"):
            st.session_state.pop(k, None)
        st.rerun()
