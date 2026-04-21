#!/usr/bin/env python3
"""Official homepage detector (single-file edition).

- API free / Google free
- Default search: DuckDuckGo, fallback Bing -> Yahoo
- Offline self validation: --selftest / --fixture-test / --sample-run
"""
from __future__ import annotations

import argparse
import csv
import html
import json
import logging
import os
import re
import sys
import time
import unicodedata
from dataclasses import dataclass
from enum import Enum
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional
from urllib.error import URLError
from urllib.parse import parse_qs, quote_plus, unquote, urljoin, urlparse
from urllib.request import Request, urlopen

# Optional dependencies (not required for selftest/fixture/sample)
try:
    import pandas as pd  # type: ignore
except Exception:  # noqa: BLE001
    pd = None  # type: ignore

try:
    import requests  # type: ignore
except Exception:  # noqa: BLE001
    requests = None  # type: ignore

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/126.0 Safari/537.36"
TIMEOUT = 20
RETRY = 2
MAX_DEPTH = 3
MAX_PAGES = 40
MAX_SEARCH_RESULTS_PER_QUERY = 10
SAVE_INTERVAL = 10
DEFAULT_STATE = "run_state.json"
DEFAULT_OUTPUT = "results.xlsx"
LOG_DIR = Path("logs")

SEARCH_ENGINES = ["duckduckgo", "bing", "yahoo"]
ZERO_LINK_ROTATE_THRESHOLD = 3
SEARCH_FAIL_ROTATE_THRESHOLD = 2

OUTPUT_COLS = ["公式HP", "判定詳細", "スコア", "採用理由"]

ALLOW_HOSTS = {"kensetumap.com"}
BLACKLIST_BASE_DOMAINS = {
    # search/map/portal/db
    "search.yahoo.co.jp",
    "map.yahoo.co.jp",
    "google.com",
    "google.co.jp",
    "24u.jp",
    "navitime.co.jp",
    "travel.navitime.com",
    "jpnumber.com",
    "houjin.goo.to",
    "houjin.info",
    "houjin.jp",
    "kaisharesearch.com",
    "itp.ne.jp",
    "townpage.goo.ne.jp",
    "ekiten.jp",
    # sns/news/ec
    "facebook.com",
    "instagram.com",
    "x.com",
    "twitter.com",
    "youtube.com",
    "rakuten.co.jp",
    "amazon.co.jp",
    # free hosts
    "wixsite.com",
    "jimdo.com",
    "fc2.com",
    "ameblo.jp",
    "amebaownd.com",
}
LOW_TRUST_BASE_DOMAINS = {
    "ullet.com",
    "ivry.jp",
    "keisin.net",
    "suke-dachi.jp",
    "toranet.jp",
    "mynavi.jp",
    "enepi.jp",
    "architects-j.com",
    "keisin.biz",
    "baseconnect.in",
    "hatalike.jp",
    "xn--pckua2a7gp15o89zb.com",
    "kensetumap.com",
    "kitaq-wood.com",
    "chikuhou-ie.com",
    "tsukulink.net",
    "companyinformation.jp",
    "mapfan.com",
    "houzz.jp",
    "biz-maps.com",
    "iejin.com",
    "jia-9.org",
    "nikkei.com",
    "yelp.com",
    "data-link-plus.com",
    "jcarb.com",
    "kentikusi.jp",
    "gasuyanomadoguchi.com",
    "gaiheki-tatsujin.com",
    "leohouselife.com",
    "praise-arc.com",
}

CORP_WORDS = ["株式会社", "（株）", "(株)", "㈱", "有限会社", "（有）", "(有)", "合同会社"]
PROFILE_HINTS = ["会社概要", "会社案内", "事務所概要", "about", "profile", "gaiyo", "contact", "access"]
LOW_VALUE_HINTS = ["blog", "news", "works", "column", "施工事例", "お知らせ", "ニュース"]
FINAL_BAD_PATH_HINTS = ["/contact", "/recruit", "/works", "/news", "/column", "/blog", "/privacy", "/sitemap", "/search/", "/introduce/", "/companies/", "/corp/", "/location/"]
FINAL_STRICT_BAD_HINTS = ["/viewjob", "/jobinfo", "/kaiin.html", "tab.php", "/member", "/inquiry", "/form", "/kaiin/"]
ASSET_EXTS = (
    ".css", ".js", ".ico", ".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp",
    ".woff", ".woff2", ".ttf", ".pdf", ".xml", ".rss", ".zip", ".mp4", ".mp3",
)
EXTERNAL_LINK_BLOCK_HOSTS = {
    "facebook.com", "instagram.com", "x.com", "twitter.com", "youtube.com",
    "line.me", "amazon.co.jp", "rakuten.co.jp",
}

# Embedded fixtures (no external html required)
FIXTURE_SEARCH_DDG = """
<html><body>
<a class="result__a" href="/l/?kh=-1&uddg=https%3A%2F%2Fexample.co.jp%2Fcompany">公式</a>
<a class="result__a" href="/l/?kh=-1&uddg=https%3A%2F%2Fatelier-kasho.com%2Fabout">atelier</a>
<a class="result__a" href="/l/?kh=-1&uddg=https%3A%2F%2Fwww.navitime.co.jp%2Fpoi%3Fspot%3D00011-080000507">navitime</a>
<a class="result__a" href="/l/?kh=-1&uddg=https%3A%2F%2Fwww.jpnumber.com%2Fnumberinfo_092_753_9567.html">jpnumber</a>
<a class="result__a" href="/l/?kh=-1&uddg=https%3A%2F%2Fsearch.yahoo.co.jp%2Fsearch%3Fp%3Dfoo">search</a>
<a class="result__a" href="/l/?kh=-1&uddg=https%3A%2F%2Fhoujin.info%2Fdetail">db</a>
</body></html>
"""
FIXTURE_SEARCH_BING = '<html><body><li class="b_algo"><h2><a href="https://example.co.jp/company">公式</a></h2></li></body></html>'
FIXTURE_SEARCH_YAHOO = '<html><body><section id="web"><a href="https://example.co.jp/company">公式</a></section></body></html>'
FIXTURE_OFFICIAL_PAGE = """
<html><head><title>株式会社テスト設計 | 会社概要</title>
<meta name="description" content="会社概要とアクセス" />
<script type="application/ld+json">{"@context":"https://schema.org","@type":"Organization"}</script>
</head><body>
<header>株式会社テスト設計 一級建築士事務所</header>
<address>福岡県北九州市門司区本町4-14 TEL 093-331-0555</address>
<footer>代表者 山田太郎 登録番号 A-1234</footer>
<a href="/contact">お問い合わせ</a>
</body></html>
"""
FIXTURE_ATELIER_PAGE = """
<html><head><title>atelier Kasho | about</title>
<meta property="og:site_name" content="atelier-kasho"/>
</head><body>
<h1>atelier Kasho</h1>
<footer>福岡市中央区 / contact</footer>
<a href="/about">about</a>
</body></html>
"""


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------
@dataclass
class OfficeRecord:
    office_name: str
    corp_name: str
    address: str
    phone: str
    representative: str
    registration_no: str


@dataclass
class ScoreResult:
    total: int
    detail: Dict[str, int]
    instant_accept: bool
    accepted: bool
    reason: str


@dataclass
class CandidateResult:
    url: str
    score: ScoreResult


class HostClass(str, Enum):
    official_candidate = "official_candidate"
    association_member_page = "association_member_page"
    directory = "directory"
    portal = "portal"
    corporate_db = "corporate_db"
    phonebook = "phonebook"
    job_board = "job_board"
    public_site = "public_site"
    asset = "asset"
    leadgen_site = "leadgen_site"
    comparison_site = "comparison_site"
    review_site = "review_site"
    lp_site = "lp_site"
    media_profile_page = "media_profile_page"


# ---------------------------------------------------------------------------
# Logging / CLI
# ---------------------------------------------------------------------------
def setup_logger() -> logging.Logger:
    LOG_DIR.mkdir(exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = LOG_DIR / f"run_{ts}.log"
    logger = logging.getLogger("official_hp_single")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setFormatter(fmt)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(fh)
    logger.addHandler(sh)
    logger.info("log file: %s", log_path)
    return logger


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="企業公式ホームページ検出AI (single file)")
    p.add_argument("--input", help="入力 xlsx/csv")
    p.add_argument("--output", default=DEFAULT_OUTPUT, help="出力 xlsx/csv")
    p.add_argument("--state", default=DEFAULT_STATE)
    p.add_argument("--resume", action="store_true")
    p.add_argument("--save-interval", type=int, default=SAVE_INTERVAL)
    p.add_argument("--max-rows", type=int, default=0)
    p.add_argument("--check-env", action="store_true")
    p.add_argument("--selftest", action="store_true")
    p.add_argument("--fixture-test", action="store_true")
    p.add_argument("--sample-run", action="store_true")
    return p.parse_args()


# ---------------------------------------------------------------------------
# Utils
# ---------------------------------------------------------------------------
def normalize_text(text: str) -> str:
    text = html.unescape(str(text or ""))
    text = unicodedata.normalize("NFKC", text).replace("\u3000", " ")
    return re.sub(r"\s+", " ", text).strip()


def compact_text(text: str) -> str:
    return re.sub(r"\s+", "", normalize_text(text))


def strip_symbols(text: str) -> str:
    return re.sub(r"[\(\)（）\[\]【】『』「」<>＜＞・,，.。'\"!！?？:/\\\-_]+", " ", normalize_text(text)).strip()


def remove_corp_words(text: str) -> str:
    s = normalize_text(text)
    for w in CORP_WORDS:
        s = s.replace(w, "")
    return strip_symbols(s)


def normalize_phone(phone: str) -> str:
    d = re.sub(r"\D", "", unicodedata.normalize("NFKC", str(phone or "")))
    if len(d) == 10 and not d.startswith("0"):
        d = "0" + d
    return d


def address_variants(addr: str) -> tuple[str, str, str]:
    full = normalize_text(addr)
    no_banchi = re.sub(r"\d+[\-丁目番地号\d]*", "", full).strip(" ,")
    m = re.search(r"(.+?[都道府県])?(.+?[市区郡])(.+?[町村])?", full)
    city = "".join(g for g in m.groups() if g) if m else full
    return full, normalize_text(no_banchi), normalize_text(city)


def tokenize_candidates(*values: str) -> list[str]:
    out: list[str] = []

    def add(v: str) -> None:
        if v and len(v) >= 2 and v not in out:
            out.append(v)

    for v in values:
        base = strip_symbols(v)
        add(base)
        add(remove_corp_words(base))
    return out


def is_weak_query_token(token: str) -> bool:
    t = normalize_text(token)
    if not t:
        return True
    ascii_only = bool(re.fullmatch(r"[A-Za-z0-9\-_ ]+", t))
    compact = re.sub(r"[\s\-_]+", "", t)
    return ascii_only and len(compact) <= 2


def extract_core_name(name: str) -> str:
    n = normalize_text(name)
    suffixes = [
        "一級建築士事務所", "二級建築士事務所", "建築士事務所",
        "設計事務所", "アーキテクト", "architecture", "architect",
    ]
    for s in suffixes:
        n = re.sub(re.escape(s) + r"$", "", n, flags=re.IGNORECASE).strip()
        n = n.replace(s, "").strip()
    return normalize_text(n)


def ascii_variations(name: str) -> list[str]:
    n = normalize_text(name)
    if not re.search(r"[A-Za-z]", n):
        return []
    out = [n]
    out.append(n.replace(" ", "-"))
    out.append(n.replace("-", " "))
    out.append(n.replace(" ", "").replace("-", ""))
    if "&" in n:
        out.extend([n.replace("&", ""), n.replace("&", " and "), n.replace("&", "and")])
    # dedupe
    ret: list[str] = []
    for x in out:
        x = normalize_text(x)
        if x and x not in ret:
            ret.append(x)
    return ret[:8]


def symbol_name_variations(name: str) -> list[str]:
    n = normalize_text(name)
    if not n:
        return []
    out = [n]
    out.extend([
        n.replace("&", " and "),
        n.replace("&", "and"),
        n.replace("&", " "),
        n.replace("&", ""),
        n.replace("＆", "&"),
        n.replace("＆", " and "),
        re.sub(r"[^A-Za-z0-9\u3040-\u30ff\u3400-\u9fff]+", " ", n),
        re.sub(r"[^A-Za-z0-9\u3040-\u30ff\u3400-\u9fff]+", "", n),
    ])
    ret: list[str] = []
    for x in out:
        x = normalize_text(x)
        if x and x not in ret:
            ret.append(x)
    return ret[:12]


def build_queries(office: str, corp: str, addr: str, phone: str, rep: str) -> list[str]:
    office_tokens = [x for x in tokenize_candidates(office) if not is_weak_query_token(x)]
    corp_tokens = [x for x in tokenize_candidates(corp) if not is_weak_query_token(x)]
    rep_tokens = tokenize_candidates(rep)
    full, no_banchi, city = address_variants(addr)
    p = normalize_phone(phone)

    out: list[str] = []

    def add(q: str) -> None:
        q = normalize_text(q)
        if q and q not in out:
            out.append(q)

    for o in office_tokens[:4]:
        add(f"{o} {full}")
        add(f"{o} {p}") if p else None
        for r in rep_tokens[:2]:
            add(f"{o} {r}")
    for c in corp_tokens[:4]:
        add(f"{c} {full}")
        add(f"{c} {p}") if p else None
        for r in rep_tokens[:2]:
            add(f"{c} {r}")
    for o in office_tokens[:3]:
        add(f"{o} {city} 建築士事務所")
    for c in corp_tokens[:3]:
        add(f"{c} {city} 建築士事務所")
    for o in office_tokens[:3]:
        add(o)
    for c in corp_tokens[:3]:
        add(c)

    # personal/small-office friendly queries (ASCII core name)
    for src in [office, corp]:
        core = extract_core_name(src)
        if is_weak_query_token(core):
            continue
        vars_ = ascii_variations(core) + symbol_name_variations(core)
        if not vars_:
            continue
        for v in vars_:
            add(f'"{v}"')
            add(v)
            add(f"{v} {city}")

    return out


def host_slug_matches_name(url: str, office: str, corp: str) -> bool:
    host = re.sub(r"[^a-z0-9]", "", host_of(url).lower())
    if not host:
        return False
    variants: list[str] = []
    for core in [extract_core_name(office), extract_core_name(corp)]:
        variants.extend(ascii_variations(core))
        variants.extend(symbol_name_variations(core))
    for v in variants:
        s = re.sub(r"[^a-z0-9]", "", v.lower())
        if len(s) < 3:
            continue
        if s in host or host in s:
            return True
    return False


def make_search_url(engine: str, query: str) -> str:
    q = quote_plus(query)
    if engine == "duckduckgo":
        return f"https://duckduckgo.com/html/?q={q}"
    if engine == "bing":
        return f"https://www.bing.com/search?q={q}"
    return f"https://search.yahoo.co.jp/search?p={q}"


def parse_result_links(engine: str, html_text: str) -> list[str]:
    """Engine-specific extraction (no broad a[href] crawling).

    - duckduckgo: result__a + uddg restore
    - bing: li.b_algo h2 a
    - yahoo: section#web / algo blocks
    """
    links: list[str] = []

    if engine == "duckduckgo":
        # only result__a
        for href in re.findall(r'<a[^>]*class=["\'][^"\']*result__a[^"\']*["\'][^>]*href=["\']([^"\']+)["\']', html_text, flags=re.IGNORECASE):
            if href.startswith("http"):
                links.append(href)
            elif "uddg=" in href:
                qs = parse_qs(urlparse(href).query)
                if qs.get("uddg"):
                    u = unquote(qs["uddg"][0])
                    if u.startswith("http"):
                        links.append(u)

    elif engine == "bing":
        for href in re.findall(r'<li[^>]*class=["\'][^"\']*b_algo[^"\']*["\'][\s\S]*?<h2[^>]*>\s*<a[^>]*href=["\']([^"\']+)["\']', html_text, flags=re.IGNORECASE):
            if href.startswith("http"):
                links.append(href)

    else:  # yahoo
        # narrow: web result section / algo class
        for href in re.findall(r'<(?:section[^>]*id=["\']web["\'][\s\S]*?|div[^>]*class=["\'][^"\']*algo[^"\']*["\'][\s\S]*?)<a[^>]*href=["\']([^"\']+)["\']', html_text, flags=re.IGNORECASE):
            if href.startswith("http"):
                links.append(href)

    # unique
    seen: set[str] = set()
    out: list[str] = []
    for u in links:
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out


def is_blocked_response(text: str) -> bool:
    t = normalize_text(text).lower()
    return any(w in t for w in ["captcha", "unusual traffic", "アクセスが集中", "blocked"])


# ---------------------------------------------------------------------------
# Page type / blacklist
# ---------------------------------------------------------------------------
def host_of(url: str) -> str:
    return urlparse(url).netloc.lower().split(":")[0]


def host_match(host: str, base: str) -> bool:
    return host == base or host.endswith(f".{base}")


def is_blacklisted_host(url: str) -> bool:
    h = host_of(url)
    if any(host_match(h, a) for a in ALLOW_HOSTS):
        return False
    return any(host_match(h, b) for b in BLACKLIST_BASE_DOMAINS)


def is_low_trust_host(url: str) -> bool:
    h = host_of(url)
    return any(host_match(h, b) for b in LOW_TRUST_BASE_DOMAINS)


def classify_host(url: str, html_text: str, title: str) -> HostClass:
    u = url.lower()
    h = host_of(url)
    low = f"{u} {normalize_text(title).lower()} {normalize_text(strip_tags(html_text)).lower()}"
    if any(u.endswith(ext) for ext in ASSET_EXTS):
        return HostClass.asset
    if any(host_match(h, x) for x in ["yelp.com"]):
        return HostClass.review_site
    if any(host_match(h, x) for x in ["gasuyanomadoguchi.com", "gaiheki-tatsujin.com"]):
        return HostClass.leadgen_site
    if any(host_match(h, x) for x in ["jcarb.com", "kentikusi.jp", "leohouselife.com"]):
        return HostClass.media_profile_page
    if any(host_match(h, x) for x in ["nikkei.com", "data-link-plus.com"]) and any(k in u for k in ["/compass/company/", "/jp/corporation/"]):
        return HostClass.corporate_db
    if "visit_captcha" in u or "口コミ" in low or "レビュー" in low:
        return HostClass.review_site
    if any(k in low for k in ["比較", "見積", "相談", "窓口", "一括", "ランキング"]):
        return HostClass.comparison_site
    if any(k in u for k in ["/lp/", "/landing", "/campaign"]) or "lp" in title.lower():
        return HostClass.lp_site
    if any(k in u for k in ["/kaiin", "/member"]) or "kitaq-wood.com" in h:
        return HostClass.association_member_page
    if any(k in u for k in ["/viewjob", "/jobinfo", "/recruit"]) or any(host_match(h, x) for x in ["hatalike.jp", "xn--pckua2a7gp15o89zb.com", "mynavi.jp"]):
        return HostClass.job_board
    if any(host_match(h, x) for x in ["baseconnect.in", "enepi.jp", "companyinformation.jp", "biz-maps.com"]) or is_corporate_db_page(url, html_text, title):
        return HostClass.corporate_db
    if is_phone_book_page(url, html_text, title):
        return HostClass.phonebook
    if is_reference_directory_page(url, html_text, title) or any(host_match(h, x) for x in ["architects-j.com", "keisin.biz", "kensetumap.com", "chikuhou-ie.com", "tsukulink.net", "mapfan.com", "houzz.jp", "iejin.com", "jia-9.org"]):
        return HostClass.directory
    if is_portal_page(url, html_text, title):
        return HostClass.portal
    if any(host_match(h, x) for x in ["go.jp", "lg.jp", "or.jp"]):
        return HostClass.public_site
    return HostClass.official_candidate


def is_final_forbidden_class(host_class: HostClass) -> bool:
    return host_class in {
        HostClass.association_member_page,
        HostClass.directory,
        HostClass.portal,
        HostClass.corporate_db,
        HostClass.phonebook,
        HostClass.job_board,
        HostClass.public_site,
        HostClass.asset,
        HostClass.leadgen_site,
        HostClass.comparison_site,
        HostClass.review_site,
        HostClass.lp_site,
        HostClass.media_profile_page,
    }


def extract_title(html_text: str) -> str:
    m = re.search(r"<title[^>]*>(.*?)</title>", html_text, flags=re.IGNORECASE | re.DOTALL)
    return normalize_text(m.group(1) if m else "")


def is_search_result_page(url: str, html_text: str, title: str) -> bool:
    h = host_of(url)
    u = url.lower()
    t = title.lower()
    return (
        ("duckduckgo.com" in h and "/html/" in u)
        or ("bing.com" in h and "/search" in u)
        or ("search.yahoo.co.jp" in h)
        or "検索結果" in title
        or "search results" in t
    )


def is_map_or_directory_page(url: str, html_text: str, title: str) -> bool:
    h = host_of(url)
    u = url.lower()
    return any(x in h for x in ["map.yahoo.co.jp", "mapion.co.jp", "itp.ne.jp", "townpage.goo.ne.jp", "24u.jp", "navitime.co.jp"]) or any(k in u for k in ["/poi", "spot=", "spt=", "facility"])


def is_corporate_db_page(url: str, html_text: str, title: str) -> bool:
    h = host_of(url)
    u = url.lower()
    return any(x in h for x in ["houjin.", "kaisharesearch.com", "houjin.goo.to", "houjin.info", "houjin.jp", "jpnumber.com"]) or any(k in u for k in ["/numberinfo_", "directory", "listing", "detail"])


def is_portal_page(url: str, html_text: str, title: str) -> bool:
    h = host_of(url)
    t = title.lower()
    u = url.lower()
    if any(x in h for x in ["rakuten.co.jp", "facebook.com", "instagram.com", "x.com", "twitter.com"]):
        return True
    if is_low_trust_host(url):
        return True
    if "ランキング" in title or "口コミ" in title or "portal" in t:
        return True
    # '/company/' is not a universal portal signal; it is common on official sites.
    return any(k in u for k in ["/member/", "/shop/", "/association/"])


def is_reference_directory_page(url: str, html_text: str, title: str) -> bool:
    u = url.lower()
    h = host_of(url)
    # kensetumap profile is discoverable but not final-acceptable
    if "kensetumap.com" in h and re.search(r"/company/.+/profile\\.php", u):
        return True
    return any(k in u for k in ["/directory", "/listing", "/detail", "profile.php"])


def is_member_listing_page(url: str, html_text: str, title: str) -> bool:
    u = url.lower()
    return any(k in u for k in ["/member/", "/members/", "/kaiin/", "/shop/", "/company/"]) and ("profile" in u or "detail" in u)


def is_phone_book_page(url: str, html_text: str, title: str) -> bool:
    h = host_of(url)
    u = url.lower()
    return "jpnumber.com" in h or "/numberinfo_" in u or "電話帳" in title


def is_map_poi_page(url: str, html_text: str, title: str) -> bool:
    h = host_of(url)
    u = url.lower()
    return "navitime.co.jp" in h or "map.yahoo.co.jp" in h or any(k in u for k in ["/poi", "spot=", "spt="])


def is_likely_official_site(url: str, html_text: str, title: str) -> bool:
    host_class = classify_host(url, html_text, title)
    if is_final_forbidden_class(host_class):
        return False
    if is_blacklisted_host(url):
        return False
    if is_search_result_page(url, html_text, title):
        return False
    if is_map_or_directory_page(url, html_text, title):
        return False
    if is_corporate_db_page(url, html_text, title):
        return False
    if is_portal_page(url, html_text, title):
        return False
    if is_reference_directory_page(url, html_text, title):
        return False
    if is_member_listing_page(url, html_text, title):
        return False
    if is_phone_book_page(url, html_text, title):
        return False
    if is_map_poi_page(url, html_text, title):
        return False
    return True


def normalize_final_url(url: str) -> str:
    p = urlparse(url)
    path = p.path or "/"
    if path in {"/about", "/about/", "/company", "/company/", "/profile", "/profile/", "/index.html", "/index.htm"}:
        path = "/"
    return p._replace(path=path, query="", fragment="").geturl()


def final_url_score(url: str) -> int:
    host_class = classify_host(url, "", "")
    if is_final_forbidden_class(host_class):
        return -200
    if is_low_trust_host(url) or is_blacklisted_host(url):
        return -1000
    u = url.lower()
    p = urlparse(u).path or "/"
    if p in {"/", ""}:
        return 60
    if any(k in p for k in ["/company", "/profile", "/about", "/office", "/gaiyo", "/access"]):
        return 25
    if any(k in p for k in FINAL_STRICT_BAD_HINTS):
        return -120
    if any(k in p for k in FINAL_BAD_PATH_HINTS):
        if "/recruit" in p:
            return -100
        if "/contact" in p:
            return -80
        if "/location/" in p:
            return -60
        return -50
    if urlparse(u).query:
        return -40
    return 0


def extract_h1(html_text: str) -> str:
    return normalize_text(" ".join(re.findall(r"<h1[^>]*>(.*?)</h1>", html_text, flags=re.IGNORECASE | re.DOTALL)))


def key_presence_score(html_text: str, title: str, office: str, corp: str, addr: str, phone: str, rep: str, reg_no: str) -> tuple[int, Dict[str, int]]:
    detail: Dict[str, int] = {}
    body = strip_tags(html_text)
    h1 = extract_h1(html_text)
    footer = extract_section(html_text, "footer")
    address_tag = extract_section(html_text, "address")
    key_text = " ".join([title, h1, body, footer, address_tag])
    hits = 0
    if match_name(key_text, office, corp):
        hits += 1
        detail["key_name"] = 1
    if match_addr(key_text, addr):
        hits += 1
        detail["key_addr"] = 1
    if match_phone(key_text, phone):
        hits += 1
        detail["key_phone"] = 1
    if match_rep(key_text, rep):
        hits += 1
        detail["key_rep"] = 1
    if match_reg(key_text, reg_no):
        hits += 1
        detail["key_reg"] = 1
    return hits, detail


def final_evidence_gate(url: str, html_text: str, office: str, corp: str, addr: str, phone: str, rep: str, reg_no: str) -> tuple[bool, Dict[str, int], str]:
    title = extract_title(html_text)
    body = strip_tags(html_text)
    h1 = extract_h1(html_text)
    footer = extract_section(html_text, "footer")
    profile_like = " ".join([extract_section(html_text, "section"), extract_section(html_text, "article")])
    pool = " ".join([title, h1, body, footer, profile_like])

    name_ev = match_name(pool, office, corp)
    addr_ev = match_addr(pool, addr)
    rep_ev = match_rep(pool, rep)
    phone_ev = match_phone(pool, phone)
    reg_ev = match_reg(pool, reg_no)
    key_count = sum([1 if x else 0 for x in [name_ev, addr_ev, rep_ev, phone_ev, reg_ev]])
    ownership_positive = 0
    ownership_negative = 0
    if re.search(r"(会社概要|会社案内|事務所概要|about|company|profile|access)", pool, flags=re.IGNORECASE):
        ownership_positive += 1
    if re.search(r"(お問い合わせ|contact|所在地|住所|電話|tel|copyright|©)", pool, flags=re.IGNORECASE):
        ownership_positive += 1
    if re.search(r"(portal|directory|review|口コミ|レビュー|compass|corporation|portfolio|users|visit_captcha|窓口|比較)", pool, flags=re.IGNORECASE):
        ownership_negative += 1
    if any(k in url.lower() for k in ["/compass/company/", "/jp/corporation/", "/portfolio", "/dr/users/", "/visit_captcha"]):
        ownership_negative += 2
    if host_slug_matches_name(url, office, corp):
        ownership_positive += 1

    ownership_ok = ownership_positive >= 1 and ownership_negative == 0

    detail = {
        "ev_name": 1 if name_ev else 0,
        "ev_addr": 1 if addr_ev else 0,
        "ev_rep": 1 if rep_ev else 0,
        "ev_phone": 1 if phone_ev else 0,
        "ev_reg": 1 if reg_ev else 0,
        "ev_count": key_count,
        "ownership_pos": ownership_positive,
        "ownership_neg": ownership_negative,
        "ownership_ok": 1 if ownership_ok else 0,
    }

    host_class = classify_host(url, html_text, title)
    if is_final_forbidden_class(host_class):
        return False, detail, f"forbidden_host_class:{host_class.value}"
    p = urlparse(url.lower()).path or "/"
    if any(k in p for k in FINAL_STRICT_BAD_HINTS):
        return False, detail, "forbidden_final_path"
    if key_count < 2:
        return False, detail, "evidence不足(2系統未満)"
    if not name_ev:
        return False, detail, "name evidence不足"
    if not (addr_ev or rep_ev):
        return False, detail, "address/rep evidence不足"
    if not ownership_ok:
        return False, detail, "official ownership evidence不足"
    return True, detail, "evidence gate pass"


def domain_score(url: str) -> int:
    h = host_of(url)
    p = urlparse(url).path or "/"
    score = 50 if h.endswith(".co.jp") else 40
    if len([x for x in p.split("/") if x]) >= 2:
        score -= 30
    if h.count(".") >= 3:
        score -= 20
    return score


# ---------------------------------------------------------------------------
# HTML extraction / scoring
# ---------------------------------------------------------------------------
def strip_tags(src: str) -> str:
    src = re.sub(r"<script.*?</script>|<style.*?</style>|<noscript.*?</noscript>", " ", src, flags=re.IGNORECASE | re.DOTALL)
    return normalize_text(re.sub(r"<[^>]+>", " ", src))


def extract_section(html_text: str, tag: str) -> str:
    return normalize_text(" ".join(re.findall(rf"<{tag}[^>]*>(.*?)</{tag}>", html_text, flags=re.IGNORECASE | re.DOTALL)))


def match_name(text: str, office: str, corp: str) -> bool:
    t = compact_text(text)
    core_office = extract_core_name(office)
    core_corp = extract_core_name(corp)
    ascii_office = ascii_variations(core_office)
    ascii_corp = ascii_variations(core_corp)
    cands = [
        compact_text(strip_symbols(office)),
        compact_text(remove_corp_words(office)),
        compact_text(strip_symbols(corp)),
        compact_text(remove_corp_words(corp)),
        compact_text(core_office),
        compact_text(core_corp),
    ]
    cands.extend(compact_text(x) for x in ascii_office + ascii_corp)
    return any(c and c in t for c in cands)


def match_phone(text: str, phone: str) -> bool:
    p = normalize_phone(phone)
    if not p:
        return False
    d = re.sub(r"\D", "", text)
    return p in d


def match_addr(text: str, addr: str) -> bool:
    full, no_banchi, city = address_variants(addr)
    t = normalize_text(text)
    city_no_pref = re.sub(r"^[^都道府県]+[都道府県]", "", city).strip()
    full_no_pref = re.sub(r"^[^都道府県]+[都道府県]", "", full).strip()
    return any(x and len(x) >= 4 and x in t for x in [full, no_banchi, city, city_no_pref, full_no_pref])


def match_rep(text: str, rep: str) -> bool:
    r = compact_text(rep)
    return bool(r and r in compact_text(text))


def match_reg(text: str, reg_no: str) -> bool:
    rr = compact_text(reg_no)
    return bool(rr and len(rr) >= 4 and rr in compact_text(text))


def calc_score(url: str, html_text: str, office: str, corp: str, addr: str, phone: str, rep: str, reg_no: str) -> ScoreResult:
    detail: Dict[str, int] = {}
    title = extract_title(html_text)

    if not is_likely_official_site(url, html_text, title):
        return ScoreResult(-999, {"page_type_reject": -999}, False, False, "ページ種別除外")

    body = strip_tags(html_text)
    header = extract_section(html_text, "header")
    footer = extract_section(html_text, "footer")
    address_tag = extract_section(html_text, "address")
    page_text = " ".join([body, header, footer, address_tag, title])

    key_hits, key_detail = key_presence_score(html_text, title, office, corp, addr, phone, rep, reg_no)
    detail.update(key_detail)
    gate_ok, gate_detail, gate_reason = final_evidence_gate(url, html_text, office, corp, addr, phone, rep, reg_no)
    detail.update(gate_detail)
    name_ok = match_name(page_text, office, corp)
    phone_ok = match_phone(page_text, phone)
    addr_ok = match_addr(page_text, addr)
    rep_ok = match_rep(page_text, rep)
    reg_ok = match_reg(page_text, reg_no)

    score = 0
    ds = domain_score(url)
    score += ds
    detail["domain"] = ds

    if name_ok:
        score += 40
        detail["name_match"] = 40
    else:
        score -= 80
        detail["name_miss_penalty"] = -80

    if phone_ok:
        score += 40
        detail["phone_match"] = 40
    city_match = address_variants(addr)[2] in normalize_text(page_text)
    if addr_ok:
        score += 30
        detail["address_match"] = 30
    else:
        # personal-office friendly: relax when strong name/domain/about evidence
        relax = False
        if any(k in url.lower() for k in ["/about", "/company", "/profile", "/contact"]) and name_ok:
            relax = True
        if city_match and name_ok:
            relax = True
        if relax:
            score -= 10
            detail["address_miss_penalty_relaxed"] = -10
        else:
            score -= 45
            detail["address_miss_penalty"] = -45

    if rep_ok:
        score += 40
        detail["rep_match"] = 40
    elif rep:
        score -= 25
        detail["rep_miss_penalty"] = -25
    if reg_ok:
        score += 20
        detail["reg_match"] = 20

    structure_bonus = 0
    for s in [header, footer, address_tag]:
        if match_name(s, office, corp):
            structure_bonus += 10
        if match_phone(s, phone):
            structure_bonus += 10
        if match_addr(s, addr):
            structure_bonus += 10
    if structure_bonus:
        score += structure_bonus
        detail["structure_bonus"] = structure_bonus

    low = f"{url} {title}".lower()
    if any(k.lower() in low for k in PROFILE_HINTS):
        score += 20
        detail["profile_bonus"] = 20
    if any(k.lower() in low for k in LOW_VALUE_HINTS):
        score -= 30
        detail["low_value_penalty"] = -30

    schema = ("schema.org" in html_text.lower() and "organization" in html_text.lower())
    if schema:
        score += 10
        detail["schema_bonus"] = 10

    mobile = normalize_phone(phone).startswith(("070", "080", "090"))
    strong_official_path = any(k in url.lower() for k in ["/about", "/company", "/profile", "/contact", "/access"])
    # dual instant rules
    instant_corp = bool(name_ok and phone_ok and addr_ok and is_likely_official_site(url, html_text, title))
    instant_personal = bool(
        name_ok
        and is_likely_official_site(url, html_text, title)
        and strong_official_path
        and (city_match or structure_bonus >= 10)
        and (mobile or not phone)  # mobile mismatch is not hard fail
    )
    instant = instant_corp or instant_personal
    if key_hits <= 0:
        detail["key_presence_reject"] = -1
        return ScoreResult(-999, detail, False, False, "検索キー未出現")
    if not gate_ok:
        detail["evidence_gate_reject"] = -1
        return ScoreResult(-999, detail, False, False, gate_reason)
    accepted = instant or score >= 70
    if instant_corp:
        reason = "法人強一致 instant_accept"
    elif instant_personal:
        reason = "個人事務所ルール instant_accept"
        detail["personal_office_rule"] = 1
    else:
        reason = f"スコア{score}>=70" if accepted else f"スコア不足({score})"

    return ScoreResult(score, detail, instant, accepted, reason)


# ---------------------------------------------------------------------------
# I/O (xlsx optional, csv stdlib)
# ---------------------------------------------------------------------------
COLUMN_ALIASES = {
    "office_name": ["事務所名称", "事務所名", "事業所名"],
    "corp_name": ["法人名称", "法人名", "会社名"],
    "address": ["住所", "所在地", "事務所所在地"],
    "phone": ["電話番号", "電話", "事務所電話番号"],
    "representative": ["代表者", "代表者名", "申請人", "代表"],
    "registration_no": ["登録番号", "事務所登録番号"],
}


def _clean(v: Any) -> str:
    s = normalize_text(str(v or ""))
    return "" if s.lower() == "nan" else s


def load_rows(path: str) -> list[dict[str, str]]:
    p = Path(path)
    suffix = p.suffix.lower()

    # pandas path if available
    if pd is not None and suffix in {".xlsx", ".xlsm", ".xls", ".csv"}:
        try:
            if suffix == ".csv":
                df = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
            else:
                df = pd.read_excel(path, dtype=str)
            return [{k: _clean(v) for k, v in row.items()} for row in df.to_dict(orient="records")]
        except Exception:
            pass

    # stdlib csv fallback
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                return [{k: _clean(v) for k, v in r.items()} for r in csv.DictReader(f)]
        except Exception:
            continue

    raise ValueError(f"入力読込失敗: {path} (CSVかpandas対応形式を使用してください)")


def map_record(row: dict[str, str]) -> OfficeRecord:
    def pick(key: str) -> str:
        for c in COLUMN_ALIASES[key]:
            if c in row:
                return row[c]
        return ""

    return OfficeRecord(
        office_name=pick("office_name"),
        corp_name=pick("corp_name"),
        address=pick("address"),
        phone=pick("phone"),
        representative=pick("representative"),
        registration_no=pick("registration_no"),
    )


def save_results(output_path: str, input_path: str, results: list[dict[str, Any]]) -> None:
    rows = load_rows(input_path)
    for r in rows:
        for c in OUTPUT_COLS:
            if c not in r:
                r[c] = ""

    for i, result in enumerate(results):
        if i >= len(rows):
            break
        rows[i]["公式HP"] = str(result.get("url", ""))
        rows[i]["判定詳細"] = str(result.get("detail", ""))
        rows[i]["スコア"] = str(result.get("score", ""))
        rows[i]["採用理由"] = str(result.get("reason", ""))

    p = Path(output_path)
    if p.suffix.lower() == ".xlsx" and pd is not None:
        tmp = p.with_name(f"{p.stem}.__tmp__.xlsx")
        bak = p.with_suffix(".bak.xlsx")
        import pandas as _pd  # type: ignore

        _pd.DataFrame(rows).to_excel(tmp, index=False)
        if p.exists():
            p.replace(bak)
        tmp.replace(p)
        return

    # csv fallback
    tmp = p.with_name(f"{p.stem}.__tmp__.csv")
    bak = p.with_suffix(".bak.csv")
    headers = list(rows[0].keys()) if rows else OUTPUT_COLS
    with open(tmp, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        w.writerows(rows)
    if p.exists():
        p.replace(bak)
    tmp.replace(p)


def save_state(path: str, next_index: int, results: list[dict[str, Any]]) -> None:
    Path(path).write_text(json.dumps({"next_index": next_index, "results": results}, ensure_ascii=False, indent=2), encoding="utf-8")


def load_state(path: str) -> dict[str, Any]:
    p = Path(path)
    if not p.exists():
        return {"next_index": 0, "results": []}
    return json.loads(p.read_text(encoding="utf-8"))


# ---------------------------------------------------------------------------
# Search / crawl
# ---------------------------------------------------------------------------
class HttpClient:
    def __init__(self) -> None:
        self.session = None
        if requests is not None:
            try:
                self.session = requests.Session()
                self.session.headers.update({"User-Agent": UA})
            except Exception:
                self.session = None

    def get(self, url: str) -> Optional[str]:
        if self.session is not None:
            for _ in range(RETRY + 1):
                try:
                    r = self.session.get(url, timeout=TIMEOUT)
                    if r.status_code >= 400:
                        continue
                    r.encoding = r.apparent_encoding or r.encoding
                    return r.text
                except Exception:
                    time.sleep(0.3)

        for _ in range(RETRY + 1):
            try:
                req = Request(url, headers={"User-Agent": UA})
                with urlopen(req, timeout=TIMEOUT) as resp:
                    raw = resp.read()
                for enc in ("utf-8", "cp932", "shift_jis", "latin1"):
                    try:
                        return raw.decode(enc)
                    except Exception:
                        pass
            except URLError:
                time.sleep(0.3)
            except Exception:
                time.sleep(0.3)
        return None


def extract_internal_links(base_url: str, html_text: str) -> list[str]:
    hrefs = re.findall(r'href=["\']([^"\']+)["\']', html_text, flags=re.IGNORECASE)
    out: list[str] = []
    seen: set[str] = set()
    base_host = host_of(base_url)
    for h in hrefs:
        h_low = h.lower()
        if h_low.startswith(("data:", "mailto:", "tel:")):
            continue
        u = urljoin(base_url, h).split("#")[0]
        u_low = u.lower()
        if any(u_low.endswith(ext) for ext in ASSET_EXTS):
            continue
        if u.startswith("http") and host_of(u) == base_host and u not in seen:
            seen.add(u)
            out.append(u)
    return out


def extract_external_official_links(base_url: str, html_text: str) -> list[str]:
    links: list[str] = []
    base_host = host_of(base_url)
    for m in re.finditer(r'<a[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', html_text, flags=re.IGNORECASE | re.DOTALL):
        href = m.group(1)
        anchor = strip_tags(m.group(2)).lower()
        u = urljoin(base_url, href).split("#")[0]
        if not u.startswith("http"):
            continue
        if host_of(u) == base_host:
            continue
        ext_host = host_of(u)
        if any(host_match(ext_host, b) for b in EXTERNAL_LINK_BLOCK_HOSTS):
            continue
        hint = any(k in anchor for k in ["公式", "ホームページ", "website", "web site", "home", "company", "about"])
        broad_ok = len(anchor) >= 1 and not any(k in anchor for k in ["facebook", "instagram", "twitter", "x.com", "youtube"])
        if hint or broad_ok:
            cls = classify_host(u, "", "")
            if not is_final_forbidden_class(cls):
                links.append(u)
    seen: set[str] = set()
    out: list[str] = []
    for u in links:
        if u not in seen and not is_blacklisted_host(u):
            seen.add(u)
            out.append(u)
    return out


def page_priority(url: str) -> int:
    u = url.lower()
    if any(k in u for k in ["/company", "/about", "/profile", "/gaiyo", "/overview", "/access", "/contact"]):
        return 100
    if any(k in u for k in ["/blog", "/news", "/works", "/column"]):
        return 5
    return 40


def find_official_site(record: OfficeRecord, logger: logging.Logger, fetcher: Optional[Callable[[str], Optional[str]]] = None) -> Optional[CandidateResult]:
    get_html = fetcher or HttpClient().get
    queries = build_queries(record.office_name, record.corp_name, record.address, record.phone, record.representative)
    logger.info("queries=%s", queries)

    engine_idx = 0
    consecutive_zero = 0
    consecutive_fail = 0
    candidates: list[str] = []

    for q in queries:
        retry_same_query = 0
        while retry_same_query < len(SEARCH_ENGINES):
            engine = SEARCH_ENGINES[engine_idx]
            html_text = get_html(make_search_url(engine, q))
            if not html_text:
                logger.warning("search failed engine=%s query=%s", engine, q)
                consecutive_fail += 1
                if consecutive_fail >= SEARCH_FAIL_ROTATE_THRESHOLD:
                    old = engine
                    engine_idx = (engine_idx + 1) % len(SEARCH_ENGINES)
                    consecutive_fail = 0
                    retry_same_query += 1
                    logger.info("engine switch retry same query: %s -> %s query=%s", old, SEARCH_ENGINES[engine_idx], q)
                    continue
                break
            consecutive_fail = 0

            if is_blocked_response(html_text):
                old = engine
                engine_idx = (engine_idx + 1) % len(SEARCH_ENGINES)
                retry_same_query += 1
                logger.warning("search blocked -> switch engine=%s and retry same query=%s", SEARCH_ENGINES[engine_idx], q)
                logger.info("engine switch retry same query: %s -> %s query=%s", old, SEARCH_ENGINES[engine_idx], q)
                continue

            links = parse_result_links(engine, html_text)[:MAX_SEARCH_RESULTS_PER_QUERY]
            logger.info("search engine=%s query=%s links=%d", engine, q, len(links))
            if not links:
                consecutive_zero += 1
                if consecutive_zero >= ZERO_LINK_ROTATE_THRESHOLD:
                    old = engine
                    engine_idx = (engine_idx + 1) % len(SEARCH_ENGINES)
                    consecutive_zero = 0
                    retry_same_query += 1
                    logger.info("engine switch retry same query: %s -> %s query=%s (zero links)", old, SEARCH_ENGINES[engine_idx], q)
                    continue
            else:
                consecutive_zero = 0
                candidates.extend(links)
            time.sleep(0.2)
            break

    # filter by page type
    uniq: list[str] = []
    seen: set[str] = set()
    for u in candidates:
        if u in seen:
            continue
        seen.add(u)
        title = ""
        if is_blacklisted_host(u):
            continue
        uniq.append(u)
    logger.info("candidate_urls=%d", len(uniq))

    best: Optional[CandidateResult] = None
    for seed in uniq:
        low_trust_host = is_reference_directory_page(seed, "", "") or is_member_listing_page(seed, "", "") or is_portal_page(seed, "", "")
        host_cap_pages = 1 if low_trust_host else MAX_PAGES
        host_cap_depth = 1 if low_trust_host else MAX_DEPTH
        queue: list[tuple[int, int, str]] = [(-page_priority(seed), 0, seed)]
        visited: set[str] = set()
        while queue and len(visited) < host_cap_pages:
            queue.sort(key=lambda x: x[0])
            _, depth, url = queue.pop(0)
            if url in visited or depth > host_cap_depth:
                continue
            visited.add(url)
            h = get_html(url)
            if not h:
                continue
            title = extract_title(h)
            host_class = classify_host(url, h, title)
            if is_final_forbidden_class(host_class):
                for ext in extract_external_official_links(url, h):
                    if ext not in visited:
                        queue.append((-120, depth + 1, ext))
                logger.info("final rejected url=%s reason=host_class=%s", url, host_class.value)
                continue
            if not is_likely_official_site(url, h, title):
                logger.info("final rejected url=%s reason=page_type_reject", url)
                continue

            score = calc_score(url, h, record.office_name, record.corp_name, record.address, record.phone, record.representative, record.registration_no)
            logger.info("score url=%s total=%s accepted=%s reason=%s detail=%s", url, score.total, score.accepted, score.reason, score.detail)

            cand = CandidateResult(url, score)
            if best is None or cand.score.total > best.score.total:
                best = cand
            if score.instant_accept:
                final_norm = normalize_final_url(url)
                gate_ok, gate_detail, gate_reason = final_evidence_gate(final_norm, h, record.office_name, record.corp_name, record.address, record.phone, record.representative, record.registration_no)
                score.detail.update(gate_detail)
                if not gate_ok:
                    logger.info("instant candidate rejected by evidence_gate url=%s reason=%s", final_norm, gate_reason)
                    continue
                fscore = final_url_score(final_norm)
                if fscore > -1000 and score.total + fscore >= 70:
                    score.detail["final_url_score"] = fscore
                    score.detail["final_total"] = score.total + fscore
                    logger.info("final accepted url=%s via=%s", final_norm, score.reason)
                    return CandidateResult(final_norm, score)
                logger.info("instant candidate rejected by final_url_score url=%s fscore=%s", final_norm, fscore)

            if depth < host_cap_depth:
                for link in extract_internal_links(url, h):
                    if link not in visited:
                        queue.append((-page_priority(link), depth + 1, link))

    if best and best.score.accepted:
        final_norm = normalize_final_url(best.url)
        best_html = get_html(best.url) or ""
        gate_ok, gate_detail, gate_reason = final_evidence_gate(final_norm, best_html, record.office_name, record.corp_name, record.address, record.phone, record.representative, record.registration_no)
        best.score.detail.update(gate_detail)
        if not gate_ok:
            logger.info("final rejected best url=%s reason=%s", final_norm, gate_reason)
            return None
        fscore = final_url_score(final_norm)
        if fscore <= -1000:
            logger.info("final rejected best url=%s reason=low_trust_host", final_norm)
            return None
        final_total = best.score.total + fscore
        best.score.detail["final_url_score"] = fscore
        best.score.detail["final_total"] = final_total
        if final_total < 70:
            logger.info("final rejected best url=%s reason=final_total不足(%s)", final_norm, final_total)
            return None
        logger.info("final accepted url=%s via=%s", final_norm, best.score.reason)
        return CandidateResult(final_norm, best.score)
    if best:
        logger.info("final rejected best url=%s reason=score不足 total=%s", best.url, best.score.total)
    return None


# ---------------------------------------------------------------------------
# Fixture/Sample modes
# ---------------------------------------------------------------------------
def fixture_fetcher(url: str) -> Optional[str]:
    if "duckduckgo.com" in url:
        return FIXTURE_SEARCH_DDG
    if "bing.com" in url:
        return FIXTURE_SEARCH_BING
    if "yahoo.co.jp" in url:
        return FIXTURE_SEARCH_YAHOO
    if "example.co.jp" in url:
        return FIXTURE_OFFICIAL_PAGE
    if "atelier-kasho.com" in url:
        return FIXTURE_ATELIER_PAGE
    return None


def run_selftest(logger: logging.Logger) -> int:
    ok = True
    try:
        Path("selftest.tmp").write_text("ok", encoding="utf-8")
        Path("selftest.tmp").unlink()
        logger.info("selftest writable: OK")
    except Exception as e:  # noqa: BLE001
        logger.error("selftest writable: NG %s", e)
        ok = False

    try:
        save_state("selftest_state.json", 1, [{"url": "x"}])
        st = load_state("selftest_state.json")
        Path("selftest_state.json").unlink(missing_ok=True)
        logger.info("selftest state: %s", "OK" if st.get("next_index") == 1 else "NG")
        ok = ok and (st.get("next_index") == 1)
    except Exception as e:  # noqa: BLE001
        logger.error("selftest state: NG %s", e)
        ok = False

    return 0 if ok else 1


def run_fixture_test(logger: logging.Logger) -> int:
    links = parse_result_links("duckduckgo", FIXTURE_SEARCH_DDG)
    logger.info("fixture links=%s", links)
    if not links:
        return 1

    # noisy URLs must be excluded
    noisy = [
        "https://search.yahoo.co.jp/search?p=x",
        "https://24u.jp/a",
        "https://houjin.info/x",
        "https://map.yahoo.co.jp/x",
        "https://kaisharesearch.com/x",
        "https://www.navitime.co.jp/poi?spot=00011-080000507",
        "https://www.jpnumber.com/numberinfo_092_753_9567.html",
    ]
    for u in noisy:
        if not (is_search_result_page(u, "", "") or is_map_or_directory_page(u, "", "") or is_corporate_db_page(u, "", "") or is_portal_page(u, "", "") or is_blacklisted_host(u)):
            logger.error("fixture noise filter NG for %s", u)
            return 1

    rec = OfficeRecord("株式会社テスト設計 一級建築士事務所", "株式会社テスト設計", "福岡県北九州市門司区本町4-14", "0933310555", "山田太郎", "A-1234")
    cand = find_official_site(rec, logger, fetcher=fixture_fetcher)
    if not cand:
        logger.error("fixture candidate missing")
        return 1
    if not cand.url.startswith("https://example.co.jp"):
        logger.error("fixture candidate wrong url=%s", cand.url)
        return 1
    logger.info("fixture candidate=%s score=%s", cand.url, cand.score.total)

    # personal office style test (atelier-kasho)
    rec2 = OfficeRecord("atelier Kasho一級建築士事務所", "", "福岡県福岡市中央区", "09012345678", "", "")
    cand2 = find_official_site(rec2, logger, fetcher=fixture_fetcher)
    if not cand2 or "atelier-kasho.com" not in cand2.url:
        logger.error("fixture personal-office NG: %s", cand2.url if cand2 else "None")
        return 1
    logger.info("fixture personal-office candidate=%s score=%s reason=%s", cand2.url, cand2.score.total, cand2.score.reason)
    return 0


def run_sample(logger: logging.Logger) -> int:
    sample_input = Path("sample_input.csv")
    if not sample_input.exists():
        with sample_input.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=["事務所名称", "法人名称", "住所", "電話番号", "代表者", "登録番号"])
            w.writeheader()
            w.writerow({"事務所名称": "株式会社テスト設計 一級建築士事務所", "法人名称": "株式会社テスト設計", "住所": "福岡県北九州市門司区本町4-14", "電話番号": "0933310555", "代表者": "山田太郎", "登録番号": "A-1234"})

    rows = load_rows(str(sample_input))
    records = [map_record(r) for r in rows]
    results: list[dict[str, Any]] = []
    for rec in records:
        cand = find_official_site(rec, logger, fetcher=fixture_fetcher)
        if cand:
            results.append({"url": cand.url, "detail": json.dumps(cand.score.detail, ensure_ascii=False), "score": cand.score.total, "reason": cand.score.reason})
        else:
            results.append({"url": "", "detail": "{}", "score": 0, "reason": "有効候補なし"})

    save_results("sample_results.csv", str(sample_input), results)
    logger.info("sample-run saved=sample_results.csv")
    return 0


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------
def resolve_input_path(raw_input: Optional[str]) -> Path:
    if raw_input:
        p = Path(raw_input).expanduser()
        if p.exists():
            return p
    for c in Path(".").glob("*.xlsx"):
        return c
    for c in Path(".").glob("*.csv"):
        return c
    raise FileNotFoundError("入力ファイルが見つかりません")


def check_env() -> int:
    missing = []
    if pd is None:
        missing.append("pandas/openpyxl (xlsx入出力時のみ必要)")
    if requests is None:
        missing.append("requests (任意、urllibで代替可能)")
    print("環境チェック:")
    if missing:
        print("- 不足(ただし一部は代替経路あり):")
        for m in missing:
            print("  *", m)
    else:
        print("- 主要依存は利用可能")
    print("- selftest/fixture-test/sample-run は標準ライブラリで実行可能")
    return 0


def main() -> int:
    args = parse_args()
    logger = setup_logger()

    if args.check_env:
        return check_env()
    if args.selftest:
        return run_selftest(logger)
    if args.fixture_test:
        return run_fixture_test(logger)
    if args.sample_run:
        return run_sample(logger)

    try:
        input_path = resolve_input_path(args.input)
    except Exception as e:  # noqa: BLE001
        logger.error("input resolve failed: %s", e)
        return 2

    logger.info("start input=%s output=%s", input_path, args.output)
    try:
        rows = load_rows(str(input_path))
    except Exception as e:  # noqa: BLE001
        logger.exception("input load failed: %s", e)
        return 2

    records = [map_record(r) for r in rows]
    if args.max_rows > 0:
        records = records[: args.max_rows]
    logger.info("total records=%d", len(records))

    state = load_state(args.state) if args.resume else {"next_index": 0, "results": []}
    start = int(state.get("next_index", 0))
    results: list[dict[str, Any]] = list(state.get("results", []))
    if len(results) < len(records):
        results.extend({"url": "", "detail": "", "score": "", "reason": ""} for _ in range(len(records) - len(results)))

    for idx in range(start, len(records)):
        rec = records[idx]
        logger.info("row=%d office=%s", idx + 1, rec.office_name or rec.corp_name)
        try:
            cand = find_official_site(rec, logger)
            if cand:
                results[idx] = {"url": cand.url, "detail": json.dumps(cand.score.detail, ensure_ascii=False), "score": cand.score.total, "reason": cand.score.reason}
                logger.info("accepted row=%d url=%s", idx + 1, cand.url)
            else:
                results[idx] = {"url": "", "detail": "{}", "score": 0, "reason": "有効候補なし"}
                logger.info("rejected row=%d reason=no reliable candidate", idx + 1)
        except Exception as e:  # noqa: BLE001
            logger.exception("row=%d error=%s", idx + 1, e)
            results[idx] = {"url": "", "detail": "{}", "score": 0, "reason": f"例外: {e}"}

        if (idx + 1) % args.save_interval == 0:
            save_results(args.output, str(input_path), results)
            save_state(args.state, idx + 1, results)
            logger.info("checkpoint saved row=%d", idx + 1)

    save_results(args.output, str(input_path), results)
    save_state(args.state, len(records), results)
    logger.info("done output=%s", args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
