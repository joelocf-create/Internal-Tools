import json
import re
import sys
import hashlib
from datetime import datetime, timezone, timedelta
from html.parser import HTMLParser
from urllib.request import Request, urlopen
from urllib.parse import urljoin, quote
from urllib.error import URLError, HTTPError

DATA_PATH = "cf-us-tariff-monitor/data.json"
USER_AGENT = "CF-US-Tariff-Monitor/1.0"
MAX_BYTES = 3_000_000
LOOKBACK_DAYS = 21

SOURCE_LISTS = [
    ("White House Releases", "https://www.whitehouse.gov/releases/"),
    ("USTR Press Releases", "https://www.ustr.gov/about/policy-offices/press-office/press-releases/2026"),
    ("CBP Trade Remedies - China 301", "https://www.cbp.gov/trade/remedies/301-certain-products-china"),
    ("CBP Media Releases", "https://www.cbp.gov/newsroom/media-releases/all"),
]

KEYWORDS = [
    "tariff", "tariffs", "section 301", "section 232", "section 122",
    "forced labor", "china", "customs", "duty", "duties",
    "chapter 99", "hts", "exclusion", "trade remedy", "reciprocal",
    "first sale", "country of origin", "textile", "blanket",
    "throw", "blind", "blinds", "shade", "shades"
]

HIGH_VALUE = [
    "section 301", "section 232", "section 122", "forced labor",
    "china", "chapter 99", "exclusion", "tariff", "duty", "hts"
]

PRODUCT_TERMS = ["blind", "blinds", "shade", "shades", "blanket", "throw", "textile"]

DATE_PATTERNS = [
    re.compile(r"\b(20\d{2}-\d{2}-\d{2})\b"),
    re.compile(r"\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+20\d{2}\b", re.I),
]

class LinkParser(HTMLParser):
    def __init__(self, base):
        super().__init__()
        self.base = base
        self.links = []
        self.in_script = False
        self.in_style = False
        self.text_buffer = []
        self.last_date = None
        self.current_href = None
        self.current_text = []

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag in ("script", "style", "noscript"):
            self.in_script = True
        if tag == "a" and "href" in attrs:
            self.current_href = urljoin(self.base, attrs["href"])
            self.current_text = []
        if tag == "time" and attrs.get("datetime"):
            self._set_date(attrs["datetime"][:10])

    def handle_endtag(self, tag):
        if tag in ("script", "style", "noscript"):
            self.in_script = False
        if tag == "a" and self.current_href:
            text = " ".join(self.current_text).strip()
            if text:
                self.links.append({
                    "url": self.current_href,
                    "title": re.sub(r"\s+", " ", text),
                    "date": self.last_date
                })
            self.current_href = None
            self.current_text = []

    def handle_data(self, data):
        if self.in_script or self.in_style:
            return
        data = re.sub(r"\s+", " ", data).strip()
        if not data:
            return
        for p in DATE_PATTERNS:
            m = p.search(data)
            if m:
                raw = m.group(1) if p.pattern.startswith(r"\\b(20") else m.group(0)
                try:
                    if re.match(r"^20\d{2}-", raw):
                        self._set_date(raw)
                    else:
                        self._set_date(datetime.strptime(raw, "%B %d, %Y").date().isoformat())
                except ValueError:
                    pass
        if self.current_href:
            self.current_text.append(data)

    def _set_date(self, value):
        if re.match(r"^20\d{2}-\d{2}-\d{2}$", value):
            self.last_date = value

def fetch(url):
    req = Request(url, headers={"User-Agent": USER_AGENT, "Accept": "text/html,application/json"})
    with urlopen(req, timeout=30) as r:
        raw = r.read(MAX_BYTES)
        return raw.decode("utf-8", errors="ignore")

def normalize_text(text):
    text = re.sub(r"<script[\s\S]*?</script>", " ", text, flags=re.I)
    text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()

def relevant(title):
    t = title.lower()
    hits = [k for k in KEYWORDS if k in t]
    return hits

def relevance(title):
    t = title.lower()
    high = [k for k in HIGH_VALUE if k in t]
    products = [k for k in PRODUCT_TERMS if k in t]
    if products and high:
        return "PRODUCT-RELEVANT"
    if high:
        return "POLICY"
    return "MONITOR"

def allowed_link(source_name, url):
    u = url.lower()
    if source_name == "White House Releases":
        return "/releases/" in u or "/presidential-actions/" in u
    if source_name == "USTR Press Releases":
        return "/press-releases/" in u and "/2026/" in u
    if source_name == "CBP Media Releases":
        return "/newsroom/" in u and (
            "/national-media-release/" in u or "/media-release/" in u
        )
    return False

def should_monitor_title(title):
    t = title.lower()
    excluded = [
        "career", "privacy", "contact", "importer tips", "importing a car",
        "intellectual property rights", "lab leak", "covid"
    ]
    return not any(x in t for x in excluded)

def canonical_id(url):
    return hashlib.sha256(url.split("#")[0].rstrip("/").encode()).hexdigest()[:20]

def load_data():
    with open(DATA_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def save_data(data):
    with open(DATA_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write("\n")

def main():
    data = load_data()
    today = datetime.now(timezone.utc).date()
    cutoff = today - timedelta(days=LOOKBACK_DAYS)

    monitor = data.setdefault("monitor", {})
    known = set(monitor.get("known_ids", []))

    # Seed the known set from existing dashboard events so the first automated
    # run does not re-add items already shown to the user.
    for event in data.get("events", []):
        if event.get("url"):
            known.add(canonical_id(event["url"]))

    candidates = []
    source_errors = []

    for source_name, url in SOURCE_LISTS:
        try:
            html = fetch(url)
            parser = LinkParser(url)
            parser.feed(html)
            seen_urls = set()
            for item in parser.links:
                u = item["url"]
                title = item["title"]
                if not u.startswith("http") or u in seen_urls:
                    continue
                seen_urls.add(u)
                if not allowed_link(source_name, u):
                    continue
                if not should_monitor_title(title):
                    continue
                hits = relevant(title)
                if not hits:
                    continue
                if relevance(title) == "MONITOR":
                    continue
                date = item.get("date")
                if date:
                    try:
                        if datetime.fromisoformat(date).date() < cutoff:
                            continue
                    except ValueError:
                        pass
                candidates.append({
                    "id": canonical_id(u),
                    "source": source_name,
                    "title": title,
                    "date": date or today.isoformat(),
                    "url": u,
                    "hits": hits,
                    "relevance": relevance(title)
                })
        except Exception as e:
            source_errors.append(f"{source_name}: {type(e).__name__}: {e}")

    # Federal Register API: machine-readable official publication index.
    fr_terms = [
        "Section 301 China tariff",
        "Section 232 tariff",
        "Section 122 tariff",
        "forced labor tariff",
        "China tariff exclusion",
        "HTS Chapter 99 tariff"
    ]
    for term in fr_terms:
        try:
            api = (
                "https://www.federalregister.gov/api/v1/documents.json"
                "?per_page=20&order=newest&conditions[term]=" + quote(term)
            )
            raw = fetch(api)
            payload = json.loads(raw)
            for doc in payload.get("results", []):
                title = doc.get("title", "")
                url = doc.get("html_url") or doc.get("document_number_url")
                pub = doc.get("publication_date")
                if not url or not title:
                    continue
                if pub:
                    try:
                        if datetime.fromisoformat(pub).date() < cutoff:
                            continue
                    except ValueError:
                        pass
                t = title.lower()
                fr_focus = [
                    "section 301", "section 232", "section 122",
                    "forced labor", "chapter 99", "htsus",
                    "china tariff", "tariff exclusion", "product exclusion"
                ]
                if "section 337" in t or "antidumping" in t or "countervailing" in t:
                    continue
                if not any(k in t for k in fr_focus):
                    continue
                hits = relevant(title)
                if not hits:
                    continue
                if relevance(title) == "MONITOR":
                    continue
                candidates.append({
                    "id": canonical_id(url),
                    "source": "Federal Register",
                    "title": title,
                    "date": pub or today.isoformat(),
                    "url": url,
                    "hits": hits,
                    "relevance": relevance(title)
                })
        except Exception as e:
            source_errors.append(f"Federal Register API ({term}): {type(e).__name__}: {e}")

    # De-duplicate by URL/id and keep the most informative hit set.
    dedup = {}
    for c in candidates:
        old = dedup.get(c["id"])
        if old is None or len(c["hits"]) > len(old["hits"]):
            dedup[c["id"]] = c

    new_items = [c for c in dedup.values() if c["id"] not in known]
    new_items.sort(key=lambda x: (x.get("date") or "", x["title"]), reverse=True)

    # Always retain the IDs we have inspected, but only change the dashboard
    # when there is a genuinely new candidate or a source error.
    monitor["last_scan_utc"] = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    monitor["source_errors"] = source_errors

    if not new_items:
        print("NO_MATERIAL_CHANGE")
        return 0

    monitor["known_ids"] = sorted(known.union(dedup.keys()))

    for item in new_items[:20]:
        data.setdefault("events", []).insert(0, {
            "status": "WATCH",
            "title": f"[自動偵測] {item['title']}",
            "date": item["date"],
            "summary": (
                f"{item['source']} 發現新的官方文件；關鍵字："
                + ", ".join(item["hits"])
                + "。目前先標記 WATCH，需進一步核對 HTS / Chapter 99 / effective date / filing impact。"
            ),
            "detail": (
                f"自動監控分類：{item['relevance']}。這是官方來源的新文件偵測，"
                "不是自動法律結論。若文件涉及 CF 的 Window Blinds/Shades、Throw/Blanket、"
                "China COO、Section 122/232/301、Forced Labor、Chapter 99 或 entry filing，"
                "再進行人工/模型深度解析。"
            ),
            "url": item["url"]
        })

    data["events"] = data["events"][:30]
    data["cards"][3]["value"] = today.isoformat()
    data["cards"][2]["value"] = data["cards"][2].get("value", "WATCH")
    data["automation_note"] = (
        "GitHub Actions 每日自動查核 White House / USTR / CBP / Federal Register。"
        "只有偵測到新官方文件或來源異常時才修改 data.json；Dashboard 由 data.json 驅動。"
    )
    save_data(data)

    print(f"NEW_ITEMS={len(new_items)}")
    if source_errors:
        print("SOURCE_ERRORS=" + " | ".join(source_errors))
    return 0

if __name__ == "__main__":
    sys.exit(main())
