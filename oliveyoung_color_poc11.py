"""
올리브영 메이크업(립/베이스/아이) 옵션별 대표 색상 수집 PoC  (v11)
- 비로그인, API key 없음, requests 계열(curl_cffi) + BeautifulSoup, Selenium/Playwright 없음
- 중분류당 상품 10개, 페이지 요청 간격 0.5초, 이미지 6개 병렬 다운로드
- 상세 페이지 HTML 안의 Next.js 데이터(self.__next_f.push)에서 옵션/브랜드/소분류를 읽음
  (화면의 옵션 목록은 JS가 그리기 때문에 HTML 태그 셀렉터로는 비어 보임)
- 옵션별 이미지(컬러칩/색상비교/옵션이미지)를 모두 분석해 '신뢰도 점수'가 가장 높은 색을 채택
  · 컬러칩 방식: 이미지 전체 유채색 비율 (제품 사진을 잘 걸러냄, 꽉 찬 컬러칩에 강함)
  · 스와치 방식: 배경 제거 후 남은 영역의 색 균일도 (흰 배경 스와치, 저채도 색에 강함)
  · 여러 호수가 같은 이미지를 쓰면(세트 사진) 제외 → 같은 호수의 다른 옵션 색을 가져옴
- 이미지는 메모리에서만 열어 색을 뽑고 저장하지 않음
- 결과: 19열 엑셀, '색상(배경)' 칸을 HEX 색으로 채움

pip install curl_cffi beautifulsoup4 pillow openpyxl
같은 폴더에 cookie.txt (cookie / user-agent 형식) 필요
"""
import io, os, re, json, time, random, threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

from bs4 import BeautifulSoup
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

try:
    from curl_cffi import requests as http
    from curl_cffi.requests.impersonate import BrowserType, REAL_TARGET_MAP
except ImportError as e:
    raise SystemExit("curl_cffi가 필요합니다: python -m pip install curl_cffi") from e

BASE = "https://www.oliveyoung.co.kr"
MAKEUP_URL = f"{BASE}/store/display/getCategoryShop.do?dispCatNo=10000010002"
TARGETS = {"립메이크업": "100000100020006",
           "베이스메이크업": "100000100020001",
           "아이메이크업": "100000100020007"}
PER_CATEGORY = 10
PAGE_DELAY = 0.5          # 페이지 요청의 최소 간격(초)
IMG_DELAY = 0             # 순차 다운로드가 필요한 경우에만 사용
IMG_WORKERS = 6           # 상품별 이미지 동시 다운로드 수
IMG_SIZE = 300            # 분석용 컬러칩 크기
MIN_COLOR_FRAC = 0.5      # 유채색 비율이 이보다 낮으면 제품 사진으로 보고 HEX 제외 (테스트 스크립트와 동일)
HERE = os.path.dirname(os.path.abspath(__file__))
OUT_XLSX = os.path.join(HERE, f"oliveyoung_makeup_colors_{datetime.now():%Y%m%d_%H%M}.xlsx")

COLUMNS = ["상품번호", "중분류", "제품유형(소분류)", "브랜드명", "상품명", "호수", "호수이름",
           "올리브영 옵션명", "올리브영 HEX", "색상(배경)", "L*", "a*", "b*",
           "옵션 이미지 URL", "수집일시", "제형", "판정 근거", "색 출처", "색 신뢰도"]


# ───────── 세션 (cookie.txt) ─────────
def load_cookie_file():
    for name in ("cookie.txt", "cookie.txt.txt", "cookie"):
        p = os.path.join(HERE, name)
        if os.path.isfile(p):
            raw = open(p, "rb").read()
            for enc in ("utf-8-sig", "utf-16", "cp949"):
                try:
                    text = raw.decode(enc); break
                except UnicodeError:
                    continue
            lines = [l.strip() for l in text.splitlines() if l.strip()]
            vals = {lines[i].lower(): lines[i + 1] for i in range(len(lines) - 1)
                    if lines[i].lower() in ("cookie", "user-agent")}
            jar = {}
            for part in vals.get("cookie", "").split(";"):
                if "=" in part:
                    k, v = part.strip().split("=", 1)
                    jar.pop(k, None); jar[k] = v
            print(f"[쿠키] {name} 읽음 → {len(jar)}개, cf_clearance {'있음' if 'cf_clearance' in jar else '없음'}")
            return "; ".join(f"{k}={v}" for k, v in jar.items()), vals.get("user-agent", "")
    print(f"[쿠키] cookie.txt 없음 ({HERE})")
    return "", ""


COOKIE, UA = load_cookie_file()
PROFILE_DEFAULT = REAL_TARGET_MAP["chrome"]
PROFILE_MAJOR = int(re.search(r"chrome(\d+)", PROFILE_DEFAULT).group(1))
UA = UA or ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            f"Chrome/{PROFILE_MAJOR}.0.0.0 Safari/537.36")
UA_MAJOR_MATCH = re.search(r"Chrome/(\d+)", UA)
if not UA_MAJOR_MATCH:
    raise SystemExit("cookie.txt의 user-agent가 Chrome UA 형식이 아닙니다")
UA_MAJOR = int(UA_MAJOR_MATCH.group(1))
PROFILE = f"chrome{UA_MAJOR}" if f"chrome{UA_MAJOR}" in {p.value for p in BrowserType} else PROFILE_DEFAULT
IMPERSONATE = {"impersonate": PROFILE}
PLATFORM = '"Windows"' if "Windows NT" in UA else '"macOS"'
BROWSER_HEADERS = {
    "User-Agent": UA,
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8",
    "Sec-CH-UA": f'"Not;A=Brand";v="8", "Chromium";v="{UA_MAJOR}", "Google Chrome";v="{UA_MAJOR}"',
    "Sec-CH-UA-Mobile": "?0",
    "Sec-CH-UA-Platform": PLATFORM,
    "Referer": f"{BASE}/store/main/main.do",
}
PAGE_HEADERS = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "same-origin",
    "Upgrade-Insecure-Requests": "1",
}
IMAGE_HEADERS = {
    "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
    "Sec-Fetch-Dest": "image",
    "Sec-Fetch-Mode": "no-cors",
    "Sec-Fetch-Site": "same-site",
}
session = http.Session()
session.headers.update(BROWSER_HEADERS)
if COOKIE:
    session.headers["Cookie"] = COOKIE
_last = {"page": 0.0, "img": 0.0}


RETRY_429 = [60, 120, 180]   # 429(요청 과다) 시 쉬는 시간(초)


def get(url, kind="page", **kw):
    default_headers = {} if "/review/api/" in url else PAGE_HEADERS
    headers = {**default_headers, **kw.pop("headers", {})}
    for attempt in range(len(RETRY_429) + 1):
        wait = (PAGE_DELAY if kind == "page" else IMG_DELAY) - (time.time() - _last[kind])
        if wait > 0:
            time.sleep(wait + random.uniform(0, 0.5))
        r = session.get(url, timeout=20, headers=headers, **IMPERSONATE, **kw)
        _last[kind] = time.time()
        if r.status_code == 429 and attempt < len(RETRY_429):
            print(f"\n      [429 요청 과다] {RETRY_429[attempt]}초 쉬었다가 다시 시도합니다...")
            time.sleep(RETRY_429[attempt])
            continue
        if r.status_code == 403:
            raise RuntimeError(f"차단됨(403, {r.headers.get('server', '?')}) → cookie.txt 갱신 필요")
        r.raise_for_status()
        return r


# ───────── 1. 목록 페이지 → 상품번호 ─────────
def list_goods(cat_no, limit):
    params = {"dispCatNo": cat_no, "fltDispCatNo": "", "prdSort": "01", "pageIdx": 1, "rowsPerPage": 24}
    soup = BeautifulSoup(get(f"{BASE}/store/display/getMCategoryList.do", params=params,
                             headers={"Referer": MAKEUP_URL}).text, "html.parser")
    goods = []
    for a in soup.select("a[data-ref-goodsno], a[href*='goodsNo=']"):
        no = a.get("data-ref-goodsno") or (re.search(r"goodsNo=(\w+)", a.get("href", "")) or [None, None])[1]
        if no and no not in goods:
            goods.append(no)
        if len(goods) >= limit:
            break
    return goods


# ───────── 2. 상세 페이지 → Next.js 데이터 파싱 ─────────
PUSH_RE = re.compile(r'self\.__next_f\.push\(\[1,("(?:[^"\\]|\\.)*")\]\)')


def next_data_text(html):
    """self.__next_f.push([1,"..."]) 조각들을 이어 붙여 하나의 텍스트로 복원"""
    return "".join(json.loads(m.group(1)) for m in PUSH_RE.finditer(html))


def json_after(txt, key):
    """txt 안에서 "key": 바로 뒤의 JSON 값(배열/객체)을 파싱"""
    i = txt.find(f'"{key}":')
    if i < 0:
        return None
    try:
        return json.JSONDecoder().raw_decode(txt, i + len(key) + 3)[0]
    except ValueError:
        return None


def str_after(txt, key):
    m = re.search(rf'"{key}":"((?:[^"\\]|\\.)*)"', txt)
    return json.loads(f'"{m.group(1)}"') if m else ""


def parse_detail(html, goods_no):
    txt = next_data_text(html)
    cat = json_after(txt, "displayCategory") or {}
    options = [o for o in (json_after(txt, "options") or []) if o.get("goodsNumber") == goods_no]
    seen, uniq = set(), []
    for o in options:                                  # 같은 옵션이 두 번 나오는 경우 제거
        if o.get("optionNumber") not in seen:
            seen.add(o.get("optionNumber")); uniq.append(o)
    name = (uniq[0].get("goodsName") if uniq else "") or str_after(txt, "goodsName")
    return {"sub_cat": cat.get("leafCategoryName") or cat.get("lowerCategoryName", ""),
            "brand": str_after(txt, "onlineBrandName"),
            "name": name, "options": uniq}


# ───────── 3. 호수 분리 / 색 추출 / Lab ─────────
HOSU_RE = re.compile(r"^#?([A-Za-z]{0,3}\d+[A-Za-z0-9.\-]*)[\s._)\]:-]*(.*)$")


def split_hosu(option_name):
    s = re.sub(r"\[[^\]]*\]|\([^)]*\)", " ", option_name).strip()   # [기획], (단품) 태그 제거
    m = HOSU_RE.match(s)
    return (m.group(1), m.group(2).strip()) if m else ("", s)


MIN_CONF = 0.5            # 신뢰도가 이보다 낮으면 색으로 채택하지 않음
SWATCH_BG_DIST = 20       # 스와치 방식: 테두리 색(배경)과 이 거리 이내면 배경
SWATCH_NEAR = 45          # 스와치 방식: 대표색과 RGB 거리 45 이내인 픽셀 비율 = 균일도
MIN_OBJECT_FRAC = 0.02    # 배경 제거 후 남은 부분이 2% 미만이면 전체를 색 영역으로 사용


def _lum(p):
    return 0.2126 * p[0] + 0.7152 * p[1] + 0.0722 * p[2]


def open_rgb(img_bytes, max_side=200):
    im = Image.open(io.BytesIO(img_bytes))
    if im.mode in ("RGBA", "LA", "P"):                       # 투명 배경은 흰색으로 합성
        im = Image.alpha_composite(Image.new("RGBA", im.size, (255, 255, 255, 255)), im.convert("RGBA"))
    im = im.convert("RGB")
    im.thumbnail((max_side, max_side))
    return im


def score_chip(im):
    """컬러칩 방식 (v4 원래 방식): 이미지 전체 중 유채색 픽셀 비율 = 신뢰도
    제품 사진(흰 배경·검정 용기)은 비율이 낮아 자동으로 걸러짐"""
    px = list(im.getdata())[::4]
    col = [p + (_lum(p),) for p in px if max(p) - min(p) >= 25]
    frac = len(col) / len(px) if px else 0
    if len(col) < 50:
        return None, 0.0
    col.sort(key=lambda p: p[3])
    sel = col[int(len(col) * 0.3):int(len(col) * 0.7)] or col
    return [sum(p[j] for p in sel) / len(sel) for j in range(3)], frac


def score_swatch(im):
    """스와치 방식: 테두리 색을 배경으로 보고 제거 → 남은 영역의 대표색(밝기 30~70% 평균)
    신뢰도 = 남은 영역 중 대표색과 비슷한 픽셀 비율(균일도). 제품 사진·팔레트는 낮게 나옴"""
    w, h = im.size
    pix = im.load()
    border = [pix[x, 0] for x in range(w)] + [pix[x, h - 1] for x in range(w)] + \
             [pix[0, y] for y in range(h)] + [pix[w - 1, y] for y in range(h)]
    bg = [sorted(p[k] for p in border)[len(border) // 2] for k in range(3)]
    allp = [pix[x, y] for y in range(0, h, 2) for x in range(0, w, 2)]
    obj = [p for p in allp if sum((p[k] - bg[k]) ** 2 for k in range(3)) ** 0.5 > SWATCH_BG_DIST]
    if len(obj) < len(allp) * MIN_OBJECT_FRAC:                # 화면 전체가 한 색인 컬러칩
        obj = allp
    s = sorted(obj, key=_lum)
    sel = s[int(len(s) * 0.3):int(len(s) * 0.7)] or s
    rgb = [sum(p[j] for p in sel) / len(sel) for j in range(3)]
    near = sum(1 for p in obj if sum((p[k] - rgb[k]) ** 2 for k in range(3)) ** 0.5 < SWATCH_NEAR)
    return rgb, near / len(obj)


COMPARE_PENALTY = 0.9     # 색상비교 이미지만 있는 경우 점수 ×0.9


TIP_BG_DIST = 40          # 펜슬 심: 배경과 이 거리 넘으면 제품 픽셀
TIP_END_FRAC = 0.08       # 펜슬 심: 가장 가는 끝에서 제품 길이의 8% 구간을 심으로 봄
TIP_NARROW_RATIO = 0.6    # 심 쪽 끝 폭이 다른 끝들보다 이 비율 이하로 가늘어야 펜슬 사진으로 인정
TIP_MIN_ASPECT = 2.5      # 제품 영역이 가로·세로 중 한쪽으로 2.5배 이상 길어야 펜슬 사진으로 인정


def score_pencil_tip(im):
    """브로우·라이너 펜슬 사진에서 '심 끝'만 측정
    1) 배경 제거 후 제품 영역의 상하좌우 끝 폭 비교 → 이미지 테두리에 닿지 않은 가장 가는 끝 = 심
    2) 그 끝에서 제품 길이의 8% 구간 픽셀을 밝기 30~70% 평균
    신뢰도 = 그 구간 픽셀이 대표색과 비슷한 비율(균일도)"""
    w, h = im.size
    pix = im.load()
    border = [pix[x, 0] for x in range(w)] + [pix[x, h - 1] for x in range(w)] + \
             [pix[0, y] for y in range(h)] + [pix[w - 1, y] for y in range(h)]
    bg = [sorted(p[k] for p in border)[len(border) // 2] for k in range(3)]
    obj = [[sum((pix[x, y][k] - bg[k]) ** 2 for k in range(3)) ** 0.5 > TIP_BG_DIST for x in range(w)] for y in range(h)]
    ys = [y for y in range(h) if any(obj[y])]
    xs = [x for x in range(w) if any(obj[y][x] for y in range(h))]
    if not ys or not xs:
        return None, 0.0
    x0, x1, y0, y1 = xs[0], xs[-1], ys[0], ys[-1]
    bw, bh = x1 - x0 + 1, y1 - y0 + 1
    if max(bw, bh) < TIP_MIN_ASPECT * min(bw, bh):                    # 길쭉하지 않으면 펜슬 사진 아님
        return None, 0.0

    def band(side, frac):
        ext = (y1 - y0 + 1) if side in ("top", "bottom") else (x1 - x0 + 1)
        n = max(3, round(ext * frac))
        pts = []
        for i in range(n):
            if side == "top":
                pts += [pix[x, y0 + i] for x in range(x0, x1 + 1) if obj[y0 + i][x]]
            elif side == "bottom":
                pts += [pix[x, y1 - i] for x in range(x0, x1 + 1) if obj[y1 - i][x]]
            elif side == "left":
                pts += [pix[x0 + i, y] for y in range(y0, y1 + 1) if obj[y][x0 + i]]
            else:
                pts += [pix[x1 - i, y] for y in range(y0, y1 + 1) if obj[y][x1 - i]]
        return pts, n

    touches = {"top": y0 == 0, "bottom": y1 == h - 1, "left": x0 == 0, "right": x1 == w - 1}
    widths = {s: len(band(s, 0.1)[0]) / band(s, 0.1)[1] for s in touches}
    free = [s for s in touches if not touches[s] and widths[s] > 0]
    if not free:
        return None, 0.0
    tip = min(free, key=lambda s: widths[s])
    others = [widths[s] for s in widths if s != tip]
    if not others or widths[tip] > TIP_NARROW_RATIO * max(others):   # 뾰족한 끝이 없으면 펜슬 사진 아님
        return None, 0.0
    pts, _ = band(tip, TIP_END_FRAC)
    if len(pts) < 30:
        return None, 0.0
    s = sorted(pts, key=_lum)
    sel = s[int(len(s) * 0.3):int(len(s) * 0.7)] or s
    rgb = [sum(p[j] for p in sel) / len(sel) for j in range(3)]
    near = sum(1 for p in pts if sum((p[k] - rgb[k]) ** 2 for k in range(3)) ** 0.5 < SWATCH_NEAR)
    return rgb, near / len(pts)


PALE_CHROMA = 25          # 립 패키지 판별: 결과 색의 (R,G,B 최댓값-최솟값)이 이 값 미만 = 거의 무채색
PACKAGE_EDGE = 0.12       # 립 패키지 판별: 내부 무늬 밀도가 이 값 초과면 패키지 사진


def inner_edge_density(im):
    """배경을 뺀 제품 영역 안에서 밝기가 급하게 바뀌는 픽셀 비율
    (패키지 사진: 글자·캐릭터·모서리로 0.18~0.52 / 매끈한 스와치: 0.01~0.08 – 실측)"""
    w, h = im.size
    pix = im.load()
    border = [pix[x, 0] for x in range(w)] + [pix[x, h - 1] for x in range(w)] + \
             [pix[0, y] for y in range(h)] + [pix[w - 1, y] for y in range(h)]
    bg = [sorted(p[k] for p in border)[len(border) // 2] for k in range(3)]
    is_obj = lambda p: sum((p[k] - bg[k]) ** 2 for k in range(3)) ** 0.5 > SWATCH_BG_DIST
    inner = edges = 0
    for y in range(1, h - 1):
        for x in range(1, w - 1):
            p = pix[x, y]
            if not is_obj(p):
                continue
            nb = (pix[x + 1, y], pix[x - 1, y], pix[x, y + 1], pix[x, y - 1])
            if not all(is_obj(q) for q in nb):
                continue
            inner += 1
            lp = _lum(p)
            if max(abs(_lum(q) - lp) for q in nb) > 18:
                edges += 1
    return edges / inner if inner else 0.0


def conf_label(c):
    return f"{c:.2f} " + ("높음" if c >= 0.8 else "보통" if c >= 0.6 else "낮음(검수)")


# ───────── 품목별 규칙 ─────────
# 립메이크업: v8 규칙 그대로 (v8 결과가 Aside보다 정확했음 – 바꾸지 않음)
# 브로우·라이너: 색상비교는 얼굴에 바른 사진(피부가 대부분)이라 제외, 옵션 이미지(그은 선 스와치 등)로만 실측
# 마스카라: 옵션 이미지가 패키지 사진이라 패키지 색과 실제 색을 구분할 수 없음 → 비우고 검수 대상
def item_rule(mid_name, sub_cat, goods_name):
    text = f"{sub_cat or ''} {goods_name or ''}"
    if mid_name == "립메이크업":
        return "lip"
    if "마스카라" in text:
        return "mascara"
    if "브로우" in text or "라이너" in text:
        return "line"
    return "default"


def to_lab(rgb):
    def lin(v):
        v /= 255
        return v / 12.92 if v <= 0.04045 else ((v + 0.055) / 1.055) ** 2.4
    R, G, B = (lin(c) for c in rgb)
    X = (R * 0.4124 + G * 0.3576 + B * 0.1805) / 0.95047
    Y = R * 0.2126 + G * 0.7152 + B * 0.0722
    Z = (R * 0.0193 + G * 0.1192 + B * 0.9505) / 1.08883
    f = lambda t: t ** (1 / 3) if t > 0.008856 else 7.787 * t + 16 / 116
    return round(116 * f(Y) - 16, 1), round(500 * (f(X) - f(Y)), 1), round(200 * (f(Y) - f(Z)), 1)


IMG_KEYS = [("colorChipImage", "컬러칩"), ("colorCompareImage", "색상비교"), ("optionImage", "옵션이미지")]
_img_cache = {}
_img_errors = {}
_image_thread = threading.local()


def image_url(o, key="colorChipImage"):
    img = o.get(key) or {}
    return (img.get("url", "") + img.get("path", "")) if img.get("path") else ""


def candidates(o, rule="default"):
    """옵션의 이미지 후보 {url: '컬러칩·옵션이미지' 같은 라벨} (같은 파일은 1번만)
    rule='line'(브로우·라이너)이면 색상비교(얼굴 사진) 제외"""
    out = {}
    for key, label in IMG_KEYS:
        if rule == "line" and key == "colorCompareImage":
            continue
        u = image_url(o, key)
        if u:
            out[u] = f"{out[u]}·{label}" if u in out else label
    return out


def image_session():
    """이미지 세션은 스레드마다 따로 두고 페이지 쿠키는 CDN에 보내지 않는다."""
    if not hasattr(_image_thread, "session"):
        _image_thread.session = http.Session()
        _image_thread.session.headers.update(BROWSER_HEADERS)
    return _image_thread.session


def download_img(url):
    """CDN 이미지를 받고 429에만 제한적으로 재시도한다."""
    for attempt in range(4):
        r = image_session().get(f"{url}?RS={IMG_SIZE}x0&QT=100", timeout=20,
                                headers=IMAGE_HEADERS, **IMPERSONATE)
        if r.status_code == 429 and attempt < 3:
            retry_after = r.headers.get("Retry-After", "")
            wait = int(retry_after) if retry_after.isdigit() else (3, 9, 27)[attempt]
            time.sleep(wait + random.uniform(0, 0.5))
            continue
        if r.status_code == 403:
            raise RuntimeError("이미지 CDN 차단(403) – 쿠키 또는 접속 상태 확인 필요")
        r.raise_for_status()
        return r.content


def prefetch_images(opts, shared_urls, rule):
    """한 상품에서 실제 분석할 이미지 URL만 중복 제거 후 병렬로 받는다."""
    if len(_img_cache) > 300:
        _img_cache.clear()
    urls = list(dict.fromkeys(
        url for o in opts for url in candidates(o, rule)
        if url not in shared_urls and url not in _img_cache
    ))
    if not urls:
        return
    with ThreadPoolExecutor(max_workers=min(IMG_WORKERS, len(urls))) as pool:
        tasks = {pool.submit(download_img, url): url for url in urls}
        for task in as_completed(tasks):
            url = tasks[task]
            try:
                _img_cache[url] = task.result()
                _img_errors.pop(url, None)
            except Exception as e:
                _img_errors[url] = str(e)


def fetch_img(url):
    if url in _img_errors:
        raise RuntimeError(_img_errors[url])
    if url not in _img_cache:
        _img_cache[url] = download_img(url)
    return _img_cache[url]


def best_color(o, shared_urls, rule="default"):
    """후보 이미지마다 두 방식으로 점수를 매겨 가장 높은 것을 채택
    반환: (rgb|None, 사용한 url, 출처 설명, 신뢰도)"""
    cands = candidates(o, rule)
    if not cands:
        return None, "", "색 이미지 없음", 0.0
    best = (None, next(iter(cands)), "", 0.0)
    skipped, rejected_pkg = [], []
    for url, label in cands.items():
        if url in shared_urls:
            skipped.append(label)
            continue
        im = open_rgb(fetch_img(url))
        # 컬러칩 방식: 립은 v8대로 컬러칩 이미지에만,
        #   그 외(베이스·아이)는 옵션이미지에도 적용 (화면 가득 찬 텍스처 사진 – 클리오 리필 사례)
        # 색상비교 이미지에는 적용하지 않음 (모델 사진은 피부까지 유채색이라 피부색이 섞임)
        methods = [("스와치 방식", score_swatch)]
        if "컬러칩" in label or (rule != "lip" and "옵션이미지" in label):
            methods.insert(0, ("컬러칩 방식", score_chip))
        only_compare = label == "색상비교" and rule != "lip"
        for method, fn in methods:
            rgb, conf = fn(im)
            if only_compare:
                conf *= COMPARE_PENALTY       # 베이스·아이: 모델 사진일 수 있어 한 단계 낮게 평가
            # 립: 스와치 방식 결과가 거의 무채색인데 내부 무늬(글자·캐릭터·모서리)가 많으면 패키지 사진 → 제외
            #     (포켓몬 에디션 파우치, 이니미니 상자 사례 / 핑크링·유리멘탈·강철멘탈 같은 연한 스와치는 무늬가 적어 유지)
            if rgb and rule == "lip" and method == "스와치 방식" and max(rgb) - min(rgb) < PALE_CHROMA \
                    and inner_edge_density(im) > PACKAGE_EDGE:
                rejected_pkg.append(label)
                continue
            if rgb and conf > best[3]:
                best = (rgb, url, f"{label}({method})", conf)
    if best[0] and best[3] >= MIN_CONF:
        return best
    if rejected_pkg:
        return None, next(iter(cands)), f"패키지 사진으로 판단해 제외 ({'/'.join(rejected_pkg)}), 다른 색 이미지 없음", 0.0
    # 브로우·라이너: 그은 선 스와치가 없으면 펜슬 사진의 '심 끝'만 측정
    if rule == "line":
        for url, label in cands.items():
            if url in shared_urls:
                continue
            rgb, conf = score_pencil_tip(open_rgb(fetch_img(url)))
            if rgb:   # 펜슬 모양이 확인되면 기준값과 관계없이 실측값 사용, 신뢰도는 그대로 표시
                return rgb, url, f"{label}(펜슬 심 측정 – 몸체 색이 섞일 수 있음)", conf
    if skipped and not best[2]:
        return None, best[1], f"여러 호수가 같은 이미지 사용(세트 사진) – {'/'.join(skipped)} 제외", 0.0
    return None, best[1], f"제품 사진·다색 이미지로 판단 (최고 점수 {best[3]:.2f})", best[3]




# ───────── 4. 제형 판정 (1차 상품명 → 2차 리뷰 AI 요약 → 3차 리뷰 설문 통계) ─────────
REVIEW_API = "https://m.oliveyoung.co.kr/review/api"   # www 도메인에는 없음(404)

# 7종: matte / velvet / cream / satin / sheer / glossy / shimmer   (제형.md 기준)
# '글로우·윤광·촉촉'은 7종 중 어디인지 단정할 수 없어 일부러 키워드에서 뺌 → 다음 단계로 넘어감
TEXTURE_WORDS = [
    ("shimmer", ["쉬머", "글리터", "스파클", "메탈릭", "펄"]),
    ("matte",   ["세미매트", "매트", "마뜨", "무광", "보송"]),   # '세미매트'는 아래에서 velvet으로 바꿈
    ("velvet",  ["벨벳", "블러"]),
    ("cream",   ["크리미", "크림"]),
    ("satin",   ["새틴"]),
    ("sheer",   ["시어", "쉬어", "워터리", "워터"]),
    ("glossy",  ["글로시", "글로스", "물광", "유리알", "유광", "광택"]),
]
NAME_EXCLUDE = ["워터프루프", "크림치즈"]            # 제형이 아닌데 키워드를 포함하는 말
NEGATION = re.compile(r"^.{0,6}(없|않|아닌|안 )")      # '펄감이 없어요', '매트하지 않고'


def find_texture(text, exclude=()):
    """텍스트에서 가장 먼저 나오는 제형 키워드 → (제형, 키워드)"""
    for w in exclude:
        text = text.replace(w, " ")
    best = None
    for tex, words in TEXTURE_WORDS:
        for w in words:
            for m in re.finditer(re.escape(w), text):
                if NEGATION.match(text[m.end():]):
                    continue
                if best is None or m.start() < best[0]:
                    best = (m.start(), "velvet" if w == "세미매트" else tex, w)
                break
    return (best[1], best[2]) if best else (None, None)


def review_json(path, goods_no):
    r = get(f"{REVIEW_API}/{path}", headers={
        "Accept": "application/json",
        "Origin": BASE,
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-site",
        "Referer": f"{BASE}/store/goods/getGoodsDetail.do?goodsNo={goods_no}"})
    return (r.json() or {}).get("data") or {}


# 3차 기준값 (필요하면 조정)
def texture_from_stats(stats):
    q = {s.get("questionName"): {a.get("answerName", ""): a.get("answerPercentage", 0)
                                 for a in s.get("answerDtos", [])}
         for s in stats.get("satisfactionStats") or []}
    pearl, pearl_note = q.get("펄감"), None
    if pearl:
        strong = sum(v for k, v in pearl.items() if "강해" in k)
        none_ = sum(v for k, v in pearl.items() if "없어" in k)
        note = f"리뷰 펄감(강함 {strong}/보통 {pearl.get('보통이에요', 0)}/없음 {none_}%)"
        if strong >= 40:
            return "shimmer", note
        if none_ >= 60:
            return "matte", note
        pearl_note = note
    moist = q.get("수분감")
    if moist:
        wet = sum(v for k, v in moist.items() if "촉촉" in k)
        mat = sum(v for k, v in moist.items() if "매트" in k)
        note = f"리뷰 수분감(촉촉 {wet}/보통 {moist.get('보통이에요', 0)}/매트 {mat}%)"
        if mat >= 50:
            return "matte", note
        if mat >= 25:
            return "velvet", note
        if wet >= 60:
            return "glossy", note
        if wet >= 40:
            return "satin", note
        return "cream", note
    return None, pearl_note


def judge_texture(goods_no, goods_name):
    # 1차: 상품명
    tex, word = find_texture(goods_name, NAME_EXCLUDE)
    if tex:
        return tex, f"상품명 '{word}'"
    # 2차: 리뷰 AI 요약에 제형이 명시된 경우
    try:
        s = review_json(f"v1/reviews/{goods_no}/summary", goods_no)
        for i in (1, 2, 3):
            title = s.get(f"feature{i}Title") or ""
            desc = s.get(f"feature{i}Description") or ""
            tex, word = find_texture(f"{title} {desc}")
            if tex:
                return tex, f"리뷰 언급 '{title.strip()}' ('{word}')"
    except Exception as e:
        print(f"      리뷰 요약 실패: {e}")
    # 3차: 리뷰 설문 통계(수분감 / 아이섀도우는 펄감)
    try:
        tex, note = texture_from_stats(review_json(f"v2/reviews/{goods_no}/stats", goods_no))
        if tex:
            return tex, note
        if note:
            return "", f"판정 불가 ({note})"
    except Exception as e:
        print(f"      리뷰 통계 실패: {e}")
        return "", "판정 불가 (리뷰 API 실패)"
    return "", "판정 불가 (근거 없음)"


# ───────── 5. 상품 1개 → 행 목록 ─────────
def crawl_goods(mid_name, goods_no):
    d = parse_detail(get(f"{BASE}/store/goods/getGoodsDetail.do?goodsNo={goods_no}").text, goods_no)
    texture, reason = judge_texture(goods_no, d["name"])
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    base = [goods_no, mid_name, d["sub_cat"], d["brand"], d["name"]]
    tail = [texture, reason]
    if not d["options"]:
        return [base + ["", "", "단일상품(옵션없음)", "", "", "", "", "", "", now] + tail + ["", ""]]
    opts = d["options"]
    is_palette = "팔레트" in d["name"]
    rule = item_rule(mid_name, d["sub_cat"], d["name"])

    # (1) 여러 호수가 함께 쓰는 이미지 찾기 → 세트/대표 사진이라 색 근거로 쓰지 않음
    users = {}
    for o in opts:
        hosu, _ = split_hosu(o.get("optionName", ""))
        for u in candidates(o, rule):
            users.setdefault(u, set()).add(hosu or o.get("optionName", ""))
    shared = {u for u, who in users.items() if len(who) >= 2}

    if not is_palette and rule != "mascara":
        prefetch_images(opts, shared, rule)

    # (2) 옵션별 최고 점수 색  (팔레트는 옵션 1개에 여러 색이라 대표색을 만들지 않음)
    results = []
    for k, o in enumerate(opts, 1):
        print(f"\r      이미지 분석 {k}/{len(opts)}...", end="", flush=True)
        if is_palette:
            results.append((None, next(iter(candidates(o)), ""), "팔레트(여러 색) – 대표색 1개로 표현 불가, 별도 처리 필요", 0.0))
            continue
        if rule == "mascara":
            results.append((None, next(iter(candidates(o)), ""),
                            "마스카라 – 옵션 이미지가 패키지 사진이라 실제 색과 구분 불가 (검수 필요)", 0.0))
            continue
        try:
            r = best_color(o, shared, rule)
            if rule == "line" and not r[0]:
                r = (r[0], r[1], f"브로우·라이너 – 옵션 이미지에 발색(그은 선) 없음, 얼굴 사진은 제외 / {r[2]}", r[3])
            results.append(r)
        except Exception as e:
            print(f"\n      이미지 실패({o.get('optionName', '')}): {e}")
            results.append((None, "", f"이미지 오류: {e}", 0.0))
    print("\r" + " " * 40 + "\r", end="")

    # (3) 같은 호수의 다른 옵션 색을 가져옴 (예: 세트 기획 옵션 ← 본품 옵션)
    #     - 색이 없는 옵션, 또는
    #     - 같은 호수 옵션의 점수(×0.9)가 더 높은 경우 (예: 입술 모델 사진에서 뽑은 색 → 같은 호수 컬러칩 색으로 교체)
    by_hosu = {}
    for o, r in zip(opts, results):
        hosu, _ = split_hosu(o.get("optionName", ""))
        if hosu and r[0] and (hosu.upper() not in by_hosu or r[3] > by_hosu[hosu.upper()][1][3]):
            by_hosu[hosu.upper()] = (o.get("optionName", ""), r)
    for i, (o, r) in enumerate(zip(opts, results)):
        name = o.get("optionName", "")
        hosu, _ = split_hosu(name)
        if not hosu or hosu.upper() not in by_hosu:
            continue
        src_name, src = by_hosu[hosu.upper()]
        # 립은 v8대로 '색이 없을 때만' 가져옴 / 베이스·아이는 점수가 더 높으면 교체
        if src_name != name and (not r[0] or (rule != "lip" and src[3] * 0.9 > r[3])):
            results[i] = (src[0], src[1], f"같은 호수 옵션 '{src_name}'에서 가져옴 ← {src[2]}", src[3] * 0.9)

    rows = []
    for o, (rgb, url, source, conf) in zip(opts, results):
        opt = o.get("optionName", "")
        hosu, hosu_name = split_hosu(opt)
        hexv, lab = "", ("", "", "")
        if rgb:
            hexv = "#%02X%02X%02X" % tuple(round(c) for c in rgb)
            lab = to_lab(rgb)
        rows.append(base + [hosu, hosu_name, opt, hexv, "", *lab, url, now] + tail +
                    [source, conf_label(conf) if rgb else ""])
    return rows


def fmt_dur(sec):
    sec = int(max(sec, 0))
    return f"{sec // 3600}시간 {sec % 3600 // 60}분" if sec >= 3600 else f"{sec // 60}분 {sec % 60}초"


def eta_line(done, total, started, durations, rough_per_goods):
    """지금까지 상품 1개당 실제 걸린 시간 평균 × 남은 상품 수 (처음엔 대략값)"""
    per = sum(durations) / len(durations) if durations else rough_per_goods
    remain = per * (total - done)
    end = datetime.fromtimestamp(time.time() + remain)
    return (f"진행 {done}/{total} | 경과 {fmt_dur(time.time() - started)} | "
            f"남은 시간 약 {fmt_dur(remain)} (예상 완료 {end:%H:%M})")


def save_excel(rows):
    wb = Workbook(); ws = wb.active; ws.title = "oliveyoung_colors"
    ws.append(COLUMNS)
    for c in ws[1]:
        c.font = Font(bold=True)
    hex_i, fill_i = COLUMNS.index("올리브영 HEX"), COLUMNS.index("색상(배경)") + 1
    for r in rows:
        ws.append(r)
        if r[hex_i]:
            ws.cell(ws.max_row, fill_i).fill = PatternFill("solid", fgColor=r[hex_i][1:])
    for i, w in enumerate([15, 13, 14, 12, 45, 7, 18, 32, 12, 10, 7, 7, 7, 70, 19, 10, 45, 50, 12], 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    ws.freeze_panes = "B2"
    wb.save(OUT_XLSX)
    print(f"\n저장 완료: {OUT_XLSX} ({len(rows)}행, HEX 추출 {sum(1 for r in rows if r[hex_i])}개)")


def main():
    print(f"HTTP: curl_cffi({PROFILE}) | UA: Chrome {UA_MAJOR} {PLATFORM} | 쿠키: {bool(COOKIE)}")
    if UA_MAJOR != int(re.search(r"chrome(\d+)", PROFILE).group(1)):
        print("[주의] UA와 curl_cffi TLS 프로필 버전이 다릅니다. 현재 브라우저와 같은 프로필이 필요합니다.")
    started = time.time()

    # 1단계: 세 카테고리 목록을 먼저 받아 전체 작업량 확정
    jobs = []
    for mid_name, cat_no in TARGETS.items():
        try:
            goods = list_goods(cat_no, PER_CATEGORY)
            jobs += [(mid_name, no) for no in goods]
            print(f"[{mid_name}] 상품 {len(goods)}개 확인")
        except Exception as e:
            print(f"[{mid_name}] 목록 실패: {e}")

    # 대략값: 상세 페이지 5초 + 옵션 약 20개 × 이미지 1초+α  (첫 상품 이후엔 실측 평균으로 교체)
    rough = PAGE_DELAY * 2 + 20 * 1.5 * (IMG_DELAY + 0.3)   # 상세 + (필요 시) 리뷰 API + 옵션 이미지(1~3장)
    print(f"\n총 {len(jobs)}개 상품 | 첫 예상: 약 {fmt_dur(rough * len(jobs))} "
          f"(옵션 수에 따라 달라지며 진행하면서 보정됩니다)")

    # 2단계: 상품별 수집
    rows, durations, cur_mid = [], [], None
    for i, (mid_name, no) in enumerate(jobs, 1):
        if mid_name != cur_mid:
            print(f"\n[{mid_name}]"); cur_mid = mid_name
        t0 = time.time()
        try:
            got = crawl_goods(mid_name, no)
            rows += got
            print(f"   {i:02d}. {no} [{got[0][2]}] {got[0][3]} {got[0][4][:28]} → 제형 {got[0][15] or '?'} "
                  f"({got[0][16]}) | 옵션 {len(got)}개, "
                  f"HEX {sum(1 for r in got if r[8])}개")
        except Exception as e:
            print(f"   {i:02d}. {no} 실패: {e}")
        durations.append(time.time() - t0)
        print("       " + eta_line(i, len(jobs), started, durations, rough))
    print(f"\n총 소요 시간: {fmt_dur(time.time() - started)}")
    if rows:
        save_excel(rows)
    else:
        print("수집된 데이터 없음")


if __name__ == "__main__":
    main()
