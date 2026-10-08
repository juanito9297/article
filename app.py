# -*- coding: utf-8 -*-
"""Windows desktop clipping report. No AI service/API key is used."""
import ipaddress
import html
import json
from datetime import datetime
from pathlib import Path
import queue
import re
import socket
import threading
import tkinter as tk
import webbrowser
from collections import Counter, OrderedDict
from tkinter import filedialog, messagebox, ttk
from urllib.parse import urlparse, urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

CATEGORIES = ("정치", "경제", "사회일반", "해외동향", "오피니언")
PRIORITY = ("조선일보", "중앙일보", "동아일보", "한겨레", "경향신문", "한국경제", "매일경제")
ALIASES = {"조선":"조선일보", "중앙":"중앙일보", "동아":"동아일보", "경향":"경향신문", "한경":"한국경제", "매경":"매일경제"}
MAX_BYTES = 2_000_000
ARTICLE_DEADLINE = 25  # seconds, including DNS/redirect/body processing
MAX_LINKS_PER_CARD = 2  # do not let a single cluster hold the entire queue indefinitely


def media_key(name):
    name = re.sub(r"\s+", "", name or "")
    return ALIASES.get(name, name)


def media_rank(name):
    key = media_key(name)
    return PRIORITY.index(key) if key in PRIORITY else len(PRIORITY)


def normalize_title(title):
    return re.sub(r"[^\w가-힣]", "", title.casefold())


def group_key(title):
    """Conservative title-based grouping, preserving example spreadsheet clusters."""
    t = normalize_title(title)
    if re.search(r"dmz|비무장지대|목함지뢰", t, re.I) and re.search(r"폭발|지뢰|현장조사|엿새|이틀", t):
        return "사안:dmz폭발"
    if re.search(r"미중|美中|g2|트럼프|시진핑|무역휴전|빅딜", t, re.I) and re.search(r"회담|정상|무역|휴전|빅딜|갈등|합의|친분|관세", t):
        return "사안:미중회담"
    if re.search(r"北포로|북포로|북한군포로", t) and re.search(r"송환|한국행|정부|치료|조사", t):
        return "사안:북포로"
    return "제목:" + t


def read_report(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if '기사 목록' not in wb:
            raise ValueError("'기사 목록' 시트를 찾지 못했습니다.")
        rows = wb['기사 목록'].iter_rows(values_only=True)
        headers = next(rows, None)
        if not headers:
            raise ValueError("첫 행이 비어 있습니다.")
        cols = {re.sub(r"\s+", "", str(v or "")): i for i, v in enumerate(headers)}
        def idx(*names):
            return next((cols[n] for n in names if n in cols), None)
        title_i, pub_i = idx('기사제목', '제목'), idx('매체명', '보도매체')
        url_i, sum_i = idx('기사URL', 'URL'), idx('요약', '기사요약')
        page_i = idx('지면', '지면정보', '면', '신문지면')
        author_i = idx('기자명', '기자', '기자이름', '작성자')
        date_i = idx('발행일', '기사날짜', '날짜', '일자', '작성일', '보도일', '등록일')
        if title_i is None or pub_i is None:
            raise ValueError("첫 행에 '기사 제목'과 '매체명' 열이 필요합니다.")
        def cell(row, i):
            return str(row[i]).strip() if i is not None and i < len(row) and row[i] is not None else ''
        result = {c: OrderedDict() for c in CATEGORIES}
        current = None
        count = 0
        for number, row in enumerate(rows, 2):
            title = cell(row, title_i)
            if not title:
                continue
            if title in CATEGORIES:
                current = title
                continue
            if current is None:
                raise ValueError(f'{number}행 앞에 분야 구분 행이 없습니다.')
            raw_date = row[date_i] if date_i is not None and date_i < len(row) else None
            date = raw_date.strftime('%Y-%m-%d') if isinstance(raw_date, datetime) else cell(row, date_i)
            item = dict(title=title, publisher=cell(row,pub_i), url=cell(row,url_i),
                        summary=cell(row,sum_i), page=cell(row,page_i),
                        author=cell(row,author_i), date=date, row=number)
            result[current].setdefault(group_key(title), []).append(item)
            count += 1
        return result, count
    finally:
        wb.close()


def validate_url(url):
    url = url.strip()
    parsed = urlparse(url)
    if parsed.scheme not in ('http', 'https') or not parsed.hostname or parsed.username or parsed.password:
        raise ValueError('공개 HTTP(S) 기사 링크가 아닙니다.')
    if parsed.port not in (None, 80, 443):
        raise ValueError('일반 웹 포트가 아닙니다.')
    host = parsed.hostname
    if host.lower() == 'localhost' or host.lower().endswith(('.localhost','.local','.internal','.lan')):
        raise ValueError('로컬 주소는 사용할 수 없습니다.')
    try:
        addresses = socket.getaddrinfo(host, parsed.port or (443 if parsed.scheme == 'https' else 80))
        if not addresses or any(not ipaddress.ip_address(info[4][0]).is_global for info in addresses):
            raise ValueError('사설망 또는 로컬 주소는 사용할 수 없습니다.')
    except socket.gaierror as exc:
        # Proxy servers can resolve public hosts that a corporate PC cannot resolve locally.
        if not requests.utils.get_environ_proxies(url):
            raise ValueError('기사 주소를 확인할 수 없습니다.') from exc
    return url


def download_article(url):
    # Use OS proxy / certificate settings. Browser and EXE may still differ.
    session = requests.Session()
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'ko-KR,ko;q=0.9,en;q=0.7',
    }
    try:
        for _ in range(6):
            validate_url(url)
            with session.get(url, headers=headers, timeout=(7, 18), allow_redirects=False, stream=True) as response:
                if response.status_code in (301, 302, 303, 307, 308):
                    location = response.headers.get('Location')
                    if not location:
                        raise ValueError('이동할 페이지 주소가 없는 리디렉션입니다.')
                    url = urljoin(url, location)
                    continue
                response.raise_for_status()
                content_type = response.headers.get('Content-Type', '').lower()
                if content_type and not any(t in content_type for t in ('html', 'xhtml')):
                    raise ValueError('HTML 기사 페이지가 아닙니다.')
                pieces, total = [], 0
                for part in response.iter_content(chunk_size=16384):
                    total += len(part)
                    if total > MAX_BYTES:
                        raise ValueError('기사 페이지가 너무 큽니다.')
                    pieces.append(part)
                response._content = b''.join(pieces)
                response.encoding = response.apparent_encoding or response.encoding or 'utf-8'
                text = response.text
                if '<html' not in text[:3000].lower() and '<article' not in text.lower() and '<meta' not in text.lower():
                    raise ValueError('응답에서 기사 HTML을 찾지 못했습니다.')
                return text
        raise ValueError('리디렉션이 너무 많습니다.')
    finally:
        session.close()

def fetch_with_deadline(url, seconds=ARTICLE_DEADLINE):
    """Bound an article attempt even when DNS/proxy/streaming never returns.

    A timed-out daemon thread cannot be forcibly stopped. It may finish later,
    but it cannot update the UI or block program exit.
    """
    answers = queue.Queue(maxsize=1)
    def attempt():
        try:
            page = download_article(url)
            answers.put((True, extract_article(page)))
        except Exception as exc:
            answers.put((False, exc))
    threading.Thread(target=attempt, daemon=True).start()
    try:
        ok, value = answers.get(timeout=seconds)
    except queue.Empty as exc:
        raise TimeoutError(f'{seconds}초 안에 응답하지 않았습니다(DNS·접속·본문 처리 포함).') from exc
    if not ok:
        raise value
    return value


def clean_text(value):
    return re.sub(r'\s+', ' ', value or '').strip()


def extract_article(html_text):
    soup = BeautifulSoup(html_text, 'html.parser')
    # Keep metadata before deleting unrelated layout elements.
    description = ''
    for attrs in ({'property':'og:description'}, {'name':'description'}, {'name':'twitter:description'}):
        meta = soup.find('meta', attrs=attrs)
        candidate = clean_text(meta.get('content')) if meta else ''
        if len(candidate) >= 60 and len(candidate) > len(description):
            description = candidate
    for tag in soup.select('script[type="application/ld+json"]'):
        try:
            structured = json.loads(tag.string or tag.get_text())
        except (ValueError, TypeError):
            continue
        def bodies(obj):
            if isinstance(obj, dict):
                value = obj.get('articleBody')
                if isinstance(value, str):
                    yield clean_text(value)
                for child in obj.values():
                    if isinstance(child, (dict, list)):
                        yield from bodies(child)
            elif isinstance(obj, list):
                for child in obj:
                    yield from bodies(child)
        matches = [v for v in bodies(structured) if len(v) >= 240]
        if matches:
            return max(matches, key=len), '본문 발췌'
    for tag in soup.select('script,style,nav,footer,header,aside,form,iframe,button, .ad, .advertisement, .related, .comments, [class*="advert"], [class*="share"]'):
        tag.decompose()
    selectors = ['[itemprop="articleBody"]','[data-article-body]','.article_body','.article-body',
                 '#articleBody','#article_body','#articleView','.article_view','.news_body','.news-body',
                 '#newsView','.news_view','#article_txt','.article_txt','[class*="article-content"]',
                 '[class*="articleContent"]', '.news_end', '#dic_area', '#articeBody',
                 '[class*="article_body"]', '[id*="articleBody"]', 'article']
    for selector in selectors:
        matches = soup.select(selector)
        if matches:
            candidates = [clean_text(node.get_text(' ', strip=True)) for node in matches]
            longest = max(candidates, key=len)
            if len(longest) >= 240:
                return longest, '본문 발췌'
    # Last resort: sufficiently long paragraphs only, not the whole page/navigation.
    for parent in soup.select('main, [role="main"]'):
        paragraphs = [clean_text(p.get_text(' ', strip=True)) for p in parent.select('p')]
        paragraphs = [p for p in paragraphs if len(p) >= 35]
        joined = ' '.join(paragraphs)
        if len(joined) >= 300:
            return joined, '본문 발췌'
    if description:
        return description, '공개 설명문'
    raise ValueError('기사 본문과 공개 설명문을 찾지 못했습니다.')

def excerpt(text, source):
    if source == '공개 설명문':
        return text[:450]
    sentences = [s.strip() for s in re.split(r'(?<=[.!?。])\s+|(?<=다\.)\s*|\n+', text) if len(s.strip()) >= 25]
    if len(sentences) < 2:
        # Do not claim a semantic summary when the site has no reliable sentence boundaries.
        return text[:450].rstrip() + ('…' if len(text)>450 else '')
    words = re.findall(r'[가-힣]{2,}|[A-Za-z]{3,}', text.lower())
    stop = {'있다','했다','것으로','위해','대한','기자','때문','이번','지난','관련','등의','그리고','article','news'}
    counts = Counter(w for w in words if w not in stop)
    ranked = []
    for i, sentence in enumerate(sentences[:45]):
        if len(sentence)>280:
            continue
        tokens = set(re.findall(r'[가-힣]{2,}|[A-Za-z]{3,}', sentence.lower()))
        score = sum(min(counts[w], 7) for w in tokens) / max(len(tokens)**0.5, 1) + 3/(i+1)
        ranked.append((score,i,sentence))
    if not ranked:
        return text[:450].rstrip() + '…'
    chosen = sorted(sorted(ranked, reverse=True)[:3], key=lambda x:x[1])
    return ' '.join(s for _,_,s in chosen)[:700]




REPORT_STYLE = """
:root {--primary:#1a365d;--accent:#2b6cb0;--bg:#f7fafc;--border:#e2e8f0;--text:#2d3748;--muted:#718096}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font:13px/1.6 'Malgun Gothic',-apple-system,BlinkMacSystemFont,sans-serif}
.container{max-width:960px;margin:0 auto;padding:20px 16px 50px}
header{background:var(--primary);color:white;border-radius:8px;padding:20px;margin-bottom:18px}
h1{font-size:23px;margin:0}header p{margin:3px 0 0;font-size:12px;color:#dbeafe}
.toolbar{display:flex;align-items:center;gap:12px;background:white;border:1px solid var(--border);border-radius:8px;padding:12px 16px;margin-bottom:16px}
.toolbar p{margin:0;flex:1;color:#52647c}button{font:inherit;cursor:pointer;border:1px solid #b9c8da;border-radius:5px;background:white;color:var(--primary);padding:6px 11px}button:hover{background:#eff6ff}
.category{margin:18px 0}.category-header{display:flex;align-items:center;gap:10px;border-bottom:2px solid var(--primary);margin:0 0 12px;padding:4px 0}
h2{font-size:17px;margin:0;color:var(--primary)}.badge{font-size:11px;color:#2563a4}
.theme-group{border:1px solid var(--border);background:#fff;border-radius:9px;margin:10px 0;overflow:hidden}
.theme-title{font-weight:700;background:#edf3f9;color:var(--primary);padding:10px 15px}
.group-media{font-weight:400;font-size:11px;color:#62748c;margin-left:10px}
.article{padding:11px 15px 14px;border-top:1px solid var(--border)}
.article-title{font-weight:700;margin:0 0 5px}.article-title a{color:#172e4b;text-decoration:none}.article-title a:hover{text-decoration:underline}
.article-meta{display:flex;gap:8px;flex-wrap:wrap;color:#607494;font-size:11px;margin-bottom:8px}.article-meta .media{color:#235e9d;font-weight:700}
.summary{background:#f7fafc;border-left:3px solid #2b6cb0;border-radius:5px;padding:10px 12px;white-space:pre-wrap;overflow-wrap:anywhere;font-size:12px;line-height:1.7}
.edit-controls{margin-top:5px}.edit-controls>summary{cursor:pointer;color:#1f5b9b;font-size:12px}.edit{padding:10px 0}.edit-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px}
.edit label{display:block;color:#607494;font-size:11px}.edit input,.edit textarea{display:block;width:100%;padding:7px;border:1px solid #becddd;border-radius:4px;color:#203654;font:13px 'Malgun Gothic',sans-serif}
.edit textarea{min-height:100px;resize:vertical;margin-top:3px}.edit .summary-label{margin:9px 0}.edit button{margin-top:5px}
.empty{color:var(--muted);font-size:12px}.note{color:#607494;font-size:11px;margin:14px 0}footer{color:#607494;font-size:11px;margin-top:20px}
@media(max-width:600px){.toolbar{flex-wrap:wrap}.edit-grid{grid-template-columns:1fr}.group-media{display:block;margin:2px 0 0}}
@media print{body{background:white}.toolbar,.edit-controls{display:none!important}.theme-group{break-inside:avoid}header{print-color-adjust:exact}}
"""


def safe_report_url(value):
    """Only HTTP(S) URLs become clickable; content is escaped at the call site."""
    try:
        p = urlparse((value or '').strip())
        if p.scheme.lower() in ('https', 'http') and p.hostname and not p.username and not p.password:
            return p.geturl()
    except (ValueError, AttributeError):
        pass
    return None


REPORT_SCRIPT = r"""

const XLSX_STYLES = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac x16r2 xr\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:x16r2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/02/main\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\"><fonts count=\"6\" x14ac:knownFonts=\"1\"><font><sz val=\"11\"/><color theme=\"1\"/><name val=\"맑은 고딕\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><b/><sz val=\"22\"/><name val=\"맑은 고딕\"/><family val=\"3\"/><charset val=\"129\"/></font><font><sz val=\"10\"/><name val=\"맑은 고딕\"/><family val=\"3\"/><charset val=\"129\"/></font><font><b/><sz val=\"11\"/><name val=\"맑은 고딕\"/><family val=\"3\"/><charset val=\"129\"/></font><font><u/><sz val=\"10\"/><color rgb=\"FF0563C1\"/><name val=\"맑은 고딕\"/><family val=\"3\"/><charset val=\"129\"/></font><font><sz val=\"8\"/><name val=\"맑은 고딕\"/><family val=\"3\"/><charset val=\"129\"/><scheme val=\"minor\"/></font></fonts><fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFBFBFBF\"/></patternFill></fill></fills><borders count=\"2\"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style=\"thin\"><color rgb=\"FF7F7F7F\"/></left><right style=\"thin\"><color rgb=\"FF7F7F7F\"/></right><top style=\"thin\"><color rgb=\"FF7F7F7F\"/></top><bottom style=\"thin\"><color rgb=\"FF7F7F7F\"/></bottom><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"8\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"3\" fillId=\"2\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf><xf numFmtId=\"0\" fontId=\"3\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\"/></xf><xf numFmtId=\"0\" fontId=\"4\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment vertical=\"center\" wrapText=\"1\"/></xf><xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf><xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf><xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyAlignment=\"1\"><alignment horizontal=\"right\" vertical=\"center\"/></xf><xf numFmtId=\"0\" fontId=\"3\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\"/></xf></cellXfs><cellStyles count=\"1\"><cellStyle name=\"표준\" xfId=\"0\" builtinId=\"0\"/></cellStyles><dxfs count=\"0\"/><tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/><extLst><ext uri=\"{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:slicerStyles defaultSlicerStyle=\"SlicerStyleLight1\"/></ext><ext uri=\"{9260A510-F301-46a8-8635-F512D64BE5F5}\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"><x15:timelineStyles defaultTimelineStyle=\"TimeSlicerStyleLight1\"/></ext></extLst></styleSheet>";
const xml = value => String(value == null ? '' : value).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
function safeLink(value) {
  try {
    const u = new URL(value);
    return ['http:', 'https:'].includes(u.protocol) && !u.username && !u.password ? u.href : '';
  } catch (_) { return ''; }
}
function zipFiles(files) {
  const enc = new TextEncoder();
  const crcTable = Array.from({length:256}, (_, n) => {
    for (let k=0; k<8; k++) n = n & 1 ? 0xedb88320 ^ (n >>> 1) : n >>> 1;
    return n >>> 0;
  });
  const crc32 = data => {
    let c = 0xffffffff;
    for (const b of data) c = crcTable[(c ^ b) & 255] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const parts = [], directory = [];
  let offset = 0, centralSize = 0;
  for (const [name, content] of Object.entries(files)) {
    const path = enc.encode(name), data = enc.encode(content), crc = crc32(data);
    const local = new Uint8Array(30 + path.length), lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true); lv.setUint16(4, 20, true);
    lv.setUint16(6, 0x800, true); lv.setUint16(12, 33, true);
    lv.setUint32(14, crc, true); lv.setUint32(18, data.length, true); lv.setUint32(22, data.length, true);
    lv.setUint16(26, path.length, true); local.set(path, 30);
    const central = new Uint8Array(46 + path.length), cv = new DataView(central.buffer);
    cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true);
    cv.setUint16(8, 0x800, true); cv.setUint16(14, 33, true);
    cv.setUint32(16, crc, true); cv.setUint32(20, data.length, true); cv.setUint32(24, data.length, true);
    cv.setUint16(28, path.length, true); cv.setUint32(42, offset, true); central.set(path, 46);
    parts.push(local, data); directory.push(central);
    offset += local.length + data.length; centralSize += central.length;
  }
  const end = new Uint8Array(22), ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, directory.length, true); ev.setUint16(10, directory.length, true);
  ev.setUint32(12, centralSize, true); ev.setUint32(16, offset, true);
  return new Blob([...parts, ...directory, end], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function makeIssueXlsx(groups, dateValue) {
  const date = /^\d{4}-\d{2}-\d{2}$/.test(dateValue) ? new Date(dateValue+'T12:00:00') : new Date();
  const dateText = date.getFullYear()+'. '+(date.getMonth()+1)+'. '+date.getDate()+'. ('+'일월화수목금토'[date.getDay()]+')';
  const cell = (ref, value, style) => '<c r="'+ref+'" s="'+style+'" t="inlineStr"><is><t xml:space="preserve">'+xml(value)+'</t></is></c>';
  const row = (n, cells, height) => '<row r="'+n+'"'+(height == null ? '' : ' ht="'+height+'" customHeight="1"')+'>'+cells+'</row>';
  const rows = [row(1, cell('A1','중앙부처 오늘의 이슈',5),44.1),row(2,cell('A2',dateText,6),null),row(3,cell('A3','부처명',1)+cell('B3','제목',1)+cell('C3','매체명',1),24)];
  const merges = ['A1:C1','A2:C2'], links = [], rels = [];
  let n = 4;
  for (const group of groups) {
    const start = n;
    for (const item of group.items) {
      rows.push(row(n,cell('A'+n,n===start?group.name:'',7)+cell('B'+n,item.title,3)+cell('C'+n,item.publisher,4),33.95));
      const url = safeLink(item.url);
      if (url) {
        const id = 'rId'+(rels.length+1);
        links.push('<hyperlink ref="B'+n+'" r:id="'+id+'"/>');
        rels.push('<Relationship Id="'+id+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="'+xml(url)+'" TargetMode="External"/>');
      }
      n++;
    }
    if (n-start>1) merges.push('A'+start+':A'+(n-1));
  }
  const pre = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const files = {
    '[Content_Types].xml':pre+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
    '_rels/.rels':pre+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml':pre+'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="오늘의 이슈" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels':pre+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    'xl/styles.xml':XLSX_STYLES,
    'xl/worksheets/sheet1.xml':pre+'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:C'+(n-1)+'"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="16.5"/><cols><col min="1" max="1" width="18" customWidth="1"/><col min="2" max="2" width="80" customWidth="1"/><col min="3" max="3" width="14" customWidth="1"/></cols><sheetData>'+rows.join('')+'</sheetData><mergeCells count="'+merges.length+'">'+merges.map(m=>'<mergeCell ref="'+m+'"/>').join('')+'</mergeCells>'+(links.length?'<hyperlinks>'+links.join('')+'</hyperlinks>':'')+'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup orientation="portrait"/></worksheet>',
    'xl/worksheets/_rels/sheet1.xml.rels':pre+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+rels.join('')+'</Relationships>'
  };
  return zipFiles(files);
}


(function () {
  'use strict';
  const items = JSON.parse(document.getElementById('report-state').textContent);
  const cards = Array.from(document.querySelectorAll('[data-card]'));
  const main = document.querySelector('main');
  const status = message => { document.getElementById('save-status').textContent = message; };
  const sections = () => Array.from(document.querySelectorAll('section.category'));
  const nameOf = section => section.querySelector('h2').textContent;
  const dataOf = card => items[Number(card.dataset.card)];
  function button(label, action, parent) {
    const b = document.createElement('button'); b.type = 'button'; b.textContent = label;
    b.addEventListener('click', action); parent.appendChild(b); return b;
  }
  function download(blob, filename) {
    const url = URL.createObjectURL(blob), a = document.createElement('a');
    a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    status('다운로드를 요청했습니다. 브라우저 다운로드 목록을 확인하세요.');
  }
  function apply(card) {
    const d = dataOf(card);
    card.querySelectorAll('[data-field]').forEach(input => { d[input.dataset.field] = input.value.trim(); });
    const title = card.querySelector('.article-title'); title.textContent = '';
    const url = safeLink(d.url);
    if (url) {
      const a = document.createElement('a'); a.href=url; a.target='_blank'; a.rel='noopener noreferrer'; a.textContent=d.title; title.appendChild(a);
    } else { title.textContent=d.title; }
    const heading = card.closest('.theme-group').querySelector('.theme-title');
    if (heading.firstChild && heading.firstChild.nodeType===3) heading.firstChild.textContent=d.title;
    card.querySelector('.media').textContent = d.publisher;
    ['page','author','date'].forEach(field => {
      const span = card.querySelector('[data-meta="'+field+'"]'); span.textContent=d[field]; span.hidden=!d[field];
    });
    card.querySelector('[data-summary]').textContent=d.summary || '요약 미입력 — 내용을 확인한 후 요약을 입력하세요.';
  }
  function refresh() {
    const all=sections();
    all.forEach((section,index) => {
      const count=section.querySelectorAll('[data-card]').length;
      section.querySelector('.badge').textContent=count+'개 대표 기사';
      section.querySelectorAll('.empty').forEach(e=>e.remove());
      section.querySelectorAll('[data-card]').forEach(card => { dataOf(card).category=nameOf(section); });
      section.querySelectorAll('[data-move-select]').forEach(select => {
        select.replaceChildren();
        all.forEach((target,i) => { const o=document.createElement('option'); o.value=String(i); o.textContent=nameOf(target); select.appendChild(o); });
        select.value=String(index);
      });
    });
  }
  function setupSection(section) {
    const controls=document.createElement('div'); controls.className='dynamic-tools';
    const header=section.querySelector('.category-header'); header.appendChild(controls);
    button('이름 변경', () => {
      const name=prompt('목차 이름을 입력하세요.',nameOf(section));
      if (!name || !name.trim()) return;
      if (sections().some(s=>s!==section && nameOf(s)===name.trim())) { alert('같은 이름의 목차가 있습니다.'); return; }
      section.querySelector('h2').textContent=name.trim(); refresh();
    }, controls);
    button('목차 위로', () => { const prev=section.previousElementSibling; if (prev && prev.matches('section.category')) { main.insertBefore(section,prev); refresh(); } },controls);
    button('목차 아래로', () => { const next=section.nextElementSibling; if (next && next.matches('section.category')) { main.insertBefore(next,section); refresh(); } },controls);
    button('빈 목차 삭제', () => { if (section.querySelector('[data-card]')) { alert('기사를 다른 목차로 먼저 옮겨주세요.'); return; } section.remove(); refresh(); },controls);
  }
  cards.forEach(card => {
    const d=dataOf(card);
    // Existing reports store the article metadata in HTML, not the JSON state.
    if (d.title===undefined) d.title=card.querySelector('.article-title').textContent.trim();
    if (d.publisher===undefined) d.publisher=card.querySelector('.media').textContent.trim();
    if (d.url===undefined) d.url=card.querySelector('.article-title a')?.getAttribute('href') || '';
    const grid=card.querySelector('.edit-grid');
    [['title','기사 제목'],['publisher','매체명'],['url','원문 URL']].forEach(([field,label]) => {
      if (grid.querySelector('[data-field="'+field+'"]')) return;
      const l=document.createElement('label'), input=document.createElement('input');
      l.textContent=label; input.dataset.field=field; input.value=d[field] || ''; l.appendChild(input); grid.appendChild(l);
    });
    card.querySelector('[data-apply]').addEventListener('click', () => { apply(card); card.querySelector('details').open=false; status('수정 사항을 반영했습니다. 파일에 남기려면 HTML을 저장하세요.'); });
    const controls=document.createElement('div'); controls.className='dynamic-tools'; card.appendChild(controls);
    button('기사 위로', () => { const group=card.closest('.theme-group'), prev=group.previousElementSibling; if(prev && prev.matches('.theme-group')) group.parentNode.insertBefore(group,prev); },controls);
    button('기사 아래로', () => { const group=card.closest('.theme-group'), next=group.nextElementSibling; if(next && next.matches('.theme-group')) group.parentNode.insertBefore(next,group); },controls);
    const select=document.createElement('select'); select.dataset.moveSelect=''; select.setAttribute('aria-label','이동할 목차'); controls.appendChild(select);
    button('목차로 이동', () => { const target=sections()[Number(select.value)]; if(target) target.appendChild(card.closest('.theme-group')); refresh(); },controls);
  });
  sections().forEach(setupSection);
  const toolbar=document.querySelector('.toolbar');
  button('목차 추가', () => {
    const name=prompt('새 목차 이름을 입력하세요.'); if(!name || !name.trim()) return;
    if(sections().some(s=>nameOf(s)===name.trim())) { alert('같은 이름의 목차가 있습니다.'); return; }
    const section=document.createElement('section'); section.className='category';
    const header=document.createElement('div'); header.className='category-header';
    const h=document.createElement('h2'); h.textContent=name.trim();
    const badge=document.createElement('span'); badge.className='badge'; header.append(h,badge); section.appendChild(header);
    main.insertBefore(section,main.querySelector('footer')); setupSection(section); refresh();
  },toolbar).classList.add('runtime-button');
  button('엑셀로 내보내기', () => {
    try {
      cards.forEach(apply); refresh();
      const groups=sections().map(section=>({name:nameOf(section),items:Array.from(section.querySelectorAll('[data-card]')).map(dataOf)}));
      const date=document.getElementById('report-date').value;
      if(!date) { alert('보고서 날짜를 입력하세요.'); return; }
      download(makeIssueXlsx(groups,date),'중앙부처_오늘의이슈_'+date+'.xlsx');
    } catch (e) { status('엑셀 저장 실패: '+e.message); alert('엑셀 저장 실패: '+e.message); }
  },toolbar).classList.add('runtime-button');
  document.getElementById('save-html').addEventListener('click', () => {
    try {
      cards.forEach(apply); refresh();
      const clone=document.documentElement.cloneNode(true);
      clone.querySelector('#report-state').textContent=JSON.stringify(items).replace(/</g,'\\u003c').replace(/>/g,'\\u003e').replace(/&/g,'\\u0026');
      // Serialize every live form value; card IDs stay stable after DOM moves.
      document.querySelectorAll('[data-card]').forEach(card => {
        const target=clone.querySelector('[data-card="'+card.dataset.card+'"]');
        card.querySelectorAll('[data-field]').forEach(input => {
          const copy=target.querySelector('[data-field="'+input.dataset.field+'"]');
          if(input.tagName==='TEXTAREA') copy.textContent=input.value; else copy.setAttribute('value',input.value);
        });
        target.querySelector('details').removeAttribute('open');
      });
      clone.querySelector('#report-date').setAttribute('value',document.getElementById('report-date').value);
      clone.querySelectorAll('.dynamic-tools,.runtime-button').forEach(e=>e.remove());
      clone.querySelector('#save-status').textContent='저장된 보고서입니다. 편집 후 다시 저장할 수 있습니다.';
      download(new Blob(['<!DOCTYPE html>\n'+clone.outerHTML],{type:'text/html;charset=utf-8'}),'중앙부처동향_'+document.getElementById('report-date').value+'_편집본.html');
    } catch(e) { status('HTML 저장 실패: '+e.message); alert('HTML 저장 실패: '+e.message); }
  });
  refresh();
}());

"""

def report_html(report_groups, summaries, generated_at=None):
    """Offline editable report in the supplied grouped article-card style."""
    esc = lambda value: html.escape(str(value or ''), quote=True)
    now = generated_at or datetime.now().strftime('%Y-%m-%d %H:%M')
    categories = list(report_groups)
    count = sum(len(articles) for cat in categories for articles in report_groups.get(cat, {}).values())
    cards = sum(len(report_groups.get(cat, {})) for cat in categories)
    pieces = ['<!DOCTYPE html>', '<html lang="ko"><head><meta charset="UTF-8">',
              '<meta name="viewport" content="width=device-width, initial-scale=1">',
              '<title>중앙부처동향 · 기사 보고서</title><style>', REPORT_STYLE + "\n.toolbar,.category-header{flex-wrap:wrap}.dynamic-tools{display:flex;flex-wrap:wrap;gap:5px;margin:8px 0}select,input[type=date]{font:inherit;padding:5px;border:1px solid #b9c8da}@media print{.dynamic-tools{display:none!important}}",
              '</style></head><body><main class="container">',
              '<header><h1>중앙부처동향</h1><p>생성: '+esc(now)+' · 전체 '+str(count)+'건 · '+str(cards)+'개 묶음</p></header>',
              '<div class="toolbar"><label>보고서 날짜 <input id="report-date" type="date" value="'+esc(str(now)[:10])+'"></label><p id="save-status">「내용 수정」 후 「수정한 HTML 저장」을 누르면 새 HTML 파일을 내려받습니다.</p><button type="button" id="save-html">수정한 HTML 저장</button></div>',
              '<p class="note">본문 발췌는 AI 요약이 아닙니다. 원문과 대조해 주세요. 지면·기자·날짜는 엑셀에 해당 열이 있을 때만 표시됩니다.</p>']
    data = []
    for cat in categories:
        groups = report_groups.get(cat, {})
        pieces.extend(['<section class="category"><div class="category-header"><h2>',esc(cat),
                       '</h2><span class="badge">',str(sum(len(v) for v in groups.values())),'건</span></div>'])
        if not groups:
            pieces.append('<p class="empty">해당 기사 없음</p>')
        for key, articles in groups.items():
            first = articles[0]
            others = sorted(articles[1:], key=lambda a:(media_rank(a.get('publisher','')), a.get('row', 0)))
            others_names = []
            seen = {media_key(first.get('publisher',''))}
            for a in others:
                name = a.get('publisher','')
                if name and media_key(name) not in seen:
                    others_names.append(name)
                    seen.add(media_key(name))
            pieces.extend(['<div class="theme-group"><div class="theme-title">',esc(first['title'])])
            if others_names:
                pieces.extend(['<span class="group-media">함께 보도한 매체: ',esc(' · '.join(others_names)),'</span>'])
            pieces.extend(['</div><article class="article" data-card="',str(len(data)),'"><div class="article-title">'])
            link = safe_report_url(first.get('url'))
            if link:
                pieces.extend(['<a href="',esc(link),'" target="_blank" rel="noopener noreferrer">',esc(first['title']),'</a>'])
            else:
                pieces.append(esc(first['title']))
            summary = summaries.get((cat,key), first.get('summary','')) or ''
            if summary == '본문 발췌 전입니다. 기사 원문을 확인하거나 아래 버튼을 누르세요.':
                summary = '[미발췌] 아직 기사를 읽지 않았습니다.'
            data.append({field: str(first.get(field,'') or '') for field in ('page','author','date')})
            data[-1]['summary'] = summary
            pieces.append('</div><div class="article-meta"><span class="media">'+esc(first.get('publisher',''))+'</span>')
            for field in ('page','author','date'):
                value = data[-1][field]
                pieces.extend(['<span data-meta="',field,'"',('' if value else ' hidden'),'>'+esc(value)+'</span>'])
            pieces.extend(['</div><div class="summary" data-summary>',esc(summary or '요약 미입력 — 내용을 확인한 후 요약을 입력하세요.'),
                           '</div><details class="edit-controls"><summary>내용 수정</summary><div class="edit"><div class="edit-grid">'])
            for field,label in (('page','지면'),('author','기자 이름'),('date','발행일')):
                pieces.extend(['<label>',label,'<input data-field="',field,'" value="',esc(data[-1][field]),'"></label>'])
            pieces.extend(['</div><label class="summary-label">요약<textarea data-field="summary">',esc(summary),
                           '</textarea></label><button type="button" data-apply>수정 반영</button></div></details></article></div>'])
        pieces.append('</section>')
    safe_json = json.dumps(data, ensure_ascii=False).replace('<', chr(92)+'u003c').replace('>', chr(92)+'u003e').replace('&', chr(92)+'u0026')
    pieces.extend(['<footer>편집 내용은 「수정한 HTML 저장」을 누르면 새 파일로 저장됩니다. 기존 파일은 변경되지 않습니다.</footer></main>',
                   '<script id="report-state" type="application/json">',safe_json,'</script><script>',REPORT_SCRIPT,'</script></body></html>'])
    return ''.join(pieces)

class ReportApp:
    def __init__(self, root):
        self.root = root
        root.title('중앙부처동향 — 기사 본문 발췌')
        root.geometry('930x720')
        self.events = queue.Queue()
        self.cards = []
        self.report_groups = None
        self.busy = False
        header = ttk.Frame(root, padding=12); header.pack(fill='x')
        ttk.Label(header,text='중앙부처동향',font=('맑은 고딕',18,'bold')).pack(side='left', padx=(0,18))
        self.open_button = ttk.Button(header,text='기사 엑셀 열기',command=self.open_excel)
        self.open_button.pack(side='left',padx=5)
        self.all_button = ttk.Button(header,text='모든 대표 기사 발췌',command=self.fetch_all)
        self.all_button.pack(side='left',padx=5)
        self.export_button = ttk.Button(header,text='HTML로 내보내기',command=self.export_html)
        self.export_button.pack(side='left',padx=5)
        self.status = ttk.Label(root,text='기사 엑셀 파일을 선택하세요. 발췌는 인터넷 연결이 필요합니다.',padding=(12,2))
        self.status.pack(fill='x')
        self.tabs = ttk.Notebook(root); self.tabs.pack(fill='both',expand=True,padx=12,pady=8)
        self.panes = {}
        for category in CATEGORIES:
            frame=ttk.Frame(self.tabs); self.tabs.add(frame,text=category)
            canvas=tk.Canvas(frame,highlightthickness=0)
            bar=ttk.Scrollbar(frame,orient='vertical',command=canvas.yview)
            body=ttk.Frame(canvas)
            body.bind('<Configure>',lambda e,c=canvas:c.configure(scrollregion=c.bbox('all')))
            window=canvas.create_window((0,0),window=body,anchor='nw')
            canvas.bind('<Configure>',lambda e,c=canvas,w=window:c.itemconfigure(w,width=e.width))
            canvas.configure(yscrollcommand=bar.set)
            canvas.pack(side='left',fill='both',expand=True)
            bar.pack(side='right',fill='y')
            self.panes[category]=body
        root.after(100,self.poll)

    def open_excel(self):
        if self.busy:
            messagebox.showinfo('처리 중', '새 엑셀은 현재 기사 처리가 끝난 후 열어 주세요.')
            return
        path = filedialog.askopenfilename(title='기사 엑셀 선택',filetypes=[('Excel 파일','*.xlsx')])
        if not path:
            return
        try:
            report,count=read_report(path)
        except Exception as exc:
            messagebox.showerror('파일 오류',str(exc)); return
        for body in self.panes.values():
            for child in body.winfo_children(): child.destroy()
        self.cards=[]
        self.report_groups = report
        self.export_button.configure(state='normal')
        for cat in CATEGORIES:
            groups=report[cat]
            if not groups:
                ttk.Label(self.panes[cat],text='해당 기사 없음',padding=15).pack(anchor='w')
            for articles in groups.values():
                first=articles[0]
                frame=ttk.LabelFrame(self.panes[cat],text=f"{first['title']}   [{first['publisher']}]",padding=10)
                frame.pack(fill='x',padx=5,pady=6)
                source_url=next((a['url'] for a in articles if a['url']), '')
                if source_url:
                    ttk.Button(frame,text='원문 열기',command=lambda u=source_url:webbrowser.open(u)).pack(anchor='w')
                others=sorted(articles[1:],key=lambda a:(media_rank(a['publisher']),a['row']))
                names=[]; seen={media_key(first['publisher'])}
                for article in others:
                    name=article['publisher']; key=media_key(name)
                    if name and key not in seen: names.append(name);seen.add(key)
                if names: ttk.Label(frame,text='보도매체  '+' · '.join(names),foreground='#245c9b',wraplength=820).pack(anchor='w',pady=5)
                text=tk.Text(frame,height=4,wrap='word',font=('맑은 고딕',10),background='#f6f8fc')
                text.pack(fill='x',pady=5)
                initial = first['summary'] or '본문 발췌 전입니다. 기사 원문을 확인하거나 아래 버튼을 누르세요.'
                self.set_text(text,initial)
                btn=ttk.Button(frame,text='이 기사 발췌',command=lambda a=articles,t=text:self.fetch_one(a,t))
                btn.pack(anchor='w')
                self.cards.append((articles,text))
        self.status.configure(text=f'기사 {count}건을 {len(self.cards)}개 카드로 표시했습니다. 발췌 내용은 저장되지 않습니다.')

    def export_html(self):
        if self.report_groups is None:
            messagebox.showinfo('엑셀 선택', '먼저 기사 엑셀을 여세요.')
            return
        summaries={}
        # self.cards follows CATEGORIES / OrderedDict group order from read_report.
        card_iter=iter(self.cards)
        for category in CATEGORIES:
            for key in self.report_groups[category]:
                _,widget=next(card_iter)
                summaries[(category,key)]=widget.get('1.0','end-1c').strip()
        filename='중앙부처동향_'+datetime.now().strftime('%Y-%m-%d')+'.html'
        target=filedialog.asksaveasfilename(title='HTML 보고서 저장',defaultextension='.html',
                 initialfile=filename,filetypes=[('HTML 파일','*.html')])
        if not target:
            return
        try:
            Path(target).write_text(report_html(self.report_groups,summaries),encoding='utf-8')
        except Exception as exc:
            messagebox.showerror('저장 실패',str(exc))
            return
        self.status.configure(text=('발췌 진행 중 · 현재 결과를 저장했습니다: ' if self.busy else 'HTML 보고서를 저장했습니다: ')+target)
        messagebox.showinfo('저장 완료', ('진행 중 결과를 저장했습니다. 완료 후 다시 저장하면 최종 결과를 반영할 수 있습니다.\n' if self.busy else 'HTML 보고서를 저장했습니다. 파일을 브라우저에서 열 수 있습니다.\n')+target)

    @staticmethod
    def set_text(widget,value):
        widget.configure(state='normal'); widget.delete('1.0','end');widget.insert('1.0',value);widget.configure(state='disabled')

    def worker(self, tasks):
        try:
            for n, (articles, widget) in enumerate(tasks, 1):
                try:
                    failures = []
                    tried = set()
                    result = None
                    candidates = []
                    for article in articles:
                        url = (article.get('url') or '').strip()
                        if url and url not in tried:
                            tried.add(url)
                            candidates.append((article, url))
                    for index, (article, url) in enumerate(candidates[:MAX_LINKS_PER_CARD], 1):
                        publisher = article.get('publisher') or '매체 미상'
                        self.events.put(('status', f'{n}/{len(tasks)}번 기사 · {publisher} 링크 {index}/{min(len(candidates), MAX_LINKS_PER_CARD)} 확인 중 (최대 {ARTICLE_DEADLINE}초)…', None))
                        try:
                            text, source = fetch_with_deadline(url)
                            result = f'[{source} · {publisher}] {excerpt(text, source)}'
                            break
                        except Exception as exc:
                            failures.append(f'{publisher}: {type(exc).__name__}: {str(exc)[:110]}')
                    if result is None:
                        if not candidates:
                            result = '[요약 불가] 이 기사 묶음의 URL이 비어 있습니다. 엑셀의 기사 URL 열을 확인해 주세요.'
                        else:
                            suffix = f' / 다른 링크 {len(candidates)-MAX_LINKS_PER_CARD}개는 미시도' if len(candidates)>MAX_LINKS_PER_CARD else ''
                            result = '[요약 불가] ' + (' / '.join(failures)+suffix)[:650]
                except Exception as exc:
                    result = f'[요약 불가] 처리 오류: {type(exc).__name__}: {str(exc)[:180]}'
                self.events.put(('article', widget, result))
                self.events.put(('status', f'{n}/{len(tasks)}개 기사 처리 완료', None))
        finally:
            self.events.put(('done', None, None))

    def start(self, tasks):
        if self.busy or not tasks:
            return
        self.busy = True
        self.all_button.configure(state='disabled')
        self.open_button.configure(state='disabled')
        # HTML output is a snapshot and remains available during processing.
        self.export_button.configure(state='normal')
        self.status.configure(text=f'0/{len(tasks)}개 기사 처리 완료 · 첫 링크 확인 중…')
        threading.Thread(target=self.worker, args=(tasks,), daemon=True).start()

    def fetch_one(self, articles, widget):
        self.start([(articles, widget)])

    def fetch_all(self):
        self.start(self.cards.copy())

    def poll(self):
        while True:
            try:
                kind, a, b = self.events.get_nowait()
            except queue.Empty:
                break
            if kind == 'article':
                try:
                    self.set_text(a, b)
                except tk.TclError:
                    pass
            elif kind == 'status':
                self.status.configure(text=a)
            elif kind == 'done':
                self.busy = False
                self.open_button.configure(state='normal')
                self.all_button.configure(state='normal')
                self.export_button.configure(state='normal')
                self.status.configure(text='발췌 처리가 끝났습니다. 실패한 기사는 카드의 오류 문구를 확인해 주세요.')
        self.root.after(100, self.poll)


if __name__=='__main__':
    root=tk.Tk(); ReportApp(root);root.mainloop()
