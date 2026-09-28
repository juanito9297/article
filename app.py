# -*- coding: utf-8 -*-
"""Windows desktop clipping report. No AI service/API key is used."""
import ipaddress
import queue
import re
import socket
import threading
import tkinter as tk
import webbrowser
from collections import Counter, OrderedDict
from tkinter import filedialog, messagebox, ttk
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

CATEGORIES = ("정치", "경제", "사회일반", "해외동향", "오피니언")
PRIORITY = ("조선일보", "중앙일보", "동아일보", "한겨레", "경향신문", "한국경제", "매일경제")
ALIASES = {"조선":"조선일보", "중앙":"중앙일보", "동아":"동아일보", "경향":"경향신문", "한경":"한국경제", "매경":"매일경제"}
MAX_BYTES = 2_000_000


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
            item = dict(title=title, publisher=cell(row,pub_i), url=cell(row,url_i), summary=cell(row,sum_i), row=number)
            result[current].setdefault(group_key(title), []).append(item)
            count += 1
        return result, count
    finally:
        wb.close()


def validate_url(url):
    parsed = urlparse(url.strip())
    if parsed.scheme not in ('http', 'https') or not parsed.hostname or parsed.username or parsed.password:
        raise ValueError('공개 HTTP(S) 기사 링크가 아닙니다.')
    if parsed.port not in (None, 80, 443):
        raise ValueError('일반 웹 포트가 아닙니다.')
    host = parsed.hostname
    try:
        addresses = socket.getaddrinfo(host, parsed.port or (443 if parsed.scheme == 'https' else 80))
        if not addresses or any(not ipaddress.ip_address(info[4][0]).is_global for info in addresses):
            raise ValueError('사설망 또는 로컬 주소는 사용할 수 없습니다.')
    except socket.gaierror as exc:
        raise ValueError('기사 주소를 확인할 수 없습니다.') from exc
    return url


def download_article(url):
    session = requests.Session()
    session.trust_env = False
    headers = {'User-Agent':'Mozilla/5.0 (compatible; ArticleExcerpt/1.0)', 'Accept':'text/html,application/xhtml+xml'}
    try:
        for _ in range(5):
            validate_url(url)
            with session.get(url, headers=headers, timeout=(5, 12), allow_redirects=False, stream=True) as response:
                if response.status_code in (301,302,303,307,308):
                    from urllib.parse import urljoin
                    url = urljoin(url, response.headers.get('Location', ''))
                    continue
                response.raise_for_status()
                content_type = response.headers.get('Content-Type','').lower()
                if 'html' not in content_type:
                    raise ValueError('HTML 기사 페이지가 아닙니다.')
                pieces, total = [], 0
                for part in response.iter_content(chunk_size=16384):
                    total += len(part)
                    if total > MAX_BYTES:
                        raise ValueError('기사 페이지가 너무 큽니다.')
                    pieces.append(part)
                response._content = b''.join(pieces)
                response.encoding = response.apparent_encoding or response.encoding or 'utf-8'
                return response.text
        raise ValueError('리디렉션이 너무 많습니다.')
    finally:
        session.close()


def clean_text(value):
    return re.sub(r'\s+', ' ', value or '').strip()


def extract_article(html):
    soup = BeautifulSoup(html, 'html.parser')
    for tag in soup.select('script,style,nav,footer,header,aside,form,iframe,button, .ad, .advertisement, .related, .comments, [class*="advert"], [class*="share"]'):
        tag.decompose()
    selectors = ['[itemprop="articleBody"]','[data-article-body]','.article_body','.article-body',
                 '#articleBody','#article_body','#articleView','.article_view','.news_body','.news-body',
                 '#newsView','.news_view','#article_txt','.article_txt','article']
    for selector in selectors:
        matches = soup.select(selector)
        if matches:
            candidates = [clean_text(node.get_text(' ', strip=True)) for node in matches]
            longest = max(candidates, key=len)
            if len(longest) >= 240:
                return longest, '본문 발췌'
    for attrs in ({'property':'og:description'}, {'name':'description'}, {'name':'twitter:description'}):
        meta = soup.find('meta', attrs=attrs)
        if meta and len(clean_text(meta.get('content'))) >= 60:
            return clean_text(meta.get('content')), '공개 설명문'
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


class ReportApp:
    def __init__(self, root):
        self.root = root
        root.title('중앙부처동향 — 기사 본문 발췌')
        root.geometry('930x720')
        self.events = queue.Queue()
        self.cards = []
        self.busy = False
        header = ttk.Frame(root, padding=12); header.pack(fill='x')
        ttk.Label(header,text='중앙부처동향',font=('맑은 고딕',18,'bold')).pack(side='left', padx=(0,18))
        ttk.Button(header,text='기사 엑셀 열기',command=self.open_excel).pack(side='left',padx=5)
        self.all_button = ttk.Button(header,text='모든 대표 기사 발췌',command=self.fetch_all)
        self.all_button.pack(side='left',padx=5)
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
        for cat in CATEGORIES:
            groups=report[cat]
            if not groups:
                ttk.Label(self.panes[cat],text='해당 기사 없음',padding=15).pack(anchor='w')
            for articles in groups.values():
                first=articles[0]
                frame=ttk.LabelFrame(self.panes[cat],text=f"{first['title']}   [{first['publisher']}]",padding=10)
                frame.pack(fill='x',padx=5,pady=6)
                if first['url']:
                    ttk.Button(frame,text='원문 열기',command=lambda u=first['url']:webbrowser.open(u)).pack(anchor='w')
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
                btn=ttk.Button(frame,text='이 기사 발췌',command=lambda a=first,t=text:self.fetch_one(a,t))
                btn.pack(anchor='w')
                self.cards.append((first,text))
        self.status.configure(text=f'기사 {count}건을 {len(self.cards)}개 카드로 표시했습니다. 발췌 내용은 저장되지 않습니다.')

    @staticmethod
    def set_text(widget,value):
        widget.configure(state='normal'); widget.delete('1.0','end');widget.insert('1.0',value);widget.configure(state='disabled')

    def worker(self, tasks):
        for n,(article,widget) in enumerate(tasks,1):
            try:
                if not article['url']: raise ValueError('기사 URL이 없습니다.')
                html=download_article(article['url'])
                text,source=extract_article(html)
                result=f'[{source}] {excerpt(text,source)}'
            except Exception as exc:
                result=f'[요약 불가] {str(exc)[:180]}'
            self.events.put(('article',widget,result))
            self.events.put(('status',f'{n}/{len(tasks)}개 기사 처리 완료'))
        self.events.put(('done',None,None))

    def start(self,tasks):
        if self.busy or not tasks: return
        self.busy=True;self.all_button.configure(state='disabled')
        self.status.configure(text=f'{len(tasks)}개 기사 확인 중…')
        threading.Thread(target=self.worker,args=(tasks,),daemon=True).start()

    def fetch_one(self,article,widget): self.start([(article,widget)])
    def fetch_all(self): self.start(self.cards.copy())

    def poll(self):
        while True:
            try: kind,a,b=self.events.get_nowait()
            except queue.Empty: break
            if kind=='article':
                try: self.set_text(a,b)
                except tk.TclError: pass
            elif kind=='status': self.status.configure(text=a)
            else: self.busy=False;self.all_button.configure(state='normal')
        self.root.after(100,self.poll)


if __name__=='__main__':
    root=tk.Tk(); ReportApp(root);root.mainloop()
