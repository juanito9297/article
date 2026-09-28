# -*- coding: utf-8 -*-
"""Windows desktop clipping report. No AI service/API key is used."""
import ipaddress
import html
from datetime import datetime
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




REPORT_STYLE = '\n    * {\n      box-sizing: border-box;\n    }\n\n    body {\n      margin: 0;\n      background: #f3f6fb;\n      color: #102b50;\n      font-family: Arial, "Malgun Gothic", sans-serif;\n      font-size: 13px;\n    }\n\n    .page {\n      width: min(100% - 32px, 735px);\n      margin: 16px auto 40px;\n    }\n\n    .upload-card,\n    .article-card {\n      background: #fff;\n      border: 1px solid #d5e1f0;\n      border-radius: 11px;\n    }\n\n    .upload-card {\n      padding: 17px 16px;\n    }\n\n    .upload-card h2 {\n      margin: 0 0 7px;\n      font-size: 13px;\n    }\n\n    .description,\n    .status {\n      margin: 0;\n      color: #6380a7;\n      font-size: 11px;\n      line-height: 1.6;\n    }\n\n    .upload-row {\n      display: flex;\n      gap: 8px;\n      margin: 13px 0 10px;\n    }\n\n    .upload-row input {\n      min-width: 0;\n      flex: 1;\n      padding: 8px;\n      border: 1px solid #c5d6ed;\n      border-radius: 6px;\n      background: #f9fbfe;\n    }\n\n    .upload-row button {\n      flex: none;\n      padding: 0 13px;\n      border: 0;\n      border-radius: 6px;\n      background: #245c9b;\n      color: #fff;\n      font-weight: 700;\n      cursor: pointer;\n    }\n\n    .upload-row button:disabled {\n      opacity: 0.6;\n      cursor: wait;\n    }\n\n    .report {\n      padding-top: 25px;\n    }\n\n    .report h1 {\n      margin: 0 0 25px;\n      font-size: 23px;\n    }\n\n    .category {\n      margin-bottom: 25px;\n    }\n\n    .category h2 {\n      margin: 0 0 13px;\n      padding-bottom: 11px;\n      border-bottom: 1px solid #b6cbe5;\n      color: #173d73;\n      font-size: 16px;\n    }\n\n    .article-card {\n      margin-bottom: 12px;\n      padding: 17px 18px 16px;\n    }\n\n    .article-card h3 {\n      margin: 0 0 14px;\n      font-size: 16px;\n      line-height: 1.5;\n    }\n\n    .article-card h3 a {\n      color: #102b50;\n      text-decoration: none;\n    }\n\n    .article-card h3 a:hover {\n      text-decoration: underline;\n    }\n\n    .headline-media {\n      margin-left: 8px;\n      color: #6380a7;\n      font-size: 11px;\n      font-weight: 400;\n      white-space: nowrap;\n    }\n\n    .media-line {\n      margin: 0 0 14px;\n      font-size: 11px;\n      line-height: 1.6;\n    }\n\n    .media-line strong {\n      color: #245c9b;\n    }\n\n    .summary {\n      margin: 0;\n      padding: 14px;\n      border-radius: 6px;\n      background: #f6f8fc;\n      font-size: 12px;\n      line-height: 1.75;\n      white-space: pre-wrap;\n      overflow-wrap: anywhere;\n    }\n\n    .summary.empty,\n    .empty-category {\n      color: #6380a7;\n    }\n\n    .empty-category {\n      margin: 0 4px;\n      font-size: 12px;\n    }\n\n    @media (max-width: 480px) {\n      .upload-row {\n        flex-direction: column;\n      }\n\n      .upload-row button {\n        min-height: 36px;\n      }\n\n      .headline-media {\n        white-space: normal;\n      }\n    }\n  \n    .report-note {font-size:11px;color:#6380a7;line-height:1.6;margin:0 0 22px;}\n'


def safe_report_url(value):
    """Only public-looking HTTP(S) absolute URLs become clickable links; no network request here."""
    try:
        p = urlparse((value or '').strip())
        if p.scheme.lower() in ('https', 'http') and p.hostname and not p.username and not p.password:
            return p.geturl()
    except (ValueError, AttributeError):
        pass
    return None


def report_html(report_groups, summaries, generated_at=None):
    """Offline, self-contained read-only snapshot matching the original report layout.

    summaries maps (category, group_key) to the current on-screen text.
    """
    esc = lambda value: html.escape(str(value or ''), quote=True)
    now = generated_at or datetime.now().strftime('%Y-%m-%d %H:%M')
    pieces = ['<!DOCTYPE html>', '<html lang="ko">', '<head>', '<meta charset="UTF-8">',
              '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
              '<title>중앙부처동향</title>', '<style>', REPORT_STYLE, '</style>', '</head>',
              '<body>', '<main class="page">', '<section class="report">',
              '<h1>중앙부처동향</h1>',
              '<p class="report-note">생성: '+esc(now)+' · EXE에서 저장한 결과입니다. 본문 발췌는 AI 요약이 아닙니다. 원문과 대조해 주세요.</p>',
              '<div id="categories">']
    for category in CATEGORIES:
        pieces.extend(['<section class="category">', '<h2>'+esc(category)+'</h2>'])
        groups=report_groups.get(category, {})
        if not groups:
            pieces.append('<p class="empty-category">해당 기사 없음</p>')
        for key, articles in groups.items():
            first=articles[0]
            pieces.extend(['<article class="article-card">', '<h3>'])
            link=safe_report_url(first.get('url'))
            if link:
                pieces.append('<a href="'+esc(link)+'" target="_blank" rel="noopener noreferrer">'+esc(first['title'])+'</a>')
            else:
                pieces.append(esc(first['title']))
            if first.get('publisher'):
                pieces.append('<span class="headline-media">'+esc(first['publisher'])+'</span>')
            pieces.append('</h3>')
            seen={media_key(first.get('publisher',''))}
            names=[]
            for a in sorted(articles[1:], key=lambda a:(media_rank(a.get('publisher','')),a['row'])):
                name=a.get('publisher',''); norm=media_key(name)
                if name and norm not in seen:
                    names.append(name);seen.add(norm)
            if names:
                pieces.append('<p class="media-line"><strong>보도매체</strong>  '+esc(' · '.join(names))+'</p>')
            content=summaries.get((category,key), first.get('summary',''))
            if not content or content == '본문 발췌 전입니다. 기사 원문을 확인하거나 아래 버튼을 누르세요.':
                content='[미발췌] 아직 기사를 읽지 않았습니다.'
            pieces.append('<p class="summary'+(' empty' if content.startswith(('[요약 불가]','[미발췌]')) else '')+'">'+esc(content)+'</p>')
            pieces.append('</article>')
        pieces.append('</section>')
    pieces.extend(['</div>', '</section>', '</main>', '</body>', '</html>'])
    return '\n'.join(pieces)


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
        ttk.Button(header,text='기사 엑셀 열기',command=self.open_excel).pack(side='left',padx=5)
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

    def export_html(self):
        if self.busy:
            messagebox.showinfo('처리 중', '기사 발췌가 끝난 뒤 저장해 주세요.')
            return
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
        except OSError as exc:
            messagebox.showerror('저장 실패',str(exc))
            return
        self.status.configure(text='HTML 보고서를 저장했습니다: '+target)
        messagebox.showinfo('저장 완료', 'HTML 보고서를 저장했습니다. 파일을 브라우저에서 열 수 있습니다.\n'+target)

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
        self.busy=True;self.all_button.configure(state='disabled');self.export_button.configure(state='disabled')
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
            else: self.busy=False;self.all_button.configure(state='normal');self.export_button.configure(state='normal')
        self.root.after(100,self.poll)


if __name__=='__main__':
    root=tk.Tk(); ReportApp(root);root.mainloop()
