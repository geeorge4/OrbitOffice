"""KalibriOffice desktop. All documents stay on this computer."""
import json
import os
import sys
import subprocess
import tempfile
from pathlib import Path
from html.parser import HTMLParser

DATA = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'OrbitOffice'
KINDS = {'write': 'Kalibri Write', 'sheet': 'Kalibri Calc', 'slides': 'Kalibri Slides'}
VERSION = '1.2.0'

def read_settings():
    try:
        settings=json.loads((DATA/'settings.json').read_text('utf-8'))
        if not isinstance(settings,dict): settings={}
    except (OSError,ValueError): settings={}
    try: seconds=max(0,min(30,float(settings.get('splash_seconds',5))))
    except (TypeError,ValueError): seconds=5
    settings['splash_seconds']=seconds
    return settings

def write_json(path,value):
    path.parent.mkdir(parents=True,exist_ok=True)
    with tempfile.NamedTemporaryFile(mode='w',encoding='utf-8',dir=path.parent,delete=False,suffix='.tmp') as f:
        json.dump(value,f,ensure_ascii=False); tmp=Path(f.name)
    try: tmp.replace(path)
    finally: tmp.unlink(missing_ok=True)

class Paragraphs(HTMLParser):
    def __init__(self):
        super().__init__(); self.lines=[]; self.text=''
    def handle_data(self, data): self.text += data
    def handle_endtag(self, tag):
        if tag in ('p','div','h1','h2','h3','li','tr'):
            self.lines.append(self.text); self.text=''
    def handle_starttag(self, tag, attrs):
        if tag == 'br': self.text += '\n'

def export_file(path, state):
    kind=state['kind']; path=Path(path)
    if kind=='write':
        from docx import Document
        d=Document(); parser=Paragraphs(); parser.feed(state['html'])
        for text in parser.lines+[parser.text]:
            if text: d.add_paragraph(text)
        d.save(path)
    elif kind=='sheet':
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        w=Workbook(); s=w.active; s.title='Лист 1'
        for address,value in state['cells'].items():
            if not value: continue
            if not value.startswith('='):
                try: value=float(value) if '.' in value else int(value)
                except ValueError: pass
            s[address]=value
            s[address].font=Font(name='Calibri',size=11)
            s[address].number_format={'number':'0.00','percent':'0.00%','currency':'#,##0.00 "₽"'}.get(state.get('formats',{}).get(address),'General')
        for col in range(1,27): s.column_dimensions[chr(64+col)].width=16
        w.save(path)
    else:
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        p=Presentation(); p.slide_width=Inches(13.333); p.slide_height=Inches(7.5)
        for item in state['slides']:
            s=p.slides.add_slide(p.slide_layouts[6]); fill=s.background.fill; fill.solid(); fill.fore_color.rgb=RGBColor.from_string(item['color'].lstrip('#'))
            for text,y,size in [(item['title'],1.3,40),(item['body'],2.7,24)]:
                box=s.shapes.add_textbox(Inches(1),Inches(y), Inches(11), Inches(3)).text_frame
                box.word_wrap=True; box.text=text
                for para in box.paragraphs: para.font.size=Pt(size); para.font.color.rgb=RGBColor(255,255,255)
            if item.get('notes'): s.notes_slide.notes_text_frame.text=item['notes']
        p.save(path)

class Api:
    def __init__(self,kind='write'):
        self._window=None; self._kind=kind; self._maximized=False
    def minimize(self):
        self._window.minimize()
    def toggle_maximize(self):
        if self._maximized: self._window.restore()
        else: self._window.maximize()
        self._maximized=not self._maximized
    def config(self):
        return {'kind':self._kind,'settings':read_settings(),'version':VERSION}
    def settings(self,seconds):
        seconds=float(seconds)
        if not 0<=seconds<=30: raise ValueError('Выбери от 0 до 30 секунд.')
        settings=read_settings(); settings['splash_seconds']=seconds
        write_json(DATA/'settings.json',settings)
        return settings
    def open_editor(self,kind):
        if kind not in KINDS: raise ValueError('Неизвестный редактор.')
        command=[sys.executable]
        if not getattr(sys,'frozen',False): command.append(str(Path(__file__).resolve()))
        subprocess.Popen(command+['--kind',kind],cwd=str(Path(sys.executable).parent))
        return True
    def close(self):
        self._window.destroy()
    def restore(self):
        try:
            state=json.loads((DATA/f'recovery-{self._kind}.json').read_text('utf-8'))
            return {self._kind:state}
        except (OSError,ValueError):
            try:
                old=json.loads((DATA/'recovery.json').read_text('utf-8'))
                return {self._kind:old[self._kind]} if self._kind in old else None
            except (OSError,ValueError,TypeError): return None
    def autosave(self,state):
        write_json(DATA/f'recovery-{self._kind}.json',state[self._kind])
        return True
    def save(self,state,native=False):
        import webview
        ext={'write':'docx','sheet':'xlsx','slides':'pptx'}[state['kind']] if native else 'orbit'
        paths=self._window.create_file_dialog(webview.FileDialog.SAVE,save_filename=state['name']+'.'+ext,file_types=(f'Документ (*.{ext})',))
        if not paths: return None
        path=Path(paths if isinstance(paths,str) else paths[0])
        if path.suffix.lower()!='.'+ext: path=Path(str(path)+'.'+ext)
        if native: export_file(path,state)
        else: path.write_text(json.dumps(state,ensure_ascii=False),'utf-8')
        return str(path)
    def open(self):
        import webview
        paths=self._window.create_file_dialog(webview.FileDialog.OPEN,file_types=('Документы (*.orbit;*.docx;*.xlsx;*.pptx;*.txt)',))
        if not paths:return None
        path=Path(paths[0]); ext=path.suffix.lower()
        if ext=='.orbit': return json.loads(path.read_text('utf-8'))
        from html import escape
        if ext in ('.docx','.txt'):
            if ext=='.docx':
                from docx import Document
                texts=[p.text for p in Document(path).paragraphs]
            else: texts=path.read_text('utf-8-sig').splitlines()
            return dict(kind='write',name=path.stem,html=''.join('<p>'+escape(t)+'</p>' for t in texts))
        if ext=='.xlsx':
            from openpyxl import load_workbook
            wb=load_workbook(path); s=wb.active
            if s.max_row>100 or s.max_column>26: raise ValueError('В первой версии поддерживается до 100 строк и 26 столбцов. Исходный файл не изменён.')
            formats={}
            for row in s:
                for c in row:
                    if '%' in c.number_format: formats[c.coordinate]='percent'
                    elif '₽' in c.number_format: formats[c.coordinate]='currency'
                    elif c.number_format=='0.00': formats[c.coordinate]='number'
            return dict(kind='sheet',name=path.stem,cells={c.coordinate:str(c.value) for row in s for c in row if c.value is not None},formats=formats)
        if ext=='.pptx':
            from pptx import Presentation
            slides=[]
            for s in Presentation(path).slides:
                texts=[sh.text for sh in s.shapes if sh.has_text_frame]
                notes=s.notes_slide.notes_text_frame.text if s.has_notes_slide else ''
                slides.append(dict(title=texts[0] if texts else '',body='\n'.join(texts[1:]),color='#234c75',notes=notes))
            return dict(kind='slides',name=path.stem,slides=slides or [dict(title='',body='',color='#234c75')])
        raise ValueError('Этот формат пока не поддерживается.')

if __name__=='__main__':
    import argparse
    import webview
    import time
    parser=argparse.ArgumentParser(); parser.add_argument('--kind',choices=KINDS,default='write')
    kind=parser.parse_args().kind
    api=Api(kind); root=Path(getattr(sys,'_MEIPASS',Path(__file__).parent))
    # Feed the bundled page directly to WebView2. Navigating from an inline
    # splash to a file:// URL can mis-handle the local path and fragment.
    editor_html=(root/'index.html').read_text('utf-8')
    editor_html=editor_html.replace('</html>','<script>'+(root/'features.js').read_text('utf-8')+'</script></html>')
    editor_html=editor_html.replace('<script>','<script>window.ORBIT_KIND='+json.dumps(kind)+';',1)

    splash_html='''<!doctype html><meta charset="utf-8"><style>
    *{box-sizing:border-box}body{margin:0;background:#315bc6;color:white;font-family:Segoe UI,sans-serif;overflow:hidden;user-select:none}
    main{height:340px;padding:40px 45px;position:relative;animation:show .55s ease}
    .brand{font-size:14px;letter-spacing:1px;opacity:.75}.mark{font-size:54px;margin:21px 0 0;line-height:1.2}.title{font-size:35px;font-weight:300;letter-spacing:-1px;margin-top:2px}
    .bottom{position:absolute;bottom:34px;left:45px;right:45px;font-size:12px;color:#d6e1ff}.track{height:16px;margin-top:12px;overflow:hidden;position:relative}.dot{position:absolute;left:0;top:5px;width:5px;height:5px;background:white;border-radius:50%;opacity:0;animation:fly 2.2s infinite;animation-timing-function:linear}.dot:nth-child(2){animation-delay:.13s}.dot:nth-child(3){animation-delay:.26s}.dot:nth-child(4){animation-delay:.39s}.dot:nth-child(5){animation-delay:.52s}
    .version{position:absolute;right:0;top:0;opacity:.6}button{position:absolute;right:12px;top:8px;border:0;background:none;color:#d6e1ff;font-size:21px;cursor:pointer}button:hover{color:white}
    @keyframes show{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}@keyframes fly{0%{left:-3%;opacity:0}8%{opacity:1}25%{left:37%}65%{left:63%}86%{left:103%;opacity:1}87%,100%{left:103%;opacity:0}}
    </style><main><button onclick="window.pywebview.api.close()" aria-label="Закрыть">×</button><div class="brand pywebview-drag-region">KALIBRI</div><div class="mark pywebview-drag-region">◉</div><div class="title pywebview-drag-region">Office</div><div class="bottom"><span>Подготовка редактора…</span><span class="version">1.2</span><div class="track" role="progressbar" aria-label="Загрузка"><i class="dot"></i><i class="dot"></i><i class="dot"></i><i class="dot"></i><i class="dot"></i></div></div></main>'''
    color={'write':'#315bc6','sheet':'#187b57','slides':'#bb4c30'}[kind]
    splash_html=splash_html.replace('#315bc6',color).replace('>Office</div>',f'>{KINDS[kind].split()[1]}</div>')
    api._window=webview.create_window(KINDS[kind],html=splash_html,js_api=api,width=520,height=340,resizable=True,frameless=True,easy_drag=False,background_color=color)
    loading=[True]
    def reveal():
        if not loading[0]: return
        loading[0]=False
        time.sleep(read_settings()['splash_seconds'])
        api._window.resize(1320,860)
        api._window.load_html(editor_html)
    api._window.events.loaded += reveal
    webview.start()
