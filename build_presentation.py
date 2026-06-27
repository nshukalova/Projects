import zipfile, html, re
from pathlib import Path
from xml.etree import ElementTree as ET

ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
DOC=Path('Программа_тренинга_ИИ_для_ЭкоИндустрия.docx')
OUT=Path('Тренинг_ИИ_для_ЭкоИндустрия.pptx')
MD=Path('Тренинг_ИИ_для_ЭкоИндустрия.md')

with zipfile.ZipFile(DOC) as z:
    root=ET.fromstring(z.read('word/document.xml'))
paras=[]
for p in root.findall('.//w:body/w:p',ns):
    txt=''.join((t.text or '') for t in p.findall('.//w:t',ns)).strip()
    if txt:
        paras.append(txt)

def is_heading(t):
    return bool(re.match(r'^(\d+\. |Модуль \d+\.|Тема \d+\.|Практическая часть|Теоретическая основа|Рекомендации по слайдам|Ожидаемые результаты|Цель блока|Продолжительность|Предлагаемый график|Онлайн-инструменты|Ключевые понятия|Что важно подчеркнуть|Защита проектов|Шаблон проекта|Финальное сообщение)',t)) or len(t)<70 and t.endswith(('тренинга','модуля','PPT','работы','инструменты'))

def visual_kind(title, kind):
    low=title.lower()
    if kind in ('title','agenda','section'):
        return kind
    if any(w in low for w in ['карта','матрица','схема','этап','дорожная карта','структура','алгоритм']):
        return 'diagram'
    if any(w in low for w in ['упражнение','практика','задание','групповая работа']):
        return 'exercise'
    return 'content'

slides=[]
def add(title, bullets=None, notes='', kind='content'):
    slides.append({'title':title[:90], 'bullets':bullets or [], 'notes':notes, 'kind':visual_kind(title,kind)})

add('Искусственный интеллект как драйвер эффективности', ['Практикум для коммерческого директора и команды','Компания: ЭкоИндустрия','Формат: онлайн‑тренинг с практикой','Продолжительность: 16 академических часов'], '\n'.join(paras[:3]), 'title')
add('Повестка тренинга', ['Вводный блок и правила безопасной работы','Модуль 1. Основы ИИ и промт‑инжиниринг','Модуль 2. ИИ для управления, коммуникаций и продаж','Модуль 3. ИИ в технологиях, логистике и тендерах','Модуль 4. Внедрение, банк промптов и дорожная карта'], '', 'agenda')

cur_title='Общая концепция тренинга'; cur=[]; notes=[]
for t in paras[3:]:
    if is_heading(t) and (cur or notes):
        items=cur if cur else notes[:5]
        for i in range(0,len(items),6):
            add(cur_title if i==0 else cur_title+' — продолжение', items[i:i+6], '\n'.join(notes))
        cur_title=t; cur=[]; notes=[t]
        if re.match(r'^(Модуль \d+\.|\d+\. )', t):
            add(t, ['Цели блока','Ключевые темы','Практические результаты'], t, 'section')
    else:
        notes.append(t)
        if not t.startswith('Слайд') and len(t)<220:
            cur.append(t)
if cur or notes:
    for i in range(0,len(cur),6):
        add(cur_title if i==0 else cur_title+' — продолжение', cur[i:i+6], '\n'.join(notes))

md=[]
for idx,s in enumerate(slides,1):
    md.append(f'---\n\n# {s["title"]}\n')
    for b in s['bullets']:
        md.append(f'- {b}')
    if s['notes']:
        md.append('\n<!-- Заметки тренера:\n'+s['notes'].replace('--','—')+'\n-->')
MD.write_text('\n'.join(md),encoding='utf-8')

W,H=12192000,6858000

def tx_box(x,y,cx,cy,text,size=28,bold=False,color='1F4E79',align='l'):
    text=html.escape(text)
    b='<a:b/>' if bold else ''
    algn=f'<a:pPr algn="{align}"/>' if align!='l' else ''
    tx_box.i+=1
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{tx_box.i}" name="Text {tx_box.i}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p>{algn}<a:r><a:rPr lang="ru-RU" sz="{size*100}">{b}<a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:rPr><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>'''
tx_box.i=1

def rect_box(x,y,cx,cy,text,fill='DDEBF7',line='5B9BD5',size=16,bold=True,color='1F4E79'):
    text=html.escape(text)
    b='<a:b/>' if bold else ''
    tx_box.i+=1
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{tx_box.i}" name="Box {tx_box.i}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" anchor="mid"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="ru-RU" sz="{size*100}">{b}<a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:rPr><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>'''

def slide_xml(s,idx):
    tx_box.i=1
    kind=s.get('kind','content')
    bg='F7FAFC'; title_color='1F4E79'; shapes=[]
    if kind in ('title','section'):
        bg='0F4C5C' if kind=='title' else '1F4E79'; title_color='FFFFFF'
        shapes.append(tx_box(650000,1150000,10800000,950000,s['title'],38,True,title_color))
        y=2550000
        for b in s['bullets'][:6]:
            shapes.append(tx_box(1000000,y,9800000,470000,'• '+b,22,False,'EAF2F8')); y+=560000
    elif kind in ('agenda','diagram'):
        shapes.append(tx_box(500000,350000,11200000,700000,s['title'],32,True,title_color))
        xs=[750000,3450000,6150000,8850000]; y0=1700000
        for n,b in enumerate(s['bullets'][:8]):
            shapes.append(rect_box(xs[n%4],y0+(n//4)*1700000,2450000,1150000,b,fill='EAF2F8' if n%2 else 'DDEBF7',size=15))
    elif kind=='exercise':
        shapes.append(tx_box(500000,350000,11200000,700000,s['title'],32,True,title_color))
        shapes.append(rect_box(700000,1300000,2700000,900000,'Инструкция',fill='FFF2CC',line='D6B656'))
        shapes.append(rect_box(4200000,1300000,2700000,900000,'Работа участников',fill='E2F0D9',line='70AD47'))
        shapes.append(rect_box(7700000,1300000,2700000,900000,'Разбор и выводы',fill='DDEBF7',line='5B9BD5'))
        y=2750000
        for b in s['bullets'][:5]:
            shapes.append(tx_box(900000,y,10300000,480000,'• '+b,19,False,'263238')); y+=580000
    else:
        shapes.append(tx_box(500000,350000,11200000,700000,s['title'],32,True,title_color))
        y=1250000
        for b in s['bullets'][:7]:
            shapes.append(tx_box(850000,y,10450000,520000,'• '+b,20,False,'263238')); y+=650000
    shapes.append(tx_box(10600000,6380000,1000000,300000,str(idx),10,False,'7F7F7F'))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="{bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>{''.join(shapes)}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'''

def notes_xml(s,idx):
    tx_box.i=1
    note = s.get('notes') or 'Заметки тренера: раскрыть ключевые тезисы слайда и связать их с рабочими процессами ЭкоИндустрии.'
    shapes=[tx_box(600000,600000,5600000,500000,'Заметки тренера',22,True), tx_box(600000,1200000,5600000,7200000,note,12,False,'263238')]
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'''+''.join(shapes)+'''</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>'''

def rels(ids):
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+''.join(ids)+'</Relationships>'

ct='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'''
for i in range(1,len(slides)+1):
    ct+=f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
    ct+=f'<Override PartName="/ppt/notesSlides/notesSlide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>'
ct+='</Types>'
pres_ids=''.join(f'<p:sldId id="{255+i}" r:id="rId{i}"/>' for i in range(1,len(slides)+1))
pres=f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{W}" cy="{H}" type="wide"/><p:notesSz cx="6858000" cy="9144000"/><p:sldIdLst>{pres_ids}</p:sldIdLst></p:presentation>'''
pres_rels=rels([f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>' for i in range(1,len(slides)+1)])
root_rels=rels(['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'])
with zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml',ct); z.writestr('_rels/.rels',root_rels); z.writestr('ppt/presentation.xml',pres); z.writestr('ppt/_rels/presentation.xml.rels',pres_rels)
    for i,s in enumerate(slides,1):
        z.writestr(f'ppt/slides/slide{i}.xml',slide_xml(s,i))
        z.writestr(f'ppt/slides/_rels/slide{i}.xml.rels', rels([f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide{i}.xml"/>']))
        z.writestr(f'ppt/notesSlides/notesSlide{i}.xml',notes_xml(s,i))
        z.writestr(f'ppt/notesSlides/_rels/notesSlide{i}.xml.rels', rels([f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide{i}.xml"/>']))
print(f'Создано слайдов: {len(slides)}')
