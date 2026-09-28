import json, collections, openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter as L
inst=json.load(open('inst.json'))
H48=4820
S={'Вн-1':(4750,3870),'Вн-2':(4750,3870),'Вн-3':(6100,3870),'Вн-4':(5400,3470),'Вн-5':(7800,3470),
'Вн-6':(4200,H48),'Вн-7':(3020,H48),'Вн-8':(2600,H48),'Вн-9':(3870,H48),'Вн-10':(4160,H48),'Вн-11':(2450,H48),
'Вн-12':(3370,H48),'Вн-13':(3200,H48),'Вн-14':(2485,H48),'Вн-15':(1400,H48),'Вн-16':(1200,H48),'Вн-17':(4200,H48),
'Вн-18':(2600,H48),'Вн-19':(3000,H48),'Вн-20':(3900,H48),'Вн-21':(2920,H48),'Вн-22':(2150,2720),'Вн-23':(3540,3020),'Вн-24':(2610,3020)}
LITS=['А','Б','В','Г','Д']
key=lambda m:int(m.split('-')[1])
drw=collections.Counter((i['mark'],i['lit']) for i in inst)
# количество по КП (колонка «по КП или Изготовлено» таблицы заказчика)
KP={('Вн-6','А'):4,('Вн-8','А'):1,('Вн-14','А'):1,('Вн-15','А'):1,('Вн-22','А'):5,('Вн-23','А'):0,('Вн-24','А'):0,
('Вн-6','Б'):4,('Вн-15','Б'):1,('Вн-17','Б'):1,('Вн-18','Б'):1,('Вн-22','Б'):5,('Вн-23','Б'):0,('Вн-24','Б'):0,
('Вн-6','В'):1,('Вн-13','В'):1,('Вн-16','В'):3,('Вн-17','В'):1,('Вн-19','В'):2,('Вн-22','В'):1,
('Вн-6','Г'):1,('Вн-10','Г'):1,('Вн-11','Г'):1,('Вн-12','Г'):1,('Вн-17','Г'):1,('Вн-20','Г'):1,('Вн-21','Г'):1,('Вн-22','Г'):0,
('Вн-6','Д'):2,('Вн-7','Д'):1,('Вн-9','Д'):1,('Вн-17','Д'):1,('Вн-10','Д'):1,
('Вн-1','Ж'):0,('Вн-2','Ж'):0,('Вн-3','Ж'):0,('Вн-4','Ж'):0,('Вн-5','Ж'):0}
# Вн-6: по 1 шт. сверх чертежа из КП литеров А и Б перенесены в литер Г (на чертеже Г — 3 шт., в КП Г — 1)
KP[('Вн-6','А')]=3; KP[('Вн-6','Б')]=3; KP[('Вн-6','Г')]=3
# площади по КП (колонка «по КП или Изготовлено», площадь), с учётом переноса Вн-6
MOVED={'А':19.90,'Б':19.96}   # 79.58/4 и 79.82/4
KPA={('Вн-6','А'):round(79.58-MOVED['А'],2),('Вн-8','А'):12.34,('Вн-14','А'):11.14,('Вн-15','А'):5.62,('Вн-22','А'):28.57,
('Вн-6','Б'):round(79.82-MOVED['Б'],2),('Вн-15','Б'):6.00,('Вн-17','Б'):19.49,('Вн-18','Б'):11.90,('Вн-22','Б'):28.57,
('Вн-6','В'):19.63,('Вн-13','В'):14.74,('Вн-16','В'):18.57,('Вн-17','В'):19.68,('Вн-19','В'):24.67,('Вн-22','В'):5.71,
('Вн-6','Г'):20.45,('Вн-10','Г'):19.97,('Вн-11','Г'):11.09,('Вн-12','Г'):15.60,('Вн-17','Г'):19.82,('Вн-20','Г'):18.43,('Вн-21','Г'):14.02,
('Вн-6','Д'):39.98,('Вн-7','Д'):19.97,('Вн-9','Д'):19.68,('Вн-17','Д'):18.14,('Вн-10','Д'):20.16}
def area(m,l):
    if (m,l)==('Вн-6','Г'): return round(KPA[(m,l)]+MOVED['А']+MOVED['Б'],2)
    return KPA[(m,l)]
cnt=collections.Counter({k:min(v,KP[k]) for k,v in drw.items() if KP[k]>0})
inst=[i for i in inst if cnt[(i['mark'],i['lit'])]]
marks=sorted({i['mark'] for i in inst},key=key)
lvl={}; drs=collections.defaultdict(set)
for i in inst:
    lvl.setdefault(i['mark'],set()).add(i['file'].replace('План на отм. ',''))
    if i['door']: drs[i['mark']].add(i['door'])
NOTE={'Вн-22':'КП: 2070×2760. На чертеже 18 шт., в КП 11 (А5, Б5, В1)','Вн-6':'Лит. Г: 1 шт. по КП Г + по 1 шт. из КП лит. А и Б (там в КП по 4, на чертеже по 3)','Вн-17':'Лит. Г: на чертеже 2 шт., в КП 1',
      'Вн-16':'1 шт. у лестницы на границе литеров В/Г отнесён к литеру В'}
thin=Side(style='thin'); B=Border(left=thin,right=thin,top=thin,bottom=thin)
hf=Font(bold=True); fill=PatternFill('solid',fgColor='DDEBF7'); tf=PatternFill('solid',fgColor='F2F2F2')
C=Alignment(horizontal='center',vertical='center',wrap_text=True)
W=Alignment(vertical='center',wrap_text=True)
def box(ws,r1,r2,c1,c2):
    for r in range(r1,r2+1):
        for c in range(c1,c2+1): ws.cell(r,c).border=B

wb=openpyxl.Workbook()
# ---- Sheet 1
ws=wb.active; ws.title='Спецификация'
ws['A1']='Спецификация витражей (по планам на отм. -6.000 и 0.000, марки выделены красным) — только позиции, входящие в КП'; ws['A1'].font=Font(bold=True,size=12)
hdr=['№','Марка','Длина, мм','Высота, мм','Площадь 1 шт по КП, м²']+[f'Литер {l}' for l in LITS]+['Всего, шт','Площадь по КП, м²','Отметка','Дверь в составе (по чертежу)','Примечание']
ws.append([]); ws.append(hdr)
for c in range(1,len(hdr)+1): ws.cell(3,c).font=hf; ws.cell(3,c).alignment=C; ws.cell(3,c).fill=fill
r=4
for n,m in enumerate(marks,1):
    w,h=S[m]
    row=[n,m,w,h,f'=ROUND(L{r}/K{r},2)']+[cnt[(m,l)] or None for l in LITS]
    row+=[f'=SUM(F{r}:J{r})',round(sum(area(m,l) for l in LITS if cnt[(m,l)]),2),', '.join(sorted(lvl[m])),', '.join(sorted(drs[m])),NOTE.get(m,'')]
    ws.append(row); r+=1
last=r-1
ws.cell(r,2,'ИТОГО').font=hf
for c in range(6,13):
    col=L(c); ws.cell(r,c,f'=SUM({col}4:{col}{last})').font=hf
for c in range(1,len(hdr)+1): ws.cell(r,c).fill=tf
box(ws,3,r,1,len(hdr))
for rr in range(4,r+1):
    for c in range(1,len(hdr)+1):
        ws.cell(rr,c).alignment=W if c in(14,15) else C
    ws.cell(rr,5).number_format='0.00'; ws.cell(rr,12).number_format='#,##0.00'
ws.cell(r+2,1,'Размеры — из таблицы заказчика (длина × высота). Количество — по чертежам, но не больше, чем в КП. Площади — по КП (площадь 1 шт — средняя).')
for i,wd in enumerate([4,8,9,9,10,7,7,7,7,7,8,11,10,13,34],1): ws.column_dimensions[L(i)].width=wd
ws.row_dimensions[3].height=62
ws.freeze_panes='C4'
ws.page_setup.orientation='landscape'; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0; ws.sheet_properties.pageSetUpPr.fitToPage=True

# ---- Sheet 2
ws2=wb.create_sheet('По литерам')
h2=['№','Марка','Длина, мм','Высота, мм','Площадь 1 шт по КП, м²','Кол-во, шт','Площадь по КП, м²','Отметка']
ws2.append(['Спецификация витражей по литерам']); ws2['A1'].font=Font(bold=True,size=12)
ws2.append([]); ws2.append(h2)
for c in range(1,9): ws2.cell(3,c).font=hf; ws2.cell(3,c).alignment=C; ws2.cell(3,c).fill=fill
r=4; n=1; subs=[]
for l in LITS:
    ms=[m for m in marks if cnt[(m,l)]]
    ws2.cell(r,2,f'Литер {l}').font=hf; ws2.merge_cells(start_row=r,start_column=2,end_row=r,end_column=8); r+=1
    s=r
    for m in ms:
        w,h=S[m]; lv=', '.join(sorted({i['file'].replace('План на отм. ','') for i in inst if i['mark']==m and i['lit']==l}))
        if (m,l)==('Вн-6','Г'):
            parts=[(m,1,KPA[(m,l)]),('Вн-6 (из КП лит. А)',1,MOVED['А']),('Вн-6 (из КП лит. Б)',1,MOVED['Б'])]
        else: parts=[(m,cnt[(m,l)],area(m,l))]
        for nm,q,a in parts:
            ws2.append([n,nm,w,h,f'=ROUND(G{r}/F{r},2)',q,a,lv]); n+=1; r+=1
    ws2.cell(r,2,f'Итого литер {l}').font=hf
    ws2.cell(r,6,f'=SUM(F{s}:F{r-1})').font=hf; ws2.cell(r,7,f'=SUM(G{s}:G{r-1})').font=hf
    for c in range(1,9): ws2.cell(r,c).fill=tf
    subs.append(r); r+=1
ws2.cell(r,2,'ВСЕГО').font=hf
ws2.cell(r,6,'='+'+'.join(f'F{x}' for x in subs)).font=hf; ws2.cell(r,7,'='+'+'.join(f'G{x}' for x in subs)).font=hf
box(ws2,3,r,1,8)
for rr in range(4,r+1):
    for c in range(1,9):
        if ws2.cell(rr,c).coordinate not in ws2.merged_cells: ws2.cell(rr,c).alignment=C
    ws2.cell(rr,5).number_format='0.00'; ws2.cell(rr,7).number_format='#,##0.00'
for i,wd in enumerate([5,21,10,10,11,10,13,11],1): ws2.column_dimensions[L(i)].width=wd
ws2.row_dimensions[3].height=32
# ---- Sheet 3
ws3=wb.create_sheet('Привязка к чертежам')
h3=['№','Чертёж','Литер','Марка','Дверь','X, мм','Y, мм','В КП']
ws3.append(['Витражи на чертежах, входящие в КП (координаты — точка выноски, мм)'])
ws3.append(['«1 из 3» — на чертеже больше, чем в КП; какие именно — по КП не определить']); ws3['A1'].font=hf
ws3.append(h3)
for c in range(1,9): ws3.cell(3,c).font=hf; ws3.cell(3,c).alignment=C; ws3.cell(3,c).fill=fill
for n,i in enumerate(sorted(inst,key=lambda i:(i['file'],LITS.index(i['lit']),key(i['mark']))),1):
    k=(i['mark'],i['lit']); ws3.append([n,i['file'],i['lit'],i['mark'],i['door'],i['x'],i['y'],'да' if drw[k]<=KP[k] else f'{KP[k]} из {drw[k]}'])
box(ws3,3,ws3.max_row,1,8)
for rr in range(4,ws3.max_row+1):
    for c in range(1,9): ws3.cell(rr,c).alignment=C
    ws3.cell(rr,6).number_format='#,##0'; ws3.cell(rr,7).number_format='#,##0'
for i,wd in enumerate([5,20,7,9,8,11,11,9],1): ws3.column_dimensions[L(i)].width=wd
# ---- Sheet 4: есть в КП, но не вошло в спецификацию
ws4=wb.create_sheet('В КП, не вошло')
ws4.append(['Позиции КП, не вошедшие в спецификацию (нет на чертежах или на чертежах меньше, чем в КП)']); ws4['A1'].font=Font(bold=True,size=12)
ws4.append([])
h4=['№','Литер','Марка','Длина, мм','Высота, мм','Площадь 1 шт по КП, м²','В КП, шт','Вошло в спецификацию, шт','Не вошло, шт','Площадь не вошедшего, м²','Причина']
ws4.append(h4)
for c in range(1,12): ws4.cell(3,c).font=hf; ws4.cell(3,c).alignment=C; ws4.cell(3,c).fill=fill
rows=[]
for (m,l),q in KP.items():
    got=cnt[(m,l)]
    if q>got:
        w,h=S[m]
        why='нет на чертежах' if drw[(m,l)]==0 else f'на чертежах {drw[(m,l)]} шт.'
        rows.append((l,m,w,h,q,got,why,round(KPA[(m,l)]/q,2) if (m,l) in KPA else None))
# остаток КП без литера (в таблице заказчика — строки «без литера»)
for m,w in [('Вн-19',2860),('Вн-19',2920),('Вн-16',1320),('Вн-16',1020),('Вн-17',4020),('Вн-9',4060),('Вн-17',3750),('Вн-6',4130)]:
    rows.append(('без литера',m,w,4800,1,0,'остаток КП без литера, по размеру к строкам договора не подходит',None))
rows.sort(key=lambda z:(LITS.index(z[0]) if z[0] in LITS else 9,key(z[1])))
r=4
for n,(l,m,w,h,q,got,why,a1) in enumerate(rows,1):
    ws4.append([n,l,m,w,h,a1 if a1 is not None else f'=ROUND(D{r}*E{r}/1000000,2)',q,got,f'=G{r}-H{r}',f'=ROUND(I{r}*F{r},2)',why]); r+=1
ws4.cell(r,3,'ИТОГО').font=hf
for c in (7,8,9,10):
    col=L(c); ws4.cell(r,c,f'=SUM({col}4:{col}{r-1})').font=hf
for c in range(1,12): ws4.cell(r,c).fill=tf
box(ws4,3,r,1,11)
for rr in range(4,r+1):
    for c in range(1,12): ws4.cell(rr,c).alignment=W if c==11 else C
    ws4.cell(rr,6).number_format='0.00'; ws4.cell(rr,10).number_format='#,##0.00'
for i,wd in enumerate([4,11,8,9,9,10,8,13,9,12,40],1): ws4.column_dimensions[L(i)].width=wd
ws4.row_dimensions[3].height=60
ws4.page_setup.orientation='landscape'; ws4.sheet_properties.pageSetUpPr.fitToPage=True; ws4.page_setup.fitToHeight=0
wb.calculation.fullCalcOnLoad=True
wb.save('Спецификация_витражей.xlsx')
