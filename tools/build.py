import json, collections, openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter as L
inst=json.load(open('inst.json'))
H48=4820
S={'Вн-1':(4750,3870),'Вн-2':(4750,3870),'Вн-3':(6100,3870),'Вн-4':(5400,3470),'Вн-5':(7800,3470),
'Вн-6':(4200,H48),'Вн-7':(3020,H48),'Вн-8':(2600,H48),'Вн-9':(3870,H48),'Вн-10':(4160,H48),'Вн-11':(2450,H48),
'Вн-12':(3370,H48),'Вн-13':(3200,H48),'Вн-14':(2485,H48),'Вн-15':(1400,H48),'Вн-16':(1200,H48),'Вн-17':(4200,H48),
'Вн-18':(2600,H48),'Вн-19':(3000,H48),'Вн-20':(3900,H48),'Вн-21':(2920,H48),'Вн-22':(2150,2720),'Вн-23':(3540,3020),'Вн-24':(2610,3020)}
LITS=['А','Б','В','Г','Д','Ж']
key=lambda m:int(m.split('-')[1])
marks=sorted({i['mark'] for i in inst},key=key)
cnt=collections.Counter((i['mark'],i['lit']) for i in inst)
lvl={}; drs=collections.defaultdict(set)
for i in inst:
    lvl.setdefault(i['mark'],set()).add(i['file'].replace('План на отм. ',''))
    if i['door']: drs[i['mark']].add(i['door'])
NOTE={'Вн-5':'По договору площадь 26,52 м² (габарит 7800×3470 = 27,07 м²)',
      'Вн-22':'КП: 2070×2760',
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
ws['A1']='Спецификация витражей (по планам на отм. -6.000 и 0.000, марки выделены красным)'; ws['A1'].font=Font(bold=True,size=12)
hdr=['№','Марка','Длина, мм','Высота, мм','Площадь 1 шт, м²']+[f'Литер {l}' for l in LITS]+['Всего, шт','Площадь всего, м²','Отметка','Дверь в составе (по чертежу)','Примечание']
ws.append([]); ws.append(hdr)
for c in range(1,len(hdr)+1): ws.cell(3,c).font=hf; ws.cell(3,c).alignment=C; ws.cell(3,c).fill=fill
r=4
for n,m in enumerate(marks,1):
    w,h=S[m]
    row=[n,m,w,h,f'=ROUND(C{r}*D{r}/1000000,2)']+[cnt[(m,l)] or None for l in LITS]
    row+=[f'=SUM(F{r}:K{r})',f'=ROUND(L{r}*E{r},2)',', '.join(sorted(lvl[m])),', '.join(sorted(drs[m])),NOTE.get(m,'')]
    ws.append(row); r+=1
last=r-1
ws.cell(r,2,'ИТОГО').font=hf
for c in range(6,14):
    col=L(c); ws.cell(r,c,f'=SUM({col}4:{col}{last})').font=hf
for c in range(1,len(hdr)+1): ws.cell(r,c).fill=tf
box(ws,3,r,1,len(hdr))
for rr in range(4,r+1):
    for c in range(1,len(hdr)+1):
        ws.cell(rr,c).alignment=W if c in(15,16) else C
    ws.cell(rr,5).number_format='0.00'; ws.cell(rr,13).number_format='#,##0.00'
ws.cell(r+2,1,'Размеры — из таблицы заказчика (длина × высота). Количество — подсчёт марок на чертежах; литер определён по привязке витража к блокам литеров в модели.')
for i,wd in enumerate([4,8,9,9,10,7,7,7,7,7,7,8,11,10,13,34],1): ws.column_dimensions[L(i)].width=wd
ws.row_dimensions[3].height=62
ws.freeze_panes='C4'
ws.page_setup.orientation='landscape'; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0; ws.sheet_properties.pageSetUpPr.fitToPage=True

# ---- Sheet 2
ws2=wb.create_sheet('По литерам')
h2=['№','Марка','Длина, мм','Высота, мм','Площадь 1 шт, м²','Кол-во, шт','Площадь всего, м²','Отметка']
ws2.append(['Спецификация витражей по литерам']); ws2['A1'].font=Font(bold=True,size=12)
ws2.append([]); ws2.append(h2)
for c in range(1,9): ws2.cell(3,c).font=hf; ws2.cell(3,c).alignment=C; ws2.cell(3,c).fill=fill
r=4; n=1; subs=[]
for l in LITS:
    ms=[m for m in marks if cnt[(m,l)]]
    ws2.cell(r,2,f'Литер {l}'+(' (паркинг)' if l=='Ж' else '')).font=hf; ws2.merge_cells(start_row=r,start_column=2,end_row=r,end_column=8); r+=1
    s=r
    for m in ms:
        w,h=S[m]; lv=', '.join(sorted({i['file'].replace('План на отм. ','') for i in inst if i['mark']==m and i['lit']==l}))
        ws2.append([n,m,w,h,f'=ROUND(C{r}*D{r}/1000000,2)',cnt[(m,l)],f'=ROUND(F{r}*E{r},2)',lv]); n+=1; r+=1
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
for i,wd in enumerate([5,16,10,10,11,10,13,11],1): ws2.column_dimensions[L(i)].width=wd
ws2.row_dimensions[3].height=32
# ---- Sheet 3
ws3=wb.create_sheet('Привязка к чертежам')
h3=['№','Чертёж','Литер','Марка','Дверь','X, мм','Y, мм']
ws3.append(['Каждая марка на чертежах (координаты — точка выноски в координатах модели, мм)']); ws3['A1'].font=hf
ws3.append([]); ws3.append(h3)
for c in range(1,8): ws3.cell(3,c).font=hf; ws3.cell(3,c).alignment=C; ws3.cell(3,c).fill=fill
for n,i in enumerate(sorted(inst,key=lambda i:(i['file'],LITS.index(i['lit']),key(i['mark']))),1):
    ws3.append([n,i['file'],i['lit'],i['mark'],i['door'],i['x'],i['y']])
box(ws3,3,ws3.max_row,1,7)
for rr in range(4,ws3.max_row+1):
    for c in range(1,8): ws3.cell(rr,c).alignment=C
    ws3.cell(rr,6).number_format='#,##0'; ws3.cell(rr,7).number_format='#,##0'
for i,wd in enumerate([5,20,7,9,8,11,11],1): ws3.column_dimensions[L(i)].width=wd
wb.calculation.fullCalcOnLoad=True
wb.save('Спецификация_витражей.xlsx')
