import ezdxf, re, collections, json
out=[]
# f1
doc=ezdxf.readfile('f1.dxf'); msp=doc.modelspace()
ins=[(e.dxf.insert.x,e.dxf.insert.y,re.search(r'--1([А-Я])$',e.dxf.name).group(1)) for e in msp.query('INSERT') if re.search(r'--1([А-Я])$',e.dxf.name)]
doors=[(e.dxf.text,e.dxf.insert.x,e.dxf.insert.y) for e in msp.query('TEXT') if e.dxf.layer=='ДВЕРИ']
lines=[(e.dxf.start,e.dxf.end) for e in msp.query('LINE') if e.dxf.layer=='МАРКИ']
M=[(e.dxf.text,e.dxf.insert.x,e.dxf.insert.y) for e in msp.query('TEXT') if e.dxf.layer=='МАРКИ']
seen=set()
for e in msp.query('TEXT'):
    if e.dxf.layer!='МАРКИ': continue
    t,x,y=e.dxf.text,e.dxf.insert.x,e.dxf.insert.y
    k=(t,round(x),round(y))
    if k in seen: continue
    seen.add(k)
    ln=min(lines,key=lambda L:min((L[0].x-x)**2+(L[0].y-y)**2,(L[1].x-x)**2+(L[1].y-y)**2))
    a,b=ln; tip=a if (a.x-x)**2+(a.y-y)**2>(b.x-x)**2+(b.y-y)**2 else b
    near=sorted(ins,key=lambda i:(i[0]-tip.x)**2+(i[1]-tip.y)**2)[:5]
    lit=collections.Counter(n[2] for n in near).most_common(1)[0][0]
    if t=='Вн-16' and abs(tip.x-291.1)<1: lit='В'
    d=[]
    for dt,dx,dy in doors:
        dd=((dx-x)**2+(dy-y)**2)**.5
        if dd<7.5 and dd<=min(((dx-mx)**2+(dy-my)**2)**.5 for mt,mx,my in M)+0.01: d.append(dt)
    out.append(dict(file='План на отм. -6.000',lit=lit,mark=t,door=d[0] if d else '',x=round(tip.x*200),y=round(tip.y*200)))
# f2
doc=ezdxf.readfile('f2.dxf'); msp=doc.modelspace()
T=[(e.dxf.text,e.dxf.insert.x,e.dxf.insert.y) for e in msp.query('TEXT') if e.dxf.layer=='МАРКИ']
for t,x,y in T:
    if not t.startswith('Вн'): continue
    d=[dt for dt,dx,dy in T if dt.startswith('Д') and abs(dx-x)<200 and 0<dy-y<900]
    n=int(t[3:])
    if n<=5: lit='Ж'
    elif n in(23,24): lit='А' if x>80000 else 'Б'
    else:
        lit='А' if x>=86000 else 'Б' if x>=50000 else 'В' if x>=20000 else 'Г'
    out.append(dict(file='План на отм. 0.000',lit=lit,mark=t,door=d[0].replace('-','') if d else '',x=round(x),y=round(y)))
json.dump(out,open('inst.json','w'),ensure_ascii=False,indent=0)
c=collections.Counter((o['lit'],o['mark'],o['door']) for o in out)
for k,v in sorted(c.items()): print(k,v)
print(len(out))
