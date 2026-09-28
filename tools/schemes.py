# Исполнительные схемы витражей на отм. -6.000 и 0.000 по образцу «Исполнительная схема №5»:
# убираются марки незакрываемых витражей и дверей, марки дверей ставятся по КП/ВДЦ,
# добавляются рамка А1, спецификация (Марка / Кол-во / Площадь по ВДЦ), штамп и подписи.
# Запуск: python tools/schemes.py <план -6.000.dxf> <план 0.000.dxf> <closing.json> <inst.json>
import sys, json, math, collections, ezdxf
from ezdxf import bbox


def to_gray(e):
    """Элементы блоков лежат на красных слоях, цвет вставки на них не влияет:
    вставка разбивается и её части красятся в серый."""
    if not e.is_alive: return
    if e.dxftype() != 'INSERT':
        e.dxf.color = 8; return
    for x in e.explode(): to_gray(x)

F1, F2, CLOSING, INST = sys.argv[1:5]
CL = json.load(open(CLOSING))
RED, GRAY = 1, 8

SIGN = [('Производитель работ ООО «Новая оконная компания»', 'Лазарев А.А.'),
        ('Начальник участка ООО «ПЛУТОС»', 'Коробко А.В.'),
        ('Начальник участка ООО «ПЛУТОС» (строительный контроль)', 'Мельниченко О.А.'),
        ('Ведущий инженер строительного контроля ООО «ТЕХСТРОЙ»', 'Зайцев А.Г.')]
OBJ = ('«Строительство многоэтажного жилого дома со встроенно-пристроенными помещениями '
       'и автостоянкой по адресу: г. Краснодар, ул. Уральская, 100/8»')
CODE = '20002-4-АР.ГЧ'
ORG = 'ООО «Новая оконная компания»'


def fmt(x):
    return f'{x:.2f}'.replace('.', ',')


def spec_rows(level):
    """Строки спецификации: суммы «к закрытию» по марке без разбивки на секции."""
    rows = collections.OrderedDict()
    for x in CL['vitr']:
        on0 = x['m'] in ('Вн-22', 'Вн-23', 'Вн-24') or x['l'] == 'Ж'
        if x['close'] and on0 == (level == 0):
            r = rows.setdefault(x['m'], [0, 0.0]); r[0] += x['close']; r[1] += x['area']
    for x in sorted(CL['doors'], key=lambda d: d['m']):
        if x['close'] and (x['m'] == 'Д22') == (level == 0):
            r = rows.setdefault(x['m'], [0, 0.0]); r[0] += x['close']; r[1] += x['area']
    key = lambda m: (m.startswith('Д'), int(''.join(c for c in m if c.isdigit())))
    return [(m, q, round(a, 2)) for m, (q, a) in sorted(rows.items(), key=lambda kv: key(kv[0]))]


# ---------------------------------------------------------------- оформление листа
class Sheet:
    """Рамка А1 с таблицей подписей, штампом и спецификацией. Координаты в мм листа,
    в модель переводятся как origin + k * (x, y) (k — масштаб модели к листу)."""
    W, H = 841, 594

    def __init__(self, doc, origin, k):
        self.doc, self.msp, self.o, self.k = doc, doc.modelspace(), origin, k
        for name, color in (('РАМКА', 7), ('ТЕКСТ', 7)):
            if name not in doc.layers: doc.layers.add(name, color=color)
        if 'ИД' not in doc.styles:
            doc.styles.add('ИД', font='isocpeur.ttf').dxf.width = 0.8

    def p(self, x, y):
        return (self.o[0] + self.k * x, self.o[1] + self.k * y)

    def line(self, x1, y1, x2, y2, lw=25):
        self.msp.add_line(self.p(x1, y1), self.p(x2, y2), dxfattribs={'layer': 'РАМКА', 'lineweight': lw})

    def rect(self, x1, y1, x2, y2, lw=25):
        self.msp.add_lwpolyline([self.p(x1, y1), self.p(x2, y1), self.p(x2, y2), self.p(x1, y2)], close=True,
                                dxfattribs={'layer': 'РАМКА', 'lineweight': lw})

    def text(self, s, x, y, h, align='MIDDLE_CENTER', color=None):
        a = {'layer': 'ТЕКСТ', 'height': h * self.k, 'style': 'ИД', 'width': 0.8}
        if color: a['color'] = color
        self.msp.add_text(s, dxfattribs=a).set_placement(self.p(x, y), align=ezdxf.enums.TextEntityAlignment[align])

    def mtext(self, s, x, y, h, width, attach=5):
        self.msp.add_mtext(s, dxfattribs={'layer': 'ТЕКСТ', 'char_height': h * self.k, 'width': width * self.k,
                                          'insert': self.p(x, y), 'attachment_point': attach, 'style': 'ИД'})

    def draw(self, title, sheet_no, sheets, rows, notes):
        W, H = self.W, self.H
        self.rect(0, 0, W, H, 25)
        self.rect(20, 5, W - 5, H - 5, 70)
        self.text('Формат А1', W - 7, 2.5, 2.2, 'MIDDLE_RIGHT')
        # заголовок
        self.text(title.replace('\n', ' '), (20 + W - 5) / 2, H - 11, 3.5)
        self.text('М 1:200', (20 + W - 5) / 2, H - 16, 2.5)
        # штамп и подписи (как в образце: 185 × 55 справа внизу)
        R = W - 5; L0 = R - 185
        self.rect(L0, 5, R, 60, 70)
        self.line(L0 + 85, 5, L0 + 85, 60, 70)
        for y in (16, 27, 38, 49): self.line(L0, y, L0 + 85, y)
        for x in (34, 55, 70): self.line(L0 + x, 5, L0 + x, 60)
        for s, x in (('Должность', 17), ('Фамилия', 44.5), ('Подпись', 62.5), ('Дата', 77.5)):
            self.text(s, L0 + x, 54.5, 2.2)
        for i, (pos, name) in enumerate(SIGN):
            y = 43.5 - 11 * i
            self.mtext(pos, L0 + 17, y, 1.8, 31)
            self.text(name, L0 + 44.5, y, 2.0)
        S0 = L0 + 85
        self.line(S0, 17, R, 17); self.line(S0, 33, R, 33); self.line(S0, 47, R, 47, 70)
        for x in (62, 76, 86): self.line(S0 + x, 17, S0 + x, 33)
        self.line(S0 + 62, 25, R, 25)
        self.text(CODE, S0 + 50, 53.5, 5)
        self.mtext(OBJ, S0 + 50, 40, 2.0, 95)
        self.mtext(title.replace('\n', '\\P'), S0 + 31, 25, 2.3, 58)
        for s, v, x in (('Стадия', 'ИД', 69), ('Лист', str(sheet_no), 81), ('Листов', str(sheets), 93)):
            self.text(s, S0 + x, 28.8, 2.2); self.text(v, S0 + x, 21, 2.8)
        self.text(ORG, S0 + 50, 11, 3.2)
        # спецификация (как в образце: 22 + 14 + 20 мм, строка 5,5 мм)
        X0, Y0 = R - 60, H - 21
        n = len(rows) + 2
        yb = Y0 - 5.5 * n
        for i in range(n + 1): self.line(X0, Y0 - 5.5 * i, X0 + 56, Y0 - 5.5 * i, 35 if i in (0, 1, n) else 18)
        for x in (0, 22, 36, 56): self.line(X0 + x, Y0, X0 + x, yb, 35)
        self.text('Спецификация', X0 + 28, Y0 + 3, 2.5)
        for s, x in (('Марка', 11), ('Кол-во, шт', 29), ('Площадь, м²', 46)): self.text(s, X0 + x, Y0 - 2.75, 2.0)
        tq = ta = 0
        for i, (m, q, a) in enumerate(rows):
            y = Y0 - 5.5 * (i + 1.5)
            self.text(m, X0 + 2.5, y, 2.4, 'MIDDLE_LEFT'); self.text(str(q), X0 + 29, y, 2.4); self.text(fmt(a), X0 + 46, y, 2.4)
            tq += q; ta += a
        y = Y0 - 5.5 * (n - 0.5)
        self.text('Итого', X0 + 2.5, y, 2.4, 'MIDDLE_LEFT'); self.text(str(tq), X0 + 29, y, 2.4); self.text(fmt(ta), X0 + 46, y, 2.4)
        self.mtext(notes, X0 + 28, yb - 4, 2.0, 56, attach=2)
        return tq, round(ta, 2)


NOTES = ('Условные обозначения соответствуют проекту.\\P'
         'Работы выполнены в соответствии с требованиями нормативных документов и рабочей документации.\\P'
         'Витражи и двери показаны с указанием марки (красным).\\P'
         'Площади — по ВДЦ (Приложение № 2 к договору).')


def box_of(pl):
    p = list(pl.get_points('xy')); xs = [a for a, b in p]; ys = [b for a, b in p]
    return min(xs), min(ys), max(xs), max(ys)


def inside(b, x, y, tol=0.2):
    return b[0] - tol <= x <= b[2] + tol and b[1] - tol <= y <= b[3] + tol


# ================================================================ отм. -6.000
def scheme_m6(path_in, path_out):
    doc = ezdxf.readfile(path_in); msp = doc.modelspace()
    inst = [i for i in json.load(open(INST)) if i['file'].endswith('-6.000')]
    red = lambda e: e.dxf.get('color', 256) == RED
    texts = [e for e in msp.query('TEXT') if e.dxf.layer == 'МАРКИ' and red(e)]
    boxes = [e for e in msp.query('LWPOLYLINE') if e.dxf.layer == 'МАРКИ' and red(e)]
    leads = [e for e in msp.query('LINE') if e.dxf.layer == 'МАРКИ' and red(e)]
    # марки: текст + рамка + выноска
    groups, seen, used = [], set(), set()
    for t in texts:
        ap = t.dxf.align_point
        bx = next(b for b in boxes if inside(box_of(b), ap.x, ap.y) and id(b) not in used)
        bb = box_of(bx)
        ld = min([l for l in leads if id(l) not in used], key=lambda l: min(math.dist((l.dxf.start.x, l.dxf.start.y), (ap.x, ap.y)),
                                          math.dist((l.dxf.end.x, l.dxf.end.y), (ap.x, ap.y))))
        s, e = ld.dxf.start, ld.dxf.end
        tip = s if not inside(bb, s.x, s.y, 1) else e
        key = (t.dxf.text, round(tip.x, 1), round(tip.y, 1))
        used.update((id(bx), id(ld)))
        g = dict(mark=t.dxf.text, t=t, box=bx, bb=bb, lead=ld, tip=(tip.x, tip.y))
        if key in seen:                       # дубль подписи на одном месте — удалить
            for x in (t, bx, ld): msp.delete_entity(x)
            continue
        seen.add(key)
        i = min(inst, key=lambda i: math.dist((i['x'] / 200, i['y'] / 200), g['tip']))
        g['lit'] = i['lit']; groups.append(g)
    # что не закрывается (сверх ВДЦ): В — Вн-17, два Вн-16 из трёх; Г — один Вн-17 из двух
    def drop(g):
        m, l, (x, y) = g['mark'], g['lit'], g['tip']
        if l == 'В' and m == 'Вн-17': return True
        if l == 'В' and m == 'Вн-16' and abs(x - 376.9) > 1: return True
        if l == 'Г' and m == 'Вн-17' and abs(x - 239.2) > 1: return True
        return False
    width = {x['m']: x['w'] for x in CL['vitr'] if x['w']}
    vitr = [e for e in msp.query('INSERT') if e.dxf.layer == 'ВИТРАЖИ' and red(e)]
    kept, removed = [], []
    for g in groups:
        if drop(g):
            removed.append(g)
            tx, ty = g['tip']; half = width.get(g['mark'], 4200) / 400 + 0.5
            vert = abs(g['lead'].dxf.start.y - g['lead'].dxf.end.y) < 0.01   # выноска горизонтальна → стена вертикальна
            for e in vitr:
                if not e.is_alive: continue
                c = bbox.extents([e], fast=True).center
                along, across = (c.y - ty, c.x - tx) if vert else (c.x - tx, c.y - ty)
                if abs(along) <= half and abs(across) <= 2: to_gray(e)
            for x in (g['t'], g['box'], g['lead']): msp.delete_entity(x)
        else:
            kept.append(g)
    # двери: снять все старые марки дверей и поставить по КП/ВДЦ
    dtexts = [e for e in msp.query('TEXT') if e.dxf.layer == 'ДВЕРИ' and red(e)]
    tmpl = dtexts[0].copy()
    for e in list(dtexts) + [e for e in msp.query('LWPOLYLINE') if e.dxf.layer == 'ДВЕРИ' and red(e)]:
        msp.delete_entity(e)
    DOORS = {('А', 'Вн-6'): 'Д17*', ('А', 'Вн-14'): 'Д17*', ('А', 'Вн-15'): 'Д35',
             ('Б', 'Вн-6'): 'Д17*', ('Б', 'Вн-18'): 'Д17*', ('Б', 'Вн-15'): 'Д35',
             ('В', 'Вн-6'): 'Д17*', ('В', 'Вн-19'): 'Д17*', ('В', 'Вн-16'): 'Д35',
             ('Г', 'Вн-10'): 'Д17*', ('Д', 'Вн-6'): 'Д17*', ('Д', 'Вн-9'): 'Д17*'}
    dcount = collections.Counter()
    for g in kept:
        dm = DOORS.get((g['lit'], g['mark']))
        if not dm: continue
        dcount[dm] += 1
        x0, y0, x1, y1 = g['bb']; cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        dx, dy = cx - g['tip'][0], cy - g['tip'][1]
        if abs(dx) > abs(dy):   # вертикальная стена: марка правее/левее
            s = 1 if dx > 0 else -1; xa = (x1 if s > 0 else x0) + 0.5 * s
            db = (min(xa, xa + 4.4 * s), cy - 4.4, max(xa, xa + 4.4 * s), cy + 4.4); rot = 90
        else:
            s = 1 if dy > 0 else -1; ya = (y1 if s > 0 else y0) + 0.5 * s
            db = (cx - 4.4, min(ya, ya + 4.4 * s), cx + 4.4, max(ya, ya + 4.4 * s)); rot = 0
        msp.add_lwpolyline([(db[0], db[1]), (db[2], db[1]), (db[2], db[3]), (db[0], db[3])], close=True,
                           dxfattribs={'layer': 'ДВЕРИ', 'color': RED, 'lineweight': 35})
        t = tmpl.copy(); msp.add_entity(t)
        t.dxf.text = dm; t.dxf.rotation = rot; t.dxf.color = RED
        t.set_placement(((db[0] + db[2]) / 2, (db[1] + db[3]) / 2), align=ezdxf.enums.TextEntityAlignment.MIDDLE_CENTER)
    # Вн-6 в секции Г: два из трёх закрываются в объёмах секций А и Б
    g6 = sorted([g for g in kept if g['lit'] == 'Г' and g['mark'] == 'Вн-6'], key=lambda g: g['tip'][0])
    notes6 = {216: 'в объёме лит. А', 268: 'в объёме лит. Б'}
    for g in g6:
        lab = notes6.get(round(g['tip'][0]))
        if lab:
            x0, y0, x1, y1 = g['bb']
            msp.add_text(lab, dxfattribs={'layer': 'МАРКИ', 'color': RED, 'height': 1.6, 'width': 0.8}).set_placement(
                ((x0 + x1) / 2, y1 + 2.5), align=ezdxf.enums.TextEntityAlignment.MIDDLE_CENTER)
    for g in kept:
        for x in (g['t'], g['box'], g['lead']): x.dxf.color = RED
    rows = spec_rows(-6)
    sh = Sheet(doc, (50, -1294), 1)
    tot = sh.draw('Исполнительная схема №1. Плановое положение витражных блоков из алюминиевого профиля\nна отм. −6.000',
                  1, 2, rows, NOTES)
    doc.saveas(path_out)
    return dict(kept=collections.Counter(g['mark'] for g in kept), removed=[(g['lit'], g['mark']) for g in removed],
                doors=dcount, rows=rows, total=tot)


# ================================================================ отм. 0.000
def scheme_0(path_in, path_out):
    doc = ezdxf.readfile(path_in); msp = doc.modelspace()
    texts = [e for e in msp.query('TEXT') if e.dxf.layer == 'МАРКИ']
    boxes = [e for e in msp.query('LWPOLYLINE') if e.dxf.layer == 'МАРКИ']
    def bx_for(t):
        p = t.dxf.insert
        def dist(b):
            x0, y0, x1, y1 = box_of(b)
            return 0 if inside((x0, y0, x1, y1), p.x, p.y) else 1e9
        cand = [b for b in boxes if dist(b) == 0]
        return min(cand, key=lambda b: math.dist(((box_of(b)[0] + box_of(b)[2]) / 2, (box_of(b)[1] + box_of(b)[3]) / 2), (p.x, p.y)))
    marks = [(t, bx_for(t)) for t in texts]
    vn = [(t, b) for t, b in marks if t.dxf.text.startswith('Вн')]
    dr = [(t, b) for t, b in marks if t.dxf.text.startswith('Д')]
    def lit22(x): return 'А' if x >= 86000 else 'Б' if x >= 50000 else 'В' if x >= 20000 else 'Г'
    KEEP_B = 45611            # из четырёх Вн-22 секции В закрывается один (КП: А5, Б5, В1)
    DOOR_X = (112864, 106184, 79305, 45611)   # 4 двери Д22 закрываются по ВДЦ (строки Д22 А, Б, Г, Д)
    def is_red(e):
        c = e.dxf.get('color', 256)
        return (doc.layers.get(e.dxf.layer).color if c == 256 else c) == RED
    red = [e for e in msp if e.dxf.layer not in ('МАРКИ',) and e.dxftype() in ('INSERT', 'LINE', 'LWPOLYLINE') and is_red(e)]
    def gray_near(cx, cy, dxm, dym, only_door=False):
        for e in red:
            if not e.is_alive or (only_door and (e.dxftype() != 'INSERT' or 'Дверь' not in e.dxf.name)): continue
            if not e.is_alive: continue
            c = bbox.extents([e], fast=True).center
            if abs(c.x - cx) <= dxm and abs(c.y - cy) <= dym: to_gray(e)
    kept = collections.Counter(); removed = []
    for t, b in vn:
        m = t.dxf.text; x = t.dxf.insert.x; b0 = box_of(b); cx = (b0[0] + b0[2]) / 2
        keep = m == 'Вн-22' and (lit22(x) in 'АБ' or abs(x - KEEP_B) < 5)
        # дверная марка над этой маркой
        d = [(dt, db) for dt, db in dr if dt.is_alive and abs(dt.dxf.insert.x - x) < 200 and 0 < dt.dxf.insert.y - t.dxf.insert.y < 900]
        if keep:
            kept[m] += 1; t.dxf.color = RED; b.dxf.color = RED
            for dt, db in d:
                if any(abs(x - v) < 5 for v in DOOR_X):
                    dt.dxf.text = 'Д-22'; dt.dxf.color = RED; db.dxf.color = RED
                else:
                    msp.delete_entity(dt); msp.delete_entity(db)
                    near = (cx, 1650) if b0[1] < 10000 else (2068, (b0[1] + b0[3]) / 2 + 1000)
                    gray_near(near[0], near[1], 1300, 1300, only_door=True)
        else:
            removed.append((lit22(x) if m == 'Вн-22' else ('Ж' if int(m[3:]) <= 5 else ''), m))
            for dt, db in d: msp.delete_entity(dt); msp.delete_entity(db)
            msp.delete_entity(t); msp.delete_entity(b)
            if m == 'Вн-22':
                if b0[1] < 10000: gray_near(cx, 1800, 2300, 600)
                else: gray_near(2300, (b0[1] + b0[3]) / 2 - 300, 600, 1700)
            elif m in ('Вн-23', 'Вн-24'):
                gray_near(cx, 25200, 1400, 1500)
    # пустые рамки марок (без текста) — удалить
    alive = [t for t in msp.query('TEXT') if t.dxf.layer == 'МАРКИ']
    for bx in boxes:
        if bx.is_alive and not any(inside(box_of(bx), t.dxf.insert.x, t.dxf.insert.y, 50) for t in alive):
            msp.delete_entity(bx)
    rows = spec_rows(0)
    k = 200; o = (-4854 - 35 * k, -6344 - 65 * k)
    sh = Sheet(doc, o, k)
    tot = sh.draw('Исполнительная схема №2. Плановое положение витражных блоков из алюминиевого профиля\nна отм. 0.000',
                  2, 2, rows, NOTES)
    doc.saveas(path_out)
    return dict(kept=kept, removed=removed, rows=rows, total=tot)


if __name__ == '__main__':
    r1 = scheme_m6(F1, 'Исп_схема_1_отм_-6.000.dxf')
    r2 = scheme_0(F2, 'Исп_схема_2_отм_0.000.dxf')
    for r in (r1, r2):
        for k, v in r.items(): print(k, v)
        print('---')
