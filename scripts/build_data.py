#!/usr/bin/env python3
"""Собирает data.json (бюджет) из Excel-файлов в папке input/.

Файлы определяются по СОДЕРЖИМУ, а не по имени. Год берётся из самого файла:
  * «Исполнение»: первый лист называется «Исполнение 2026» (год в названии листа);
  * «План закупок»: первый лист называется «2026» (год = название листа);
  * «Бюджетная заявка»: на первом листе B3 = «Год», C3 = 2026, C4 = версия (PLAN05).
Если в input/ несколько файлов одного вида и года, берётся самый свежий.
Если файла какого-то вида/года нет, эта часть data.json остаётся прежней.

Колонки ищутся по ЗАГОЛОВКАМ шапки (а не по буквам): в файле 2027 года столбцы сдвинуты относительно 2026,
но заголовки те же. Для 2026 это даёт те же колонки, что заказывались по буквам:
  Исполнение:      BA BB BC R FE FF FH GH GJ GK GQ GS
  План закупок:    F G M N P Q R U Y AW AX
  Бюджетная заявка: C M N AO AR AS AV
Шапка = все строки выше первой строки данных (первая строка, где есть код товара вида 123-45678).
Если нужный заголовок не найден, файл пропускается с понятным сообщением, остальные файлы собираются.
После успешной сборки распознанные файлы удаляются из input/ (в репозитории это фиксирует workflow).
"""
import datetime, glob, json, os, re, subprocess, sys
from collections import Counter
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as gl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT = os.environ.get('BUDGET_INPUT') or os.path.join(ROOT, 'input')
OUT = os.environ.get('BUDGET_OUT') or os.path.join(ROOT, 'data.json')
MAXCOL = 300          # дальше этой колонки заголовки не ищем
HEAD_SCAN = 40        # первая строка данных ищется в первых 40 строках
CODE = re.compile(r'^\d{3}-\d{5}$')


def s(v):
    return '' if v is None else re.sub(r'[ \t\r\f\v\xa0]+', ' ', str(v)).strip()


def n(v):
    try:
        x = round(float(v), 3)
    except (TypeError, ValueError):
        return 0
    return int(x) if x == int(x) else x


def d(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime('%Y-%m-%d')
    return s(v)[:10]


def low(x):
    return re.sub(r'\s+', ' ', s(x).lower())


# ---------------------------------------------------------------------------
# Описание трёх видов файлов. Поля идут в порядке data.json.
#   field = (имя, регулярка заголовка, после какого поля искать, конвертер[, строго после])
#   Колонка = первая слева, где в шапке (любой из строк) есть заголовок по регулярке,
#   правее колонки поля «после». Поле '_grp' служебное (не пишется в data.json).
# ---------------------------------------------------------------------------
KINDS = {
    'exec': dict(
        key=['code', 'name'],
        fields=[
            ('code', r'новый код товара', None, s),
            ('name', r'^наименование', 'code', s),
            ('unit', r'^ед\.? ?изм', 'code', s),
            ('cust', r'^заказчики$', None, s),
            ('_grp', r'переданные заявки на закуп', None, None),
            ('qty', r'^кол-?во$', '=_grp', n),
            ('sum', r'^сумма', 'qty', n),
            ('meth', r'^способ закупок', 'sum', s),
            ('st1', r'^статус заявки', 'meth', s),
            ('st2', r'^статус закупки', 'st1', s),
            ('open', r'^дата вскрытия', 'st2', d),
            ('sup', r'^наименование поставщика', 'open', s),
            ('cdate', r'^дата договора', 'sup', d),
        ],
        out=['code', 'name', 'unit', 'cust', 'qty', 'sum', 'meth', 'st1', 'st2', 'open', 'sup', 'cdate'],
        canon=['meth', 'st1', 'st2'],
    ),
    'plan': dict(
        key=['code', 'name'],  # у работ и услуг кода товара нет, есть только наименование
        fields=[
            ('type', r'^тру$', None, s),
            ('cust', r'^бюджет$', None, s),
            ('code', r'^код товара', None, s),
            ('name', r'^наименование', 'code', s),
            ('qty', r'^кол-?во$', 'name', n),
            ('price', r'^маркетинговая цена', 'qty', n),
            ('sum', r'^сумма', 'price', n),
            ('rdate', r'^дата заявки', 'sum', d),
            ('meth', r'^способ закупок', None, s),
            ('st', r'^статус закупок', 'meth', s),
            ('nopl', r'^статус несост', 'st', s),
        ],
        out=['type', 'cust', 'code', 'name', 'qty', 'price', 'sum', 'rdate', 'meth', 'st', 'nopl'],
        canon=['type', 'cust', 'meth', 'st', 'nopl'],
    ),
    'print': dict(
        key=['code'],
        fields=[
            ('abp', r'^наименование$', None, s),
            ('mvz', r'^код мвз$', None, None),
            ('cex', r'^наименование$', 'mvz', s),
            ('gpz', r'^код гпз$', None, s),
            ('code', r'новый код товара', None, s),
            ('qty', r'^кол-?во$', 'code', n),
            ('price', r'^цена', 'qty', n),
            ('note', r'^примечание', 'price', s),
        ],
        out=['abp', 'cex', 'gpz', 'code', 'qty', 'price', 'note'],
        canon=['abp', 'cex'],
    ),
}


def cell(ws, ref):
    return ws[ref].value


def detect(ws):
    """Вид файла и год по содержимому первого листа, иначе (None, None)."""
    title = low(ws.title)
    m = re.match(r'исполнение\s*(\d{4})', title)
    if m:
        return 'exec', m.group(1)
    if re.fullmatch(r'\d{4}', title):
        return 'plan', title
    try:
        if low(cell(ws, 'B3')) == 'год' and re.fullmatch(r'\d{4}', s(cell(ws, 'C3'))) and low(cell(ws, 'B4')) == 'версия':
            return 'print', s(cell(ws, 'C3'))
    except Exception:
        pass
    return None, None


def locate(kind, ws):
    """Находит колонки по заголовкам. Возвращает (первая строка данных, {поле: индекс колонки}) или поднимает ValueError."""
    grid = [list(r) for r in ws.iter_rows(min_row=1, max_row=HEAD_SCAN, max_col=MAXCOL, values_only=True)]
    first = None
    for ri, r in enumerate(grid):
        if any(isinstance(v, str) and CODE.match(v.strip()) for v in r):
            first = ri + 1
            break
    if first is None:
        raise ValueError(f'не нашёл первую строку данных (код товара вида 123-45678) в первых {HEAD_SCAN} строках')
    head = grid[:first - 1]
    width = max((len(r) for r in head), default=0)
    texts = [[low(r[c]) if c < len(r) else '' for r in head] for c in range(width)]  # по колонкам: тексты шапки
    pos, missing = {}, []
    for name, rx, after, _ in KINDS[kind]['fields']:
        start = 0
        if after:
            inclusive = after.startswith('=')
            ref = pos.get(after.lstrip('='))
            if ref is None:
                missing.append(name)
                continue
            start = ref if inclusive else ref + 1
        found = next((c for c in range(start, width) if any(re.search(rx, t) for t in texts[c] if t)), None)
        if found is None:
            missing.append(name)
        else:
            pos[name] = found
    if missing:
        raise ValueError('не нашёл заголовки для полей: ' + ', '.join(missing))
    return first, pos


def read_kind(kind, ws, first, pos):
    cfg = KINDS[kind]
    conv = {f[0]: f[3] for f in cfg['fields']}
    cols = [(name, pos[name], conv[name]) for name in cfg['out']]
    keys = [pos[k] for k in cfg['key']]
    maxc = max(c for _, c, _ in cols) + 1
    out = []
    for r in ws.iter_rows(min_row=first, max_col=maxc, values_only=True):
        if not any(k < len(r) and s(r[k]) for k in keys):
            continue
        out.append([cv(r[c]) if c < len(r) else cv(None) for _, c, cv in cols])
    # одинаковые значения, записанные по-разному («Эм» / «ЭМ», «отмена» / «Отмена»), приводим к самому частому написанию
    for name in cfg['canon']:
        c = cfg['out'].index(name)
        cnt = {}
        for row in out:
            cnt.setdefault(low(row[c]), Counter())[row[c]] += 1
        best = {k: v.most_common(1)[0][0] for k, v in cnt.items()}
        for row in out:
            row[c] = best[low(row[c])]
    return out


def commit_time(path):
    try:
        t = subprocess.run(['git', 'log', '-1', '--format=%ct', '--', path], cwd=ROOT,
                           capture_output=True, text=True).stdout.strip()
        if t:
            return int(t)
    except Exception:
        pass
    return int(os.path.getmtime(path))


def find_files():
    found = {}  # (вид, год) -> [файлы]
    for f in sorted(glob.glob(os.path.join(INPUT, '*.xls*'))):
        name = os.path.basename(f)
        if name.startswith('~$'):
            continue
        try:
            wb = load_workbook(f, read_only=True, data_only=True)
            ws = wb.worksheets[0]
            kind, year = detect(ws)
        except Exception as e:
            print('пропуск (не открывается):', name, e)
            continue
        if not kind:
            print('пропуск (не похож ни на исполнение, ни на план закупок, ни на бюджетную заявку):', name)
            continue
        try:
            first, pos = locate(kind, ws)
        except ValueError as e:
            print(f'пропуск ({name}: вид «{kind}», {year}): {e}')
            continue
        print(f'{name}: вид «{kind}», {year}, данные с строки {first}, колонки: ' +
              ' '.join(f'{k}={gl(v + 1)}' for k, v in pos.items() if not k.startswith('_')))
        found.setdefault((kind, year), []).append(f)
    return found


def main():
    data = {'years': {}}
    if os.path.exists(OUT):
        old = json.load(open(OUT, encoding='utf-8'))
        if isinstance(old.get('years'), dict):
            data = old
    found = find_files()
    if not found:
        print('В input/ нет подходящих файлов, data.json не меняю.')
        return 0
    today = datetime.date.today().strftime('%d.%m.%Y')
    used = []
    for (kind, year), files in sorted(found.items()):
        f = max(files, key=commit_time)
        wb = load_workbook(f, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        first, pos = locate(kind, ws)
        rows = read_kind(kind, ws, first, pos)
        block = {'upd': today, 'rows': rows}
        if kind == 'print':
            block['ver'] = s(cell(ws, 'C4'))
        data['years'].setdefault(year, {})[kind] = block
        print(f'{kind} {year}: {os.path.basename(f)} -> строк {len(rows)}')
        used += files
    with open(OUT, 'w', encoding='utf-8') as fh:
        json.dump(data, fh, ensure_ascii=False, separators=(',', ':'))
    # после успешной сборки input/ очищается (README и нераспознанные файлы остаются)
    for f in used:
        try:
            os.remove(f)
            print('удалён из input/:', os.path.basename(f))
        except OSError as e:
            print('не удалось удалить', f, e)
    return 0


if __name__ == '__main__':
    sys.exit(main())
