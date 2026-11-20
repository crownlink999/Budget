#!/usr/bin/env python3
"""Собирает data.json (бюджет) из Excel-файлов в папке input/.

Файлы определяются по СОДЕРЖИМУ, а не по имени. Год берётся из самого файла:
  * «Исполнение»: первый лист называется «Исполнение 2026» (год в названии листа);
  * «План закупок»: первый лист называется «2026» (год = название листа);
  * «Бюджетная заявка»: на первом листе B3 = «Год», C3 = 2026, C4 = версия (PLAN05).
Если в input/ несколько файлов одного вида и года, берётся самый свежий.
Если файла какого-то вида/года нет, эта часть data.json остаётся прежней.
Колонки берутся по буквам Excel (как заказал владелец), но перед чтением проверяется,
что в шапке стоит ожидаемый заголовок: если файл другой раскладки, он пропускается с понятным сообщением.
После успешной сборки распознанные файлы удаляются из input/ (в репозитории это фиксирует workflow).
"""
import datetime, glob, json, os, re, subprocess, sys
from collections import Counter
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string as ci

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT = os.environ.get('BUDGET_INPUT') or os.path.join(ROOT, 'input')
OUT = os.environ.get('BUDGET_OUT') or os.path.join(ROOT, 'data.json')


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
# Описание трёх видов файлов.
#   head   : буква колонки -> (строка шапки, регулярка, которой должен соответствовать заголовок)
#   cols   : порядок и конвертер полей в data.json
#   first  : первая строка данных
#   key    : колонки; если все они пустые, строка пропускается
# ---------------------------------------------------------------------------
KINDS = {
    # «1 - Исполнение по обеспечению…», лист «Исполнение 2026»
    'exec': dict(
        first=9, maxcol='GS', key=['BA', 'BB'],
        head={'BA': (5, r'код товара'), 'BB': (5, r'наименование'), 'BC': (5, r'ед\.?\s*изм'),
              'R': (5, r'заказчик'), 'FE': (6, r'кол'), 'FF': (6, r'сумма'), 'FH': (5, r'способ'),
              'GH': (5, r'статус заявки'), 'GJ': (5, r'статус закуп'), 'GK': (5, r'дата вскрытия'),
              'GQ': (6, r'поставщик'), 'GS': (6, r'дата договора')},
        cols=[('BA', s), ('BB', s), ('BC', s), ('R', s), ('FE', n), ('FF', n), ('FH', s),
              ('GH', s), ('GJ', s), ('GK', d), ('GQ', s), ('GS', d)],
        canon=[6, 7, 8],
    ),
    # «1 … 2026 …», лист «2026»: план закупок
    'plan': dict(
        first=4, maxcol='AX', key=['M', 'N'],  # у работ и услуг кода товара нет, есть только наименование
        head={'F': (2, r'тру'), 'G': (2, r'бюджет'), 'M': (3, r'код товара'), 'N': (3, r'наименование'),
              'P': (3, r'кол'), 'Q': (3, r'цена'), 'R': (3, r'сумма'), 'U': (3, r'дата заявки'),
              'Y': (2, r'способ'), 'AW': (2, r'статус закуп'), 'AX': (2, r'несост')},
        cols=[('F', s), ('G', s), ('M', s), ('N', s), ('P', n), ('Q', n), ('R', n), ('U', d),
              ('Y', s), ('AW', s), ('AX', s)],
        canon=[0, 1, 8, 9, 10],
    ),
    # «3 - Бюджетная заявка …», лист «Отчет после корр (Analysis)»
    'print': dict(
        first=13, maxcol='AV', key=['AO'],
        head={'C': (10, r'наименование'), 'M': (10, r'наименование'), 'N': (10, r'код гпз'),
              'AO': (11, r'код товара'), 'AR': (11, r'кол'), 'AS': (11, r'цена'), 'AV': (11, r'примечание')},
        cols=[('C', s), ('M', s), ('N', s), ('AO', s), ('AR', n), ('AS', n), ('AV', s)],
        canon=[0, 1],
    ),
}


def cell(ws, ref):
    # ws.cell() в read_only-режиме медленный, но нужен только для шапки (несколько ячеек)
    return ws[ref].value


def detect(ws, wb):
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


def check_headers(kind, ws):
    cfg = KINDS[kind]
    top = max(r for r, _ in cfg['head'].values())
    maxc = ci(cfg['maxcol'])
    grid = [list(r) for r in ws.iter_rows(min_row=1, max_row=top, max_col=maxc, values_only=True)]
    bad = []
    for letter, (row, rx) in cfg['head'].items():
        i = ci(letter) - 1
        v = grid[row - 1][i] if row - 1 < len(grid) and i < len(grid[row - 1]) else None
        if not re.search(rx, low(v)):
            bad.append(f'{letter}{row}: ждали «{rx}», в файле «{s(v)[:40]}»')
    return bad


def read_kind(kind, ws):
    cfg = KINDS[kind]
    maxc = ci(cfg['maxcol'])
    idx = [(ci(letter) - 1, conv) for letter, conv in cfg['cols']]
    keys = [ci(k) - 1 for k in cfg['key']]
    out = []
    for r in ws.iter_rows(min_row=cfg['first'], max_col=maxc, values_only=True):
        if not any(k < len(r) and s(r[k]) for k in keys):
            continue
        out.append([conv(r[i]) if i < len(r) else conv(None) for i, conv in idx])
    # одинаковые значения, записанные по-разному («Эм» / «ЭМ», «отмена» / «Отмена»), приводим к самому частому написанию
    for c in cfg['canon']:
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
    found = {}  # (вид, год) -> [(файл, лист-версия)]
    seen = []
    for f in sorted(glob.glob(os.path.join(INPUT, '*.xls*'))):
        if os.path.basename(f).startswith('~$'):
            continue
        try:
            wb = load_workbook(f, read_only=True, data_only=True)
            ws = wb.worksheets[0]
            kind, year = detect(ws, wb)
        except Exception as e:
            print('пропуск (не открывается):', os.path.basename(f), e)
            continue
        if not kind:
            print('пропуск (не похож ни на исполнение, ни на план закупок, ни на бюджетную заявку):', os.path.basename(f))
            continue
        bad = check_headers(kind, ws)
        if bad:
            print(f'пропуск ({os.path.basename(f)}: вид «{kind}», {year}), шапка не совпадает с ожидаемой:')
            for b in bad:
                print('   ', b)
            continue
        found.setdefault((kind, year), []).append(f)
        seen.append(f)
    return found, seen


def main():
    data = {'years': {}}
    if os.path.exists(OUT):
        old = json.load(open(OUT, encoding='utf-8'))
        if isinstance(old.get('years'), dict):
            data = old
    found, used = find_files()
    if not found:
        print('В input/ нет подходящих файлов, data.json не меняю.')
        return 0
    today = datetime.date.today().strftime('%d.%m.%Y')
    for (kind, year), files in sorted(found.items()):
        f = max(files, key=commit_time)
        wb = load_workbook(f, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        rows = read_kind(kind, ws)
        block = {'upd': today, 'rows': rows}
        if kind == 'print':
            block['ver'] = s(cell(ws, 'C4'))
        data['years'].setdefault(year, {})[kind] = block
        print(f'{kind} {year}: {os.path.basename(f)} -> строк {len(rows)}')
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
