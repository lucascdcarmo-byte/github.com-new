import json, os, re, html, datetime, time, unicodedata
import requests, xlrd

S = requests.Session()
S.headers['User-Agent'] = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/156.0 Safari/537.36'
WP = 'https://goias.gov.br/goinfra/wp-json/wp/v2/pages'
ARQ = 'https://goias.gov.br/goinfra/wp-content/uploads/sites/77/tabela-de-composicao/%s/Custo_Referencial_de_Servicos.xls'
MESES = {'janeiro': 1, 'fevereiro': 2, 'marco': 3, 'abril': 4, 'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8,
         'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12}


def norm(t):
    t = unicodedata.normalize('NFD', str(t or '')).encode('ascii', 'ignore').decode().lower()
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9 ]', ' ', t)).strip()


def get(url, **kw):
    for t in range(5):
        try:
            r = S.get(url, timeout=300, **kw)
            r.raise_for_status()
            return r
        except Exception as e:
            print('  tentativa', t + 1, 'falhou:', e)
            if t == 4:
                raise
            time.sleep(15 * (t + 1))


def tabelas():
    achadas = {}
    for termo in ['desoneracao', 'desoneração']:
        for pg in range(1, 6):
            r = S.get(WP, params={'search': termo, 'per_page': 100, 'page': pg, '_fields': 'title'}, timeout=120)
            if r.status_code != 200:
                break
            lote = r.json()
            if not lote:
                break
            for p in lote:
                t = norm(html.unescape(p['title']['rendered']))
                m = re.search(r'tabela (\d+) (com|sem) desoneracao (\w+) de (\d{4})', t)
                if not m or m.group(3) not in MESES:
                    continue
                num, des, ano, mes = m.group(1), 'cd' if m.group(2) == 'com' else 'sd', int(m.group(4)), MESES[m.group(3)]
                achadas[num] = (des, ano, mes)
    return achadas


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v or '').strip()
    if not s:
        return None
    if re.search(r',\d{1,4}$', s):
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def ler_servicos(conteudo):
    sh = xlrd.open_workbook(file_contents=conteudo).sheet_by_index(0)
    titulo = ''
    cab = None
    itens = []
    for i in range(sh.nrows):
        row = [c.value for c in sh.row(i)]
        txt = norm(' '.join(str(x) for x in row))
        if 'tabela de custos' in txt and not titulo:
            titulo = ' '.join(str(x).strip() for x in row if str(x).strip())
        if cab is None:
            n = [norm(x) for x in row]
            if any(x.startswith('codigo') for x in n) and 'unidade' in n and 'total' in n:
                cab = {
                    'c': next(j for j, x in enumerate(n) if x.startswith('codigo')),
                    'd': next(j for j, x in enumerate(n) if x.startswith('servico')),
                    'u': n.index('unidade'),
                    'm': next(j for j, x in enumerate(n) if x.startswith('material')),
                    's': next(j for j, x in enumerate(n) if x.startswith('mao de obra')),
                    'p': n.index('total')
                }
            continue
        c = str(row[cab['c']]).strip()
        d = re.sub(r'\s+', ' ', str(row[cab['d']]).strip())
        u = str(row[cab['u']]).strip()
        p = num(row[cab['p']])
        if not re.fullmatch(r'\d{4,8}', c) or not d or not u or u.upper() == 'S/U' or not p:
            continue
        m, s = num(row[cab['m']]) or 0.0, num(row[cab['s']]) or 0.0
        m = round(m, 2)
        itens.append({'c': c, 'd': d, 'u': u, 'p': round(p, 2), 'm': m, 's': round(p - m, 2)})
    return titulo, itens


def main():
    os.makedirs('dados/goinfra', exist_ok=True)
    ip = 'dados/goinfra/indice.json'
    indice = json.load(open(ip)) if os.path.exists(ip) else {}
    achadas = tabelas()
    print('tabelas com desoneracao no titulo:', achadas)
    melhor = {}
    for numt, (des, ano, mes) in achadas.items():
        if des not in melhor or (ano, mes, int(numt)) > melhor[des][0]:
            melhor[des] = ((ano, mes, int(numt)), numt)
    for des, ((ano, mes, _), numt) in sorted(melhor.items()):
        ref = '%04d-%02d' % (ano, mes)
        if indice.get(des, {}).get('tabela') == numt and os.path.exists('dados/goinfra/%s.json' % des):
            print(des, 'T' + numt, ref, 'ja atualizado')
            continue
        url = ARQ % numt
        print(des, 'T' + numt, ref, url)
        titulo, itens = ler_servicos(get(url).content)
        print('  titulo:', titulo, '| itens:', len(itens))
        if 'OBRAS CIVIS' not in titulo.upper() or len(itens) < 200:
            print('  ignorada: nao e a tabela de obras civis')
            continue
        json.dump(itens, open('dados/goinfra/%s.json' % des, 'w'), ensure_ascii=False, separators=(',', ':'))
        indice[des] = {'tabela': numt, 'ref': ref, 'titulo': titulo, 'itens': len(itens), 'fonte': url, 'atualizado': datetime.date.today().isoformat()}
        json.dump(indice, open(ip, 'w'), ensure_ascii=False, indent=1)
    print('feito')


if __name__ == '__main__':
    main()
