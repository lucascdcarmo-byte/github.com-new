import json, os, re, sys, shutil, tempfile, unicodedata, datetime, time
import requests, py7zr, openpyxl
from bs4 import BeautifulSoup

BASE = 'https://www.gov.br/dnit/pt-br/assuntos/planejamento-e-pesquisa/custos-referenciais/sistemas-de-custos/sicro/relatorios/relatorios-sicro/'
REGIOES = ['norte/norte', 'nordeste/nordeste', 'centro-oeste/centro-oeste', 'sudeste/sudeste', 'sul/sul']
UFS = {'acre': 'AC', 'alagoas': 'AL', 'amapa': 'AP', 'amazonas': 'AM', 'bahia': 'BA', 'ceara': 'CE', 'distrito-federal': 'DF',
       'espirito-santo': 'ES', 'goias': 'GO', 'maranhao': 'MA', 'mato-grosso': 'MT', 'mato-grosso-do-sul': 'MS', 'minas-gerais': 'MG',
       'para': 'PA', 'paraiba': 'PB', 'parana': 'PR', 'pernambuco': 'PE', 'piaui': 'PI', 'rio-de-janeiro': 'RJ', 'rio-grande-do-norte': 'RN',
       'rio-grande-do-sul': 'RS', 'rondonia': 'RO', 'roraima': 'RR', 'santa-catarina': 'SC', 'sao-paulo': 'SP', 'sergipe': 'SE', 'tocantins': 'TO'}
MESES = {'janeiro': 1, 'fevereiro': 2, 'marco': 3, 'abril': 4, 'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8,
         'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12}
S = requests.Session()
S.headers['User-Agent'] = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/156.0 Safari/537.36'


def norm(t):
    t = unicodedata.normalize('NFD', str(t or '')).encode('ascii', 'ignore').decode().lower()
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9 ]', ' ', t)).strip()


def links(url):
    r = S.get(url, timeout=60)
    r.raise_for_status()
    return [a['href'] for a in BeautifulSoup(r.text, 'html.parser').find_all('a', href=True)]


def referencias():
    achados = {}
    for reg in REGIOES:
        for h in links(BASE + reg):
            m = re.search(r'/relatorios-sicro/[^/]+/([a-z-]+)/(\d{4})/[^/]+/([a-z]+)-(\d{4})(-revisado)?/?$', h)
            if not m or m.group(1) not in UFS:
                continue
            uf, ano, mes, rev = UFS[m.group(1)], int(m.group(2)), MESES.get(m.group(3)), 1 if m.group(5) else 0
            if not mes:
                continue
            chave = (ano, mes, rev)
            if uf not in achados or chave > achados[uf][0]:
                achados[uf] = (chave, h if h.startswith('http') else 'https://www.gov.br' + h)
    return achados


def arquivo7z(pagina):
    for h in links(pagina):
        if h.lower().endswith('.7z') or h.lower().endswith('.zip'):
            return h if h.startswith('http') else 'https://www.gov.br' + h
    return None


def num(v):
    if v is None or v == '':
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace('R$', '').replace(' ', '')
    if re.search(r',\d{1,4}$', s):
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def ler_planilhas(pasta):
    melhor = []
    inventario = []
    for raiz, _, nomes in os.walk(pasta):
        for n in nomes:
            caminho = os.path.join(raiz, n)
            inventario.append(os.path.relpath(caminho, pasta))
            if not n.lower().endswith(('.xlsx', '.xlsm')):
                continue
            try:
                wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
            except Exception as e:
                print('  nao abriu', n, e)
                continue
            for ws in wb.worksheets:
                linhas = list(ws.iter_rows(values_only=True))
                for i, row in enumerate(linhas[:40]):
                    cab = [norm(c) for c in row]
                    ic = next((j for j, c in enumerate(cab) if c.startswith('codigo')), -1)
                    idd = next((j for j, c in enumerate(cab) if 'descricao' in c), -1)
                    iu = next((j for j, c in enumerate(cab) if c in ('unidade', 'und', 'un') or c.startswith('unid')), -1)
                    ip = next((j for j, c in enumerate(cab) if 'custo unitario' in c or 'preco unitario' in c or c.startswith('custo')), -1)
                    if ic < 0 or idd < 0 or ip < 0:
                        continue
                    itens = []
                    for r in linhas[i + 1:]:
                        if not r or len(r) <= max(ic, idd, ip):
                            continue
                        c = str(r[ic] or '').strip()
                        d = str(r[idd] or '').strip()
                        p = num(r[ip])
                        if not c or not d or p is None:
                            continue
                        itens.append({'c': c, 'd': re.sub(r'\s+', ' ', d), 'u': (str(r[iu]).strip() if iu >= 0 and r[iu] else ''), 'p': round(p, 4)})
                    print('  planilha', n, '| aba', ws.title, '| cabecalho', [str(x) for x in row if x], '| itens', len(itens))
                    if 'composic' in norm(n) and len(itens) > len(melhor):
                        melhor = itens
                    break
    return melhor, inventario



def ultimo_num(row):
    for v in reversed(row):
        n = num(v)
        if n is not None:
            return n
    return None


def ler_analitico(linhas):
    comps = {}
    atual = None
    secao = ''
    ant = ''
    for row in linhas:
        row = list(row or [])
        c0 = str(row[0] if row else '' or '').strip()
        txt = norm(' '.join(str(x) for x in row if x is not None))
        if re.fullmatch(r'\d{7}', c0) and 'custo unitario de referencia' in ant:
            atual = {'exec': 0.0, 'fic': 0.0, 'mat': 0.0, 'aux': [], 'tf': 0.0, 'tr': 0.0}
            comps[c0] = atual
            secao = ''
            ant = txt
            continue
        ant = txt
        if atual is None:
            continue
        if re.match(r'^[a-f] ', norm(c0) + ' ') and ' - ' in c0:
            secao = norm(c0)[0]
            continue
        v = ultimo_num(row)
        if 'custo unitario de execucao' in txt:
            atual['exec'] = v or 0.0
        elif 'custo do fic' in txt:
            atual['fic'] = v or 0.0
        elif 'custo unitario total de material' in txt:
            atual['mat'] = v or 0.0
        elif 'custo unitario total de tempo fixo' in txt:
            atual['tf'] = v or 0.0
        elif 'custo unitario total de transporte' in txt:
            atual['tr'] = v or 0.0
        elif secao == 'd' and re.fullmatch(r'\d{7}', c0) and v:
            atual['aux'].append((c0, v))
    return comps


def fracao_material(comps, c, memo, pilha):
    if c in memo:
        return memo[c]
    x = comps.get(c)
    if x is None or c in pilha:
        return None
    pilha.add(c)
    mat = x['mat']
    tot = x['exec'] + x['fic'] + x['mat'] + x['tf'] + x['tr']
    for ca, v in x['aux']:
        f = fracao_material(comps, ca, memo, pilha)
        mat += v * (f if f is not None else 0.0)
        tot += v
    pilha.discard(c)
    memo[c] = (mat / tot) if tot > 0 else None
    return memo[c]


def dividir(itens, comps):
    memo = {}
    n = 0
    for it in itens:
        f = fracao_material(comps, it['c'], memo, set())
        if f is None:
            continue
        m = round(it['p'] * f, 2)
        it['m'] = m
        it['s'] = round(it['p'] - m, 2)
        n += 1
    return n

def main():
    os.makedirs('dados/sicro', exist_ok=True)
    indice_path = 'dados/sicro/indice.json'
    indice = json.load(open(indice_path)) if os.path.exists(indice_path) else {}
    refs = referencias()
    print('UFs encontradas:', len(refs))
    for uf, ((ano, mes, rev), pagina) in sorted(refs.items()):
        ref = '%04d-%02d' % (ano, mes) + ('R' if rev else '')
        if indice.get(uf, {}).get('ref') == ref and indice.get(uf, {}).get('v') == 2 and os.path.exists('dados/sicro/%s.json' % uf):
            print(uf, ref, 'ja atualizado')
            continue
        print(uf, ref, pagina)
        url = arquivo7z(pagina)
        if not url:
            print('  sem arquivo na pagina')
            continue
        tmp = tempfile.mkdtemp()
        try:
            dest = os.path.join(tmp, 'a.7z')
            for tentativa in range(5):
                try:
                    with S.get(url, stream=True, timeout=300) as r:
                        r.raise_for_status()
                        with open(dest, 'wb') as f:
                            for bloco in r.iter_content(1 << 20):
                                f.write(bloco)
                    break
                except Exception as e:
                    print('  tentativa', tentativa + 1, 'falhou:', e)
                    if tentativa == 4:
                        raise
                    time.sleep(20 * (tentativa + 1))
            ext = os.path.join(tmp, 'x')
            if url.lower().endswith('.zip'):
                shutil.unpack_archive(dest, ext)
            else:
                with py7zr.SevenZipFile(dest) as z:
                    z.extractall(ext)
            itens, inventario = ler_planilhas(ext)
            print('  arquivos:', inventario)
            if not itens:
                print('  nenhuma composicao lida')
                continue
            for raiz, _, nomes in os.walk(ext):
                for n in nomes:
                    if 'analitico de composic' in norm(n) and n.lower().endswith('.xlsx'):
                        wb = openpyxl.load_workbook(os.path.join(raiz, n), read_only=True, data_only=True)
                        comps = ler_analitico(wb.worksheets[0].iter_rows(values_only=True))
                        print('  analitico:', len(comps), 'composicoes | MAT/SERV em', dividir(itens, comps), 'itens')
            json.dump(itens, open('dados/sicro/%s.json' % uf, 'w'), ensure_ascii=False, separators=(',', ':'))
            indice[uf] = {'v': 2, 'ref': ref, 'itens': len(itens), 'fonte': url, 'atualizado': datetime.date.today().isoformat()}
            json.dump(indice, open(indice_path, 'w'), ensure_ascii=False, indent=1)
        except Exception as e:
            print('  erro', uf, e)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
    print('feito')


if __name__ == '__main__':
    main()
