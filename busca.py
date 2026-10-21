import requests, re
from bs4 import BeautifulSoup
S = requests.Session()
S.headers['User-Agent'] = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/156.0 Safari/537.36'
out = open('dados/busca_goinfra.txt', 'w')
vistos = set()
fila = ['https://goias.gov.br/goinfra/', 'https://goias.gov.br/goinfra/wp-content/uploads/sites/77/tabela-de-composicao/85/arquivos/', 'https://goias.gov.br/goinfra/wp-content/uploads/sites/77/tabela-de-composicao/']
n = 0
while fila and n < 40:
    u = fila.pop(0)
    if u in vistos:
        continue
    vistos.add(u); n += 1
    try:
        r = S.get(u, timeout=60)
    except Exception as e:
        out.write('ERRO %s %s\n' % (u, e)); continue
    out.write('### %s %s %s\n' % (r.status_code, u, r.headers.get('content-type')))
    if 'html' not in (r.headers.get('content-type') or ''):
        continue
    for a in BeautifulSoup(r.text, 'html.parser').find_all('a', href=True):
        h = a['href']; t = ' '.join(a.get_text().split())
        if re.search(r'tabela|custo|composi|preco|preço|referenc|\.xls|\.zip', h + ' ' + t, re.I):
            out.write('%s | %s\n' % (t[:120], h))
            if h.startswith('https://goias.gov.br/goinfra') and not re.search(r'\.(pdf|xls|xlsx|zip)$', h, re.I):
                fila.append(h)
out.close()
