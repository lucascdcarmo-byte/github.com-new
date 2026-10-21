import requests, re, json
S = requests.Session()
S.headers['User-Agent'] = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/156.0 Safari/537.36'
out = open('dados/busca_goinfra.txt', 'w')
B = 'https://goias.gov.br/goinfra/wp-json/wp/v2/'
paginas = {}
for termo in ['tabela', 'custo', 'composi', 'preço', 'referencial']:
    for tipo in ['pages', 'posts']:
        try:
            r = S.get(B + tipo, params={'search': termo, 'per_page': 50, '_fields': 'id,link,title'}, timeout=60)
            for p in r.json():
                paginas[p['link']] = p['title']['rendered']
        except Exception as e:
            out.write('ERRO %s %s %s\n' % (tipo, termo, e))
for l, t in paginas.items():
    out.write('PAG %s | %s\n' % (t, l))
for l in list(paginas):
    try:
        h = S.get(l, timeout=60).text
    except Exception:
        continue
    achados = set(re.findall(r'(https?://[^"\'\s<>]*(?:tabela-de-composicao|\.xlsx?|\.zip)[^"\'\s<>]*)', h, re.I))
    achados |= set(re.findall(r'<iframe[^>]+src="([^"]+)"', h, re.I))
    for a in sorted(achados):
        out.write('  %s -> %s\n' % (l, a))
out.close()
