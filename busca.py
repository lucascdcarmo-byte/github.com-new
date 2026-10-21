import requests, json, re
S = requests.Session()
S.headers['User-Agent'] = 'Mozilla/5.0'
S.cookies.set('security', 'true', domain='www.caixa.gov.br')
B = 'https://www.caixa.gov.br/poder-publico/modernizacao-gestao/sinapi/Documents/busca/'
datas = sorted(S.get(B + 'data-preco.json', timeout=120).json()['datas_preco'])
mes = datas[-1]
d = S.get(B + 'dados/%s/%s-descricoes.json' % (mes, mes), timeout=300).json()
an = S.get(B + 'dados/%s/%s-analitico.json' % (mes, mes), timeout=300).json()
comps = set(x['c'] for x in an)
out = open('dados/busca_sinapi.txt', 'w')
out.write('mes %s\n' % mes)
pad = re.compile(r'PONTO|TOMADA|INTERRUPTOR|DISJUNTOR|QUADRO DE DISTRIBUI|LUMINARIA|LUMINÁRIA|DR |DPS|ELETRODUTO|CABO DE COBRE', re.I)
for x in d:
    if x['c'] in comps and pad.search(x['d']):
        out.write('%s\t%s\t%s\n' % (x['c'], x.get('u', ''), x['d']))
out.close()
