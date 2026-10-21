import requests, xlrd, io
S = requests.Session()
S.headers['User-Agent'] = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/156.0 Safari/537.36'
B = 'https://goias.gov.br/goinfra/wp-content/uploads/sites/77/tabela-de-composicao/337/'
out = open('dados/amostra_goinfra.txt', 'w')
for nome, lim in [('Custo_Referencial_de_Servicos.xls', 60), ('Relatorio_de_Composicao_do_Servi.xls', 160), ('Material.xls', 25), ('Mao_de_obra.xls', 25)]:
    try:
        r = S.get(B + nome, timeout=300)
        out.write('### %s %s %d bytes %s\n' % (nome, r.status_code, len(r.content), r.headers.get('content-type')))
        wb = xlrd.open_workbook(file_contents=r.content)
        for sh in wb.sheets():
            out.write('## aba %s linhas %d colunas %d\n' % (sh.name, sh.nrows, sh.ncols))
            for i in range(min(lim, sh.nrows)):
                out.write('%d\t%s\n' % (i, '\t'.join(str(c.value).replace('\n', ' ') for c in sh.row(i))))
            if sh.nrows > lim:
                out.write('...\n')
                for i in range(max(lim, sh.nrows - 8), sh.nrows):
                    out.write('%d\t%s\n' % (i, '\t'.join(str(c.value).replace('\n', ' ') for c in sh.row(i))))
    except Exception as e:
        out.write('ERRO %s %s\n' % (nome, e))
out.close()
