import json, os, py7zr, requests, openpyxl, tempfile, time
url = json.load(open('dados/sicro/indice.json'))['GO']['fonte']
S = requests.Session()
S.headers['User-Agent'] = 'Mozilla/5.0'
tmp = tempfile.mkdtemp()
dest = os.path.join(tmp, 'a.7z')
for t in range(5):
    try:
        with S.get(url, stream=True, timeout=300) as r:
            r.raise_for_status()
            with open(dest, 'wb') as f:
                for b in r.iter_content(1 << 20):
                    f.write(b)
        break
    except Exception as e:
        time.sleep(20)
with py7zr.SevenZipFile(dest) as z:
    z.extractall(os.path.join(tmp, 'x'))
out = open('dados/sicro/amostra.txt', 'w')
for raiz, _, nomes in os.walk(os.path.join(tmp, 'x')):
    for n in nomes:
        if 'Analítico de Composições' in n and n.endswith('.xlsx'):
            wb = openpyxl.load_workbook(os.path.join(raiz, n), read_only=True, data_only=True)
            for ws in wb.worksheets:
                out.write('### %s | %s | linhas %s\n' % (n, ws.title, ws.max_row))
                for i, row in enumerate(ws.iter_rows(values_only=True)):
                    if i >= 260:
                        break
                    out.write('%d\t%s\n' % (i, '\t'.join('' if c is None else str(c).replace('\n', ' ') for c in row)))
out.close()
