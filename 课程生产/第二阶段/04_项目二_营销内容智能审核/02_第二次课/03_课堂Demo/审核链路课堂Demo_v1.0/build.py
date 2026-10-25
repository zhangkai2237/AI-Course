from pathlib import Path
p=Path(__file__).resolve().parent
s=(p/'shell.html').read_text()
for token,name in [('/*__CSS__*/','style.css'),('/*__ENGINE__*/','engine.js'),('/*__APP__*/','app.js')]:s=s.replace(token,(p/name).read_text())
(p/'index.html').write_text(s)
print('Built standalone index.html:',len(s.encode()),'bytes')
