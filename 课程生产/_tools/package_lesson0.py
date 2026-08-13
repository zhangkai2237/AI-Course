from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import sys

folder=Path(sys.argv[1]); output=Path(sys.argv[2])
files=[]
files += sorted(folder.glob('*.md'))
files += sorted(folder.glob('*.docx'))
files += [folder/'06_正式PPT.pptx']
with ZipFile(output,'w',ZIP_DEFLATED) as z:
    for f in files:
        z.write(f,f.name)
