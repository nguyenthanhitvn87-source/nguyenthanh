"""Guard against the bug that made PowerPoint show 'Click to add title':
inside <a:p>, <a:pPr> must come first and <a:endParaRPr> must come last."""
import re
import sys
import zipfile

NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
bad = 0
z = zipfile.ZipFile(sys.argv[1])
for name in z.namelist():
    if not re.match(r"ppt/(slides|slideLayouts|slideMasters)/[^/]+\.xml$", name):
        continue
    xml = z.read(name).decode("utf8")
    for m in re.finditer(r"<a:p>(.*?)</a:p>", xml, re.S):
        body = m.group(1)
        tags = re.findall(r"<a:(pPr|r|br|fld|endParaRPr)[ />]", body)
        if "endParaRPr" in tags and tags[-1] != "endParaRPr":
            print(f"{name}: endParaRPr not last -> {tags}")
            bad += 1
        if "pPr" in tags and tags[0] != "pPr":
            print(f"{name}: pPr not first -> {tags}")
            bad += 1
print("paragraph order:", "FAIL %d" % bad if bad else "OK")
sys.exit(1 if bad else 0)
