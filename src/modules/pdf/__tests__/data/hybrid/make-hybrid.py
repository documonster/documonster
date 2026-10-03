"""Writes hybrid-reference.pdf: a hybrid-reference file (ISO 32000-1 §7.5.8.4).

The classic xref table lists only the uncompressed objects (1 catalog, 2 pages,
4 content stream, 6 object stream, 7 xref stream) and omits 3 and 5; the
trailer's /XRefStm names xref stream 7, which lists the page (3) and its font
(5) as compressed objects inside object stream 6. A reader that ignores
/XRefStm cannot find the page. qpdf cannot write this layout, hence by hand.

Run: python3 make-hybrid.py   (writes next to this script)
"""
import os, struct

out = bytearray(b"%PDF-1.5\n%\xe2\xe3\xcf\xd3\n")
off = {}

def obj(n, body):
    off[n] = len(out)
    out.extend(b"%d 0 obj\n" % n + body + b"\nendobj\n")

obj(1, b"<< /Type /Catalog /Pages 2 0 R >>")
obj(2, b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>")
content = b"BT /F1 24 Tf 72 700 Td (Original hybrid text) Tj ET"
obj(4, b"<< /Length %d >>\nstream\n" % len(content) + content + b"\nendstream")
p3 = b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>"
p5 = b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
header = b"3 0 5 %d " % (len(p3) + 1)
data = header + p3 + b" " + p5
obj(6, b"<< /Type /ObjStm /N 2 /First %d /Length %d >>\nstream\n" % (len(header), len(data)) + data + b"\nendstream")

rows = b""
for n in range(8):
    if n == 0:
        rows += struct.pack(">BHB", 0, 0, 255)
    elif n == 3:
        rows += struct.pack(">BHB", 2, 6, 0)
    elif n == 5:
        rows += struct.pack(">BHB", 2, 6, 1)
    elif n == 7:
        rows += struct.pack(">BHB", 1, len(out), 0)
    else:
        rows += struct.pack(">BHB", 1, off[n], 0)
xs = len(out)
obj(7, b"<< /Type /XRef /Size 8 /W [1 2 1] /Length %d >>\nstream\n" % len(rows) + rows + b"\nendstream")

xref = len(out)
# Subsections skip the compressed objects 3 and 5, so the table says nothing
# about them and a reader must consult /XRefStm.
out.extend(b"xref\n0 3\n0000000000 65535 f \n")
for n in (1, 2):
    out.extend(b"%010d 00000 n \n" % off[n])
out.extend(b"4 1\n%010d 00000 n \n" % off[4])
out.extend(b"6 2\n")
for n in (6, 7):
    out.extend(b"%010d 00000 n \n" % off[n])
out.extend(b"trailer\n<< /Size 8 /Root 1 0 R /XRefStm %d >>\nstartxref\n%d\n%%%%EOF\n" % (xs, xref))
open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "hybrid-reference.pdf"), "wb").write(out)
