"""Writes objstm.pdf: three pages whose catalog, page tree, pages and font all
live in object streams (only content streams, the object stream and the xref
stream are uncompressed), and there is no classic trailer.

Run: python3 make-objstm.py   (needs qpdf on PATH; writes next to this script)
A plain PDF is written by hand, then compressed with
`qpdf --object-streams=generate --deterministic-id plain.pdf objstm.pdf`.
"""
import os, subprocess, tempfile

here = os.path.dirname(os.path.abspath(__file__))
out = bytearray(b"%PDF-1.5\n")
off = {}

def obj(n, body):
    off[n] = len(out)
    out.extend(b"%d 0 obj\n" % n + body + b"\nendobj\n")

kids = b" ".join(b"%d 0 R" % (3 + 2 * i) for i in range(3))
obj(1, b"<< /Type /Catalog /Pages 2 0 R >>")
obj(2, b"<< /Type /Pages /Kids [" + kids + b"] /Count 3 >>")
for i in range(3):
    p, c = 3 + 2 * i, 4 + 2 * i
    obj(p, b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 9 0 R >> >> /Contents %d 0 R >>" % c)
    s = b"BT /F1 14 Tf 20 100 Td (Page %d text) Tj ET" % (i + 1)
    obj(c, b"<< /Length %d >>\nstream\n" % len(s) + s + b"\nendstream")
obj(9, b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
x = len(out)
out.extend(b"xref\n0 10\n0000000000 65535 f \n")
for n in range(1, 10):
    out.extend(b"%010d 00000 n \n" % off[n])
out.extend(b"trailer\n<< /Size 10 /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % x)

with tempfile.TemporaryDirectory() as tmp:
    plain = os.path.join(tmp, "plain.pdf")
    open(plain, "wb").write(out)
    subprocess.run(["qpdf", "--object-streams=generate", "--deterministic-id", plain,
                    os.path.join(here, "objstm.pdf")], check=True)
