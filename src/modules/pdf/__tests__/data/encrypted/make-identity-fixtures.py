# Builds V4/R4 AES-128 PDFs where /StrF or /StmF is /Identity, reusing the
# /Encrypt and /ID of aes-128.pdf (file key dc10c67b…, password "user"/"owner").
import hashlib, os, sys
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.primitives import padding
KEY=bytes.fromhex('dc10c67bf7cb5955bc9bfb81b55ea3f1')
ID='01a5b35bb8febf63d0d64a0e26833949'
O='0ba3835f88f90388e74e54584125ce142be0de24c6b0d37746e075b891756671'
U='058a33f7ee2c853b1a486d5a2037aca40021446990b9e4114071a4d9104984c1'
def objkey(n,g): return hashlib.md5(KEY+n.to_bytes(3,'little')+g.to_bytes(2,'little')+b'sAlT').digest()
def aes(data,n,g=0):
    iv=b'\x11'*16; p=padding.PKCS7(128).padder(); d=p.update(data)+p.finalize()
    e=Cipher(algorithms.AES(objkey(n,g)),modes.CBC(iv)).encryptor(); return iv+e.update(d)+e.finalize()
def build(strf, stmf, out):
    enc_str = strf!='Identity'; enc_stm = stmf!='Identity'
    def s(text,n): return b'<'+(aes(text,n) if enc_str else text).hex().encode()+b'>'
    content=b'BT /F1 24 Tf 72 700 Td (Base text) Tj ET'
    stm = aes(content,4) if enc_stm else content
    objs={
     1:b'<< /Type /Catalog /Pages 2 0 R >>',
     2:b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
     3:b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>',
     4:b'<< /Length %d >>\nstream\n'%len(stm)+stm+b'\nendstream',
     5:b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
     6:b'<< /Title '+s(b'Secret title',6)+b' >>',
     7:('<< /Filter /Standard /V 4 /R 4 /Length 128 /P -4 /O <%s> /U <%s> /CF << /StdCF << /AuthEvent /DocOpen /CFM /AESV2 /Length 16 >> >> /StrF /%s /StmF /%s >>'%(O,U,strf,stmf)).encode(),
    }
    b=bytearray(b'%PDF-1.6\n%\xe2\xe3\xcf\xd3\n'); off={}
    for n,body in objs.items():
        off[n]=len(b); b+=b'%d 0 obj\n'%n+body+b'\nendobj\n'
    x=len(b); b+=b'xref\n0 8\n0000000000 65535 f \n'
    for n in range(1,8): b+=b'%010d 00000 n \n'%off[n]
    b+=('trailer << /Size 8 /Root 1 0 R /Info 6 0 R /Encrypt 7 0 R /ID [<%s><%s>] >>\nstartxref\n%d\n%%%%EOF\n'%(ID,ID,x)).encode()
    open(out,'wb').write(b)
d=sys.argv[1]
build('StdCF','Identity',os.path.join(d,'aes-128-identity-streams.pdf'))
build('Identity','StdCF',os.path.join(d,'aes-128-identity-strings.pdf'))
