"""fix_bas_encoding.py
Corrige codificacao de .bas VBA para Windows-1252.
USO: python scripts/fix_bas_encoding.py [--check]
Agentes de IA salvam em UTF-8; Word exige CP1252.
"""
import os, re, sys

SRC = os.path.join(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__))), "source", "main")

A_AC=chr(0xE1);E_AC=chr(0xE9);I_AC=chr(0xED);O_AC=chr(0xF3);U_AC=chr(0xFA)
A_TL=chr(0xE3);O_TL=chr(0xF5);C_CE=chr(0xE7);A_UP=chr(0xC1)

DBL=[("a??o","a"+C_CE+A_TL+"o"),("A??O",chr(0xC7)+chr(0xC3)+"O"),
     ("a??es","a"+C_CE+O_TL+"es"),("A??ES",chr(0xC7)+chr(0xD5)+"ES"),
     ("??o",C_CE+A_TL+"o"),("??O",chr(0xC7)+chr(0xC3)+"O"),
     ("??es",C_CE+O_TL+"es"),("??ES",chr(0xC7)+chr(0xD5)+"ES")]

SGL=[("REVIS?O","REVIS"+A_TL+"O"),("revis?o","revis"+A_TL+"o"),
     ("Revis?o","Revis"+A_TL+"o"),("CONTE?DO","CONTE"+U_AC+"DO"),
     ("conte?do","conte"+U_AC+"do"),("DIAGN?STICO","DIAGN"+O_AC+"STICO"),
     ("Diagn?stico","Diagn"+O_AC+"stico"),("CONCLU?DA","CONCLU"+I_AC+"DA"),
     ("conclu?da","conclu"+I_AC+"da"),("CONSTR?I","CONSTR"+O_AC+"I"),
     ("PAR?GRAFOS","PAR"+A_UP+"GRAFOS"),("PAR?GRAFO","PAR"+A_UP+"GRAFO"),
     ("par?grafos","par"+A_AC+"grafos"),("par?grafo","par"+A_AC+"grafo")]

def fix_q(t):
    c=list(t);i=0
    while i<len(c):
        if c[i]!='?':i+=1;continue
        a=c[i+1] if i<len(c)-1 else ' '
        if a in'oO'and(i+2>=len(c)or not c[i+2].isalpha()):c[i]=A_TL
        elif a in'eE':c[i]=O_TL
        elif a in'vV':c[i]=A_AC
        elif a in'dDbBlLmM':c[i]=U_AC
        elif a in'nNgG':c[i]=I_AC
        elif a in'rRpP':c[i]=O_AC
        elif a in'tTsS':c[i]=E_AC
        elif a in'fF':c[i]=A_AC
        elif a in'cC':c[i]=chr(0xEA)
        else:c[i]=E_AC
        i+=1
    return ''.join(c)

def fix_text(t):
    t=t.replace('\ufffd','?')
    if '?'not in t:return t,0
    n=t.count('?')
    for p,r in DBL:t=t.replace(p,r)
    for p,r in SGL:t=t.replace(p,r)
    t=fix_q(t)
    return t,n-t.count('?')

def is_cp1252_ok(raw):
    if raw[:3]==b'\xef\xbb\xbf':return False
    try:
        t=raw.decode('cp1252')
        if'\ufffd'in t:return False
        return len(re.findall(r'[A-Za-z]\?[A-Za-z]',t))==0
    except:return False

def main():
    chk='--check'in sys.argv
    print('Checking...'if chk else'Fixing...')
    for n in sorted(os.listdir(SRC)):
        if not n.endswith('.bas'):continue
        p=os.path.join(SRC,n)
        raw=open(p,'rb').read()
        if is_cp1252_ok(raw):
            print('  %s: OK'%n);continue
        if chk:
            print('  %s: NEEDS FIX'%n);continue
        bom=raw[:3]==b'\xef\xbb\xbf'
        if bom:raw=raw[3:]
        txt=raw.decode('utf-8',errors='replace')
        txt,cnt=fix_text(txt)
        with open(p,'wb')as f:f.write(txt.encode('cp1252'))
        print('  %s: fixed %d chars'%(n,cnt))

if __name__=='__main__':main()