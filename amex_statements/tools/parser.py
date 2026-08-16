import pdfplumber, re, os, glob, json

U='/root/.claude/uploads/6a198810-ddf7-51b5-bd0c-94dd1d9e2703'
MONTHS={'January':1,'February':2,'March':3,'April':4,'May':5,'June':6,'July':7,
        'August':8,'September':9,'October':10,'November':11,'December':12}
MONTH_RE=r'(January|February|March|April|May|June|July|August|September|October|November|December)'
MONEY=re.compile(r'\d{1,3}(?:,\d{3})*\.\d{2}')

def money(s): return float(s.replace(',',''))

# Map the 10 card-64003 statements (exclude the 2019 duplicate/outlier)
FILES=[
 '81189b6f-20200622','72237492-20200722','30fb74a9-20200822','415e11c8-20200922',
 'd10366ca-20201022','5399fb78-20201122','49c051a6-20201222',
 '0a5e4994-20210122','306bf838-20210222','2476b152-20210322','09748a5f-20210422',
 'c158dc52-20210522','74b02585-20210622','af4ae41f-20210722','3c636159-20210822',
 'fb78491a-20210922','98e06217-20211022','c34e5402-20211122','bb21ec52-20211222',
 '2ff31cd5-20220122_1','f0e210f3-20220222','68c1a378-20220322_1','bf8d9a66-20220422',
 '3d5be94b-20220522','90a8ed32-20220622','9e62cb80-20220722','846a9f10-20220822',
 'bce499c1-20220922','0d6af30a-20221022','b15a065f-20221122','8639e389-20221222_2',
 'f0aa1a40-20230122_1','c43cde76-20230222_1','21fc6f94-20230322','eb0ec563-20230422',
 '521e8837-20230522','7e5db08b-20230622','483a6040-20230722']

def stmt_meta(text):
    d=re.search(MONTH_RE+r' (\d{1,2}), (\d{4})', text)
    mon,day,yr=d.group(1),int(d.group(2)),int(d.group(3))
    return yr, MONTHS[mon], day

def parse(path):
    with pdfplumber.open(path) as pdf:
        pages=[pg.extract_text() for pg in pdf.pages]
    text='\n'.join(pages)
    syr,smon,sday=stmt_meta(pages[0])
    source=f'{syr:04d}-{smon:02d}-{sday:02d}'
    lines=text.split('\n')

    txns=[]
    expected={'PAYMENT':None,'CHARGES':None,'DEFERRED':None}
    exp_std={}   # cardholder -> total
    section=None
    cardholder=None
    stop=False
    cur=None  # current txn dict

    def month_line(l):
        m=re.match('^'+MONTH_RE+r'\s+(\d{1,2})\b(.*)$', l)
        return m

    def flush():
        nonlocal cur
        if cur is not None:
            txns.append(cur); cur=None

    i=0
    while i < len(lines):
        raw=lines[i]; l=raw.strip()
        # stop sections
        if l.startswith('Deferred Credit Plan Summary') or l.startswith('Your Account Reward Points') or l=='Cardmember information':
            stop=True; flush()
        if stop:
            i+=1; continue
        # skip page headers / noise
        if (l.startswith('Statement of Account') or l.startswith('Page ') or
            l.startswith('Cardmember Name') or l.startswith('Details Foreign') or
            l.startswith('VICKI ZERTOPOULOS xxxx') or l.startswith('Card Number') or l==''):
            i+=1; continue
        # section headers
        if l.startswith('Payments Section'):
            flush(); section='PAYMENT'; cardholder='(account)'; i+=1; continue
        m=re.match(r'New Standard Transactions for\s*(MRS|MR|MISS|MS)?\s*(.+)$', l)
        if m:
            flush(); section='STANDARD'; cardholder=m.group(2).strip(); i+=1; continue
        if l.startswith('Account Charges'):
            flush(); section='CHARGES'; cardholder='(account)'; i+=1; continue
        m=re.match(r'New Deferred Credit Plans(?: for\s*(MRS|MR|MISS|MS)?\s*(.+))?$', l)
        if m:
            flush(); section='DEFERRED'; cardholder=(m.group(2).strip() if m.group(2) else cardholder); i+=1; continue
        # totals
        def total_cr():  # trailing CR on next line makes the printed total a net credit
            return i+1 < len(lines) and lines[i+1].strip()=='CR'
        m=re.match(r'Total of Payments received\s+('+MONEY.pattern+')', l)
        if m: flush(); expected['PAYMENT']=money(m.group(1)); i+=1; continue
        m=re.match(r'Total of Standard Transactions for\s*(MRS|MR|MISS|MS)?\s*(.+?)\s+('+MONEY.pattern+')$', l)
        if m: flush(); exp_std[m.group(2).strip()]=(-1 if total_cr() else 1)*money(m.group(3)); i+=1; continue
        m=re.match(r'Total of Account Charges\s+('+MONEY.pattern+')', l)
        if m: flush(); expected['CHARGES']=(-1 if total_cr() else 1)*money(m.group(1)); i+=1; continue
        m=re.match(r'Total New Deferred Credit Plans\s+('+MONEY.pattern+')', l)
        if m: flush(); expected['DEFERRED']=(-1 if total_cr() else 1)*money(m.group(1)); i+=1; continue

        ml=month_line(l)
        if ml and section:
            flush()
            mon=MONTHS[ml.group(1)]; day=int(ml.group(2)); rest=ml.group(3).strip()
            toks=MONEY.findall(l)
            aud=money(toks[-1]) if toks else None
            # strip ALL trailing money amounts (keep ref numbers without decimals)
            desc=re.sub(r'(?:\s+\d{1,3}(?:,\d{3})*\.\d{2})+\s*$','',l).strip()
            # date year: statement period may cross year; assign year so date<=stmt date within ~1yr
            yr=syr
            # if month > statement month, it's from previous year
            if mon>smon: yr=syr-1
            cur={'source':source,'section':section,'cardholder':cardholder,
                 'month':mon,'day':day,'year':yr,'desc':desc,'aud':aud,
                 'foreign':None,'currency':'','cr':False,'notes':[],'raw':l,'ntoks':len(toks)}
            i+=1; continue

        # continuation line for current txn
        if cur is not None:
            if l=='CR':
                cur['cr']=True; i+=1; continue
            if re.match(r'^[A-Z][A-Z .]+$', l) and 'AUD' not in l:  # currency name
                if not cur['currency']: cur['currency']=l
                i+=1; continue
            cm=re.search(r'includes conversion commission', l)
            if cm:
                # foreign amount = second-to-last money token on the txn month-line
                mt=MONEY.findall(cur['raw'])
                if len(mt)>=2: cur['foreign']=money(mt[-2])
                cur['notes'].append(l)
                i+=1; continue
            # sub-item detail (contains money) -> note; capture CR suffix inline
            if MONEY.search(l):
                cur['notes'].append(l); i+=1; continue
        i+=1
    flush()
    return source,(syr,smon,sday),txns,expected,exp_std

def reconcile(txns,expected,exp_std):
    report=[]
    # payments
    pay=sum(t['aud'] for t in txns if t['section']=='PAYMENT' and t['aud'])
    if expected['PAYMENT'] is not None:
        report.append(('Payments',round(pay,2),expected['PAYMENT'],abs(pay-expected['PAYMENT'])<0.01))
    # standard per cardholder
    for ch,exp in exp_std.items():
        net=sum((-t['aud'] if t['cr'] else t['aud']) for t in txns if t['section']=='STANDARD' and t['cardholder']==ch and t['aud'])
        report.append(('Std: '+ch,round(net,2),exp,abs(net-exp)<0.01))
    # charges
    if expected['CHARGES'] is not None:
        net=sum((-t['aud'] if t['cr'] else t['aud']) for t in txns if t['section']=='CHARGES' and t['aud'])
        report.append(('Charges',round(net,2),expected['CHARGES'],abs(net-expected['CHARGES'])<0.01))
    # deferred
    if expected['DEFERRED'] is not None:
        net=sum((-t['aud'] if t['cr'] else t['aud']) for t in txns if t['section']=='DEFERRED' and t['aud'])
        report.append(('Deferred',round(net,2),expected['DEFERRED'],abs(net-expected['DEFERRED'])<0.01))
    return report

if __name__=='__main__':
    allt=[]
    print(f"{'STATEMENT':12} {'SECTION':22} {'PARSED':>12} {'EXPECTED':>12}  OK")
    overall_ok=True
    for f in FILES:
        source,md,txns,expected,exp_std=parse(f'{U}/{f}.pdf')
        allt.extend(txns)
        rep=reconcile(txns,expected,exp_std)
        for name,got,exp,ok in rep:
            if not ok: overall_ok=False
            print(f"{source:12} {name:22} {got:12,.2f} {exp:12,.2f}  {'PASS' if ok else 'FAIL <<<'}")
        print(f"{'':12} -> {len(txns)} transactions")
    print("\nTOTAL transactions parsed:", len(allt))
    print("ALL RECONCILED:", overall_ok)
