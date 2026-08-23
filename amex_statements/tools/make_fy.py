import datetime, re, json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from parser import parse, reconcile, FILES, U

rules=json.load(open("/home/user/Private/amex_statements/reference/stylecolab_sc_rules.json"))
AUTO=[k.upper() for k in rules['auto_sc_keywords']]
REVIEW=[k.upper() for k in rules['review_keywords']]

# Ground truth: the 14 statements Vicki personally SC-marked
GT_SOURCES={'2022-06-22','2022-07-22','2022-08-22','2022-09-22','2022-10-22','2022-11-22',
 '2022-12-22','2023-01-22','2023-02-22','2023-03-22','2023-04-22','2023-05-22','2023-06-22','2023-07-22'}
def _load_gt():
    wb=openpyxl.load_workbook("edited.xlsx"); ws=wb["Vicki Zertopoulos"]
    def yellow(c): f=c.fill; return bool(f and f.patternType=='solid' and f.fgColor and f.fgColor.rgb not in (None,'00000000'))
    flags=set()
    for r in range(5,ws.max_row+1):
        d=ws.cell(row=r,column=1).value
        if not d or str(d).startswith("TOTAL"): continue
        if yellow(ws.cell(row=r,column=4)) or ws.cell(row=r,column=8).value not in (None,''):
            iso=(d.date() if hasattr(d,'date') else d).isoformat()
            deb=ws.cell(row=r,column=7).value
            deb=round(deb,2) if isinstance(deb,(int,float)) else None
            flags.add((iso,str(ws.cell(row=r,column=4).value).strip(),deb))
    return flags
GT_FLAGS=_load_gt()

def fy_of(d):
    y=d.year
    start = y if d.month>=7 else y-1
    return f"FY{start}-{str(start+1)[2:]}"

def sc_flag(t,d):
    # SC applies to Vicki's own spend only (Standard + Deferred)
    if t['cardholder']!='VICKI ZERTOPOULOS' or t['section'] not in ('STANDARD','DEFERRED'):
        return ''
    desc=re.sub(r'^(January|February|March|April|May|June|July|August|September|October|November|December)\s*\d{1,2}\s+','',t['desc'])
    if t['source'] in GT_SOURCES:                       # Vicki's authoritative marks
        deb=round(t['aud'],2) if (t['aud'] is not None and not t['cr']) else None
        return 'SC' if (d.isoformat(),desc,deb) in GT_FLAGS else ''
    u=desc.upper()                                       # earlier years: keyword pre-flag
    if any(k in u for k in AUTO): return 'SC'
    if any(k in u for k in REVIEW): return 'REVIEW'
    return ''

# gather rows + recon
rows=[]; recon_all=[]; seq=0
for f in FILES:
    source,md,txns,expected,exp_std=parse(f'{U}/{f}.pdf')
    for r in reconcile(txns,expected,exp_std): recon_all.append((source,)+r)
    for t in txns:
        d=datetime.date(t['year'],t['month'],t['day'])
        is_credit=(t['section']=='PAYMENT') or t['cr']
        sec={'PAYMENT':'Payment','STANDARD':'Standard','CHARGES':'Account Charge','DEFERRED':'Deferred Credit Plan'}[t['section']]
        desc=re.sub(r'^(January|February|March|April|May|June|July|August|September|October|November|December)\s*\d{1,2}\s+','',t['desc'])
        rows.append({'fy':fy_of(d),'date':d,'source':source,'cardholder':t['cardholder'],'section':sec,
                     'desc':desc,'foreign':t['foreign'],'cur':t['currency'],
                     'debit':None if is_credit else t['aud'],'credit':t['aud'] if is_credit else None,
                     'notes':'; '.join(t['notes']),'sc':sc_flag(t,d),'seq':seq}); seq+=1

NAVY="1F3864"; LIGHT="D9E1F2"; YEL="FFF2CC"; ORG="FCE4D6"
thin=Side(style="thin",color="BFBFBF"); border=Border(left=thin,right=thin,top=thin,bottom=thin)
money='#,##0.00;[Red]-#,##0.00'
headers=["Date","Source Statement","Cardholder","Section","Description (Merchant & Location)",
         "Foreign Amount","Foreign Currency","Debit (AUD)","SC","Credit (AUD)","Notes / Itemisation"]
# col idx: Debit=8(H), SC=9(I), Credit=10(J)

def build_fy_sheet(wb,title,fy_rows,subtitle):
    ws=wb.create_sheet(title[:31])
    ws.merge_cells("A1:K1"); ws["A1"]=subtitle; ws["A1"].font=Font(bold=True,size=13,color=NAVY)
    ws.merge_cells("A2:K2")
    ws["A2"]=f"Card xxxx-xxxxxx-64003  |  {len(fy_rows)} transactions  |  SC = Stylecolab business expense (auto-flagged); REVIEW = confirm manually"
    ws["A2"].font=Font(italic=True,size=9,color="595959")
    hr=4
    for i,h in enumerate(headers,1):
        c=ws.cell(row=hr,column=i,value=h); c.fill=PatternFill("solid",fgColor=NAVY)
        c.font=Font(bold=True,color="FFFFFF"); c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=border
    r=hr+1
    for row in sorted(fy_rows,key=lambda x:(x['date'],x['source'],x['seq'])):
        ws.cell(row=r,column=1,value=row['date']).number_format='dd/mm/yyyy'
        ws.cell(row=r,column=2,value=row['source'])
        ws.cell(row=r,column=3,value=row['cardholder'])
        ws.cell(row=r,column=4,value=row['section'])
        ws.cell(row=r,column=5,value=row['desc'])
        if row['foreign'] is not None: ws.cell(row=r,column=6,value=row['foreign']).number_format=money
        ws.cell(row=r,column=7,value=row['cur'])
        if row['debit'] is not None: ws.cell(row=r,column=8,value=row['debit']).number_format=money
        # SC column
        if row['sc']=='SC' and row['debit'] is not None:
            c=ws.cell(row=r,column=9,value=f"=H{r}"); c.number_format=money; c.fill=PatternFill("solid",fgColor=YEL)
        elif row['sc']=='REVIEW':
            c=ws.cell(row=r,column=9,value="REVIEW"); c.fill=PatternFill("solid",fgColor=ORG); c.font=Font(italic=True,color="C55A11")
        if row['credit'] is not None: ws.cell(row=r,column=10,value=row['credit']).number_format=money
        ws.cell(row=r,column=11,value=row['notes'])
        for i in range(1,12): ws.cell(row=r,column=i).border=border
        r+=1
    first=hr+1; last=r-1; tr=r
    ws.cell(row=tr,column=1,value="TOTALS").font=Font(bold=True)
    ws.merge_cells(start_row=tr,start_column=1,end_row=tr,end_column=7)
    for col,cl in ((8,'H'),(9,'I'),(10,'J')):
        c=ws.cell(row=tr,column=col,value=f"=SUM({cl}{first}:{cl}{last})"); c.number_format=money; c.font=Font(bold=True)
    ws.cell(row=tr,column=11,value="SC total = confirmed Stylecolab expenses").font=Font(italic=True,size=9)
    for i in range(1,12):
        c=ws.cell(row=tr,column=i); c.border=border; c.fill=PatternFill("solid",fgColor=LIGHT)
    for i,w in enumerate([11,15,19,18,40,13,17,13,12,13,55],1): ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes="A5"; ws.auto_filter.ref=f"A{hr}:K{last}"

def add_recon(wb):
    ws2=wb.create_sheet("Reconciliation")
    ws2["A1"]="Per-Statement Reconciliation vs Printed Totals (all 50 statements)"; ws2["A1"].font=Font(bold=True,size=12,color=NAVY)
    for i,h in enumerate(["Statement","Section","Extracted","Statement","Match"],1):
        c=ws2.cell(row=2,column=i,value=h); c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor=NAVY); c.border=border
    rr=3
    for source,name,got,exp,ok in recon_all:
        ws2.cell(row=rr,column=1,value=source); ws2.cell(row=rr,column=2,value=name)
        ws2.cell(row=rr,column=3,value=got).number_format=money
        ws2.cell(row=rr,column=4,value=exp).number_format=money
        cc=ws2.cell(row=rr,column=5,value="YES" if ok else "NO")
        if not ok: cc.fill=PatternFill("solid",fgColor="FFC7CE")
        rr+=1
    for col,w in zip("ABCDE",[13,26,14,14,7]): ws2.column_dimensions[col].width=w
    ws2.freeze_panes="A3"

fy_order=["FY2019-20","FY2020-21","FY2021-22","FY2022-23","FY2023-24","FY2024-25"]
labels={"FY2019-20":"FY2019-20 (partial — May/Jun 2020 only)",
        "FY2020-21":"Financial Year 2020-21 (1 Jul 2020 – 30 Jun 2021)",
        "FY2021-22":"Financial Year 2021-22 (1 Jul 2021 – 30 Jun 2022)",
        "FY2022-23":"Financial Year 2022-23 (1 Jul 2022 – 30 Jun 2023)",
        "FY2023-24":"Financial Year 2023-24 (1 Jul 2023 – 30 Jun 2024)",
        "FY2024-25":"FY2024-25 (partial — Jul 2024 only)"}
byfy={fy:[r for r in rows if r['fy']==fy] for fy in fy_order}

# Summary numbers
def sc_total(fr): return round(sum(r['debit'] for r in fr if r['sc']=='SC' and r['debit']),2)
def review_total(fr): return round(sum(r['debit'] for r in fr if r['sc']=='REVIEW' and r['debit']),2)

# ---- Combined workbook ----
wb=openpyxl.Workbook(); wb.remove(wb.active)
wss=wb.create_sheet("Summary")
wss["A1"]="Stylecolab Expense Summary by Financial Year"; wss["A1"].font=Font(bold=True,size=13,color=NAVY)
basis={"FY2019-20":"Auto — pending review","FY2020-21":"Auto — pending review",
       "FY2021-22":"Mixed (Jun-2022 confirmed; rest auto)","FY2022-23":"Confirmed (Vicki's marks)",
       "FY2023-24":"Mixed (Jul-2023 confirmed; rest auto)","FY2024-25":"Auto — pending review"}
for i,h in enumerate(["Financial Year","Transactions","Total Debits","Total Credits","SC confirmed/flagged","REVIEW (to confirm)","SC basis"],1):
    c=wss.cell(row=3,column=i,value=h); c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor=NAVY); c.border=border
rr=4
for fy in fy_order:
    fr=byfy[fy]
    wss.cell(row=rr,column=1,value=labels[fy])
    wss.cell(row=rr,column=2,value=len(fr))
    wss.cell(row=rr,column=3,value=round(sum(r['debit'] for r in fr if r['debit']),2)).number_format=money
    wss.cell(row=rr,column=4,value=round(sum(r['credit'] for r in fr if r['credit']),2)).number_format=money
    wss.cell(row=rr,column=5,value=sc_total(fr)).number_format=money
    wss.cell(row=rr,column=6,value=review_total(fr)).number_format=money
    wss.cell(row=rr,column=7,value=basis[fy])
    for i in range(1,8): wss.cell(row=rr,column=i).border=border
    rr+=1
wss.cell(row=rr+1,column=1,value="Yellow SC cells = Stylecolab business expense (=debit, so the SC column sums). Orange REVIEW = confirm or clear. SC applies to Vicki's card only.").font=Font(italic=True,size=9,color="595959")
for col,w in zip("ABCDEFG",[40,12,14,14,18,18,32]): wss.column_dimensions[col].width=w
for fy in fy_order:
    build_fy_sheet(wb,fy,byfy[fy],labels[fy])
add_recon(wb)
out_comb="/home/user/Private/amex_statements/DavidJones_Amex_FY-split_SC-classified_2020-2024.xlsx"
wb.save(out_comb)

# ---- Separate per-FY files (complete years) ----
for fy in ["FY2020-21","FY2021-22","FY2022-23","FY2023-24"]:
    w=openpyxl.Workbook(); w.remove(w.active)
    build_fy_sheet(w,fy,byfy[fy],labels[fy])
    w.save(f"/home/user/Private/amex_statements/DavidJones_Amex_{fy}_SC-classified.xlsx")

print("Combined saved:",out_comb)
print("\nFY | txns | debits | credits | SC confirmed | REVIEW")
for fy in fy_order:
    fr=byfy[fy]
    print(f"{fy} | {len(fr):4} | {sum(r['debit'] for r in fr if r['debit']):11,.2f} | {sum(r['credit'] for r in fr if r['credit']):11,.2f} | SC {sc_total(fr):10,.2f} | REV {review_total(fr):10,.2f}")
print("Total rows:",len(rows))
