"""يبني catalog_items.csv لموقع الكتالوج من تصدير قاعدة بيانات الخوارزمي/سكاي سوفت.

قاعدة العرض (نفس قاعدة سكربت إدارة الصور تماماً):
    يُعرض الصنف فقط إذا كان رصيد (المستودع الرئيسي 01 + فرع02) >= MIN_QTY
    فرع05 لا يدخل في القرار. الأصناف بلا رصيد تُستبعد نهائياً — لا سعر ولا كرتون.

السعر المعتمد: سعر1 فقط (سعر الجملة). سعر3 في البرنامج هو سعر القطاعي
لزوّار فرع02 ولا علاقة له بعملاء الجملة، فلا يُصدَّر ولا يُعرض في الموقع.
سعر الكرتون = سعر الحبة × عدد الحبات بالكرتون (يُحسب في الموقع، فلا يتقادم).
"""
import pandas as pd, sys, os

SRC = sys.argv[1] if len(sys.argv) > 1 else r"D:\Gama\تصدير يومي"
OUT = sys.argv[2] if len(sys.argv) > 2 else r"D:\GOOGLW DRIVE\Sm K\Claude\موقع الكتالوج\catalog_items.csv"

MIN_QTY = 2          # نفس $MinQty في ادارة_صور_الاصناف.ps1
STORES  = ['01', '02']   # المخازن المحسوبة — فرع05 مُتجاهَل عمداً

rd = lambda n: pd.read_csv(os.path.join(SRC, n), sep=';', encoding='utf-8-sig', dtype=str)
num = lambda s: pd.to_numeric(s, errors='coerce').fillna(0)

it, un, sq = rd('ItemsTb.csv'), rd('itmUnits.csv'), rd('StksQty.csv')
for c in ['uPerc', 'uPrice1', 'uPrice3', 'uPrice4']:
    un[c] = num(un[c])

# وحدة البيع الأساسية (حبه / BOX ...)
# سعر1 فقط = سعر الجملة. سعر3/سعر4 هما سعر القطاعي لزوّار فرع02 ولا يُعرضان للعملاء.
base = (un[un.uDef == '1']
        .sort_values('uPerc')
        .drop_duplicates('uItmID')[['uItmID', 'uAName', 'uPrice1']]
        .rename(columns={'uAName': 'unit', 'uPrice1': 'p1'}))

# وحدة التعبئة (كرتون)
ctn = (un[(un.uRelDef == '1') & (un.uPerc > 1)]
       .sort_values('uPerc', ascending=False)
       .drop_duplicates('uItmID')[['uItmID', 'uPerc']]
       .rename(columns={'uPerc': 'pcs'}))

# الأرصدة لكل مخزن
sq['STK_QTY'] = num(sq.STK_QTY)
piv = sq.pivot_table(index='itmID', columns='StkNo', values='STK_QTY', aggfunc='sum').fillna(0)
for s in ['01', '02', '03', '04', '05']:
    if s not in piv.columns:
        piv[s] = 0
piv['total'] = piv[STORES].sum(axis=1)   # الرصيد المعتمد = الرئيسي + فرع02
piv = piv.reset_index().rename(columns={'itmID': 'uItmID'})

df = (it[['itmID', 'itmNo', 'itmADescr', 'ifstop']]
      .rename(columns={'itmID': 'uItmID'})
      .merge(base, on='uItmID', how='left')
      .merge(ctn, on='uItmID', how='left')
      .merge(piv[['uItmID', '01', '02', '05', 'total']], on='uItmID', how='left'))

before = len(df)
df = df[(df.ifstop != 'True') & df.itmNo.notna() & df.itmADescr.notna()]
df['total'] = df['total'].fillna(0)
df = df[df['total'] >= MIN_QTY]                 # الأصناف بلا رصيد تُستبعد نهائياً
print('items in database:', before, '-> with (01+02) >=', MIN_QTY, ':', len(df))
df['p1'] = df.p1.fillna(0)
df['pcs'] = df.pcs.fillna(0).astype(int)
df = df[df.p1 > 0]

out = pd.DataFrame({
    'كود الصنف':            df.itmNo.str.strip(),
    'اسم الصنف':            df.itmADescr.str.strip(),
    'القسم':                '',
    'وحدة البيع':           df.unit.fillna('حبة').str.strip(),
    'سعر الحبة':            df.p1.round(2),
    'عدد الحبات بالكرتون':  df.pcs.replace(0, ''),
    'سعر الكرتون':          '',              # يُترك فارغاً ليُحسب دائماً = سعر الحبة × عدد الحبات
    'المخزون':              df['total'].fillna(0).astype(int),
    'الرئيسي':              df['01'].fillna(0).astype(int),
    'فرع02':                df['02'].fillna(0).astype(int),
    'فرع05':                df['05'].fillna(0).astype(int),
    'معرف الصورة':          '',
    'عرض':                  '',
    'جديد':                 '',
    'وسوم':                 '',
    'إظهار':                'نعم',
    'ترتيب':                '',
    'ملاحظة':               '',
})
out = out.sort_values('كود الصنف')
out.to_csv(OUT, index=False, encoding='utf-8-sig')

print('rows:', len(out))
print('with carton size:', (out['عدد الحبات بالكرتون'] != '').sum())
print('in stock:', (out['المخزون'] > 0).sum())
print('units:', out['وحدة البيع'].value_counts().to_dict())
print('size KB:', round(os.path.getsize(OUT) / 1024))
