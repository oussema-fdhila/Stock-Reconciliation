
import io
import warnings
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill

warnings.filterwarnings("ignore")
st.set_page_config(page_title="StockFlow | LEONI", page_icon="📊", layout="wide")

st.markdown('''
<style>
.stApp {background:#f5f7fb}
.block-container {padding-top:1.4rem;max-width:1500px}
.hero {background:linear-gradient(120deg,#102b50,#174a78);color:white;padding:1.5rem;border-radius:16px;margin-bottom:1rem}
.hero h1 {color:white;margin:0}.hero p {color:#e4edf8;margin:.3rem 0 0}
div[data-testid="stMetric"] {background:white;border:1px solid #e2e8f0;padding:12px;border-radius:12px}
</style>''', unsafe_allow_html=True)

COMMON = {
 "1_RECEIPT":1,"12_DESPATCHED":-1,"13_DESPATCHED":-1,"16_DESPATCHED":0,
 "17_CANCELED":1,"17_DESPATCHED":-1,"2_RECEIPT":0,"23_DESPATCHED":-1,
 "24_CANCELED":1,"24_RECEIPT":1,"28_UM V":0,"30_RECEIPT":1,"4_UM V":0,
 "41_DESPATCHED":-1,"5_CANCELED":1,"5_RECEIPT":1,"1_CANCELED":1,
 "16_CANCELED":0,"23_":0,"23_CANCELED":1,"4_":0,"41_CANCELED":1,
 "13_CANCELED":1
}
SIGNS = {
 "47": {**COMMON,"37_DESPATCHED":-1,"42_UM V":0,"71_RECEIPT":0,"72_DESPATCHED":0,"6_RECEIPT":0,"2_CANCELED":0},
 "48": {**COMMON,"37_DESPATCHED":-1,"42_UM V":0,"36_UM V":0,"20_DESPATCHED":0,"2_":0,"71_RECEIPT":0,"72_DESPATCHED":0},
 "66": {**COMMON,"37_DESPATCHED":0,"42_UM V":0,"36_UM V":-1,"20_DESPATCHED":0,"2_":0,"71_RECEIPT":0,"72_DESPATCHED":0,"2_CANCELED":0,"28_":0,"6_RECEIPT":1},
 "77": {**COMMON,"37_DESPATCHED":-1,"42_UM V":0,"36_UM V":-1,"20_DESPATCHED":0,"2_":0,"71_RECEIPT":0,"72_DESPATCHED":0,"21_RECEIPT":0,"6_RECEIPT":0,"2_CANCELED":0}
}
CW = {"CW","CW-VW","CW-OEM"}
PG, PI, PL, PM = 6, 8, 11, 12

def clean_text(x):
    return str(x).replace("\xa0"," ").replace("\ufeff","").strip()

def norm_wh(x):
    x = clean_text(x).upper()
    if x in ("", "NAN", "NONE"): return ""
    return x.lstrip("0") or "0" if x.isdigit() else x

def number(s):
    return pd.to_numeric(s.astype(str).str.replace("\xa0","",regex=False).str.replace(" ","",regex=False).str.replace(",",".",regex=False), errors="coerce")

def read_csv(upload, skip):
    raw = upload.getvalue().replace(b"\x00",b"")
    text = raw.decode("latin1").replace("\r\n","\n").replace("\r","\n")
    lines = text.split("\n")
    if len(lines) <= skip: raise ValueError(f"{upload.name}: fichier trop court.")
    body = "\n".join(lines[skip:])
    if not body.strip(): raise ValueError(f"{upload.name}: aucune donnée après l'en-tête.")
    df = pd.read_csv(io.StringIO(body), sep=";", dtype=str, keep_default_na=False)
    df.columns = [clean_text(c) for c in df.columns]
    cols = [df.iloc[:,i].astype(str).str.replace("\xa0"," ",regex=False).str.replace("\ufeff","",regex=False).str.strip() for i in range(df.shape[1])]
    if cols:
        out = pd.concat(cols, axis=1)
        out.columns = df.columns
        df = out
    return df

def wh_col(df):
    hits = [c for c in df.columns if c.casefold() == "warehouse"]
    if len(hits) == 1: return hits[0]
    if df.shape[1] > PI: return df.columns[PI]
    raise ValueError("Colonne Warehouse introuvable.")

def cw_count(df):
    if df.shape[1] <= PM: return 0
    i = df.iloc[:,PI].astype(str).str.strip().str.upper()
    m = df.iloc[:,PM].astype(str).str.strip().str.upper()
    return int((i.isin(CW)|m.isin(CW)).sum())

def redress(df):
    if df.shape[1] < 13: raise ValueError("13 colonnes minimum pour le redressement CW.")
    df = df.copy()
    for p in [PG,PI,PL,PM]: df.iloc[:,p] = df.iloc[:,p].astype(str).str.strip()
    i = df.iloc[:,PI].str.upper(); m = df.iloc[:,PM].str.upper()
    delete = i.isin(CW) & (df.iloc[:,PM].eq("")|m.isin(CW))
    deleted = int(delete.sum()); df = df.loc[~delete].copy()
    mask = df.iloc[:,PI].str.upper().isin(CW).to_numpy()
    dispatched = int(mask.sum())
    if dispatched:
        df.iloc[mask,PI] = df.iloc[mask,PM].to_numpy()
        df.iloc[mask,PM] = ""; df.iloc[mask,PG] = "17"; df.iloc[mask,PL] = "DESPATCHED"
    mask = df.iloc[:,PM].str.upper().isin(CW).to_numpy()
    receipt = int(mask.sum())
    if receipt:
        df.iloc[mask,PG] = "05"; df.iloc[mask,PL] = "RECEIPT"; df.iloc[mask,PM] = ""
    return df, {"CW supprimées":deleted,"CW DESPATCHED":dispatched,"CW RECEIPT":receipt,
                "CW restant I":int(df.iloc[:,PI].astype(str).str.upper().isin(CW).sum()),
                "CW restant M":int(df.iloc[:,PM].astype(str).str.upper().isin(CW).sum())}

def movement(upload, plant, excluded, do_cw):
    df = read_csv(upload,12)
    if df.shape[1] < 13: raise ValueError(f"{upload.name}: moins de 13 colonnes détectées.")
    count = cw_count(df)
    cw_stats = {"CW supprimées":0,"CW DESPATCHED":0,"CW RECEIPT":0,"CW restant I":0,"CW restant M":0}
    if do_cw and count: df,cw_stats = redress(df)
    elif count: cw_stats["CW restant I/M avant traitement"] = count
    wh = wh_col(df)
    if excluded:
        mask = df[wh].map(norm_wh).isin({norm_wh(x) for x in excluded})
        removed = int(mask.sum()); df = df.loc[~mask].copy()
    else: removed = 0
    required = ["PartNr.","BookType","Quantity","Type"]
    missing = [c for c in required if c not in df.columns]
    if missing: raise ValueError(f"{upload.name}: colonnes manquantes {missing}")
    duplicates = [c for c in required if list(df.columns).count(c)>1]
    if duplicates: raise ValueError(f"{upload.name}: colonnes requises en double {duplicates}")
    d = df[required].copy()
    d["Quantity"] = number(d["Quantity"]).fillna(0)
    d["BookType"] = d["BookType"].astype(str).str.strip().str.replace(r"^0+(\d+)$",r"\1",regex=True)
    d["Type"] = d["Type"].astype(str).str.strip()
    d["Mouv"] = d["BookType"]+"_"+d["Type"]
    unknown = sorted(set(d["Mouv"])-set(SIGNS[plant]))
    pivot = pd.pivot_table(d,index="PartNr.",columns="Mouv",values="Quantity",aggfunc="sum",fill_value=0).reset_index().rename(columns={"PartNr.":"PartNumber"})
    move_cols = [c for c in pivot.columns if c!="PartNumber"]
    for c in move_cols: pivot[c] *= SIGNS[plant].get(str(c).strip(),0)
    pivot["Total"] = pivot[move_cols].sum(axis=1) if move_cols else 0
    diag = {"Fichier":upload.name,"Lignes conservées":len(df),"Lignes de magasins supprimées":removed,
            "PartNumbers TCD":len(pivot),"Total mouvement signé":float(pivot["Total"].sum()),
            "Types inconnus":", ".join(unknown),**cw_stats}
    return pivot,diag

def stock(upload):
    df = read_csv(upload,9)
    if df.shape[1] < 10: raise ValueError(f"{upload.name}: moins de 10 colonnes détectées.")
    part = df.iloc[:,0].astype(str).str.strip()
    typ = df.iloc[:,3].astype(str).str.strip()
    qty = number(df.iloc[:,6]).fillna(0)
    value = number(df.iloc[:,9]).fillna(0)
    pu = (value/qty.replace(0,np.nan)).replace([np.inf,-np.inf],np.nan).fillna(0)
    out = pd.DataFrame({"_PartNumber":part,"_Type":typ,"_Stock":qty,"_PU":pu})
    return out[out["_PartNumber"]!=""].copy()

def first_map(df,col):
    return df.drop_duplicates("_PartNumber").set_index("_PartNumber")[col]

def reconcile(si,sf,pivots):
    parts = sorted(set(si["_PartNumber"])|set(sf["_PartNumber"]))
    r = pd.DataFrame({"PartNumber":parts})
    r["SI"] = r["PartNumber"].map(si.groupby("_PartNumber")["_Stock"].sum()).fillna(0)
    r["SF"] = r["PartNumber"].map(sf.groupby("_PartNumber")["_Stock"].sum()).fillna(0)
    mov_cols=[]
    for i,p in enumerate(pivots,1):
        col = f"Mouvement {i}" if len(pivots)==2 else "Mouvement"
        r[col] = r["PartNumber"].map(p.set_index("PartNumber")["Total"]).fillna(0); mov_cols.append(col)
    if len(pivots)==2: r["Mouvement"] = r["Mouvement 1"]+r["Mouvement 2"]
    ts,ti,ps,pi = (first_map(sf,"_Type"),first_map(si,"_Type"),first_map(sf,"_PU"),first_map(si,"_PU"))
    r["Type"] = r["PartNumber"].map(ts).fillna(r["PartNumber"].map(ti)).fillna("")
    r["PU"] = pd.to_numeric(r["PartNumber"].map(ps).fillna(r["PartNumber"].map(pi)),errors="coerce").fillna(0)
    r["Ecart"] = r["SF"]-r["SI"]-r["Mouvement"]
    r["Ecart €"] = r["Ecart"]*r["PU"]
    order=["PartNumber","Type","PU","SI"]+mov_cols
    if len(pivots)==2: order.append("Mouvement")
    return r[order+["SF","Ecart","Ecart €"]]

def excel_bytes(recon,pivots,diagnostics):
    b=io.BytesIO()
    with pd.ExcelWriter(b,engine="openpyxl") as w:
        recon.to_excel(w,sheet_name="Réconcil",index=False)
        for i,p in enumerate(pivots,1): p.to_excel(w,sheet_name=f"TCD{i}",index=False)
        pd.DataFrame(diagnostics).to_excel(w,sheet_name="Contrôles",index=False)
        for ws in w.book.worksheets:
            ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
            for cell in ws[1]:
                cell.font=Font(bold=True,color="FFFFFF")
                cell.fill=PatternFill("solid",fgColor="17365D")
            for col in ws.columns:
                vals=[len(str(c.value or "")) for c in list(col)[:120]]
                ws.column_dimensions[col[0].column_letter].width=min(max(max(vals,default=10)+2,10),32)
    b.seek(0); return b.getvalue()

st.markdown('<div class="hero"><h1>StockFlow</h1><p>LEONI · Stock reconciliation · SI + mouvements → SF</p></div>',unsafe_allow_html=True)
with st.sidebar:
    st.markdown("## ⚙️ Configuration")
    plant=st.selectbox("Plant",["47","48","66","77"],index=2)
    do_cw=st.checkbox("Redressement CW automatique",value=True,help="Appliqué uniquement si des valeurs CW sont détectées.")
    st.caption("La table des signes dépend du plant sélectionné.")

st.markdown("### 1 · Importer les fichiers")
st.caption("CSV séparés par des points-virgules. Les 12 premières lignes des mouvements et les 9 premières lignes des stocks sont ignorées.")
c1,c2=st.columns(2)
with c1:
    si_file=st.file_uploader("Stock initial (SI) · FMALS015",type=["csv","txt"],key="si")
    sf_file=st.file_uploader("Stock final (SF) · FMALS015",type=["csv","txt"],key="sf")
with c2:
    movement_files=st.file_uploader("Mouvements FMALS032 (1 ou 2 fichiers)",type=["csv","txt"],accept_multiple_files=True,key="moves")
    if len(movement_files)>2: st.error("Maximum : deux fichiers de mouvements.")
    else: st.info(f"Fichiers sélectionnés : {len(movement_files)} / 2")

# Discover available warehouses from the uploaded movements.
warehouse_raw={}
warehouse_errors=[]
if movement_files and len(movement_files)<=2:
    for f in movement_files:
        try:
            d=read_csv(f,12); col=wh_col(d)
            for raw in d[col].astype(str):
                raw=clean_text(raw); code=norm_wh(raw)
                if code: warehouse_raw.setdefault(code,set()).add(raw)
        except Exception as e: warehouse_errors.append(f"{f.name}: {e}")
for e in warehouse_errors: st.error(e)

st.markdown("### 2 · Magasins à supprimer")
selected=[]
if warehouse_raw:
    st.caption("Coche les magasins à exclure. Si tu ne coches rien, aucun magasin ne sera supprimé.")
    query=st.text_input("Rechercher un magasin",placeholder="Ex. 0452CZ ou 510500")
    codes=sorted(warehouse_raw,key=str.casefold)
    visible=[x for x in codes if query.casefold() in x.casefold() or any(query.casefold() in r.casefold() for r in warehouse_raw[x])] if query else codes
    if len(codes)>250: st.warning(f"{len(codes):,} magasins détectés : utilise la recherche pour réduire la liste.")
    cols=st.columns(3)
    for i,code in enumerate(visible):
        label=code
        originals=sorted(warehouse_raw[code])
        if originals and originals[0]!=code: label=f"{code} · {originals[0]}"
        with cols[i%3]: st.checkbox(label,key=f"wh_{plant}_{code}",help=f"Valeurs trouvées : {', '.join(originals[:4])}")
    selected=[x for x in codes if st.session_state.get(f"wh_{plant}_{x}",False)]
    st.caption(f"Magasins sélectionnés : {len(selected)}")
else:
    st.info("Importe d'abord les mouvements pour détecter les magasins.")

st.markdown("---")
st.markdown("### 3 · Lancer la réconciliation")
ready=si_file is not None and sf_file is not None and len(movement_files) in (1,2) and not warehouse_errors
run=st.button("🚀 Lancer la réconciliation",type="primary",use_container_width=True,disabled=not ready)
if run:
    try:
        with st.spinner("Calcul en cours…"):
            pivots=[]; diagnostics=[]
            for f in movement_files:
                p,diag=movement(f,plant,selected,do_cw); pivots.append(p); diagnostics.append({"Plant":plant,**diag})
            si=stock(si_file); sf=stock(sf_file)
            result=reconcile(si,sf,pivots)
            workbook=excel_bytes(result,pivots,diagnostics)
        st.session_state["result"]=result
        st.session_state["pivots"]=pivots
        st.session_state["diagnostics"]=diagnostics
        st.session_state["workbook"]=workbook
        st.session_state["result_plant"]=plant
        st.success("Réconciliation terminée.")
    except Exception as e:
        st.error(f"Erreur : {e}")
        st.exception(e)

if "result" in st.session_state:
    r=st.session_state["result"]; ps=st.session_state["pivots"]; diags=st.session_state["diagnostics"]
    st.markdown("---")
    st.markdown(f"## Résultats · Plant {st.session_state.get('result_plant',plant)}")
    a,b,c,d=st.columns(4)
    a.metric("PartNumbers",f"{len(r):,}")
    b.metric("Écart total (quantité)",f"{r['Ecart'].sum():,.2f}")
    c.metric("Écart total (€)",f"{r['Ecart €'].sum():,.2f}")
    d.metric("Références avec écart",f"{int((r['Ecart'].abs()>1e-9).sum()):,}")
    t1,t2,t3=st.tabs(["📋 Réconciliation","🔄 TCD mouvements","🧪 Contrôles"])
    with t1:
        only=st.checkbox("Afficher uniquement les lignes avec un écart",value=False)
        view=r.loc[r["Ecart"].abs()>1e-9] if only else r
        st.dataframe(view,use_container_width=True,height=500,hide_index=True)
        st.download_button("⬇️ Télécharger l'Excel complet",data=st.session_state["workbook"],file_name=f"Reconciliation_Stock_Plant_{st.session_state.get('result_plant',plant)}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
        st.download_button("Télécharger la réconciliation en CSV",data=r.to_csv(index=False,sep=";").encode("utf-8-sig"),file_name="Reconciliation.csv",mime="text/csv")
    with t2:
        for i,p in enumerate(ps,1):
            st.markdown(f"#### TCD{i}")
            st.dataframe(p,use_container_width=True,height=350,hide_index=True)
    with t3:
        st.dataframe(pd.DataFrame(diags),use_container_width=True,hide_index=True)
        unknown={x["Fichier"]:x["Types inconnus"] for x in diags if x["Types inconnus"]}
        if unknown:
            st.warning("Types inconnus : le signe par défaut 0 a été appliqué. Vérifie ces codes.")
            for name,codes in unknown.items(): st.write(f"**{name}** : {codes}")
        else: st.success("Tous les types de mouvements sont reconnus.")
st.markdown("---")
st.caption("Vérifie les types inconnus et les résultats avant toute utilisation opérationnelle. Ne téléverse pas de données d'entreprise confidentielles sur un service public sans autorisation.")
