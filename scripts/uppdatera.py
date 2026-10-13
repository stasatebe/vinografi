"""
Vinografi – automatisk uppdatering (körs av GitHub Actions varje natt).

Gör samma sak som Colab-anteckningsboken:
  1. Geokodar ifyllda googla-listor som laddats upp till data/inlamning/ och lägger in dem i registret
  2. Hämtar Systembolagets sortiment från C4illins tjänst (susbolaget.emrik.org)
  3. Räknar ut prissänkningar (prishistoriken i datan) och märker viner som är tillfälligt slut
  4. Skriver data/att_googla.xlsx + data/README.md när listan över nya producenter ändras
  5. Bygger hemsidan i _site/ (index.html + data.js) och masterfilen i master/
"""
import os, sys, glob, json, math, time, shutil, datetime
import pandas as pd
import requests

# ---------------- Inställningar ----------------
FÖNSTER_DAGAR = 92       # hur långt bakåt en prissänkning räknas (≈ 3 månader)
MIN_SÄNKNING_KR = 2      # mindre sänkningar räknas inte
DATA_URL = "https://susbolaget.emrik.org/v1/products"

ROT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REGISTER = os.path.join(ROT, 'data', 'producenter_koordinater.csv')
ATT_GOOGLA = os.path.join(ROT, 'data', 'att_googla.xlsx')
STATUS = os.path.join(ROT, 'data', 'README.md')
INLAMNING = os.path.join(ROT, 'data', 'inlamning')
SITE = os.path.join(ROT, '_site')
MASTER_MAPP = os.path.join(ROT, 'master')
IDAG = datetime.date.today().isoformat()

REG_KOL = ['Producent', 'Land', 'Latitud', 'Longitud', 'Hittad_adress', 'Koordinat_källa',
           'Manuell_sökstatus', 'Googlad_adress', 'Först_tillagd']
KOLUMNER = ['productId', 'productNumber', 'productNameBold', 'productNameThin',
            'producerName', 'supplierName', 'country', 'originLevel1', 'originLevel2',
            'categoryLevel2', 'categoryLevel3', 'price', 'volume', 'alcoholPercentage', 'vintage',
            'grapes', 'taste', 'tasteSymbols', 'assortmentText', 'lanseringsdatum', 'bildlänk',
            'pris_före_sänkning', 'tillfälligt_slut']


def nyckel(namn):
    """'REMELLURI' och 'Remelluri' ska räknas som samma producent."""
    if namn is None or (isinstance(namn, float) and math.isnan(namn)):
        return None
    return ' '.join(str(namn).replace('`', "'").replace('.', '').casefold().split())


def tom(v):
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    return isinstance(v, str) and v.strip() == ''


def ren(v):
    return None if tom(v) else v


# ---------------- 1. Inlämnade googla-listor ----------------
def geokoda(adress, api):
    for _ in range(3):
        try:
            r = requests.get("https://us1.locationiq.com/v1/search",
                             params={"key": api, "q": adress, "format": "json", "limit": 1}, timeout=15)
            if r.status_code == 200:
                d = r.json()
                return (float(d[0]['lat']), float(d[0]['lon']), d[0].get('display_name')) if d else None
            if r.status_code == 404:          # LocationIQ svarar 404 när adressen inte hittas
                return None
            time.sleep(2 if r.status_code == 429 else 1)
        except requests.RequestException:
            time.sleep(1)
    return None


def behandla_inlamning(reg):
    filer = sorted(glob.glob(os.path.join(INLAMNING, '*.xlsx')))
    if not filer:
        return reg
    api = os.environ.get('LOCATIONIQ_API_KEY', '')
    if not api:
        sys.exit("FEL: LOCATIONIQ_API_KEY saknas (lägg in den under Settings → Secrets and variables → Actions).")
    nya = []
    for fil in filer:
        df = pd.read_excel(fil)
        if 'Producent' not in df.columns or 'Adress (fyll i)' not in df.columns:
            print(f"Hoppar över {os.path.basename(fil)} – ser inte ut som en googla-lista.")
            continue
        print(f"Geokodar {os.path.basename(fil)} ({len(df)} rader)")
        for _, r in df.iterrows():
            adress = None if tom(r.get('Adress (fyll i)')) else str(r['Adress (fyll i)']).strip()
            lat_m, lon_m = r.get('Latitud (valfritt)'), r.get('Longitud (valfritt)')
            rad = {'Producent': r['Producent'], 'Land': ren(r.get('Land')), 'Latitud': None, 'Longitud': None,
                   'Hittad_adress': None, 'Koordinat_källa': 'Saknas', 'Manuell_sökstatus': 'Sökt - ej hittad',
                   'Googlad_adress': adress, 'Först_tillagd': IDAG}
            if not tom(lat_m) and not tom(lon_m):
                rad.update(Latitud=float(str(lat_m).replace(',', '.')), Longitud=float(str(lon_m).replace(',', '.')),
                           Koordinat_källa='Manuell (koordinater)', Manuell_sökstatus='Koordinater angivna manuellt')
            elif adress:
                res = geokoda(adress, api)
                time.sleep(0.55)
                if res:
                    rad.update(Latitud=res[0], Longitud=res[1], Hittad_adress=res[2],
                               Koordinat_källa='Manuell googling + LocationIQ (adress)',
                               Manuell_sökstatus='Adress hittad + geokodad')
                else:
                    rad['Manuell_sökstatus'] = 'Adress hittad, men gick ej att geokoda'
            print(f"  {rad['Manuell_sökstatus']:<40} {rad['Producent']}")
            nya.append(rad)
    if nya:
        nya_df = pd.DataFrame(nya, columns=REG_KOL)
        reg = reg[~reg['Producent'].map(nyckel).isin(set(nya_df['Producent'].map(nyckel)))]
        reg = pd.concat([reg, nya_df], ignore_index=True).sort_values(['Land', 'Producent']).reset_index(drop=True)
        reg.to_csv(REGISTER, index=False, encoding='utf-8-sig')
    for fil in filer:
        os.remove(fil)
    return reg


# ---------------- 2–3. Hämta sortimentet ----------------
def hämta_sortiment():
    r = requests.get(DATA_URL, timeout=300)
    r.raise_for_status()
    df = pd.DataFrame(r.json())
    print(f"Hämtade {len(df)} produkter")
    viner = df[(df['categoryLevel1'] == 'Vin') & (df['categoryLevel2'] != 'Glögg och Glühwein') &
               (df['isDiscontinued'] == False) & (df['isCompletelyOutOfStock'] == False)].copy()
    viner['grapes'] = viner['grapes'].apply(lambda g: ', '.join(g) if isinstance(g, list) and g else '')
    viner['tasteSymbols'] = viner['tasteSymbols'].apply(lambda x: ', '.join(x) if isinstance(x, list) else '')
    viner['bildlänk'] = viner['images'].apply(lambda x: x[0]['imageUrl'] if isinstance(x, list) and x else '')
    viner['lanseringsdatum'] = pd.to_datetime(viner['productLaunchDate']).dt.date
    viner['tillfälligt_slut'] = viner['isTemporaryOutOfStock'] == True

    start_ms = int((pd.Timestamp.now(tz='UTC') - pd.Timedelta(days=FÖNSTER_DAGAR)).timestamp() * 1000)

    def pris_före_sänkning(historik, pris):
        if not isinstance(historik, list) or not historik:
            return None
        p = sorted((h['x'], h['y']) for h in historik if h.get('x') is not None and h.get('y') is not None)
        kandidater = [y for x, y in p if x <= start_ms][-1:] + [y for x, y in p if x > start_ms]
        if not kandidater:
            return None
        högst = max(kandidater)
        return högst if högst - pris >= MIN_SÄNKNING_KR else None

    viner['pris_före_sänkning'] = [pris_före_sänkning(h, p) for h, p in zip(viner['priceHistory'], viner['price'])]
    viner = viner[KOLUMNER].reset_index(drop=True)
    print(f"Aktuella viner: {len(viner)}, varav tillfälligt slut: {int(viner['tillfälligt_slut'].sum())}, "
          f"sänkt pris: {int(viner['pris_före_sänkning'].notna().sum())}")
    return viner


# ---------------- 4. Nya producenter ----------------
def nya_producenter(viner, reg):
    kända = set(reg['Producent'].map(nyckel))
    v = viner.assign(_nyckel=viner['producerName'].map(nyckel))
    nya = v[v['_nyckel'].notna() & ~v['_nyckel'].isin(kända)]
    return (nya.groupby('_nyckel', sort=False)
               .agg(Producent=('producerName', 'first'), Land=('country', 'first'),
                    Ursprung1=('originLevel1', 'first'), Ursprung2=('originLevel2', 'first'),
                    Antal_produkter=('productNumber', 'count'))
               .reset_index(drop=True).sort_values(['Land', 'Producent']).reset_index(drop=True))


def skriv_att_googla(lista):
    """Skriver bara om filen när listan faktiskt har ändrats (annars blir det en ny commit varje natt)."""
    gamla = None
    if os.path.exists(ATT_GOOGLA):
        try:
            gamla = set(pd.read_excel(ATT_GOOGLA)['Producent'].astype(str))
        except Exception:
            pass
    if gamla is not None and set(lista['Producent'].astype(str)) == gamla and os.path.exists(STATUS):
        return False
    ut = lista.copy()
    ut['Adress (fyll i)'] = ''
    ut['Latitud (valfritt)'] = ''
    ut['Longitud (valfritt)'] = ''
    ut['Google-sök'] = ''
    with pd.ExcelWriter(ATT_GOOGLA, engine='openpyxl') as w:
        ut.to_excel(w, index=False, sheet_name='Att googla')
        ark = w.sheets['Att googla']
        kol = ut.shape[1]
        for i, r in lista.iterrows():
            q = requests.utils.quote(f"{r['Producent']} {r['Land'] if pd.notna(r['Land']) else ''} winery address")
            c = ark.cell(row=i + 2, column=kol, value='Google →')
            c.hyperlink = f'https://www.google.com/search?q={q}'
            c.style = 'Hyperlink'
        for bokstav, bredd in zip('ABCDEFGHI', [40, 16, 18, 18, 10, 60, 14, 14, 12]):
            ark.column_dimensions[bokstav].width = bredd
    rubrik = (f"## 🍷 {len(lista)} nya producenter väntar på att googlas" if len(lista)
              else "## ✅ Inga nya producenter att googla just nu")
    with open(STATUS, 'w', encoding='utf-8') as f:
        f.write(f"""{rubrik}

*Listan ändrades senast {IDAG}. Hemsidan uppdateras automatiskt varje natt.*

### Så här gör du
1. Klicka på **att_googla.xlsx** här ovanför och sedan på nedladdningsknappen (pil nedåt, "Download raw file").
2. Fyll i kolumnen **Adress (fyll i)** i Excel. Hittar du inte producenten: lämna raden tom (räknas som "sökt – ej hittad").
   Har du koordinater direkt kan du skriva dem i *Latitud (valfritt)* och *Longitud (valfritt)*.
3. Gå in i mappen **inlamning**, klicka **Add file → Upload files**, dra in filen och klicka **Commit changes**.
4. Klart! GitHub geokodar adresserna, uppdaterar registret och bygger om hemsidan (tar några minuter).
   Följ körningen under fliken **Actions** om du vill.
""")
    return True


# ---------------- 5. Master + data.js ----------------
def bygg_master(viner, reg):
    reg = reg.copy()
    reg['_nyckel'] = reg['Producent'].map(nyckel)
    reg = reg.sort_values('Latitud', na_position='last')
    exakt = reg.drop_duplicates('Producent').set_index('Producent')
    via_nyckel = reg.drop_duplicates('_nyckel').set_index('_nyckel')
    kol = ['Latitud', 'Longitud', 'Hittad_adress', 'Koordinat_källa', 'Manuell_sökstatus', 'Googlad_adress']
    rader = []
    for namn in viner['producerName']:
        k = nyckel(namn)
        if k is not None and namn in exakt.index:
            u, m = exakt.loc[namn], 'Exakt'
        elif k is not None and k in via_nyckel.index:
            u = via_nyckel.loc[k]
            m = f"Variant av '{u['Producent']}'"
        else:
            rader.append({**{c: None for c in kol}, 'Koordinat_källa': 'Saknas (ny producent)', 'Namnmatchning': None})
            continue
        rader.append({**{c: u[c] for c in kol}, 'Namnmatchning': m})
    return pd.concat([viner, pd.DataFrame(rader)], axis=1)


def bygg_data_js(master):
    kartan = master[master['Latitud'].notna() & master['producerName'].notna()].copy()
    kartan['_nyckel'] = kartan['producerName'].map(nyckel)
    producenter = []
    for _, g in kartan.groupby('_nyckel', sort=False):
        f = g.iloc[0]
        vinlista = []
        for _, w in g.iterrows():
            pris = float(w['price'])
            pris = int(pris) if pris.is_integer() else pris
            vin = {'namn': ren(w['productNameBold']), 'namn2': ren(w['productNameThin']),
                   'kategori': ren(w['categoryLevel2']), 'pris': pris,
                   'argang': None if tom(w['vintage']) else str(int(w['vintage'])),
                   'druvor': ren(w['grapes']), 'image': ren(w['bildlänk']), 'varunr': str(w['productNumber']),
                   'lanserad': None if tom(w['lanseringsdatum']) else str(pd.to_datetime(w['lanseringsdatum']).date()),
                   'sortiment': ren(w['assortmentText'])}
            gammalt = w['pris_före_sänkning']
            if not tom(gammalt) and pris < gammalt:
                gm = int(gammalt) if float(gammalt).is_integer() else float(gammalt)
                vin['extrapris'] = {'old': gm, 'new': pris, 'pct': math.floor((pris - gm) / gm * 100 + 0.5)}
            if w['tillfälligt_slut'] == True:
                vin['slut'] = True
            vinlista.append(vin)
        producenter.append({'name': f['producerName'], 'country': ren(f['country']),
                            'lat': float(f['Latitud']), 'lon': float(f['Longitud']),
                            'address': ren(f['Hittad_adress']), 'verified': True,
                            'wines': vinlista, 'wine_count': len(vinlista)})
    senaste = max(w['lanserad'] for p in producenter for w in p['wines'] if w['lanserad'])
    info = {'uttag': IDAG, 'nyhetsmanad': senaste[:7], 'skapad': IDAG}
    print(f"Hemsidan: {len(producenter)} producenter, {sum(len(p['wines']) for p in producenter)} viner, "
          f"nyheter för {senaste[:7]}")
    return ("// Vinografi-data. Skapas automatiskt varje natt av GitHub Actions - redigera inte för hand.\n"
            "const DATA_INFO = " + json.dumps(info, ensure_ascii=False) + ";\n"
            "const PRODUCERS = " + json.dumps(producenter, ensure_ascii=False, separators=(',', ':'), allow_nan=False) + ";\n")


def bygg_sajt(js):
    """Kopierar allt i repots rot (index.html, CNAME, ev. bilder) till _site/ och lägger till nya data.js."""
    if os.path.exists(SITE):
        shutil.rmtree(SITE)
    hoppa_over = {'.git', '.github', 'data', 'scripts', 'master', '_site', 'README.md', 'data.js'}
    shutil.copytree(ROT, SITE, ignore=lambda d, namn: [n for n in namn if d == ROT and n in hoppa_over])
    with open(os.path.join(SITE, 'data.js'), 'w', encoding='utf-8') as f:
        f.write(js)


def main():
    reg = pd.read_csv(REGISTER, encoding='utf-8-sig')
    reg = behandla_inlamning(reg)
    viner = hämta_sortiment()
    lista = nya_producenter(viner, reg)
    ändrad = skriv_att_googla(lista)
    print(f"Nya producenter att googla: {len(lista)}" + (" (listan uppdaterad)" if ändrad else ""))
    master = bygg_master(viner, reg)
    os.makedirs(MASTER_MAPP, exist_ok=True)
    master.to_excel(os.path.join(MASTER_MAPP, f'systembolaget_viner_{IDAG}_med_koordinater.xlsx'),
                    index=False, sheet_name='Systembolaget viner+koord')
    bygg_sajt(bygg_data_js(master))


if __name__ == '__main__':
    main()
