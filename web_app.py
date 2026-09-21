import streamlit as str_web
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime, timedelta
import os
from github import Github

# --- DEFINICE SLOUPCŮ PRO NOVA POST CSV ---
NOVAPOST_COLUMNS = [
    'order_number', 'recipient_phone *', 'recipient_email *', 'recipient_name *',
    'recipient_country_code *', 'recipient_delivery_type *', 'parcels_insurance_currency *',
    'payer_type *', 'parcel_0_description *', 'parcel_0_insurance_cost *',
    'parcel_0_weight *', 'parcel_0_width *', 'parcel_0_length *', 'parcel_0_height *',
    'parcel_1_description *', 'parcel_1_insurance_cost *', 'parcel_1_weight *',
    'parcel_1_width *', 'parcel_1_length *', 'parcel_1_height *',
    'parcel_2_description *', 'parcel_2_insurance_cost *', 'parcel_2_weight *',
    'parcel_2_width *', 'parcel_2_length *', 'parcel_2_height *',
    'parcel_3_description *', 'parcel_3_insurance_cost *', 'parcel_3_weight *',
    'parcel_3_width *', 'parcel_3_length *', 'parcel_3_height *',
    'parcel_4_description *', 'parcel_4_insurance_cost *', 'parcel_4_weight *',
    'parcel_4_width *', 'parcel_4_length *', 'parcel_4_height *',
    'sender_phone', 'sender_name', 'sender_email', 'recipient_company_tin',
    'recipient_company_name', 'recipient_division_address', 'recipient_division_digital_address',
    'recipient_city', 'recipient_post_code', 'recipient_AddressLine1', 'recipient_AddressLine2',
    'recipient_street', 'recipient_building', 'recipient_flat', 'recipient_note', 'promo_code',
    'cod_amount', 'cod_currency_code', 'cod_iban', 'cod_bank_account_owner_tin',
    'parcel_0_category', 'parcel_0_size', 'parcel_1_category', 'parcel_1_size',
    'parcel_2_category', 'parcel_2_size', 'parcel_3_category', 'parcel_3_size',
    'parcel_4_category', 'parcel_4_size', 'allowed_inspection', 'expected_backward_goods',
    'expected_backward_goods_description', 'expected_backward_credit_doc',
    'expected_backward_documents', 'invoice_customer_number', 'invoice_incoterm',
    'invoice_export_reason', 'invoice_cost', 'invoice_currency', 'invoice_payer_fees_customs',
    'invoice_item_id', 'invoice_item_packing_type', 'invoice_item_hs_code',
    'invoice_item_name', 'invoice_item_name_eng', 'invoice_item_material',
    'invoice_item_material_eng', 'invoice_item_made_in_country_code',
    'invoice_item_producer_and_model', 'invoice_item_actual_weight',
    'invoice_item_measurement_code', 'invoice_item_amount', 'invoice_item_cost'
]

str_web.set_page_config(page_title="Toptrans & Nova Post Převodník", layout="wide")
str_web.title("🌐 Převodník & Správce produktů (Toptrans & Nova Post)")

def nacist_katalog():
    if os.path.exists('products.xlsx'):
        df = pd.read_excel('products.xlsx')
        if 'ZBOZI_TYP_DOPRAVCE' not in df.columns:
            df['ZBOZI_TYP_DOPRAVCE'] = 'TOPTRANS'
        else:
            df['ZBOZI_TYP_DOPRAVCE'] = df['ZBOZI_TYP_DOPRAVCE'].fillna('TOPTRANS')
        return df
    else:
        return pd.DataFrame(columns=[
            'ZBOZI_2', 'ZBOZI_NAZEV', 'ZBOZI_HMOTNOST', 
            'ZBOZI_DELKA', 'ZBOZI_SIRKA', 'ZBOZI_VYSKA', 'ZBOZI_TYP_DOPRAVCE'
        ])

def ulozit_katalog(df):
    local_path = 'products.xlsx'
    df.to_excel(local_path, index=False)
    
    try:
        token = str_web.secrets["GITHUB_TOKEN"]
        repo_name = str_web.secrets["GITHUB_REPO"]
        
        g = Github(token)
        repo = g.get_repo(repo_name)
        
        with open(local_path, 'rb') as file:
            content = file.read()
            
        try:
            contents = repo.get_contents(local_path)
            repo.update_file(
                path=local_path,
                message=f"Aktualizace katalogu produktů - {datetime.now().strftime('%d.%m.%Y %H:%M')}",
                content=content,
                sha=contents.sha
            )
            str_web.toast("🌐 Data byla úspěšně zálohována na GitHub!")
        except Exception:
            repo.create_file(
                path=local_path,
                message="První vytvoření katalogu produktů",
                content=content
            )
            str_web.toast("🌐 Katalog byl nově vytvořen na GitHubu!")
            
    except Exception as e:
        str_web.error(f"⚠️ Nepodařilo se odeslat data na GitHub. Zkontrolujte nastavení v Secrets. Chyba: {e}")

if 'katalog' not in str_web.session_state:
    str_web.session_state.katalog = nacist_katalog()

if 'chybejici_fronta' not in str_web.session_state:
    str_web.session_state.chybejici_fronta = []

if 'neoverene_np_fronta' not in str_web.session_state:
    str_web.session_state.neoverene_np_fronta = []

def ziskat_baliky_pro_produkt(df_katalog, nazev_eshop, prepravce):
    match = df_katalog[df_katalog['ZBOZI_2'] == nazev_eshop]
    if match.empty:
        return match

    if prepravce == "Nova Post":
        np_match = match[match['ZBOZI_TYP_DOPRAVCE'] == 'NOVAPOST']
        if not np_match.empty:
            return np_match
        return match[match['ZBOZI_TYP_DOPRAVCE'].isin(['ALL', 'TOPTRANS'])]
    else:
        return match[match['ZBOZI_TYP_DOPRAVCE'].isin(['ALL', 'TOPTRANS'])]

zalozka1, zalozka2 = str_web.tabs(["🔄 Převodník objednávek", "🗂️ Správa katalogu produktů"])

# =========================================================
# ZÁLOŽKA 1: PŘEVODNÍK OBJEDNÁVEK
# =========================================================
with zalozka1:
    str_web.subheader("📅 Ruční úprava termínů dopravy (nepovinné)")
    
    vychozi_nakladka = datetime.now() + timedelta(days=1)
    vychozi_vykladka = datetime.now() + timedelta(days=2)
    
    col_datum1, col_datum2 = str_web.columns(2)
    with col_datum1:
        zvolena_nakladka = str_web.date_input("Datum nakládky", vychozi_nakladka, format="DD.MM.YYYY")
    with col_datum2:
        zvolena_vykladka = str_web.date_input("Datum vykládky (doručení)", vychozi_vykladka, format="DD.MM.YYYY")
        
    loading_date_str = zvolena_nakladka.strftime("%d.%m.%Y")
    discharge_date_str = zvolena_vykladka.strftime("%d.%m.%Y")

    str_web.divider()
    
    str_web.subheader("1. Nahrání exportu z e-shopu (Shoptet)")
    shoptet_url = str_web.text_input("Vložte URL adresu exportu objednávek ze Shoptetu", placeholder="např. https://www.vaseshop.cz/export/orders.csv")
    shoptet_format = str_web.radio("Formát dat ze Shoptetu", ["CSV", "Excel (.xlsx)"], horizontal=True)

    firma_volba = str_web.selectbox(
        "Vyberte firmu, pro kterou generujete export (nastavení dobírek a odesílatele):",
        options=["PR&PL s.r.o.", "Vomaks unit, s.r.o."]
    )

    if firma_volba == "PR&PL s.r.o.":
        banka_account2 = "1934179002"
        banka_kod = "5500"
        banka_iban = "CZ0855000000001934179002"
        banka_swift = "RZBCCZPP"
        
        sender_name = "Lukáš Pololáník"
        sender_phone = "+420725864066"
        sender_email = "import@max-interier.cz"
    else:
        banka_account2 = "9915665001"
        banka_kod = "5500"
        banka_iban = "CZ0955000000009915665001"
        banka_swift = "RZBCCZPP"
        
        sender_name = "Lukáš Pololáník"
        sender_phone = "+420728460271"
        sender_email = "import@vomaks.cz"

    prepravek_volba = str_web.radio(
        "Vyberte přepravní službu pro export:",
        options=["Toptrans", "Nova Post"],
        horizontal=True
    )

    # --- ZOBRAZENÍ FORMULÁŘE PRO SCHVÁLENÍ/ÚPRAVU NOVA POST BALÍKŮ ---
    if len(str_web.session_state.neoverene_np_fronta) > 0:
        str_web.error("🛑 Níže prosím potvrďte nebo upravte rozměry pro Nova Post před vygenerováním exportu:")
        
        neoverene_kopie = list(str_web.session_state.neoverene_np_fronta)
        for prod_nazev in neoverene_kopie:
            with str_web.expander(f"📦 Produkt ke schválení: **{prod_nazev}**", expanded=True):
                toptrans_radky = str_web.session_state.katalog[
                    (str_web.session_state.katalog['ZBOZI_2'] == prod_nazev) & 
                    (str_web.session_state.katalog['ZBOZI_TYP_DOPRAVCE'].isin(['ALL', 'TOPTRANS']))
                ]
                
                rozhodnuti = str_web.radio(
                    f"Vyberte akci pro '{prod_nazev}':",
                    options=["Ponechat stejné balíky z Toptransu", "Upravit balíky speciálně pro Nova Post"],
                    key=f"radio_np_{prod_nazev}"
                )
                
                if rozhodnuti == "Ponechat stejné balíky z Toptransu":
                    if str_web.button(f"✅ Potvrdit stejné balíky pro '{prod_nazev}'", key=f"btn_same_{prod_nazev}", type="primary"):
                        str_web.session_state.katalog.loc[str_web.session_state.katalog['ZBOZI_2'] == prod_nazev, 'ZBOZI_TYP_DOPRAVCE'] = 'ALL'
                        ulozit_katalog(str_web.session_state.katalog)
                        str_web.session_state.neoverene_np_fronta.remove(prod_nazev)
                        str_web.success(f"Produkt '{prod_nazev}' schválen se stejnými rozměry.")
                        str_web.rerun()
                        
                else:
                    predloha_list = toptrans_radky.to_dict('records')
                    vychozi_pocet = len(predloha_list) if len(predloha_list) > 0 else 1
                    
                    # MOŽNOST UBRAT / PŘIDAT BALÍKY
                    pocet_bal = str_web.number_input(
                        f"Počet balíků pro Nova Post ({prod_nazev}):", 
                        min_value=1, 
                        value=vychozi_pocet, 
                        step=1, 
                        key=f"np_e_cnt_{prod_nazev}"
                    )
                    
                    hmotnosti_np, delky_np, sirky_np, vysky_np, nazvy_np = [], [], [], [], []
                    
                    for i in range(pocet_bal):
                        def_val = predloha_list[i] if i < len(predloha_list) else {}
                        str_web.markdown(f"**Úprava Balíku {i+1} pro Nova Post:**")
                        nazev_i = str_web.text_input("Popis balíku", value=str(def_val.get('ZBOZI_NAZEV', f"Balík {i+1}")), key=f"np_e_nazev_{prod_nazev}_{i}")
                        nazvy_np.append(nazev_i)
                        
                        col_v, col_d, col_s, col_h = str_web.columns(4)
                        with col_v: hmotnosti_np.append(str_web.number_input("Hmotnost (kg)", value=float(str(def_val.get('ZBOZI_HMOTNOST', 0.0)).replace(',', '.')), step=0.1, key=f"np_e_vaha_{prod_nazev}_{i}"))
                        with col_d: delky_np.append(str_web.number_input("Délka (m)", value=float(str(def_val.get('ZBOZI_DELKA', 0.0)).replace(',', '.')), step=0.01, key=f"np_e_delka_{prod_nazev}_{i}"))
                        with col_s: sirky_np.append(str_web.number_input("Šířka (m)", value=float(str(def_val.get('ZBOZI_SIRKA', 0.0)).replace(',', '.')), step=0.01, key=f"np_e_sirka_{prod_nazev}_{i}"))
                        with col_h: vysky_np.append(str_web.number_input("Výška (m)", value=float(str(def_val.get('ZBOZI_VYSKA', 0.0)).replace(',', '.')), step=0.01, key=f"np_e_vyska_{prod_nazev}_{i}"))
                        
                    if str_web.button(f"💾 Uložit speciální rozměry pro Nova Post ({prod_nazev})", key=f"btn_save_custom_{prod_nazev}", type="primary"):
                        str_web.session_state.katalog.loc[(str_web.session_state.katalog['ZBOZI_2'] == prod_nazev) & (str_web.session_state.katalog['ZBOZI_TYP_DOPRAVCE'] == 'ALL'), 'ZBOZI_TYP_DOPRAVCE'] = 'TOPTRANS'
                        
                        # Smazat případné staré neověřené záznamy Nova Post
                        str_web.session_state.katalog = str_web.session_state.katalog[~((str_web.session_state.katalog['ZBOZI_2'] == prod_nazev) & (str_web.session_state.katalog['ZBOZI_TYP_DOPRAVCE'] == 'NOVAPOST'))]
                        
                        nove_np_radky = []
                        for i in range(pocet_bal):
                            nove_np_radky.append({
                                'ZBOZI_2': prod_nazev,
                                'ZBOZI_NAZEV': nazvy_np[i],
                                'ZBOZI_HMOTNOST': hmotnosti_np[i],
                                'ZBOZI_DELKA': delky_np[i],
                                'ZBOZI_SIRKA': sirky_np[i],
                                'ZBOZI_VYSKA': vysky_np[i],
                                'ZBOZI_TYP_DOPRAVCE': 'NOVAPOST'
                            })
                        
                        str_web.session_state.katalog = pd.concat([str_web.session_state.katalog, pd.DataFrame(nove_np_radky)], ignore_index=True)
                        ulozit_katalog(str_web.session_state.katalog)
                        str_web.session_state.neoverene_np_fronta.remove(prod_nazev)
                        str_web.success(f"Uloženo {pocet_bal} samostatných balíků pro Nova Post u '{prod_nazev}'.")
                        str_web.rerun()

    if str_web.button("🚀 Spustit kontrolu a generovat export", type="primary"):
        if not shoptet_url:
            str_web.warning("Zadejte alespoň URL Shoptet exportu.")
        else:
            with str_web.spinner('Stahuji a zpracovávám data...'):
                try:
                    if shoptet_format == "CSV":
                        try:
                            orders_df = pd.read_csv(shoptet_url, sep=';', encoding='windows-1250')
                        except Exception:
                            orders_df = pd.read_csv(shoptet_url)
                    else:
                        orders_df = pd.read_excel(shoptet_url)
                    
                    orders_df = orders_df[orders_df['orderItemType'].isin(['product', 'set'])]
                    
                    if 'orderItemStatusName' in orders_df.columns:
                        orders_df = orders_df[~orders_df['orderItemStatusName'].isin(['Vyřízena', 'Stornována'])]
                    eshop_produkty = orders_df['orderItemName'].dropna().unique()

                    products_df = str_web.session_state.katalog
                    sloupec_katalog_nazev = 'ZBOZI_2'
                    katalog_produkty = products_df[sloupec_katalog_nazev].dropna().unique()
                    
                    # 1. Kontrola chybějících produktů v databázi
                    chybejici = [p for p in eshop_produkty if p not in katalog_produkty]
                    
                    if len(chybejici) > 0:
                        str_web.session_state.chybejici_fronta = chybejici 
                        str_web.error("🛑 Generování přerušeno! Některé produkty chybí v katalogu rozměrů:")
                        for p in chybejici:
                            str_web.write(f"• {p}")
                        str_web.info("💡 Přejděte do záložky 'Správa katalogu produktů'.")
                    
                    else:
                        str_web.session_state.chybejici_fronta = []
                        
                        # 2. Kontrola neověřených produktů pro Nova Post
                        if prepravek_volba == "Nova Post":
                            neoverene = []
                            for p in eshop_produkty:
                                p_rows = products_df[products_df['ZBOZI_2'] == p]
                                typy = p_rows['ZBOZI_TYP_DOPRAVCE'].tolist()
                                if 'NOVAPOST' not in typy and 'ALL' not in typy:
                                    neoverene.append(p)
                                    
                            if len(neoverene) > 0:
                                str_web.session_state.neoverene_np_fronta = neoverene
                                str_web.rerun()

                        loading_date_str = zvolena_nakladka.strftime("%d.%m.%Y")
                        discharge_date_str = zvolena_vykladka.strftime("%d.%m.%Y")

                        # ==========================================
                        # OPTION A: TOPTRANS (XML EXPORT)
                        # ==========================================
                        if prepravek_volba == "Toptrans":
                            root = ET.Element("orders")
                            objednavky_skupiny = orders_df.groupby('label')

                            for cislo_objednavky, polozky_v_objednavce in objednavky_skupiny:
                                prvni_radek = polozky_v_objednavce.iloc[0]
                                order_el = ET.SubElement(root, "order")
                                
                                ET.SubElement(order_el, "label").text = str(cislo_objednavky)
                                var_sym = prvni_radek.get('var_symbol', cislo_objednavky)
                                if pd.isna(var_sym): var_sym = cislo_objednavky
                                ET.SubElement(order_el, "var_symbol").text = str(int(var_sym)) if isinstance(var_sym, float) else str(var_sym)
                                
                                ET.SubElement(order_el, "loading_select").text = "1"
                                ET.SubElement(order_el, "term_id").text = "1"
                                
                                ET.SubElement(order_el, "loading_date").text = loading_date_str
                                ET.SubElement(order_el, "loading_time_from").text = ""
                                ET.SubElement(order_el, "loading_time_to").text = ""
                                ET.SubElement(order_el, "discharge_date").text = discharge_date_str
                                ET.SubElement(order_el, "discharge_time_from").text = ""
                                ET.SubElement(order_el, "discharge_time_to").text = ""
                                
                                ET.SubElement(order_el, "loading_personal_branch_id").text = ""
                                ET.SubElement(order_el, "discharge_personal_branch_id").text = ""
                                
                                ET.SubElement(order_el, "loading")
                                
                                discharge_el = ET.SubElement(order_el, "discharge")
                                address_el = ET.SubElement(discharge_el, "address")
                                zeme = str(prvni_radek.get('country', 'Česká republika'))
                                ET.SubElement(address_el, "country").text = zeme
                                ET.SubElement(address_el, "region").text = ""
                                ET.SubElement(address_el, "city").text = str(prvni_radek.get('city', ''))
                                ET.SubElement(address_el, "city_part").text = ""
                                ET.SubElement(address_el, "street").text = str(prvni_radek.get('street', ''))
                                
                                dum = prvni_radek.get('house_num', '') 
                                if pd.notna(dum) and str(dum).strip() != "":
                                    ET.SubElement(address_el, "house_num").text = str(dum)
                                else:
                                    ET.SubElement(address_el, "house_num").text = ""
                                    
                                ET.SubElement(address_el, "zip").text = str(prvni_radek.get('zip', '')).replace(" ", "")

                                ET.SubElement(discharge_el, "name").text = str(prvni_radek.get('name', ''))
                                ET.SubElement(discharge_el, "registration_code").text = ""
                                ET.SubElement(discharge_el, "vat_code").text = ""
                                ET.SubElement(discharge_el, "first_name").text = str(prvni_radek.get('first_name', ''))
                                ET.SubElement(discharge_el, "last_name").text = str(prvni_radek.get('last_name', ''))
                                ET.SubElement(discharge_el, "phone").text = "+" + str(int(prvni_radek.get('phone'))) if pd.notna(prvni_radek.get('phone')) else ""
                                ET.SubElement(discharge_el, "email").text = str(prvni_radek.get('email', ''))

                                ET.SubElement(order_el, "loading_comfort_id").text = "1"
                                ET.SubElement(order_el, "discharge_comfort_id").text = "1"
                                ET.SubElement(order_el, "twoway_shipment").text = "0"
                                ET.SubElement(order_el, "twoway_shipment_description").text = ""
                                ET.SubElement(order_el, "yard").text = "0"
                                ET.SubElement(order_el, "delivery_notes_back").text = "0"
                                ET.SubElement(order_el, "euro_pallets_back").text = "0"
                                ET.SubElement(order_el, "loading_aviso").text = "0"
                                ET.SubElement(order_el, "discharge_aviso").text = "0"
                                ET.SubElement(order_el, "aviso_sms").text = "1"
                                ET.SubElement(order_el, "consider").text = "0"
                                ET.SubElement(order_el, "oversize").text = "0"
                                ET.SubElement(order_el, "label_fragile").text = "0"
                                ET.SubElement(order_el, "label_dont_tilt").text = "0"
                                ET.SubElement(order_el, "label_this_side_up").text = "0"
                                ET.SubElement(order_el, "hydraulic_front_loading").text = "0"
                                ET.SubElement(order_el, "hydraulic_front_discharge").text = "0"

                                cena_objednavky = prvni_radek.get('price', 0)
                                if pd.notna(cena_objednavky):
                                    val_str = str(cena_objednavky).replace(',', '.').strip()
                                    float_cena = float(val_str)
                                else:
                                    float_cena = 0.0

                                if float_cena > 0:
                                    cod_el = ET.SubElement(order_el, "cash_on_delivery")
                                    ET.SubElement(cod_el, "type").text = "1"
                                    ET.SubElement(cod_el, "price").text = str(int(float_cena))
                                    ET.SubElement(cod_el, "price_cur_id").text = "1"
                                    ET.SubElement(cod_el, "account1").text = ""
                                    ET.SubElement(cod_el, "account2").text = banka_account2 
                                    ET.SubElement(cod_el, "bank").text = banka_kod
                                    ET.SubElement(cod_el, "iban").text = banka_iban
                                    ET.SubElement(cod_el, "swift").text = banka_swift

                                celkova_vaha = 0
                                kg_el = ET.SubElement(order_el, "kg")
                                
                                ET.SubElement(order_el, "m3").text = ""
                                ET.SubElement(order_el, "order_value").text = str(int(float_cena))
                                ET.SubElement(order_el, "order_value_currency_id").text = "1"
                                
                                ET.SubElement(order_el, "note_loading").text = ""
                                ET.SubElement(order_el, "note_discharge").text = ""
                                
                                ET.SubElement(order_el, "return_pack_id").text = "0"
                                ET.SubElement(order_el, "return_pack_count").text = "0"
                                ET.SubElement(order_el, "return_pack_description").text = ""

                                packs_el = ET.SubElement(order_el, "packs")
                                
                                for index, row in polozky_v_objednavce.iterrows():
                                    nazev_produktu_eshop = row['orderItemName']
                                    mnozstvi_produktu = row['orderItemAmount']
                                    if pd.isna(mnozstvi_produktu): mnozstvi_produktu = 1
                                    
                                    nalezeno = ziskat_baliky_pro_produkt(products_df, nazev_produktu_eshop, "Toptrans")
                                    
                                    for _, balik in nalezeno.iterrows():
                                        pack_el = ET.SubElement(packs_el, "pack")
                                        celkove_mnozstvi = int(mnozstvi_produktu)
                                        ET.SubElement(pack_el, "quantity").text = str(celkove_mnozstvi)
                                        ET.SubElement(pack_el, "pack_id").text = "1"
                                        ET.SubElement(pack_el, "description").text = str(balik['ZBOZI_NAZEV'])[:50]
                                        
                                        vaha_baliku = float(str(balik.get('ZBOZI_HMOTNOST', 0)).replace(',', '.')) if pd.notna(balik.get('ZBOZI_HMOTNOST')) else 0.0
                                        celkova_vaha += (vaha_baliku * celkove_mnozstvi)
                                        
                                        delka_m = float(str(balik.get('ZBOZI_DELKA', 0)).replace(',', '.')) if pd.notna(balik.get('ZBOZI_DELKA')) else 0.0
                                        sirka_m = float(str(balik.get('ZBOZI_SIRKA', 0)).replace(',', '.')) if pd.notna(balik.get('ZBOZI_SIRKA')) else 0.0
                                        vyska_m = float(str(balik.get('ZBOZI_VYSKA', 0)).replace(',', '.')) if pd.notna(balik.get('ZBOZI_VYSKA')) else 0.0
                                        
                                        if delka_m > 0:
                                            ET.SubElement(pack_el, "dimensions_d").text = str(int(delka_m * 100))
                                        if sirka_m > 0:
                                            ET.SubElement(pack_el, "dimensions_s").text = str(int(sirka_m * 100))
                                        if vyska_m > 0:
                                            ET.SubElement(pack_el, "dimensions_v").text = str(int(vyska_m * 100))
                                
                                upravena_vaha = int(celkova_vaha)
                                vahove_limity = [5, 15, 30, 50, 75, 100, 150, 200, 300, 400, 500]
                                for limit in vahove_limity:
                                    if limit <= upravena_vaha <= (limit + 4):
                                        upravena_vaha = limit
                                        break
                                kg_el.text = str(upravena_vaha)

                            xml_str = ET.tostring(root, encoding='utf-8')
                            pretty_xml = minidom.parseString(xml_str).toprettyxml(indent="    ")
                            pretty_xml = '\n'.join([line for line in pretty_xml.split('\n') if line.strip()])
                            finalni_xml = '<?xml version="1.0" encoding="UTF-8"?>\n' + (pretty_xml.split('?>\n', 1)[1] if '?>' in pretty_xml else pretty_xml)

                            str_web.success("🎉 TOPTRANS XML export byl úspěšně vygenerován!")
                            str_web.download_button(
                                label="💾 Stáhnout TOPTRANS XML",
                                data=finalni_xml,
                                file_name="HOTOVY_EXPORT_TOPTRANS.xml",
                                mime="text/xml"
                            )

                        # ==========================================
                        # OPTION B: NOVA POST (CSV EXPORT)
                        # ==========================================
                        else:
                            novapost_radky = []
                            objednavky_skupiny = orders_df.groupby('label')

                            for cislo_objednavky, polozky_v_objednavce in objednavky_skupiny:
                                prvni_radek = polozky_v_objednavce.iloc[0]
                                
                                radek = {col: "" for col in NOVAPOST_COLUMNS}
                                
                                radek['sender_name'] = sender_name
                                radek['sender_phone'] = sender_phone
                                radek['sender_email'] = sender_email

                                radek['order_number'] = str(cislo_objednavky)
                                
                                tel = prvni_radek.get('phone', '')
                                radek['recipient_phone *'] = "+" + str(int(tel)) if pd.notna(tel) and str(tel) != "" else ""
                                radek['recipient_email *'] = str(prvni_radek.get('email', ''))
                                
                                jmeno = str(prvni_radek.get('first_name', '')).strip()
                                prijmeni = str(prvni_radek.get('last_name', '')).strip()
                                radek['recipient_name *'] = f"{jmeno} {prijmeni}".strip() or str(prvni_radek.get('name', ''))
                                
                                radek['recipient_country_code *'] = "CZ" if str(prvni_radek.get('country', '')).lower() in ['česká republika', 'cz', 'czechia'] else "PL"
                                radek['recipient_delivery_type *'] = 'address'
                                radek['parcels_insurance_currency *'] = 'CZK'
                                radek['payer_type *'] = 'sender'
                                
                                baliky_seznam = []
                                for index, row in polozky_v_objednavce.iterrows():
                                    nazev_produktu_eshop = row['orderItemName']
                                    mnozstvi = int(row['orderItemAmount']) if pd.notna(row['orderItemAmount']) else 1
                                    
                                    nalezeno = ziskat_baliky_pro_produkt(products_df, nazev_produktu_eshop, "Nova Post")
                                    for _, b in nalezeno.iterrows():
                                        for _ in range(mnozstvi):
                                            baliky_seznam.append(b)

                                if 'price_insurance' in prvni_radek and pd.notna(prvni_radek['price_insurance']):
                                    val_str = str(prvni_radek['price_insurance']).replace(',', '.').strip()
                                    hodnota_pojisteni = float(val_str)
                                else:
                                    cena_objednavky = prvni_radek.get('price', 0)
                                    val_str = str(cena_objednavky).replace(',', '.').strip() if pd.notna(cena_objednavky) else "0"
                                    hodnota_pojisteni = float(val_str)

                                pocet_baliku = len(baliky_seznam) if len(baliky_seznam) > 0 else 1
                                cena_na_balik = int(hodnota_pojisteni / pocet_baliku)

                                for i in range(min(5, len(baliky_seznam))):
                                    b = baliky_seznam[i]
                                    vaha = float(str(b.get('ZBOZI_HMOTNOST', 0)).replace(',', '.')) if pd.notna(b.get('ZBOZI_HMOTNOST')) else 0.0
                                    delka_cm = int(float(str(b.get('ZBOZI_DELKA', 0)).replace(',', '.')) * 100) if pd.notna(b.get('ZBOZI_DELKA')) else 20
                                    sirka_cm = int(float(str(b.get('ZBOZI_SIRKA', 0)).replace(',', '.')) * 100) if pd.notna(b.get('ZBOZI_SIRKA')) else 20
                                    vyska_cm = int(float(str(b.get('ZBOZI_VYSKA', 0)).replace(',', '.')) * 100) if pd.notna(b.get('ZBOZI_VYSKA')) else 20

                                    radek[f'parcel_{i}_description *'] = str(b.get('ZBOZI_NAZEV', 'Zboží'))[:50]
                                    radek[f'parcel_{i}_insurance_cost *'] = str(cena_na_balik)
                                    radek[f'parcel_{i}_weight *'] = str(vaha)
                                    radek[f'parcel_{i}_width *'] = str(sirka_cm)
                                    radek[f'parcel_{i}_length *'] = str(delka_cm)
                                    radek[f'parcel_{i}_height *'] = str(vyska_cm)
                                    radek[f'parcel_{i}_category'] = 'parcel'
                                    radek[f'parcel_{i}_size'] = f"{delka_cm}*{sirka_cm}*{vyska_cm}"

                                radek['recipient_city'] = str(prvni_radek.get('city', ''))
                                radek['recipient_post_code'] = str(prvni_radek.get('zip', '')).replace(" ", "")
                                
                                ulice = str(prvni_radek.get('street', '')).strip()
                                cislo_popisne = str(prvni_radek.get('house_num', '')).strip()
                                
                                radek['recipient_street'] = ulice
                                radek['recipient_building'] = cislo_popisne if pd.notna(prvni_radek.get('house_num')) else ""
                                radek['recipient_AddressLine1'] = ""
                                
                                cena_objednavky = prvni_radek.get('price', 0)
                                if pd.notna(cena_objednavky):
                                    val_str = str(cena_objednavky).replace(',', '.').strip()
                                    float_cena = float(val_str)
                                else:
                                    float_cena = 0.0

                                if float_cena > 0:
                                    radek['cod_amount'] = str(int(float_cena))
                                    radek['cod_currency_code'] = 'CZK'
                                    radek['cod_iban'] = banka_iban
                                    
                                novapost_radky.append(radek)

                            df_np = pd.DataFrame(novapost_radky, columns=NOVAPOST_COLUMNS)
                            csv_bytes = df_np.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

                            str_web.success("🎉 NOVA POST CSV export byl úspěšně vygenerován!")
                            str_web.download_button(
                                label="💾 Stáhnout NOVA POST CSV",
                                data=csv_bytes,
                                file_name="HOTOVY_EXPORT_NOVAPOST.csv",
                                mime="text/csv"
                            )

                except Exception as e:
                    str_web.error(f"❌ Chyba při zpracování dat: {e}")

# =========================================================
# ZÁLOŽKA 2: SPRÁVA KATALOGU PRODUKTŮ
# =========================================================
with zalozka2:
    str_web.header("📦 Správa a nastavení rozměrů produktů")
    
    rezim_akce = str_web.radio(
        "Zvolte požadovanou akci:",
        options=["➕ Přidat nový produkt", "✏️ Upravit / vytvořit rozměry pro Nova Post z Toptransu"],
        horizontal=True
    )

    df_kat = str_web.session_state.katalog

    # --- MOŽNOST A: PŘIDÁNÍ NOVÉHO PRODUKTU ---
    if rezim_akce == "➕ Přidat nový produkt":
        vybrany_chybejici = ""
        if len(str_web.session_state.chybejici_fronta) > 0:
            str_web.warning("⚠️ Máte nevyřešené chybějící produkty z posledního převodu!")
            vybrany_chybejici = str_web.selectbox(
                "Vyberte produkt, který chcete nyní doplnit:", 
                options=["-- Vyberte produkt z fronty --"] + str_web.session_state.chybejici_fronta
            )
        
        vychozi_nazev = vybrany_chybejici if vybrany_chybejici != "-- Vyberte produkt z fronty --" else ""

        zbozi_2 = str_web.text_input("Přesný název ze Shoptetu", value=vychozi_nazev)
        zbozi_nazev = str_web.text_input("Základní označení pro kurýra (např. Postel KOBE)")
        
        typ_dopravce_volba = str_web.radio(
            "Platnost těchto rozměrů:",
            options=["Stejné pro Toptrans i Nova Post", "Pouze pro Toptrans", "Pouze pro Nova Post"],
            index=0
        )
        
        mapovani_typu = {
            "Stejné pro Toptrans i Nova Post": "ALL",
            "Pouze pro Toptrans": "TOPTRANS",
            "Pouze pro Nova Post": "NOVAPOST"
        }
        
        pocet_baliku = str_web.number_input("Počet balíků (krabic) pro tento produkt", min_value=1, value=1, step=1)
        
        hmotnosti, delky, sirky, vysky = [], [], [], []
        for i in range(pocet_baliku):
            oznaceni = f" {i+1}/{pocet_baliku}" if pocet_baliku > 1 else ""
            with str_web.container(border=True):
                str_web.markdown(f"**📦 Balík {i+1}** (`{zbozi_nazev}{oznaceni}`)")
                col_v, col_d, col_s, col_h = str_web.columns(4)
                with col_v: hmotnosti.append(str_web.number_input("Hmotnost (kg)", min_value=0.0, step=0.1, key=f"vaha_{i}"))
                with col_d: delky.append(str_web.number_input("Délka (m)", min_value=0.0, step=0.01, key=f"delka_{i}"))
                with col_s: sirky.append(str_web.number_input("Šířka (m)", min_value=0.0, step=0.01, key=f"sirka_{i}"))
                with col_h: vysky.append(str_web.number_input("Výška (m)", min_value=0.0, step=0.01, key=f"vyska_{i}"))
                
        if str_web.button("💾 Uložit nový produkt do databáze", type="primary"):
            if zbozi_2.strip() == "" or zbozi_nazev.strip() == "":
                str_web.error("Shoptet název i označení pro kurýra musí být vyplněné!")
            else:
                nove_radky = []
                for i in range(pocet_baliku):
                    oznaceni = f" {i+1}/{pocet_baliku}" if pocet_baliku > 1 else ""
                    nove_radky.append({
                        'ZBOZI_2': zbozi_2.strip(),
                        'ZBOZI_NAZEV': f"{zbozi_nazev.strip()}{oznaceni}",
                        'ZBOZI_HMOTNOST': hmotnosti[i],
                        'ZBOZI_DELKA': delky[i],
                        'ZBOZI_SIRKA': sirky[i],
                        'ZBOZI_VYSKA': vysky[i],
                        'ZBOZI_TYP_DOPRAVCE': mapovani_typu[typ_dopravce_volba]
                    })
                
                str_web.session_state.katalog = pd.concat([str_web.session_state.katalog, pd.DataFrame(nove_radky)], ignore_index=True)
                ulozit_katalog(str_web.session_state.katalog)
                
                if zbozi_2 in str_web.session_state.chybejici_fronta:
                    str_web.session_state.chybejici_fronta.remove(zbozi_2)
                
                str_web.success(f"Úspěšně vloženo: {pocet_baliku} balík(ů) pro produkt '{zbozi_2}'.")
                str_web.rerun()

    # --- MOŽNOST B: SPECIFICKÉ BALÍKY PRO NOVA POST ---
    else:
        vsechny_produkty = sorted(df_kat['ZBOZI_2'].dropna().unique().tolist())
        if not vsechny_produkty:
            str_web.info("V katalogu zatím nemáte žádné produkty.")
        else:
            vybrany_prod = str_web.selectbox("Vyberte produkt z databáze:", options=vsechny_produkty)
            
            existujici_np = df_kat[(df_kat['ZBOZI_2'] == vybrany_prod) & (df_kat['ZBOZI_TYP_DOPRAVCE'] == 'NOVAPOST')]
            toptrans_baliky = df_kat[(df_kat['ZBOZI_2'] == vybrany_prod) & (df_kat['ZBOZI_TYP_DOPRAVCE'].isin(['ALL', 'TOPTRANS']))]
            
            if not existujici_np.empty:
                str_web.info("ℹ️ Tento produkt již MÁ samostatně definované balíky pro Nova Post. Zde je můžete upravit.")
                predloha = existujici_np
            else:
                str_web.success("💡 Tento produkt aktuálně používá stejné balíky z Toptransu. Níže jsou předvyplněné k úpravě pro Nova Post.")
                predloha = toptrans_baliky

            pocet_baliku_np = str_web.number_input("Počet balíků pro Nova Post", min_value=1, value=len(predloha) if len(predloha) > 0 else 1, step=1)
            
            hmotnosti_np, delky_np, sirky_np, vysky_np, nazvy_np = [], [], [], [], []
            
            predloha_list = predloha.to_dict('records')
            
            for i in range(pocet_baliku_np):
                default_val = predloha_list[i] if i < len(predloha_list) else {}
                with str_web.container(border=True):
                    str_web.markdown(f"**📦 Nova Post Balík {i+1}**")
                    nazev_i = str_web.text_input("Popis balíku", value=str(default_val.get('ZBOZI_NAZEV', f"Balík {i+1}")), key=f"np_nazev_{i}")
                    nazvy_np.append(nazev_i)
                    
                    col_v, col_d, col_s, col_h = str_web.columns(4)
                    with col_v: hmotnosti_np.append(str_web.number_input("Hmotnost (kg)", value=float(str(default_val.get('ZBOZI_HMOTNOST', 0.0)).replace(',', '.')), step=0.1, key=f"np_vaha_{i}"))
                    with col_d: delky_np.append(str_web.number_input("Délka (m)", value=float(str(default_val.get('ZBOZI_DELKA', 0.0)).replace(',', '.')), step=0.01, key=f"np_delka_{i}"))
                    with col_s: sirky_np.append(str_web.number_input("Šířka (m)", value=float(str(default_val.get('ZBOZI_SIRKA', 0.0)).replace(',', '.')), step=0.01, key=f"np_sirka_{i}"))
                    with col_h: vysky_np.append(str_web.number_input("Výška (m)", value=float(str(default_val.get('ZBOZI_VYSKA', 0.0)).replace(',', '.')), step=0.01, key=f"np_vyska_{i}"))
            
            col_save, col_reset = str_web.columns([2, 1])
            with col_save:
                if str_web.button("💾 Uložit speciální balíky pro Nova Post", type="primary"):
                    str_web.session_state.katalog.loc[(str_web.session_state.katalog['ZBOZI_2'] == vybrany_prod) & (str_web.session_state.katalog['ZBOZI_TYP_DOPRAVCE'] == 'ALL'), 'ZBOZI_TYP_DOPRAVCE'] = 'TOPTRANS'
                    
                    str_web.session_state.katalog = str_web.session_state.katalog[~((str_web.session_state.katalog['ZBOZI_2'] == vybrany_prod) & (str_web.session_state.katalog['ZBOZI_TYP_DOPRAVCE'] == 'NOVAPOST'))]
                    
                    nove_np_radky = []
                    for i in range(pocet_baliku_np):
                        nove_np_radky.append({
                            'ZBOZI_2': vybrany_prod,
                            'ZBOZI_NAZEV': nazvy_np[i],
                            'ZBOZI_HMOTNOST': hmotnosti_np[i],
                            'ZBOZI_DELKA': delky_np[i],
                            'ZBOZI_SIRKA': sirky_np[i],
                            'ZBOZI_VYSKA': vysky_np[i],
                            'ZBOZI_TYP_DOPRAVCE': 'NOVAPOST'
                        })
                    
                    str_web.session_state.katalog = pd.concat([str_web.session_state.katalog, pd.DataFrame(nove_np_radky)], ignore_index=True)
                    ulozit_katalog(str_web.session_state.katalog)
                    str_web.success(f"✨ Úspěšně uloženy rozměry Nova Post pro '{vybrany_prod}'. Původní rozměry pro Toptrans zůstaly netknuté!")
                    str_web.rerun()

            with col_reset:
                if not existujici_np.empty:
                    if str_web.button("🔄 Vrátit k Toptransu (Smazat Nova Post balíky)"):
                        str_web.session_state.katalog = str_web.session_state.katalog[~((str_web.session_state.katalog['ZBOZI_2'] == vybrany_prod) & (str_web.session_state.katalog['ZBOZI_TYP_DOPRAVCE'] == 'NOVAPOST'))]
                        str_web.session_state.katalog.loc[(str_web.session_state.katalog['ZBOZI_2'] == vybrany_prod) & (str_web.session_state.katalog['ZBOZI_TYP_DOPRAVCE'] == 'TOPTRANS'), 'ZBOZI_TYP_DOPRAVCE'] = 'ALL'
                        ulozit_katalog(str_web.session_state.katalog)
                        str_web.success(f"Dopravce Nova Post bude u produktu '{vybrany_prod}' opět přebírat balíky z Toptransu.")
                        str_web.rerun()

    str_web.divider()
    str_web.subheader("✏️ Přehled a manuální úprava databáze")
    str_web.caption("V sloupci 'ZBOZI_TYP_DOPRAVCE' značí **ALL** společný balík, **TOPTRANS** balík pouze pro Toptrans a **NOVAPOST** balík pouze pro Nova Post.")
    
    upravena_data = str_web.data_editor(
        str_web.session_state.katalog, 
        use_container_width=True,
        num_rows="dynamic",
        key="katalog_editor"
    )
    
    if str_web.button("💾 Definitivně uložit změny a mazání", type="primary"):
        str_web.session_state.katalog = upravena_data
        ulozit_katalog(upravena_data)
        str_web.success("✨ Všechny změny byly úspěšně uloženy do souboru products.xlsx!")
        str_web.rerun()