import streamlit as str_web
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime, timedelta
import urllib.error
import os
from github import Github  # <--- TENTO NOVÝ ŘÁDEK PŘIDEJTE

str_web.set_page_config(page_title="Toptrans Převodník", layout="wide")
str_web.title("🌐 Toptrans Převodník & Správce produktů")

def nacist_katalog():
    if os.path.exists('products.xlsx'):
        return pd.read_excel('products.xlsx')
    else:
        return pd.DataFrame(columns=['ZBOZI_2', 'ZBOZI_NAZEV', 'ZBOZI_HMOTNOST', 'ZBOZI_DELKA', 'ZBOZI_SIRKA', 'ZBOZI_VYSKA'])

def ulozit_katalog(df):
    # 1. Nejdříve uložíme Excel lokálně na disk serveru
    local_path = 'products.xlsx'
    df.to_excel(local_path, index=False)
    
    # 2. Pokusíme se odeslat soubor na GitHub pomocí tokenu z trezoru
    try:
        # Načtení přihlašovacích údajů ze Secrets trezoru
        token = str_web.secrets["GITHUB_TOKEN"]
        repo_name = str_web.secrets["GITHUB_REPO"]
        
        g = Github(token)
        repo = g.get_repo(repo_name)
        
        # Přečteme nově uložený lokální soubor jako binární data
        with open(local_path, 'rb') as file:
            content = file.read()
            
        try:
            # Zkusíme zjistit, zda už soubor na GitHubu existuje (potřebujeme jeho SHA kód pro přepis)
            contents = repo.get_contents(local_path)
            repo.update_file(
                path=local_path,
                message=f"Aktualizace katalogu produktů - {datetime.now().strftime('%d.%m.%Y %H:%M')}",
                content=content,
                sha=contents.sha
            )
            str_web.toast("🌐 Data byla úspěšně zálohována na GitHub!")
        except Exception:
            # Pokud soubor na GitHubu ještě vůbec není, vytvoříme ho jako nový
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

# Vytvoříme paměť pro "čekající" chybějící produkty
if 'chybejici_fronta' not in str_web.session_state:
    str_web.session_state.chybejici_fronta = []

zalozka1, zalozka2 = str_web.tabs(["🔄 Převodník objednávek", "🗂️ Správa katalogu produktů"])

# =========================================================
# ZÁLOŽKA 1: PŘEVODNÍK OBJEDNÁVEK
# =========================================================
with zalozka1:
    str_web.subheader("1. Zdroj dat (Shoptet Objednávky z URL)")
    with zalozka1:
    str_web.subheader("📅 Ruční úprava termínů dopravy (nepovinné)")
    
    # Výchozí výpočty (zítřek a pozítří) jako doposud
    vychozi_nakladka = datetime.now() + timedelta(days=1)
    vychozi_vykladka = datetime.now() + timedelta(days=2)
    
    # Zobrazení dvou kalendářů vedle sebe
    col_datum1, col_datum2 = str_web.columns(2)
    with col_datum1:
        zvolena_nakladka = str_web.date_input("Datum nakládky", vychozi_nakladka, format="DD.MM.YYYY")
    with col_datum2:
        zvolena_vykladka = str_web.date_input("Datum vykládky (doručení)", vychozi_vykladka, format="DD.MM.YYYY")
        
    # Převedení vybraných datumů na textový formát pro Toptrans šablonu
    loading_date_str = zvolena_nakladka.strftime("%d.%m.%Y")
    discharge_date_str = zvolena_vykladka.strftime("%d.%m.%Y")

    str_web.divider()
    
    str_web.subheader("1. Nahrání exportu z e-shopu (Shoptet)")
    shoptet_url = str_web.text_input("Vložte URL adresu exportu objednávek ze Shoptetu", placeholder="např. https://www.vaseshop.cz/export/orders.csv")
    shoptet_format = str_web.radio("Formát dat ze Shoptetu", ["CSV", "Excel (.xlsx)"], horizontal=True)

    # Výběr firmy pro správné bankovní údaje
    firma_volba = str_web.selectbox(
        "Vyberte firmu, pro kterou generujete export (nastavení dobírek):",
        options=["PR&PL s.r.o.", "Vomaks unit, s.r.o."]
    )

    # Definice bankovních údajů podle zvolené firmy
    if firma_volba == "PR&PL s.r.o.":
        banka_account2 = "1934179002"
        banka_kod = "5500"
        banka_iban = "CZ0855000000001934179002"
        banka_swift = "RZBCCZPP"
    else:
        # ZDE SI UPRAVTE ÚDAJE PRO DRUHOU FIRMU
        banka_account2 = "9915665001"       # Číslo účtu druhé firmy
        banka_kod = "5500"                # Kód banky (např. 0100 pro KB)
        banka_iban = "CZ0955000000009915665001"
        banka_swift = "RZBCCZPP"

    if str_web.button("🚀 Spustit kontrolu a generovat XML", type="primary"):
        if not shoptet_url:
            str_web.warning("Zadejte alespoň URL Shoptet exportu.")
        else:
            with str_web.spinner('Stahuji a zpracovávám data...'):
                try:
                    # Načtení Shoptet dat
                    if shoptet_format == "CSV":
                        try:
                            orders_df = pd.read_csv(shoptet_url, sep=';', encoding='windows-1250')
                        except Exception:
                            orders_df = pd.read_csv(shoptet_url)
                    else:
                        orders_df = pd.read_excel(shoptet_url)
                    
                    orders_df = orders_df[orders_df['orderItemType'].isin(['product', 'set'])]
                    
                    # OPRAVENÝ FILTR: Ignorovat vyřízené a stornované objednávky podle správného sloupce
                    if 'orderItemStatusName' in orders_df.columns:
                        orders_df = orders_df[~orders_df['orderItemStatusName'].isin(['Vyřízena', 'Stornována'])]
                    eshop_produkty = orders_df['orderItemName'].dropna().unique()

                    # Načtení katalogu (online nebo z lokálního souboru/session_state)
                    products_df = str_web.session_state.katalog

                    sloupec_katalog_nazev = 'ZBOZI_2'
                    katalog_produkty = products_df[sloupec_katalog_nazev].dropna().unique()
                    
                    chybejici = [p for p in eshop_produkty if p not in katalog_produkty]
                    
                    if len(chybejici) > 0:
                        # ULOŽENÍ DO FRONTY PRO ZÁLOŽKU 2
                        str_web.session_state.chybejici_fronta = chybejici 
                        
                        str_web.error("🛑 Generování přerušeno! Některé produkty chybí v katalogu rozměrů:")
                        for p in chybejici:
                            str_web.write(f"• {p}")
                        str_web.info("💡 Přejděte do záložky 'Správa katalogu produktů'. Názvy tam máte připravené k rychlému vyplnění.")
                    else:
                        str_web.session_state.chybejici_fronta = [] # Vymazání fronty
                        loading_date_str = zvolena_nakladka.strftime("%d.%m.%Y")
                        discharge_date_str = zvolena_vykladka.strftime("%d.%m.%Y")

                        root = ET.Element("orders")
                        objednavky_skupiny = orders_df.groupby('label')

                        # --- GENERUJEME JEDNOTLIVÉ OBJEDNÁVKY ---
                        for cislo_objednavky, polozky_v_objednavce in objednavky_skupiny:
                            prvni_radek = polozky_v_objednavce.iloc[0]
                            order_el = ET.SubElement(root, "order")
                            
                            # --- HLAVIČKA OBJEDNÁVKY ---
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
                            
                            # --- VYKLÁDKA (Příjemce z e-shopu) ---
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

                            # --- DALŠÍ SLUŽBY ---
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

                            # --- DOBÍRKA ---
                            cena_objednavky = prvni_radek.get('price', 0)
                            if pd.notna(cena_objednavky) and float(cena_objednavky) > 0:
                                cod_el = ET.SubElement(order_el, "cash_on_delivery")
                                ET.SubElement(cod_el, "type").text = "1"
                                ET.SubElement(cod_el, "price").text = str(int(float(cena_objednavky)))
                                ET.SubElement(cod_el, "price_cur_id").text = "1"
                                ET.SubElement(cod_el, "account1").text = ""
                                ET.SubElement(cod_el, "account2").text = banka_account2 
                                ET.SubElement(cod_el, "bank").text = banka_kod
                                ET.SubElement(cod_el, "iban").text = banka_iban
                                ET.SubElement(cod_el, "swift").text = banka_swift

                            # --- CELKOVÉ ÚDAJE A POZNÁMKY ---
                            celkova_vaha = 0
                            kg_el = ET.SubElement(order_el, "kg")
                            
                            ET.SubElement(order_el, "m3").text = ""
                            ET.SubElement(order_el, "order_value").text = str(int(float(cena_objednavky))) if pd.notna(cena_objednavky) else "0"
                            ET.SubElement(order_el, "order_value_currency_id").text = "1"
                            
                            ET.SubElement(order_el, "note_loading").text = ""
                            ET.SubElement(order_el, "note_discharge").text = ""
                            
                            ET.SubElement(order_el, "return_pack_id").text = "0"
                            ET.SubElement(order_el, "return_pack_count").text = "0"
                            ET.SubElement(order_el, "return_pack_description").text = ""

                            # --- BALÍKY (Packs) ---
                            packs_el = ET.SubElement(order_el, "packs")
                            
                            for index, row in polozky_v_objednavce.iterrows():
                                nazev_produktu_eshop = row['orderItemName']
                                mnozstvi_produktu = row['orderItemAmount']
                                if pd.isna(mnozstvi_produktu): mnozstvi_produktu = 1
                                
                                nalezeno = products_df[products_df['ZBOZI_2'] == nazev_produktu_eshop]
                                
                                for _, balik in nalezeno.iterrows():
                                    pack_el = ET.SubElement(packs_el, "pack")
                                    
                                    celkove_mnozstvi = int(mnozstvi_produktu)
                                    ET.SubElement(pack_el, "quantity").text = str(celkove_mnozstvi)
                                    ET.SubElement(pack_el, "pack_id").text = "1"
                                    ET.SubElement(pack_el, "description").text = str(balik['ZBOZI_NAZEV'])[:50]
                                    
                                    vaha_baliku = balik.get('ZBOZI_HMOTNOST', 0)
                                    if pd.notna(vaha_baliku):
                                        celkova_vaha += (float(vaha_baliku) * celkove_mnozstvi)
                                    
                                    delka_m = balik.get('ZBOZI_DELKA', 0)
                                    sirka_m = balik.get('ZBOZI_SIRKA', 0)
                                    vyska_m = balik.get('ZBOZI_VYSKA', 0)
                                    
                                    if pd.notna(delka_m) and float(delka_m) > 0:
                                        ET.SubElement(pack_el, "dimensions_d").text = str(int(float(delka_m) * 100))
                                    if pd.notna(sirka_m) and float(sirka_m) > 0:
                                        ET.SubElement(pack_el, "dimensions_s").text = str(int(float(sirka_m) * 100))
                                    if pd.notna(vyska_m) and float(vyska_m) > 0:
                                        ET.SubElement(pack_el, "dimensions_v").text = str(int(float(vyska_m) * 100))
                            
                            kg_el.text = str(int(celkova_vaha))

                        # --- ZPRACOVÁNÍ TEXTU XML A TLAČÍTKO KE STAŽENÍ ---
                        xml_str = ET.tostring(root, encoding='utf-8')
                        pretty_xml = minidom.parseString(xml_str).toprettyxml(indent="    ")
                        pretty_xml = '\n'.join([line for line in pretty_xml.split('\n') if line.strip()])
                        
                        finalni_xml = '<?xml version="1.0" encoding="UTF-8"?>\n' + (pretty_xml.split('?>\n', 1)[1] if '?>' in pretty_xml else pretty_xml)

                        str_web.success("🎉 XML export byl úspěšně vygenerován a připraven ke stažení!")
                        
                        str_web.download_button(
                            label="💾 Stáhnout HOTOVY_EXPORT.xml",
                            data=finalni_xml,
                            file_name="HOTOVY_EXPORT.xml",
                            mime="text/xml"
                        )
                except Exception as e:
                    str_web.error(f"❌ Chyba při zpracování dat: {e}")

# =========================================================
# ZÁLOŽKA 2: SPRÁVA KATALOGU PRODUKTŮ
# =========================================================
with zalozka2:
    str_web.header("Přidání produktu do katalogu")
    
    vybrany_chybejici = ""
    if len(str_web.session_state.chybejici_fronta) > 0:
        str_web.warning("⚠️ Máte nevyřešené chybějící produkty z posledního převodu!")
        vybrany_chybejici = str_web.selectbox(
            "Vyberte produkt, který chcete nyní doplnit:", 
            options=["-- Vyberte produkt z fronty --"] + str_web.session_state.chybejici_fronta
        )
    
    vychozi_nazev = vybrany_chybejici if vybrany_chybejici != "-- Vyberte produkt z fronty --" else ""

    # Upozornění: Pro dynamické formuláře nesmíme použít str_web.form,
    # protože klasický form nedovoluje dynamicky měnit počet prvků uvnitř na základě jiného prvku.
    # Použijeme běžné uspořádání.
    
    str_web.subheader("Základní informace")
    zbozi_2 = str_web.text_input("Přesný název ze Shoptetu", value=vychozi_nazev)
    zbozi_nazev = str_web.text_input("Základní označení pro kurýra (např. Postel KOBE)")
    
    # Výběr počtu balíků mimo formulář, aby stránka mohla ihned reagovat
    pocet_baliku = str_web.number_input("Počet různých balíků (krabic) pro tento produkt", min_value=1, value=1, step=1)
    
    str_web.divider()
    str_web.subheader("Rozměry jednotlivých balíků")
    
    # Zde si vytvoříme seznamy, do kterých budeme ukládat hodnoty z jednotlivých políček
    hmotnosti = []
    delky = []
    sirky = []
    vysky = []
    
    # Dynamicky vykreslíme políčka pro každý balík
    for i in range(pocet_baliku):
        oznaceni = f" {i+1}/{pocet_baliku}" if pocet_baliku > 1 else ""
        
        # Každý balík dostane svůj rámeček
        with str_web.container(border=True):
            str_web.markdown(f"**📦 Balík {i+1}** (bude pojmenován jako: `{zbozi_nazev}{oznaceni}`)")
            
            # Políčka dáme hezky vedle sebe
            col_v, col_d, col_s, col_h = str_web.columns(4)
            with col_v:
                # DŮLEŽITÉ: Každé políčko musí mít unikátní key (např. vaha_0, vaha_1), jinak Streamlit spadne
                h = str_web.number_input("Hmotnost (kg)", min_value=0.0, step=0.1, key=f"vaha_{i}")
                hmotnosti.append(h)
            with col_d:
                d = str_web.number_input("Délka (m)", min_value=0.0, step=0.01, key=f"delka_{i}")
                delky.append(d)
            with col_s:
                s = str_web.number_input("Šířka (m)", min_value=0.0, step=0.01, key=f"sirka_{i}")
                sirky.append(s)
            with col_h:
                v = str_web.number_input("Výška (m)", min_value=0.0, step=0.01, key=f"vyska_{i}")
                vysky.append(v)
            
    # Samostatné tlačítko pro uložení
    if str_web.button("💾 Uložit do databáze", type="primary"):
        if zbozi_2.strip() == "" or zbozi_nazev.strip() == "":
            str_web.error("Shoptet název i označení pro kurýra musí být vyplněné!")
        else:
            nove_radky = []
            for i in range(pocet_baliku):
                oznaceni = f" {i+1}/{pocet_baliku}" if pocet_baliku > 1 else ""
                konecny_nazev_baliku = f"{zbozi_nazev.strip()}{oznaceni}"
                
                # Zde bereme ty unikátní hodnoty pro daný balík (i-tý prvek ze seznamu)
                nove_radky.append({
                    'ZBOZI_2': zbozi_2,
                    'ZBOZI_NAZEV': konecny_nazev_baliku,
                    'ZBOZI_HMOTNOST': hmotnosti[i],
                    'ZBOZI_DELKA': delky[i],
                    'ZBOZI_SIRKA': sirky[i],
                    'ZBOZI_VYSKA': vysky[i]
                })
            
            str_web.session_state.katalog = pd.concat([str_web.session_state.katalog, pd.DataFrame(nove_radky)], ignore_index=True)
            ulozit_katalog(str_web.session_state.katalog)
            
            if zbozi_2 in str_web.session_state.chybejici_fronta:
                str_web.session_state.chybejici_fronta.remove(zbozi_2)
            
            str_web.success(f"Úspěšně vloženo: {pocet_baliku} balík(ů) pro produkt '{zbozi_2}'.")
            str_web.rerun()

    str_web.divider()
    str_web.subheader("✏️ Úprava a mazání v katalogu")
    str_web.info("💡 Kliknutím do buňky změníte hodnotu. Zaškrtnutím políčka vlevo a stisknutím klávesy 'Delete' (nebo ikony koše) řádek smažete.")
    
    # Místo str_web.dataframe použijeme str_web.data_editor
    # num_rows="dynamic" aktivuje možnost řádky mazat i přidávat prázdné řádky na konec
    upravena_data = str_web.data_editor(
        str_web.session_state.katalog, 
        use_container_width=True,
        num_rows="dynamic",
        key="katalog_editor"
    )
    
    # Pokud se data v editoru liší od naší paměti, nabídneme uložení
    if str_web.button("💾 Definitivně uložit změny a mazání", type="primary"):
        # Přepíšeme katalog v paměti webu novými daty z editoru
        str_web.session_state.katalog = upravena_data
        
        # Uložíme je fyzicky do souboru products.xlsx na disk
        ulozit_katalog(upravena_data)
        
        str_web.success("✨ Všechny změny a smazané řádky byly úspěšně uloženy do souboru products.xlsx!")
        # Znovu načteme stránku, aby se editor aktualizoval
        str_web.rerun()