"""
Easy script for converting Excel VAT purchase data into separate XML files (JPK_VAT) per worksheet.
"""

import sys
import re
from pathlib import Path
from datetime import datetime
import xml.etree.ElementTree as ET
import pandas


def normalize_columns(columns):
    """Strip and clean column headers from trailing commas."""
    return [col.strip().rstrip(",") for col in columns]


def find_column(columns, pattern):
    """Find exact or partial match for a column name."""
    matches = [col for col in columns if col.strip().lower() == pattern.lower()]
    if matches:
        return matches[0]
    for col in columns:
        if pattern.lower() in col.lower():
            return col
    raise KeyError(f"Nie znaleziono kolumny zawierającej: '{pattern}'")


def sanitize_filename(name):
    """Sanitize filename by replacing forbidden characters."""
    return re.sub(r'[\\/*?:"<>|]', "_", name)


def generate_xml(   # pylint: disable=too-many-arguments, too-many-locals, too-many-positional-arguments
    df,
    output_file,
    col_vat_reg,
    col_ext_doc,
    col_doc_date,
    col_rec_date,
    col_base,
    col_amount,
):
    """Generate XML file from a filtered purchase DataFrame."""
    namespace = "http://crd.gov.pl/wzor/2021/12/27/11148/"
    ET.register_namespace("", namespace)
    jpk = ET.Element("JPK", xmlns=namespace)

    naglowek = ET.SubElement(jpk, "Naglowek")
    ET.SubElement(
        naglowek, "KodFormularza", kodSystemowy="JPK_V7M (2)", wersjaSchemy="1-0E"
    ).text = "JPK_VAT"
    ET.SubElement(naglowek, "WariantFormularza").text = "2"
    ET.SubElement(naglowek, "DataWytworzeniaJPK").text = datetime.now().isoformat()
    ET.SubElement(naglowek, "NazwaSystemu").text = "Python JPK Generator"
    ET.SubElement(naglowek, "CelZlozenia", poz="P_7").text = "1"
    ET.SubElement(naglowek, "KodUrzedu").text = "0000"
    ET.SubElement(naglowek, "Rok").text = str(datetime.now().year)
    ET.SubElement(naglowek, "Miesiac").text = str(datetime.now().month)

    podmiot = ET.SubElement(jpk, "Podmiot1", rola="Podatnik")
    osoba = ET.SubElement(podmiot, "OsobaFizyczna")
    ET.SubElement(
        osoba,
        "{http://crd.gov.pl/xml/schematy/dziedzinowe/mf/2021/06/08/eD/DefinicjeTypy/}NIP",
    ).text = "0000000000"
    ET.SubElement(
        osoba,
        "{http://crd.gov.pl/xml/schematy/dziedzinowe/mf/2021/06/08/eD/DefinicjeTypy/}ImiePierwsze",
    ).text = "Imie"
    ET.SubElement(
        osoba,
        "{http://crd.gov.pl/xml/schematy/dziedzinowe/mf/2021/06/08/eD/DefinicjeTypy/}Nazwisko",
    ).text = "Nazwisko"
    ET.SubElement(
        osoba,
        "{http://crd.gov.pl/xml/schematy/dziedzinowe/mf/2021/06/08/eD/DefinicjeTypy/}DataUrodzenia",
    ).text = "1990-01-01"

    deklaracja = ET.SubElement(jpk, "Deklaracja")
    naglowek_d = ET.SubElement(deklaracja, "Naglowek")
    ET.SubElement(
        naglowek_d,
        "KodFormularzaDekl",
        kodSystemowy="VAT-7 (22)",
        kodPodatku="VAT",
        rodzajZobowiazania="Z",
        wersjaSchemy="1-0E",
    ).text = "VAT-7"
    ET.SubElement(naglowek_d, "WariantFormularzaDekl").text = "22"
    ET.SubElement(deklaracja, "Pouczenia").text = "1"

    ewidencja = ET.SubElement(jpk, "Ewidencja")
    sprzedaz_ctrl = ET.SubElement(ewidencja, "SprzedazCtrl")
    ET.SubElement(sprzedaz_ctrl, "LiczbaWierszySprzedazy").text = "0"
    ET.SubElement(sprzedaz_ctrl, "PodatekNalezny").text = "0"

    suma_vat = 0.0
    for i, (_, row) in enumerate(df.iterrows(), start=1):
        zakup = ET.SubElement(ewidencja, "ZakupWiersz")
        ET.SubElement(zakup, "LpZakupu").text = str(i)
        ET.SubElement(zakup, "NrDostawcy").text = str(row[col_vat_reg])
        ET.SubElement(zakup, "NazwaDostawcy").text = ""
        ET.SubElement(zakup, "DowodZakupu").text = str(row[col_ext_doc])
        ET.SubElement(zakup, "DataZakupu").text = str(
            pandas.to_datetime(row[col_doc_date], dayfirst=True).date()
        )
        ET.SubElement(zakup, "DataWplywu").text = str(
            pandas.to_datetime(row[col_rec_date], dayfirst=True).date()
        )

        netto = float(row[col_base])
        vat = float(row[col_amount])
        ET.SubElement(zakup, "K_42").text = f"{netto:.2f}"
        ET.SubElement(zakup, "K_43").text = f"{vat:.2f}"
        suma_vat += vat

    zakup_ctrl = ET.SubElement(ewidencja, "ZakupCtrl")
    ET.SubElement(zakup_ctrl, "LiczbaWierszyZakupow").text = str(len(df))
    ET.SubElement(zakup_ctrl, "PodatekNaliczony").text = f"{suma_vat:.2f}"

    ET.ElementTree(jpk).write(output_file, encoding="utf-8", xml_declaration=True)
    print(f"✅ Zapisano: {output_file}")


def main(): # pylint: disable=too-many-locals
    """Main entrypoint for Excel-to-XML JPK generator."""
    xlsx_file = Path(
        input("Podaj ścieżkę do pliku Excel (*.xlsx): ").strip().strip('"').strip("'")
    )

    if not xlsx_file.exists():
        print(f"Błąd: Plik nie istnieje: {xlsx_file}")
        sys.exit(1)

    all_sheets = pandas.read_excel(xlsx_file, sheet_name=None)

    for sheet_name, df in all_sheets.items():
        print(f"\n🔄 Przetwarzanie zakładki: {sheet_name}")
        df.columns = normalize_columns(df.columns)
        columns = df.columns.tolist()

        try:
            col_type = find_column(columns, "Type")
            col_vat_reg = find_column(columns, "VAT Registration")
            col_ext_doc = find_column(columns, "External Document")
            col_doc_date = find_column(columns, "Document Date")
            col_rec_date = find_column(columns, "Receipt")
            col_base = find_column(columns, "VAT Base")
            col_amount = find_column(columns, "VAT Amount")
        except KeyError as e:
            print(f"⚠️ Pomijam zakładkę '{sheet_name}': {e}")
            continue

        df["__type_clean"] = df[col_type].astype(str).str.strip().str.lower()
        purchase_df = df[df["__type_clean"] == "purchase"]

        if purchase_df.empty:
            print(
                f"ℹ️ Zakładka '{sheet_name}' nie zawiera danych typu 'Purchase'. Pomijam."
            )
            continue

        purchase_df[col_base] = pandas.to_numeric(
            purchase_df[col_base], errors="coerce"
        ).fillna(0)
        purchase_df[col_amount] = pandas.to_numeric(
            purchase_df[col_amount], errors="coerce"
        ).fillna(0)

        base_name = xlsx_file.stem
        safe_sheet = sanitize_filename(sheet_name)
        output_file = xlsx_file.with_name(f"{base_name}__{safe_sheet}.xml")

        generate_xml(
            df=purchase_df,
            output_file=output_file,
            col_vat_reg=col_vat_reg,
            col_ext_doc=col_ext_doc,
            col_doc_date=col_doc_date,
            col_rec_date=col_rec_date,
            col_base=col_base,
            col_amount=col_amount,
        )


if __name__ == "__main__":
    main()
