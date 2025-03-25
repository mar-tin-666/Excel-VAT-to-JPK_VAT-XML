"""Unit tests for Excel to JPK VAT XML converter."""

import tempfile
from pathlib import Path
import xml.etree.ElementTree as ET
from unittest import mock

import pandas as pd
import pytest

from main import normalize_columns, find_column, sanitize_filename, generate_xml, main

def test_normalize_columns():
    """Test stripping and trimming column names."""
    cols = [" VAT Base,", " Document Date ", "Type"]
    expected = ["VAT Base", "Document Date", "Type"]
    assert normalize_columns(cols) == expected

def test_find_column_exact():
    """Test exact column name matching."""
    cols = ["Type", "Document Type", "VAT Amount"]
    assert find_column(cols, "Type") == "Type"

def test_find_column_partial():
    """Test partial column name matching."""
    cols = ["VAT Registration Number", "External Doc"]
    assert find_column(cols, "VAT Registration") == "VAT Registration Number"

def test_find_column_not_found():
    """Test error raised when column not found."""
    cols = ["Col1", "Col2"]
    with pytest.raises(KeyError):
        find_column(cols, "Type")

def test_sanitize_filename():
    """Test sanitizing a filename with forbidden characters."""
    assert sanitize_filename("Test/Doc:2024|Name.xlsx") == "Test_Doc_2024_Name.xlsx"

def test_generate_xml_creates_valid_file():
    """Test XML generation from simple one-row DataFrame."""
    data = {
        "VAT Registration No": ["1234567890"],
        "External Document No": ["FV/01/2024"],
        "Document Date": ["01/03/2024"],
        "Document Receipt/Sales Date": ["02/03/2024"],
        "VAT Base": [1000.0],
        "VAT Amount": [230.0],
    }
    ns_map = {"ns": "http://crd.gov.pl/wzor/2021/12/27/11148/"}
    df = pd.DataFrame(data)

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "test_output.xml"
        generate_xml(
            df=df,
            output_file=output_path,
            col_vat_reg="VAT Registration No",
            col_ext_doc="External Document No",
            col_doc_date="Document Date",
            col_rec_date="Document Receipt/Sales Date",
            col_base="VAT Base",
            col_amount="VAT Amount",
        )

        assert output_path.exists()

        tree = ET.parse(output_path)
        root = tree.getroot()

        # Check key elements
        assert root.find("ns:Naglowek", ns_map) is not None
        assert root.find("ns:Deklaracja", ns_map) is not None
        assert root.find("ns:Ewidencja/ns:ZakupWiersz/ns:LpZakupu", ns_map).text == "1"
        assert root.find("ns:Ewidencja/ns:ZakupWiersz/ns:K_42", ns_map).text == "1000.00"
        assert root.find("ns:Ewidencja/ns:ZakupWiersz/ns:K_43", ns_map).text == "230.00"
        assert root.find("ns:Ewidencja/ns:ZakupWiersz/ns:DataZakupu", ns_map).text == "2024-03-01"
        assert root.find("ns:Ewidencja/ns:ZakupWiersz/ns:DataWplywu", ns_map).text == "2024-03-02"
        assert root.find("ns:Ewidencja/ns:ZakupWiersz/ns:NrDostawcy", ns_map).text == "1234567890"
        assert root.find("ns:Ewidencja/ns:ZakupCtrl/ns:LiczbaWierszyZakupow", ns_map).text == "1"
        assert root.find("ns:Ewidencja/ns:ZakupCtrl/ns:PodatekNaliczony", ns_map).text == "230.00"
        assert root.find("ns:Ewidencja/ns:SprzedazCtrl/ns:LiczbaWierszySprzedazy", ns_map).text == "0"
        assert root.find("ns:Ewidencja/ns:SprzedazCtrl/ns:PodatekNalezny", ns_map).text == "0"

def test_generate_xml_empty_df():
    """Test XML generation with empty DataFrame should produce valid but empty ZakupCtrl."""
    df = pd.DataFrame(
        columns=[
            "VAT Registration No",
            "External Document No",
            "Document Date",
            "Document Receipt/Sales Date",
            "VAT Base",
            "VAT Amount",
        ]
    )
    ns_map = {"ns": "http://crd.gov.pl/wzor/2021/12/27/11148/"}

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "empty_test.xml"
        generate_xml(
            df=df,
            output_file=output_path,
            col_vat_reg="VAT Registration No",
            col_ext_doc="External Document No",
            col_doc_date="Document Date",
            col_rec_date="Document Receipt/Sales Date",
            col_base="VAT Base",
            col_amount="VAT Amount",
        )

        tree = ET.parse(output_path)
        root = tree.getroot()
        assert root.find("ns:Ewidencja/ns:ZakupCtrl/ns:LiczbaWierszyZakupow", ns_map).text == "0"
        assert root.find("ns:Ewidencja/ns:ZakupCtrl/ns:PodatekNaliczony", ns_map).text == "0.00"

def test_main_with_missing_file(monkeypatch):
    """Test main exits when provided file path does not exist."""
    with pytest.raises(SystemExit):
        monkeypatch.setattr("builtins.input", lambda _: "nonexistent.xlsx")
        main()

def test_main_skips_sheet_without_required_columns(monkeypatch, tmp_path):
    """Test that main skips sheets without expected columns."""
    # Prepare a dummy Excel file with one sheet and wrong column names
    df = pd.DataFrame({"Wrong Column": [1, 2, 3]})
    excel_path = tmp_path / "test_file.xlsx"
    df.to_excel(excel_path, sheet_name="Arkusz1", index=False)

    monkeypatch.setattr("builtins.input", lambda _: str(excel_path))

    # Should skip the sheet, and not crash
    main()

    generated_files = list(tmp_path.glob("*.xml"))
    assert not generated_files  # No XML should be generated
