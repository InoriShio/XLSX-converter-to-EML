#!/usr/bin/env python3

import json
import re
import os
import sys
from datetime import datetime
from pathlib import Path

# Check for openpyxl
try:
    from openpyxl import load_workbook
except ImportError:
    print("Error: openpyxl is not installed.")
    print("Please install it by running:")
    print("    pip install openpyxl")
    sys.exit(1)

# Check for tkinter
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
except ImportError:
    print("Error: tkinter is not available.")
    print("Please ensure Python was installed with tkinter support.")
    sys.exit(1)

def load_config(config_path: str) -> dict:
    """Load configuration from JSON file."""
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"Error: Configuration file not found: {config_path}")
        print("Please create a config.json file with the following structure:")
        print(json.dumps({
            "gemeente_naam": "YourMunicipality",
            "gemeente_code": "0000",
            "election_date": "2026-03-18",
            "election_subcategory": "GR2",
            "number_of_seats": 29,
            "preference_threshold": 25,
            "nomination_date": "2026-02-02"
        }, indent=2))
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON in configuration file: {e}")
        sys.exit(1)


def find_column_index(headers: list, search_terms: list) -> int:
    """Find column index by searching for terms in headers (case-insensitive)."""
    for idx, header in enumerate(headers):
        if header:
            header_lower = str(header).lower()
            for term in search_terms:
                if term.lower() in header_lower:
                    return idx
    return -1


def parse_address(address: str) -> tuple:
    if not address:
        return ("", "")
    
    address = str(address).strip()
    
    # Regex pattern: match everything up to the last space followed by a number (with optional suffix)
    pattern = r'^(.*?)\s+(\d+.*)$'
    match = re.match(pattern, address)
    
    if match:
        street = match.group(1).strip()
        number = match.group(2).strip()
        return (street, number)
    
    # If no match, return the whole address as street name
    return (address, "")


def read_excel_data(file_path: str) -> list:
    wb = load_workbook(filename=file_path, read_only=True, data_only=True)
    ws = wb.active
    
    # Get headers from first row
    headers = [cell.value for cell in ws[1]]
    
    # Find column indices
    # Note: 'naam stembureau' must come before 'naam' to match the correct column
    # when headers are like "Nummer stembureau" and "Naam stembureau"
    nummer_col = find_column_index(headers, ['nummer', 'nr', 'number', 'id'])
    naam_col = find_column_index(headers, ['naam stembureau', 'naam', 'name'])
    straat_col = find_column_index(headers, ['straat', 'street', 'adres', 'address', 'huisnummer'])
    
    if nummer_col == -1:
        raise ValueError("Could not find 'nummer' column in Excel file. "
                        "Please ensure there is a column with 'nummer' in the header.")
    if naam_col == -1:
        raise ValueError("Could not find 'naam' column in Excel file. "
                        "Please ensure there is a column with 'naam' in the header.")
    if straat_col == -1:
        raise ValueError("Could not find 'straat' column in Excel file. "
                        "Please ensure there is a column with 'straat' or 'adres' in the header.")
    
    print(f"Found columns: nummer={headers[nummer_col]} (col {nummer_col}), naam={headers[naam_col]} (col {naam_col}), straat={headers[straat_col]} (col {straat_col})")
    
    # Read data rows
    polling_stations = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if len(row) > max(nummer_col, naam_col, straat_col):
            nummer = row[nummer_col]
            naam = row[naam_col]
            straat = row[straat_col]
            
            # Debug: print first row to check column mapping
            if row_idx == 2:
                print(f"DEBUG Row 2: nummer={nummer}, naam={naam}, straat={straat}")
                print(f"DEBUG Full row: {row}")
            
            # Skip empty rows
            if nummer is None and naam is None:
                continue
            
            # Convert nummer to int if possible
            try:
                nummer = int(nummer)
            except (ValueError, TypeError):
                print(f"Warning: Row {row_idx} has invalid nummer '{nummer}', skipping.")
                continue
            
            polling_stations.append({
                'nummer': nummer,
                'naam': str(naam) if naam else f"Stembureau {nummer}",
                'straat': str(straat) if straat else ""
            })
    
    wb.close()
    return polling_stations

def generate_eml_xml(config: dict, polling_stations: list) -> str:
    """Generate EML_NL 110b XML content (polling stations list)."""
    
    gemeente_naam = config['gemeente_naam']
    gemeente_code = config['gemeente_code']
    election_date = config['election_date']
    election_subcategory = config.get('election_subcategory', 'GR2')
    nomination_date = config.get('nomination_date', '')
    
    # Generate transaction ID
    transaction_id = f"GR{election_date[:4]}{gemeente_code}"
    
    # Current datetime for CreationDateTime
    creation_datetime = datetime.now().strftime('%Y-%m-%dT%H:%M:%S')
    
    # If no nomination date provided, use election date as fallback
    if not nomination_date:
        nomination_date = election_date
    
    # Start building XML - 110b format for polling stations
    xml_parts = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<EML xmlns="urn:oasis:names:tc:evs:schema:eml"',
        '     xmlns:kr="http://www.kiesraad.nl/extensions"',
        '     xmlns:xal="urn:oasis:names:tc:ciq:xsdschema:xAL:2.0"',
        '     xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"',
        '     Id="110b" SchemaVersion="5">',
        f'  <TransactionId>{transaction_id}</TransactionId>',
        '  <ManagingAuthority>',
        f'    <AuthorityIdentifier Id="{gemeente_code}">{gemeente_naam}</AuthorityIdentifier>',
        '    <AuthorityAddress/>',
        '  </ManagingAuthority>',
        f'  <kr:CreationDateTime>{creation_datetime}</kr:CreationDateTime>',
        '  <ElectionEvent>',
        '    <EventIdentifier/>',
        '    <Election>',
        f'      <ElectionIdentifier Id="{transaction_id}">',
        f'        <ElectionName>Gemeenteraadsverkiezingen {gemeente_naam}</ElectionName>',
        '        <ElectionCategory>GR</ElectionCategory>',
        f'        <kr:ElectionSubcategory>{election_subcategory}</kr:ElectionSubcategory>',
        f'        <kr:ElectionDomain Id="{gemeente_code}">{gemeente_naam}</kr:ElectionDomain>',
        f'        <kr:ElectionDate>{election_date}</kr:ElectionDate>',
        f'        <kr:NominationDate>{nomination_date}</kr:NominationDate>',
        '      </ElectionIdentifier>',
        '      <Contest>',
        '        <ContestIdentifier Id="geen"/>',
        '        <ReportingUnit>',
        f'          <ReportingUnitIdentifier Id="{gemeente_code}">{gemeente_naam}</ReportingUnitIdentifier>',
        '        </ReportingUnit>',
        '        <VotingMethod>SPV</VotingMethod>',
        '        <MaxVotes>0</MaxVotes>',
    ]
    
    # Add polling stations as PollingPlace elements inside Contest
    for ps in polling_stations:
        nummer = ps['nummer']  # This is the position/stemgebiedsnr
        naam = escape_xml(ps['naam'])  # This is the name of the stembureau
        
        xml_parts.extend([
            '        <PollingPlace Channel="polling">',
            '          <PhysicalLocation>',
            '            <Address>',
            '              <Locality>',
            f'                <xal:LocalityName>{naam}</xal:LocalityName>',
            '              </Locality>',
            '            </Address>',
            f'            <PollingStation Id="{nummer}">0</PollingStation>',
            '          </PhysicalLocation>',
            '        </PollingPlace>',
        ])
    
    # Close XML
    xml_parts.extend([
        '      </Contest>',
        '    </Election>',
        '  </ElectionEvent>',
        '</EML>',
    ])
    
    return '\n'.join(xml_parts)


def escape_xml(text: str) -> str:
    """Escape special XML characters."""
    if not text:
        return ""
    text = str(text)
    text = text.replace('&', '&amp;')
    text = text.replace('<', '&lt;')
    text = text.replace('>', '&gt;')
    text = text.replace('"', '&quot;')
    text = text.replace("'", '&apos;')
    return text


def main():
    """Main entry point."""
    print("=" * 60)
    print("Excel to EML_NL Converter")
    print("Converts polling station data to EML_NL 110b XML format")
    print("=" * 60)
    print()
    
    # Determine config path (same directory as script)
    script_dir = Path(__file__).parent
    config_path = script_dir / 'config.json'
    
    # Load configuration
    print(f"Loading configuration from: {config_path}")
    config = load_config(str(config_path))
    print(f"  Gemeente: {config['gemeente_naam']} (code: {config['gemeente_code']})")
    print(f"  Election date: {config['election_date']}")
    print(f"  Election type: {config.get('election_subcategory', 'GR2')}")
    print()
    
    # Create hidden Tk root window
    root = tk.Tk()
    root.withdraw()
    
    # Show file open dialog
    print("Opening file dialog to select Excel file...")
    input_file = filedialog.askopenfilename(
        title="Select Excel file with polling station data",
        filetypes=[
            ("Excel files", "*.xlsx"),
            ("All files", "*.*")
        ]
    )
    
    if not input_file:
        print("No file selected. Exiting.")
        sys.exit(0)
    
    print(f"Selected file: {input_file}")
    print()
    
    # Read Excel data
    print("Reading Excel file...")
    try:
        polling_stations = read_excel_data(input_file)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel file:\n{e}")
        print(f"Error reading Excel file: {e}")
        sys.exit(1)
    
    print(f"Found {len(polling_stations)} polling stations")
    print()
    
    # Show polling stations
    print("Polling stations found:")
    for ps in polling_stations[:5]:  # Show first 5
        print(f"  {ps['nummer']:3d}: {ps['naam'][:40]:<40} | {ps['straat']}")
    if len(polling_stations) > 5:
        print(f"  ... and {len(polling_stations) - 5} more")
    print()
    
    # Generate XML
    print("Generating EML_NL XML...")
    xml_content = generate_eml_xml(config, polling_stations)
    
    # Suggest output filename
    gemeente_clean = config['gemeente_naam'].lower().replace(' ', '_')
    default_filename = f"stembureaus_{gemeente_clean}.eml.xml"
    default_dir = Path(input_file).parent
    
    # Show save dialog
    print("Opening save dialog...")
    output_file = filedialog.asksaveasfilename(
        title="Save EML.xml file as",
        initialdir=str(default_dir),
        initialfile=default_filename,
        defaultextension=".eml.xml",
        filetypes=[
            ("EML XML files", "*.eml.xml"),
            ("XML files", "*.xml"),
            ("All files", "*.*")
        ]
    )
    
    if not output_file:
        print("No output file selected. Exiting.")
        sys.exit(0)
    
    # Write output file
    print(f"Writing output to: {output_file}")
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(xml_content)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to write output file:\n{e}")
        print(f"Error writing output file: {e}")
        sys.exit(1)
    
    print()
    print("=" * 60)
    print("SUCCESS!")
    print(f"Created: {output_file}")
    print(f"Polling stations: {len(polling_stations)}")
    print("=" * 60)
    
    # Show success message
    messagebox.showinfo(
        "Success",
        f"EML.xml file created successfully!\n\n"
        f"Output: {output_file}\n"
        f"Polling stations: {len(polling_stations)}"
    )
    
    root.destroy()


if __name__ == "__main__":
    main()
