import streamlit as st
import pandas as pd
import numpy as np
import re
import tempfile
from datetime import datetime, date
from dateutil.parser import parse as dt_parse
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from copy import copy as pycopy

# =========================
# Page config
# =========================
st.set_page_config(page_title="Bank Reconciliation", layout="wide")

# =========================
# Theme (Wells Fargo style)
# =========================
WF_RED = "#D71E28"
WF_BG = "#F3F3F3"
WF_CARD = "#FFFFFF"
WF_TEXT = "#111111"
WF_MUTED = "#666666"
WF_BORDER = "#E6E6E6"

# =========================
# Streamlit compatibility helpers
# =========================
def safe_primary_button(label, use_container_width=False, key=None):
    try:
        return st.button(label, type="primary", use_container_width=use_container_width, key=key)
    except TypeError:
        return st.button(label, use_container_width=use_container_width, key=key)

def safe_download_button(label, data, file_name, mime, use_container_width=False, key=None):
    try:
        return st.download_button(
            label,
            data=data,
            file_name=file_name,
            mime=mime,
            use_container_width=use_container_width,
            key=key,
        )
    except TypeError:
        return st.download_button(
            label,
            data=data,
            file_name=file_name,
            mime=mime,
            key=key,
        )

def status_box(label):
    try:
        return st.status(label, expanded=True)
    except Exception:
        st.info(label)
        return None

# =========================
# CSS
# =========================
st.markdown(
    f"""
<style>
html, body, [class*="css"] {{
    background-color: {WF_BG};
    color: {WF_TEXT};
}}
.block-container {{
    padding-top: 4.1rem;
    padding-bottom: 2rem;
    max-width: 1400px;
}}

.topbar {{
    background: {WF_RED};
    color: white;
    padding: 16px 18px;
    border-radius: 14px;
    margin-bottom: 16px;
    position: relative;
}}
.brand-center {{
    text-align: center;
    font-size: 30px;
    font-weight: 900;
    letter-spacing: 0.3px;
}}
.sub-center {{
    text-align: center;
    font-size: 12px;
    opacity: 0.98;
    margin-top: 2px;
}}
.topbar-right {{
    position: absolute;
    right: 18px;
    top: 18px;
    font-size: 12px;
    opacity: 0.98;
}}

.card {{
    background: {WF_CARD};
    border: 1px solid {WF_BORDER};
    border-radius: 16px;
    padding: 16px 16px;
}}
.hero {{
    background: {WF_CARD};
    border: 1px solid {WF_BORDER};
    border-radius: 16px;
    padding: 18px 18px;
    margin-bottom: 12px;
    text-align: center;
}}
.hero h2 {{
    margin: 0;
    font-size: 28px;
    font-weight: 900;
}}
.hero p {{
    margin: 8px 0 0 0;
    color: {WF_MUTED};
}}

hr {{
    border: none;
    border-top: 2px solid {WF_BORDER};
    margin: 20px 0;
}}

div.stButton > button[kind="primary"] {{
    background: {WF_RED} !important;
    color: white !important;
    border-radius: 12px !important;
}}
</style>
""",
    unsafe_allow_html=True,
)

# =========================
# Header
# =========================
st.markdown(
    f"""
<div class="topbar">
  <div class="brand-center">Bank Reconciliation</div>
  <div class="sub-center">Match. Reconcile. Download.</div>
  <div class="topbar-right">v1.0</div>
</div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<div class="hero">
  <h2>Bank Reconciliation Tool</h2>
  <p>Upload your Bank Statement and Cashbook/Ledger. The app will match transactions by amount and date.</p>
</div>
""",
    unsafe_allow_html=True,
)

# =========================
# Helper functions
# =========================
def to_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def to_num(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = to_str(x).replace(",", "")
    # Handle negative numbers in parentheses
    if '(' in s and ')' in s:
        s = '-' + s.replace('(', '').replace(')', '').strip()
    s = re.sub(r"[^\d\.\-]", "", s)
    if not s or s == '-':
        return np.nan
    try:
        return float(s)
    except Exception:
        return np.nan

def to_date(x):
    if isinstance(x, (pd.Timestamp, np.datetime64)):
        return pd.to_datetime(x, errors='coerce')
    s = to_str(x)
    if not s:
        return pd.NaT
    try:
        # Try common date formats
        for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y', '%Y/%m/%d', '%d/%m/%y', '%d-%m-%y']:
            try:
                return pd.to_datetime(s, format=fmt)
            except:
                continue
        return pd.to_datetime(dt_parse(s, fuzzy=True))
    except Exception:
        return pd.NaT

def format_date(date_val):
    """Safely format a date value to string - handles NaT, None, and empty values"""
    if date_val is None:
        return ''
    if pd.isna(date_val):
        return ''
    if isinstance(date_val, str):
        # If it's already a string, try to parse it
        if date_val == '':
            return ''
        try:
            # Try to convert string to datetime
            parsed = pd.to_datetime(date_val, errors='coerce')
            if pd.notna(parsed):
                return parsed.strftime('%d/%m/%y')
            else:
                # If can't parse, return the original string truncated
                return date_val[:10] if len(date_val) > 10 else date_val
        except:
            return date_val[:10] if len(date_val) > 10 else date_val
    if isinstance(date_val, (pd.Timestamp, datetime, date)):
        # Check if it's NaT
        if pd.isna(date_val):
            return ''
        try:
            return date_val.strftime('%d/%m/%y')
        except:
            return ''
    return str(date_val)

def detect_table(df_raw, max_rows=100):
    """Detect where the actual table starts in an Excel sheet"""
    if df_raw is None or df_raw.empty:
        return None, 0
    
    best_row = 0
    best_score = 0
    
    for r in range(min(max_rows, len(df_raw))):
        row = df_raw.iloc[r]
        non_empty = row.notna().sum()
        if non_empty < 2:
            continue
        
        # Check for header-like keywords
        header_score = 0
        for cell in row:
            if pd.isna(cell):
                continue
            s = str(cell).lower()
            if any(kw in s for kw in ['date', 'amount', 'credit', 'debit', 'ref', 'description', 'posting', 'document', 'narration']):
                header_score += 2
        
        score = non_empty + header_score
        if score > best_score:
            best_score = score
            best_row = r
    
    if best_row >= 0 and best_row < len(df_raw):
        # Use the detected row as header
        headers = df_raw.iloc[best_row].fillna('').astype(str).tolist()
        # Clean up headers - remove empty strings and duplicates
        headers = [h.strip() if h.strip() else f"Column_{i}" for i, h in enumerate(headers)]
        data = df_raw.iloc[best_row + 1:].reset_index(drop=True)
        # Only keep columns that have headers
        data = data.iloc[:, :len(headers)]
        data.columns = headers
        # Remove rows that are completely empty
        data = data.dropna(how='all')
        return data, best_row
    
    return df_raw, 0

# =========================
# Load and normalize files
# =========================
def load_bank_statement(file):
    """Load bank statement and normalize columns"""
    try:
        # Try to read all sheets and find the first valid one
        excel_file = pd.ExcelFile(file)
        sheet_names = excel_file.sheet_names
        
        # Filter out hidden sheets (sheets starting with _ or containing 'hidden' in name)
        visible_sheets = [s for s in sheet_names if not s.startswith('_') and 'hidden' not in s.lower()]
        
        if not visible_sheets:
            visible_sheets = sheet_names  # fallback to all sheets
        
        df_bank = None
        header_row_used = 0
        
        for sheet in visible_sheets:
            try:
                df_raw = pd.read_excel(file, sheet_name=sheet, header=None)
                if df_raw is not None and not df_raw.empty:
                    df, header_row = detect_table(df_raw)
                    if df is not None and len(df) > 3:  # At least a few rows of data
                        df_bank = df
                        header_row_used = header_row
                        break
            except Exception:
                continue
        
        if df_bank is None or df_bank.empty:
            # Fallback: try reading without header detection
            df_bank = pd.read_excel(file)
            if df_bank.empty:
                raise ValueError("Could not read bank statement file")
        
    except Exception as e:
        # Last resort: try simple read
        st.warning(f"Auto-detection failed, trying simple read: {str(e)}")
        df_bank = pd.read_excel(file)
    
    # Find relevant columns
    date_col = None
    credit_col = None
    debit_col = None
    ref_col = None
    desc_col = None
    balance_col = None
    
    for col in df_bank.columns:
        col_lower = str(col).lower()
        if 'date' in col_lower and not date_col:
            date_col = col
        if 'credit' in col_lower and not credit_col:
            credit_col = col
        if 'debit' in col_lower and not debit_col:
            debit_col = col
        if 'ref' in col_lower or 'reference' in col_lower:
            ref_col = col
        if 'narrative' in col_lower or 'description' in col_lower or 'particulars' in col_lower or 'details' in col_lower:
            desc_col = col
        if 'balance' in col_lower and not balance_col:
            balance_col = col
    
    # If no separate credit/debit, look for amount column
    if not credit_col and not debit_col:
        for col in df_bank.columns:
            if 'amount' in str(col).lower() or 'value' in str(col).lower():
                credit_col = col
                debit_col = col
                break
    
    # Extract opening balance
    opening_balance = None
    for idx, row in df_bank.iterrows():
        row_str = ' '.join(str(v).lower() for v in row.values if pd.notna(v))
        if 'balance at period start' in row_str or 'opening balance' in row_str or 'balance brought forward' in row_str:
            for col in df_bank.columns:
                val = to_num(row[col])
                if not pd.isna(val) and val != 0:
                    opening_balance = val
                    break
            break
    
    # Build normalized dataframe
    transactions = []
    for idx, row in df_bank.iterrows():
        credit = 0
        debit = 0
        
        if credit_col and credit_col in df_bank.columns:
            credit = to_num(row[credit_col]) if pd.notna(row[credit_col]) else 0
        if debit_col and debit_col in df_bank.columns:
            debit = to_num(row[debit_col]) if pd.notna(row[debit_col]) else 0
        
        # If amount column only
        if credit == 0 and debit == 0 and credit_col == debit_col and credit_col:
            amount = to_num(row[credit_col]) if pd.notna(row[credit_col]) else 0
            if amount == 0:
                continue
            credit = amount if amount > 0 else 0
            debit = -amount if amount < 0 else 0
        
        if credit == 0 and debit == 0:
            continue
        
        trans_date = to_date(row[date_col]) if date_col and date_col in df_bank.columns else pd.NaT
        if pd.isna(trans_date):
            continue
        
        # Get description
        desc = ''
        if desc_col and desc_col in df_bank.columns:
            desc = to_str(row[desc_col])
        if not desc and ref_col and ref_col in df_bank.columns:
            desc = to_str(row[ref_col])
        
        transactions.append({
            'date': trans_date,
            'reference': to_str(row[ref_col]) if ref_col and ref_col in df_bank.columns else '',
            'description': desc,
            'credit': credit,
            'debit': debit,
            'amount': credit - debit,
            'abs_amount': abs(credit - debit),
            'source': 'BANK'
        })
    
    df_bank_normalized = pd.DataFrame(transactions)
    
    if df_bank_normalized.empty:
        st.error("No valid transactions found in bank statement. Please check the file format.")
        st.stop()
    
    # Get closing balance from last row if balance column exists
    closing_balance = None
    if balance_col and balance_col in df_bank.columns:
        last_valid = df_bank[balance_col].dropna().iloc[-1] if not df_bank[balance_col].dropna().empty else None
        closing_balance = to_num(last_valid) if last_valid else None
    
    return df_bank_normalized, opening_balance, closing_balance, header_row_used

def load_ledger(file):
    """Load Yellowcob ledger and normalize columns"""
    try:
        # Try to read all sheets and find the first valid one
        excel_file = pd.ExcelFile(file)
        sheet_names = excel_file.sheet_names
        
        # Filter out hidden sheets
        visible_sheets = [s for s in sheet_names if not s.startswith('_') and 'hidden' not in s.lower()]
        
        if not visible_sheets:
            visible_sheets = sheet_names
        
        df_ledger = None
        header_row_used = 0
        
        for sheet in visible_sheets:
            try:
                df_raw = pd.read_excel(file, sheet_name=sheet, header=None)
                if df_raw is not None and not df_raw.empty:
                    df, header_row = detect_table(df_raw)
                    if df is not None and len(df) > 3:
                        df_ledger = df
                        header_row_used = header_row
                        break
            except Exception:
                continue
        
        if df_ledger is None or df_ledger.empty:
            df_ledger = pd.read_excel(file)
            if df_ledger.empty:
                raise ValueError("Could not read ledger file")
        
    except Exception as e:
        st.warning(f"Auto-detection failed, trying simple read: {str(e)}")
        df_ledger = pd.read_excel(file)
    
    # Find relevant columns
    date_col = None
    amount_col = None
    desc_col = None
    ref_col = None
    
    for col in df_ledger.columns:
        col_lower = str(col).lower()
        if 'date' in col_lower and not date_col:
            date_col = col
        if 'amount' in col_lower and not amount_col:
            amount_col = col
        if 'description' in col_lower or 'desc' in col_lower or 'particulars' in col_lower or 'details' in col_lower or 'narration' in col_lower:
            desc_col = col
        if 'document' in col_lower or 'ref' in col_lower or 'external' in col_lower or 'document no' in col_lower or 'reference' in col_lower:
            ref_col = col
    
    # Build normalized dataframe
    transactions = []
    for idx, row in df_ledger.iterrows():
        amount = to_num(row[amount_col]) if amount_col and amount_col in df_ledger.columns else np.nan
        if pd.isna(amount):
            continue
        
        trans_date = to_date(row[date_col]) if date_col and date_col in df_ledger.columns else pd.NaT
        if pd.isna(trans_date):
            continue
        
        # Positive amount = Credit (deposit), Negative amount = Debit (payment)
        transactions.append({
            'date': trans_date,
            'reference': to_str(row[ref_col]) if ref_col and ref_col in df_ledger.columns else '',
            'description': to_str(row[desc_col]) if desc_col and desc_col in df_ledger.columns else '',
            'amount': amount,
            'abs_amount': abs(amount),
            'type': 'CREDIT' if amount > 0 else 'DEBIT',
            'source': 'LEDGER'
        })
    
    df_ledger_normalized = pd.DataFrame(transactions)
    
    if df_ledger_normalized.empty:
        st.error("No valid transactions found in ledger. Please check the file format.")
        st.stop()
    
    # Calculate closing balance (sum of all amounts)
    closing_balance = df_ledger_normalized['amount'].sum() if not df_ledger_normalized.empty else 0
    
    return df_ledger_normalized, closing_balance, header_row_used

# =========================
# Matching logic
# =========================
def match_transactions(bank_df, ledger_df):
    """Match transactions by amount first, then date for ties"""
    
    # Make copies and add IDs
    bank_df_copy = bank_df.copy()
    ledger_df_copy = ledger_df.copy()
    
    # Separate credits and debits
    bank_credits = bank_df_copy[bank_df_copy['credit'] > 0].copy() if 'credit' in bank_df_copy.columns else bank_df_copy[bank_df_copy['amount'] > 0].copy()
    bank_debits = bank_df_copy[bank_df_copy['debit'] > 0].copy() if 'debit' in bank_df_copy.columns else bank_df_copy[bank_df_copy['amount'] < 0].copy()
    
    # For bank, if we have amount column only
    if 'credit' not in bank_df_copy.columns:
        bank_credits = bank_df_copy[bank_df_copy['amount'] > 0].copy()
        bank_debits = bank_df_copy[bank_df_copy['amount'] < 0].copy()
    
    ledger_credits = ledger_df_copy[ledger_df_copy['amount'] > 0].copy()
    ledger_debits = ledger_df_copy[ledger_df_copy['amount'] < 0].copy()
    
    # Add unique IDs
    bank_credits['id'] = [f'B_C_{i}' for i in range(len(bank_credits))]
    bank_debits['id'] = [f'B_D_{i}' for i in range(len(bank_debits))]
    ledger_credits['id'] = [f'L_C_{i}' for i in range(len(ledger_credits))]
    ledger_debits['id'] = [f'L_D_{i}' for i in range(len(ledger_debits))]
    
    matches = []
    
    # Match Credits
    matched_ledger_credit_ids = set()
    matched_bank_credit_ids = set()
    
    all_amounts = set()
    if not ledger_credits.empty:
        all_amounts.update(ledger_credits['abs_amount'].unique())
    if not bank_credits.empty:
        all_amounts.update(bank_credits['abs_amount'].unique())
    
    for amount in all_amounts:
        ledger_items = ledger_credits[ledger_credits['abs_amount'] == amount] if not ledger_credits.empty else pd.DataFrame()
        bank_items = bank_credits[bank_credits['abs_amount'] == amount] if not bank_credits.empty else pd.DataFrame()
        
        if len(ledger_items) == 1 and len(bank_items) == 1:
            # Perfect match by amount only
            matches.append({
                'ledger_id': ledger_items.iloc[0]['id'],
                'bank_id': bank_items.iloc[0]['id'],
                'amount': amount,
                'type': 'CREDIT',
                'match_method': 'amount_only',
                'ledger_date': ledger_items.iloc[0]['date'],
                'bank_date': bank_items.iloc[0]['date'],
                'date_match': ledger_items.iloc[0]['date'] == bank_items.iloc[0]['date'],
                'ledger_desc': ledger_items.iloc[0]['description'],
                'ledger_ref': ledger_items.iloc[0]['reference'],
                'bank_desc': bank_items.iloc[0]['description'],
                'bank_ref': bank_items.iloc[0]['reference']
            })
            matched_ledger_credit_ids.add(ledger_items.iloc[0]['id'])
            matched_bank_credit_ids.add(bank_items.iloc[0]['id'])
        
        elif len(ledger_items) > 0 and len(bank_items) > 0:
            # Multiple same amounts - need date matching
            ledger_list = ledger_items.to_dict('records')
            bank_list = bank_items.to_dict('records')
            
            # Try to match by date
            for l_item in ledger_list:
                if l_item['id'] in matched_ledger_credit_ids:
                    continue
                
                # Find bank item with same date
                matched_bank = None
                for b_item in bank_list:
                    if b_item['id'] in matched_bank_credit_ids:
                        continue
                    if l_item['date'].date() == b_item['date'].date():
                        matched_bank = b_item
                        break
                
                if matched_bank:
                    matches.append({
                        'ledger_id': l_item['id'],
                        'bank_id': matched_bank['id'],
                        'amount': amount,
                        'type': 'CREDIT',
                        'match_method': 'amount_and_date',
                        'ledger_date': l_item['date'],
                        'bank_date': matched_bank['date'],
                        'date_match': True,
                        'ledger_desc': l_item['description'],
                        'ledger_ref': l_item['reference'],
                        'bank_desc': matched_bank['description'],
                        'bank_ref': matched_bank['reference']
                    })
                    matched_ledger_credit_ids.add(l_item['id'])
                    matched_bank_credit_ids.add(matched_bank['id'])
    
    # Unmatched credits
    unmatched_ledger_credits = []
    for _, item in ledger_credits.iterrows():
        if item['id'] not in matched_ledger_credit_ids:
            unmatched_ledger_credits.append(item.to_dict())
    
    unmatched_bank_credits = []
    for _, item in bank_credits.iterrows():
        if item['id'] not in matched_bank_credit_ids:
            unmatched_bank_credits.append(item.to_dict())
    
    # Match Debits (same logic)
    matched_ledger_debit_ids = set()
    matched_bank_debit_ids = set()
    
    all_debit_amounts = set()
    if not ledger_debits.empty:
        all_debit_amounts.update(ledger_debits['abs_amount'].unique())
    if not bank_debits.empty:
        all_debit_amounts.update(bank_debits['abs_amount'].unique())
    
    for amount in all_debit_amounts:
        ledger_items = ledger_debits[ledger_debits['abs_amount'] == amount] if not ledger_debits.empty else pd.DataFrame()
        bank_items = bank_debits[bank_debits['abs_amount'] == amount] if not bank_debits.empty else pd.DataFrame()
        
        if len(ledger_items) == 1 and len(bank_items) == 1:
            matches.append({
                'ledger_id': ledger_items.iloc[0]['id'],
                'bank_id': bank_items.iloc[0]['id'],
                'amount': amount,
                'type': 'DEBIT',
                'match_method': 'amount_only',
                'ledger_date': ledger_items.iloc[0]['date'],
                'bank_date': bank_items.iloc[0]['date'],
                'date_match': ledger_items.iloc[0]['date'] == bank_items.iloc[0]['date'],
                'ledger_desc': ledger_items.iloc[0]['description'],
                'ledger_ref': ledger_items.iloc[0]['reference'],
                'bank_desc': bank_items.iloc[0]['description'],
                'bank_ref': bank_items.iloc[0]['reference']
            })
            matched_ledger_debit_ids.add(ledger_items.iloc[0]['id'])
            matched_bank_debit_ids.add(bank_items.iloc[0]['id'])
        
        elif len(ledger_items) > 0 and len(bank_items) > 0:
            ledger_list = ledger_items.to_dict('records')
            bank_list = bank_items.to_dict('records')
            
            for l_item in ledger_list:
                if l_item['id'] in matched_ledger_debit_ids:
                    continue
                
                matched_bank = None
                for b_item in bank_list:
                    if b_item['id'] in matched_bank_debit_ids:
                        continue
                    if l_item['date'].date() == b_item['date'].date():
                        matched_bank = b_item
                        break
                
                if matched_bank:
                    matches.append({
                        'ledger_id': l_item['id'],
                        'bank_id': matched_bank['id'],
                        'amount': amount,
                        'type': 'DEBIT',
                        'match_method': 'amount_and_date',
                        'ledger_date': l_item['date'],
                        'bank_date': matched_bank['date'],
                        'date_match': True,
                        'ledger_desc': l_item['description'],
                        'ledger_ref': l_item['reference'],
                        'bank_desc': matched_bank['description'],
                        'bank_ref': matched_bank['reference']
                    })
                    matched_ledger_debit_ids.add(l_item['id'])
                    matched_bank_debit_ids.add(matched_bank['id'])
    
    # Unmatched debits
    unmatched_ledger_debits = []
    for _, item in ledger_debits.iterrows():
        if item['id'] not in matched_ledger_debit_ids:
            unmatched_ledger_debits.append(item.to_dict())
    
    unmatched_bank_debits = []
    for _, item in bank_debits.iterrows():
        if item['id'] not in matched_bank_debit_ids:
            unmatched_bank_debits.append(item.to_dict())
    
    return {
        'matches': matches,
        'unmatched_ledger_credits': unmatched_ledger_credits,
        'unmatched_ledger_debits': unmatched_ledger_debits,
        'unmatched_bank_credits': unmatched_bank_credits,
        'unmatched_bank_debits': unmatched_bank_debits
    }

# =========================
# Build working paper with new structure
# =========================
def build_working_paper(match_results):
    """
    Build the working paper with the exact column structure:
    SECTION | LEDGER_DATE | LEDGER_DESC | LEDGER_REF | LEDGER_AMOUNT | MATCH_STATUS | BANK_AMOUNT | BANK_DATE | BANK_REF | BANK_DESC
    """
    
    working_paper = []
    
    # 1. Add MATCHED items first
    for match in match_results['matches']:
        working_paper.append({
            'SECTION': 'MATCHED',
            'LEDGER_DATE': match['ledger_date'],
            'LEDGER_DESC': match.get('ledger_desc', ''),
            'LEDGER_REF': match.get('ledger_ref', ''),
            'LEDGER_AMOUNT': match['amount'] if match['type'] == 'CREDIT' else -match['amount'],
            'MATCH_STATUS': f"MATCHED - {match['match_method'].replace('_', ' ').upper()}",
            'BANK_AMOUNT': match['amount'] if match['type'] == 'CREDIT' else -match['amount'],
            'BANK_DATE': match['bank_date'],
            'BANK_REF': match.get('bank_ref', ''),
            'BANK_DESC': match.get('bank_desc', '')
        })
    
    # 2. Add LEDGER ONLY items (in ledger but not in bank)
    for item in match_results['unmatched_ledger_credits']:
        working_paper.append({
            'SECTION': 'UNMATCHED - LEDGER ONLY',
            'LEDGER_DATE': item['date'],
            'LEDGER_DESC': item['description'],
            'LEDGER_REF': item['reference'],
            'LEDGER_AMOUNT': item['amount'],
            'MATCH_STATUS': 'NO BANK MATCH',
            'BANK_AMOUNT': '',
            'BANK_DATE': '',
            'BANK_REF': '',
            'BANK_DESC': ''
        })
    
    for item in match_results['unmatched_ledger_debits']:
        working_paper.append({
            'SECTION': 'UNMATCHED - LEDGER ONLY',
            'LEDGER_DATE': item['date'],
            'LEDGER_DESC': item['description'],
            'LEDGER_REF': item['reference'],
            'LEDGER_AMOUNT': item['amount'],
            'MATCH_STATUS': 'NO BANK MATCH',
            'BANK_AMOUNT': '',
            'BANK_DATE': '',
            'BANK_REF': '',
            'BANK_DESC': ''
        })
    
    # 3. Add BANK ONLY items (in bank but not in ledger)
    for item in match_results['unmatched_bank_credits']:
        working_paper.append({
            'SECTION': 'UNMATCHED - BANK ONLY',
            'LEDGER_DATE': '',
            'LEDGER_DESC': '',
            'LEDGER_REF': '',
            'LEDGER_AMOUNT': '',
            'MATCH_STATUS': 'NO LEDGER MATCH',
            'BANK_AMOUNT': item['amount'],
            'BANK_DATE': item['date'],
            'BANK_REF': item['reference'],
            'BANK_DESC': item['description']
        })
    
    for item in match_results['unmatched_bank_debits']:
        working_paper.append({
            'SECTION': 'UNMATCHED - BANK ONLY',
            'LEDGER_DATE': '',
            'LEDGER_DESC': '',
            'LEDGER_REF': '',
            'LEDGER_AMOUNT': '',
            'MATCH_STATUS': 'NO LEDGER MATCH',
            'BANK_AMOUNT': item['amount'],
            'BANK_DATE': item['date'],
            'BANK_REF': item['reference'],
            'BANK_DESC': item['description']
        })
    
    return pd.DataFrame(working_paper)

# =========================
# Build final reconciliation statement
# =========================
def build_recon_statement(opening_balance, ledger_closing_balance, match_results):
    """Build final reconciliation statement"""
    
    recon_items = []
    total_adjustment = 0
    
    # Bank-only items (need to add/subtract)
    for item in match_results['unmatched_bank_credits']:
        recon_items.append({
            'date': item['date'],
            'description': f"BANK ONLY: {item['description']} ({item['reference']})",
            'amount': item['amount'],
            'adjustment': item['amount']
        })
        total_adjustment += item['amount']
    
    for item in match_results['unmatched_bank_debits']:
        recon_items.append({
            'date': item['date'],
            'description': f"BANK ONLY: {item['description']} ({item['reference']})",
            'amount': item['amount'],
            'adjustment': item['amount']
        })
        total_adjustment += item['amount']
    
    # Ledger-only items
    for item in match_results['unmatched_ledger_credits']:
        recon_items.append({
            'date': item['date'],
            'description': f"LEDGER ONLY - DEPOSIT IN TRANSIT: {item['description']} ({item['reference']})",
            'amount': item['amount'],
            'adjustment': item['amount']
        })
        total_adjustment += item['amount']
    
    for item in match_results['unmatched_ledger_debits']:
        recon_items.append({
            'date': item['date'],
            'description': f"LEDGER ONLY - UNPRESENTED CHEQUE: {item['description']} ({item['reference']})",
            'amount': item['amount'],
            'adjustment': item['amount']
        })
        total_adjustment += item['amount']
    
    if opening_balance is None or opening_balance == 0:
        adjusted_balance = total_adjustment
    else:
        adjusted_balance = opening_balance + total_adjustment
    
    difference = adjusted_balance - ledger_closing_balance if ledger_closing_balance else 0
    
    return {
        'opening_balance': opening_balance if opening_balance else 0,
        'recon_items': pd.DataFrame(recon_items),
        'total_adjustment': total_adjustment,
        'adjusted_balance': adjusted_balance,
        'ledger_balance': ledger_closing_balance if ledger_closing_balance else 0,
        'difference': difference
    }

# =========================
# Export to Excel with proper formatting
# =========================
def export_to_excel(working_paper_df, recon_statement, bank_df, ledger_df, match_results):
    """Export all results to Excel with proper formatting"""
    
    output_path = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx').name
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # ============================================================
        # Sheet 1: WORKING PAPER - New structure
        # ============================================================
        
        # Create a copy for display with formatted dates
        wp_display = working_paper_df.copy()
        
        # Format dates safely
        wp_display['LEDGER_DATE'] = wp_display['LEDGER_DATE'].apply(lambda x: format_date(x) if pd.notna(x) else '')
        wp_display['BANK_DATE'] = wp_display['BANK_DATE'].apply(lambda x: format_date(x) if pd.notna(x) else '')
        
        # Format amounts - handle empty strings and NaN
        wp_display['LEDGER_AMOUNT'] = wp_display['LEDGER_AMOUNT'].apply(
            lambda x: f"{float(x):,.2f}" if pd.notna(x) and x != '' and x != 0 else ('' if x == '' else '')
        )
        wp_display['BANK_AMOUNT'] = wp_display['BANK_AMOUNT'].apply(
            lambda x: f"{float(x):,.2f}" if pd.notna(x) and x != '' and x != 0 else ('' if x == '' else '')
        )
        
        # Write to Excel with headers
        wp_display.to_excel(writer, sheet_name='WORKING_PAPER', index=False)
        
        # ============================================================
        # Sheet 2: RECONCILIATION STATEMENT
        # ============================================================
        
        recon_data = []
        
        recon_data.append(['Bank Reconciliation Statement', ''])
        recon_data.append(['', ''])
        recon_data.append(['Statement Balance (from bank)', '', f"{recon_statement['opening_balance']:,.2f}"])
        recon_data.append(['', '', ''])
        recon_data.append(['RECONCILIATION ITEMS:', '', ''])
        recon_data.append(['Date', 'Description', 'Amount'])
        recon_data.append(['', '', ''])
        
        for _, item in recon_statement['recon_items'].iterrows():
            date_str = format_date(item['date']) if pd.notna(item['date']) else ''
            recon_data.append([date_str, item['description'][:80], f"{item['adjustment']:,.2f}"])
        
        recon_data.append(['', '', ''])
        recon_data.append(['', 'Total Reconciling Items', f"{recon_statement['total_adjustment']:,.2f}"])
        recon_data.append(['', '', ''])
        recon_data.append(['', 'Total Cash (Adjusted Balance)', f"{recon_statement['adjusted_balance']:,.2f}"])
        recon_data.append(['', '', ''])
        recon_data.append(['', 'Ledger Balance (from Yellowcob)', f"{recon_statement['ledger_balance']:,.2f}"])
        recon_data.append(['', '', ''])
        recon_data.append(['', 'Difference', f"{recon_statement['difference']:,.2f}"])
        recon_data.append(['', '', ''])
        recon_data.append(['', '', ''])
        recon_data.append(['PREPARED BY:', '', ''])
        recon_data.append(['Name:', '', ''])
        recon_data.append(['Signature:', '', ''])
        recon_data.append(['Date:', '', datetime.now().strftime('%d/%m/%y')])
        
        recon_df = pd.DataFrame(recon_data)
        recon_df.to_excel(writer, sheet_name='RECON_STATEMENT', index=False, header=False)
        
        # ============================================================
        # Sheet 3: MATCHED TRANSACTIONS DETAIL
        # ============================================================
        
        if match_results['matches']:
            matched_detail = []
            matched_detail.append(['Matched Transactions - Detailed View', '', '', '', ''])
            matched_detail.append(['', '', '', '', ''])
            matched_detail.append(['Ledger Date', 'Ledger Description', 'Ledger Ref', 'Ledger Amount', 
                                   'Bank Date', 'Bank Ref', 'Bank Description', 'Bank Amount', 'Match Method'])
            
            for match in match_results['matches']:
                ledger_date_str = format_date(match['ledger_date']) if pd.notna(match['ledger_date']) else ''
                bank_date_str = format_date(match['bank_date']) if pd.notna(match['bank_date']) else ''
                
                matched_detail.append([
                    ledger_date_str,
                    (match.get('ledger_desc', '')[:50] if match.get('ledger_desc') else ''),
                    (match.get('ledger_ref', '')[:30] if match.get('ledger_ref') else ''),
                    f"{match['amount']:,.2f}",
                    bank_date_str,
                    (match.get('bank_ref', '')[:30] if match.get('bank_ref') else ''),
                    (match.get('bank_desc', '')[:50] if match.get('bank_desc') else ''),
                    f"{match['amount']:,.2f}",
                    match['match_method']
                ])
            
            matched_df = pd.DataFrame(matched_detail)
            matched_df.to_excel(writer, sheet_name='MATCHED_DETAIL', index=False, header=False)
        
        # ============================================================
        # Sheet 4: UNMATCHED - LEDGER ONLY
        # ============================================================
        
        if match_results['unmatched_ledger_credits'] or match_results['unmatched_ledger_debits']:
            unmatch_ledger = []
            unmatch_ledger.append(['Ledger Transactions with No Bank Match', '', '', ''])
            unmatch_ledger.append(['', '', '', ''])
            unmatch_ledger.append(['Date', 'Description', 'Reference', 'Amount', 'Type'])
            
            for item in match_results['unmatched_ledger_credits']:
                date_str = format_date(item['date']) if pd.notna(item['date']) else ''
                unmatch_ledger.append([
                    date_str,
                    (item['description'][:60] if item['description'] else ''),
                    (item['reference'][:30] if item['reference'] else ''),
                    f"{item['amount']:,.2f}",
                    'CREDIT (Deposit)'
                ])
            
            for item in match_results['unmatched_ledger_debits']:
                date_str = format_date(item['date']) if pd.notna(item['date']) else ''
                unmatch_ledger.append([
                    date_str,
                    (item['description'][:60] if item['description'] else ''),
                    (item['reference'][:30] if item['reference'] else ''),
                    f"{item['amount']:,.2f}",
                    'DEBIT (Payment)'
                ])
            
            unmatch_ledger_df = pd.DataFrame(unmatch_ledger)
            unmatch_ledger_df.to_excel(writer, sheet_name='UNMATCHED_LEDGER', index=False, header=False)
        
        # ============================================================
        # Sheet 5: UNMATCHED - BANK ONLY
        # ============================================================
        
        if match_results['unmatched_bank_credits'] or match_results['unmatched_bank_debits']:
            unmatch_bank = []
            unmatch_bank.append(['Bank Transactions with No Ledger Match', '', '', ''])
            unmatch_bank.append(['', '', '', ''])
            unmatch_bank.append(['Date', 'Description', 'Reference', 'Amount', 'Type'])
            
            for item in match_results['unmatched_bank_credits']:
                date_str = format_date(item['date']) if pd.notna(item['date']) else ''
                unmatch_bank.append([
                    date_str,
                    (item['description'][:60] if item['description'] else ''),
                    (item['reference'][:30] if item['reference'] else ''),
                    f"{item['amount']:,.2f}",
                    'CREDIT (Deposit)'
                ])
            
            for item in match_results['unmatched_bank_debits']:
                date_str = format_date(item['date']) if pd.notna(item['date']) else ''
                unmatch_bank.append([
                    date_str,
                    (item['description'][:60] if item['description'] else ''),
                    (item['reference'][:30] if item['reference'] else ''),
                    f"{item['amount']:,.2f}",
                    'DEBIT (Payment)'
                ])
            
            unmatch_bank_df = pd.DataFrame(unmatch_bank)
            unmatch_bank_df.to_excel(writer, sheet_name='UNMATCHED_BANK', index=False, header=False)
        
        # ============================================================
        # Sheet 6: SUMMARY
        # ============================================================
        
        summary_data = [
            ['RECONCILIATION SUMMARY', ''],
            ['', ''],
            ['Total Bank Transactions', f"{len(bank_df)}"],
            ['Total Ledger Transactions', f"{len(ledger_df)}"],
            ['Matched Transactions', f"{len(match_results['matches'])}"],
            ['', ''],
            ['Unmatched - Ledger Only', f"{len(match_results['unmatched_ledger_credits']) + len(match_results['unmatched_ledger_debits'])}"],
            ['Unmatched - Bank Only', f"{len(match_results['unmatched_bank_credits']) + len(match_results['unmatched_bank_debits'])}"],
            ['', ''],
            ['Opening Bank Balance', f"{recon_statement['opening_balance']:,.2f}"],
            ['Total Reconciling Items', f"{recon_statement['total_adjustment']:,.2f}"],
            ['Adjusted Balance', f"{recon_statement['adjusted_balance']:,.2f}"],
            ['Ledger Balance', f"{recon_statement['ledger_balance']:,.2f}"],
            ['Difference', f"{recon_statement['difference']:,.2f}"],
        ]
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='SUMMARY', index=False, header=False)
        
        # Auto-adjust column widths
        for sheetname in writer.sheets:
            worksheet = writer.sheets[sheetname]
            for column in worksheet.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if cell.value and len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
    
    return output_path

# =========================
# Main UI
# =========================
col1, col2 = st.columns(2)

with col1:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("🏦 Bank Statement")
    bank_file = st.file_uploader(
        "Upload Bank Statement",
        type=["xlsx", "xls"],
        help="Excel file with bank statement (Date, Credit, Debit, Description columns)",
        key="bank"
    )
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("📒 Cashbook / Ledger")
    ledger_file = st.file_uploader(
        "Upload Yellowcob Ledger",
        type=["xlsx", "xls"],
        help="Excel file with ledger (Date, Amount, Description columns)",
        key="ledger"
    )
    st.markdown('</div>', unsafe_allow_html=True)

# Settings
st.markdown('<div class="card">', unsafe_allow_html=True)
st.subheader("⚙️ Reconciliation Settings")

tolerance_col1, tolerance_col2 = st.columns(2)
with tolerance_col1:
    opening_balance_manual = st.number_input(
        "Opening Bank Balance (if not auto-detected)",
        value=0.00,
        step=100.00,
        format="%.2f"
    )
with tolerance_col2:
    use_manual_opening = st.checkbox("Use manual opening balance", value=False)

st.markdown('</div>', unsafe_allow_html=True)

# Run button
run_recon = safe_primary_button("🔄 RUN RECONCILIATION", use_container_width=True)

# =========================
# Main processing
# =========================
if run_recon:
    if not bank_file or not ledger_file:
        st.error("Please upload both Bank Statement and Ledger files")
        st.stop()
    
    sbox = status_box("Processing files...")
    
    try:
        # Load bank statement
        if sbox:
            sbox.write("Loading bank statement...")
        bank_df, bank_opening, bank_closing, bank_header = load_bank_statement(bank_file)
        
        if bank_df.empty:
            st.error("Could not detect transactions in bank statement. Check file format.")
            st.stop()
        
        if sbox:
            sbox.write(f"Found {len(bank_df)} bank transactions")
        
        # Load ledger
        if sbox:
            sbox.write("Loading ledger...")
        ledger_df, ledger_closing, ledger_header = load_ledger(ledger_file)
        
        if ledger_df.empty:
            st.error("Could not detect transactions in ledger. Check file format.")
            st.stop()
        
        if sbox:
            sbox.write(f"Found {len(ledger_df)} ledger transactions")
        
        # Determine opening balance
        opening_balance = opening_balance_manual if use_manual_opening else (bank_opening or 0)
        
        # Match transactions
        if sbox:
            sbox.write("Matching transactions (amount first, then date)...")
        match_results = match_transactions(bank_df, ledger_df)
        
        if sbox:
            sbox.write(f"Found {len(match_results['matches'])} matched pairs")
            sbox.write("Building working paper...")
        
        # Build outputs with new structure
        working_paper = build_working_paper(match_results)
        recon_statement = build_recon_statement(opening_balance, ledger_closing, match_results)
        
        # Export
        if sbox:
            sbox.write("Generating Excel output...")
        output_file = export_to_excel(working_paper, recon_statement, bank_df, ledger_df, match_results)
        
        if sbox:
            sbox.update(label="Reconciliation complete!", state="complete", expanded=False)
        
        # Display results
        st.success(f"✅ Reconciliation complete! {len(match_results['matches'])} transactions matched.")
        
        # Summary metrics
        metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
        with metric_col1:
            st.metric("Matched Transactions", len(match_results['matches']))
        with metric_col2:
            st.metric("Unmatched - Ledger", len(match_results['unmatched_ledger_credits']) + len(match_results['unmatched_ledger_debits']))
        with metric_col3:
            st.metric("Unmatched - Bank", len(match_results['unmatched_bank_credits']) + len(match_results['unmatched_bank_debits']))
        with metric_col4:
            diff = recon_statement['difference']
            st.metric("Balance Difference", f"{diff:,.2f}", delta="Zero" if abs(diff) < 0.01 else "Check")
        
        # Show working paper preview
        st.markdown("---")
        st.subheader("📋 Working Paper Preview")
        
        # Display the new structure
        display_cols = ['SECTION', 'LEDGER_DATE', 'LEDGER_DESC', 'LEDGER_REF', 'LEDGER_AMOUNT', 
                        'MATCH_STATUS', 'BANK_AMOUNT', 'BANK_DATE', 'BANK_REF', 'BANK_DESC']
        
        wp_preview = working_paper[display_cols].head(30).copy()
        
        # Format dates and amounts for display
        wp_preview['LEDGER_DATE'] = wp_preview['LEDGER_DATE'].apply(lambda x: format_date(x) if pd.notna(x) else '')
        wp_preview['BANK_DATE'] = wp_preview['BANK_DATE'].apply(lambda x: format_date(x) if pd.notna(x) else '')
        wp_preview['LEDGER_AMOUNT'] = wp_preview['LEDGER_AMOUNT'].apply(lambda x: f"{float(x):,.2f}" if pd.notna(x) and x != '' else '')
        wp_preview['BANK_AMOUNT'] = wp_preview['BANK_AMOUNT'].apply(lambda x: f"{float(x):,.2f}" if pd.notna(x) and x != '' else '')
        
        st.dataframe(wp_preview, use_container_width=True)
        
        # Show reconciliation statement preview
        st.markdown("---")
        st.subheader("📄 Reconciliation Statement Preview")
        
        recon_preview = []
        recon_preview.append({"Item": "Opening Balance (Bank)", "Amount": f"{recon_statement['opening_balance']:,.2f}"})
        for _, item in recon_statement['recon_items'].head(10).iterrows():
            recon_preview.append({"Item": item['description'][:60], "Amount": f"{item['adjustment']:,.2f}"})
        if len(recon_statement['recon_items']) > 10:
            recon_preview.append({"Item": f"... and {len(recon_statement['recon_items']) - 10} more items", "Amount": ""})
        recon_preview.append({"Item": "--- TOTAL ADJUSTMENTS ---", "Amount": f"{recon_statement['total_adjustment']:,.2f}"})
        recon_preview.append({"Item": "Adjusted Balance", "Amount": f"{recon_statement['adjusted_balance']:,.2f}"})
        recon_preview.append({"Item": "Ledger Balance", "Amount": f"{recon_statement['ledger_balance']:,.2f}"})
        recon_preview.append({"Item": "Difference", "Amount": f"{recon_statement['difference']:,.2f}"})
        
        st.dataframe(pd.DataFrame(recon_preview), use_container_width=True)
        
        # Download button
        with open(output_file, 'rb') as f:
            safe_download_button(
                "📥 Download Excel Report",
                data=f,
                file_name=f"bank_reconciliation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        
    except Exception as e:
        st.error(f"Error processing files: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        if sbox:
            sbox.update(label="Error occurred", state="error")

# =========================
# Footer
# =========================
st.markdown(
    """
    <hr>
    <div style="text-align: center; color: #666666; font-size: 12px; padding: 20px;">
        Bank Reconciliation Tool | Matches by Amount, then Date | Outputs Working Paper & Recon Statement
    </div>
    """,
    unsafe_allow_html=True
)