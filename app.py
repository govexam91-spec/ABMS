import os
import io
import sqlite3
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from flask import Flask, render_template, request, redirect, url_for, session, send_file
from docx import Document
from docx.shared import Pt, Mm, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from werkzeug.utils import secure_filename

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
GTR_TEMPLATE_PATH = os.path.join(BASE_DIR, 'GTR63_template.xlsx')
app.secret_key = 'abms_secret_key_here'

# Database Initialization with UNIQUE constraint on company_name and claim_period
def init_db():
    conn = sqlite3.connect('database.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS verified_bills (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            company_name TEXT,
            claim_period TEXT,
            apprentices INTEGER,
            total_mats REAL,
            UNIQUE(company_name, claim_period)
        )
    ''')
    conn.commit()
    conn.close()

init_db()

@app.route('/')
def home():
    if 'user' in session:
        return redirect(url_for('dashboard'))
    return render_template('login.html')

@app.route('/login', methods=['POST'])
def login():
    email = request.form.get('email')
    password = request.form.get('password')
    session['user'] = email if email else "admin@kknowledgesathi.com"
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('home'))

@app.route('/dashboard')
def dashboard():
    if 'user' in session:
        return render_template('dashboard.html', email=session['user'])
    return redirect(url_for('home'))

@app.route('/mats_stipend')
def mats_stipend():
    if 'user' in session:
        return render_template('mats_stipend.html', email=session['user'], bill_data=None)
    else:
        return redirect(url_for('home'))


def get_all_bills():
    """All verified bills from database (with id)."""
    conn = sqlite3.connect('database.db')
    cursor = conn.cursor()
    cursor.execute('SELECT id, filename, company_name, claim_period, apprentices, total_mats FROM verified_bills ORDER BY id')
    rows = cursor.fetchall()
    conn.close()
    return [{
        'id': r[0],
        'filename': r[1],
        'company_name': r[2],
        'claim_period': r[3],
        'apprentices': r[4],
        'total_mats': r[5]
    } for r in rows]


def get_selected_bills():
    """Only the bills that the user ticked in the popup (stored in session)."""
    selected = set(session.get('selected_bill_ids', []))
    if not selected:
        return []
    return [b for b in get_all_bills() if b['id'] in selected]


@app.route('/fetch_bills')
def fetch_bills():
    if 'user' not in session:
        return redirect(url_for('home'))

    all_bills = get_all_bills()
    valid_ids = {b['id'] for b in all_bills}

    if 'selected_bill_ids' in session:
        selected_ids = [i for i in session['selected_bill_ids'] if i in valid_ids]
    else:
        # first time: everything is ticked
        selected_ids = [b['id'] for b in all_bills]

    # Table behind the popup shows the already selected records (if any)
    bill_data = get_selected_bills() if 'selected_bill_ids' in session else None

    return render_template(
        'mats_stipend.html',
        email=session.get('user'),
        bill_data=bill_data,
        all_bills=all_bills,
        selected_ids=selected_ids,
        show_picker=True,
        active_tab='generate'
    )


@app.route('/select_bills', methods=['POST'])
def select_bills():
    if 'user' not in session:
        return redirect(url_for('home'))

    ids = []
    for value in request.form.getlist('bill_ids'):
        if value.isdigit():
            ids.append(int(value))
    session['selected_bill_ids'] = ids
    return redirect(url_for('show_selected'))


@app.route('/show_selected')
def show_selected():
    if 'user' not in session:
        return redirect(url_for('home'))

    return render_template(
        'mats_stipend.html',
        email=session.get('user'),
        bill_data=get_selected_bills(),
        selection_made='selected_bill_ids' in session,
        active_tab='generate'
    )


@app.route('/check_claim', methods=['POST'])
def check_claim():
    if 'user' not in session:
        return redirect(url_for('home'))
        
    if 'excel_file' not in request.files:
        return redirect(url_for('mats_stipend'))
    
    file = request.files['excel_file']
    messages = []
    success_status = False
    
    if file and file.filename != '':
        filename = secure_filename(file.filename)
        
        if not os.path.exists('uploads'):
            os.makedirs('uploads')
        filepath = os.path.join('uploads', filename)
        file.save(filepath)
        
        try:
            if filename.endswith('.xls'):
                sheets = pd.read_excel(filepath, sheet_name=None, engine='xlrd', header=None)
            else:
                sheets = pd.read_excel(filepath, sheet_name=None, engine='openpyxl', header=None)
            
            sheet_names = list(sheets.keys())
            
            claim_period_str = "JAN-MAR-2026"
            if len(sheet_names) >= 2:
                sheet2 = sheets[sheet_names[1]]
                raw_cp = str(sheet2.iloc[1, 10] if sheet2.shape[0] > 1 and sheet2.shape[1] > 10 else "")
                if ":" in raw_cp:
                    claim_period_str = raw_cp.split(':')[-1].strip()
                else:
                    claim_period_str = raw_cp.strip()

            if len(sheet_names) >= 3:
                sheet3 = sheets[sheet_names[2]]
                
                for idx in range(3, len(sheet3)):
                    row = idx + 1
                    try:
                        F_val = sheet3.iloc[idx, 5] if sheet3.shape[1] > 5 else 0
                        G_val = sheet3.iloc[idx, 6] if sheet3.shape[1] > 6 else 0
                        K_val = sheet3.iloc[idx, 10] if sheet3.shape[1] > 10 else 0
                        L_val = sheet3.iloc[idx, 11] if sheet3.shape[1] > 11 else 0
                        
                        if pd.isna(F_val) and pd.isna(G_val) and pd.isna(K_val) and pd.isna(L_val):
                            continue
                            
                        F = float(0 if pd.isna(F_val) else F_val)
                        G = float(0 if pd.isna(G_val) else G_val)
                        K = float(0 if pd.isna(K_val) else K_val)
                        L = float(0 if pd.isna(L_val) else L_val)
                        
                        if G != 0 and L == 0:
                            messages.append(f"Sheet 3, Row {row}: Stipend Not Pay (G 0 નથી પરંતુ L 0 છે)")
                            
                        if F != 0:
                            calc = (K / F) * G
                            if calc > L:
                                messages.append(f"Sheet 3, Row {row}: False Error (ગણતરી {round(calc, 2)} છે, જે L={L} કરતા મોટી છે)")
                    except Exception as ex:
                        pass

            if len(messages) == 0 and len(sheet_names) >= 4:
                sheet4 = sheets[sheet_names[3]]
                
                raw_comp = str(sheet4.iloc[1, 8] if sheet4.shape[1] > 8 else "")
                if ":" in raw_comp:
                    company_name = raw_comp.split(':')[-1].strip()
                else:
                    company_name = raw_comp.strip()
                
                valid_rows = 0
                sum_val = 0
                
                for idx in range(4, len(sheet4)):
                    sr_val = sheet4.iloc[idx, 1] if sheet4.shape[1] > 1 else None
                    
                    if pd.notna(sr_val) and "Grand Total" in str(sr_val):
                        mats_val = sheet4.iloc[idx, 16] if sheet4.shape[1] > 16 else 0
                        try:
                            sum_val = float(mats_val)
                        except:
                            pass
                        break
                        
                    if pd.notna(sr_val) and str(sr_val).strip().isdigit():
                        valid_rows += 1
                        mats_val = sheet4.iloc[idx, 16] if sheet4.shape[1] > 16 else 0
                        try:
                            sum_val += float(mats_val)
                        except:
                            pass
                            
                total_apprentices = valid_rows
                mats_total_sum = round(sum_val, 2)
                
                conn = sqlite3.connect('database.db')
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT id FROM verified_bills WHERE company_name = ? AND claim_period = ?
                ''', (company_name, claim_period_str))
                existing_record = cursor.fetchone()
                
                if existing_record:
                    messages.append(f"ડુપ્લિકેટ એરર: '{company_name}' અને ક્લેઈમ પિરિયડ '{claim_period_str}' નો રેકોર્ડ પહેલેથી જ ડેટાબેઝમાં મોજૂદ છે!")
                else:
                    try:
                        cursor.execute('''
                            INSERT INTO verified_bills (filename, company_name, claim_period, apprentices, total_mats)
                            VALUES (?, ?, ?, ?, ?)
                        ''', (filename, company_name, claim_period_str, total_apprentices, mats_total_sum))
                        conn.commit()
                        success_status = True
                    except sqlite3.IntegrityError:
                        messages.append(f"ડુપ્લિકેટ એરર: આ રેકોર્ડ પહેલેથી જ ડેટાબેઝમાં સાચવેલો છે.")
                
                conn.close()

        except Exception as e:
            messages.append(f"ફાઈલ પ્રોસેસ કરવામાં એરર આવી: {str(e)}")
            
        finally:
            if os.path.exists(filepath):
                os.remove(filepath)
                
    return render_template('mats_stipend.html', email=session.get('user'), messages=messages, success=success_status, bill_data=None)


@app.route('/reset_history')
def reset_history():
    conn = sqlite3.connect('database.db')
    cursor = conn.cursor()
    cursor.execute('DELETE FROM verified_bills')
    conn.commit()
    conn.close()
    session.pop('selected_bill_ids', None)
    return redirect(url_for('mats_stipend'))


# Amount in words (Indian numbering: Crore, Lac, Thousand, Hundred)
ONES = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
        "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
        "Seventeen", "Eighteen", "Nineteen"]
TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]

def _two_digits(n):
    if n < 20:
        return ONES[n]
    return (TENS[n // 10] + (" " + ONES[n % 10] if n % 10 else "")).strip()

def _three_digits(n):
    words = []
    if n >= 100:
        words.append(ONES[n // 100] + " Hundred")
        n %= 100
    if n:
        words.append(_two_digits(n))
    return " ".join(words)

def number_to_words(num):
    try:
        num = round(float(num), 2)
    except (TypeError, ValueError):
        return ""
    rupees = int(num)
    paise = int(round((num - rupees) * 100))

    if rupees == 0 and paise == 0:
        return "Zero Only"

    parts = []
    crore, rupees = divmod(rupees, 10000000)
    lac, rupees = divmod(rupees, 100000)
    thousand, rupees = divmod(rupees, 1000)

    if crore:
        parts.append(_three_digits(crore) + " Crore")
    if lac:
        parts.append(_two_digits(lac) + " Lac")
    if thousand:
        parts.append(_two_digits(thousand) + " Thousand")
    if rupees:
        parts.append(_three_digits(rupees))

    result = ", ".join(parts)
    if paise:
        result = (result + " and " if result else "") + _two_digits(paise) + " Paise"
    return result + " Only"


# Stipend Schedule Excel Generation Function with smart earliest and latest date range finder
def generate_stipend_schedule(company_data_list, wb=None):
    if wb is None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Stipend Schedule"
    else:
        ws = wb.create_sheet("Stipend Schedule")
    
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    
    ws.views.sheetView[0].showGridLines = True
    
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    
    base_font = Font(name='Hind Vadodara', size=11)
    bold_font = Font(name='Hind Vadodara', size=11, bold=True)
    
    # Smart Date Range Finder across ALL records in the database table
    month_map_rev = {1: 'JAN', 2: 'FEB', 3: 'MAR', 4: 'APR', 5: 'MAY', 6: 'JUN', 
                     7: 'JUL', 8: 'AUG', 9: 'SEP', 10: 'OCT', 11: 'NOV', 12: 'DEC'}
    month_map = {
        'JAN': 1, 'FEB': 2, 'MAR': 3, 'APR': 4, 'MAY': 5, 'JUN': 6,
        'JUL': 7, 'AUG': 8, 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DEC': 12
    }
    
    earliest_date = None
    latest_date = None
    
    for item in company_data_list:
        cp = str(item.get('claim_period', '')).strip().upper()
        clean_cp = cp.replace(' ', '').replace('/', '-').replace('TO', '-')
        parts = [p for p in clean_cp.split('-') if p]
        if len(parts) >= 3:
            s_m_str = parts[0][:3]
            e_m_str = parts[1][:3]
            y_str = parts[2]
            if len(y_str) == 2:
                y_str = '20' + y_str
            try:
                s_m = month_map.get(s_m_str, 1)
                e_m = month_map.get(e_m_str, s_m)
                year = int(y_str)
                
                s_date = datetime(year, s_m, 1)
                e_date = datetime(year, e_m, 1)
                
                if earliest_date is None or s_date < earliest_date:
                    earliest_date = s_date
                if latest_date is None or e_date > latest_date:
                    latest_date = e_date
            except:
                pass
                
    if earliest_date and latest_date:
        start_str = f"{month_map_rev[earliest_date.month]}-{earliest_date.year}"
        end_str = f"{month_map_rev[latest_date.month]}-{latest_date.year}"
        period_title_str = f"{start_str} TO {end_str}"
    else:
        period_title_str = "JAN-2026 TO JUN-2026"

    ws.merge_cells('A1:H1')
    ws['A1'] = f"BILL OF SCHOLARSHIP TENABLE AT Industrial Training Institute, Surendranagar DURING THE MONTH OF {period_title_str}"
    ws['A1'].font = bold_font
    ws['A1'].alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[1].height = 30

    headers = [
        "Sr. No.", "Year of Award", "Kind of Scholarship / Stipend", 
        "Name of Scholarship/ Stipend Holder", "Monthly Value of Stipend", 
        "No. of Days for Which Drawn", "Amount Drawn", "Remarks"
    ]
    
    ws.append([]) 
    ws.append(["1", "2", "3", "4", "5", "6", "7", "8"]) 
    ws.append(headers) 
    
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=4, column=col_num)
        cell.font = bold_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[4].height = 35
        
    for col_num in range(1, 9):
        cell = ws.cell(row=3, column=col_num)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
        cell.font = bold_font
    ws.row_dimensions[3].height = 22

    current_row = 5
    total_amount = 0
    
    for idx, item in enumerate(company_data_list, start=1):
        ws.cell(row=current_row, column=1, value=idx)
        ws.cell(row=current_row, column=2, value="2026-27")
        ws.cell(row=current_row, column=3, value="STIPEND BILL OF THE MATS")
        ws.cell(row=current_row, column=4, value=item.get('company_name', ''))
        ws.cell(row=current_row, column=5, value=item.get('apprentices', 0))
        ws.cell(row=current_row, column=6, value=item.get('claim_period', ''))
        ws.cell(row=current_row, column=7, value=item.get('total_mats', 0))
        ws.cell(row=current_row, column=8, value="")
        
        total_amount += float(item.get('total_mats', 0))
        
        for col in range(1, 9):
            c = ws.cell(row=current_row, column=col)
            c.border = thin_border
            c.font = base_font
            if col in [1, 2, 5, 6, 7]:
                c.alignment = Alignment(horizontal='center', vertical='center')
            else:
                c.alignment = Alignment(horizontal='left', vertical='center')
                
        ws.row_dimensions[current_row].height = 26
        current_row += 1

    r1 = current_row
    ws.merge_cells(start_row=r1, start_column=1, end_row=r1, end_column=2)
    ws.cell(row=r1, column=1, value="IN WARDS RUPEES.").font = bold_font
    
    ws.merge_cells(start_row=r1, start_column=3, end_row=r1, end_column=5)
    total_amount = round(total_amount, 2)
    ws.cell(row=r1, column=3, value=number_to_words(total_amount)).font = bold_font
    ws.cell(row=r1, column=3).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    
    ws.cell(row=r1, column=6, value="TOTAL AMOUNT").font = bold_font
    ws.cell(row=r1, column=6).alignment = Alignment(horizontal='center', vertical='center')
    
    ws.cell(row=r1, column=7, value=total_amount).font = bold_font
    ws.cell(row=r1, column=7).alignment = Alignment(horizontal='center', vertical='center')
    
    for col in range(1, 9):
        ws.cell(row=r1, column=col).border = thin_border
    ws.row_dimensions[r1].height = 40

    r2 = r1 + 1
    ws.merge_cells(start_row=r2, start_column=3, end_row=r2, end_column=6)
    ws.cell(row=r2, column=3, value="DEDUCT AMOUNT UNDISBURSED FROM LAST MONTH").font = bold_font
    ws.cell(row=r2, column=7, value=0.00).font = bold_font
    ws.cell(row=r2, column=7).alignment = Alignment(horizontal='center', vertical='center')
    
    for col in range(1, 9):
        ws.cell(row=r2, column=col).border = thin_border
    ws.row_dimensions[r2].height = 26

    r3 = r2 + 1
    ws.merge_cells(start_row=r3, start_column=3, end_row=r3, end_column=6)
    ws.cell(row=r3, column=3, value="BALANCE DUE").font = bold_font
    ws.cell(row=r3, column=7, value=total_amount).font = bold_font
    ws.cell(row=r3, column=7).alignment = Alignment(horizontal='center', vertical='center')
    
    for col in range(1, 9):
        ws.cell(row=r3, column=col).border = thin_border
    ws.row_dimensions[r3].height = 28

    r4 = r3 + 2
    ws.merge_cells(start_row=r4, start_column=1, end_row=r4, end_column=8)
    ws.cell(row=r4, column=1, value="I Hereby certify that the scholarship or stipend holders named in this bill have been regular in attendance and have confirmed to the rule under which their scholarship or stipend are tenable.").font = base_font
    ws.row_dimensions[r4].height = 28

    r5 = r4 + 1
    ws.merge_cells(start_row=r5, start_column=1, end_row=r5, end_column=8)
    ws.cell(row=r5, column=1, value="Certified also that the scholarship or stipends drawn on the last bill, with exception of these refunded by deduction have paid to the proper persons and their receipts taken in acquittace rolls kept in my office.").font = base_font
    ws.row_dimensions[r5].height = 32

    r6 = r5 + 3
    ws.merge_cells(start_row=r6, start_column=7, end_row=r6, end_column=8)
    cell_sign = ws.cell(row=r6, column=7, value="Received Payment")
    cell_sign.font = bold_font
    cell_sign.alignment = Alignment(horizontal='center', vertical='center')
    
    for col in range(7, 9):
        ws.cell(row=r6, column=col).border = thin_border
    ws.row_dimensions[r6].height = 35

    column_widths = {
        'A': 10,
        'B': 15,
        'C': 30,
        'D': 35,
        'E': 18,
        'F': 18,
        'G': 18,
        'H': 18
    }
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    return wb

def get_financial_year(today=None):
    """Current financial year (April to March), e.g. Oct-2026 -> '2026-27'."""
    today = today or datetime.now()
    start_year = today.year if today.month >= 4 else today.year - 1
    return f"{start_year}-{str(start_year + 1)[-2:]}"


def fill_gtr_sheet(ws, total_amount):
    """Fill the GTR-63 sheet: month, amount, amount in words, budget allotment year."""
    now = datetime.now()

    # During the month of : -> current month & year (cell format shows e.g. October-26)
    ws['P11'] = datetime(now.year, now.month, 1)

    # Amount (Rs.) in the Object of Expenditure table
    ws['AK59'] = round(float(total_amount), 2)

    # In words: Rupees...
    ws['AB65'] = number_to_words(total_amount)

    # Budget Allotment for : current financial year
    ws['U84'] = f"Budget Allotment for   {get_financial_year(now)}"


def build_workbook(history):
    """GTR-63 sheet first, then the Stipend Schedule sheet."""
    total_amount = round(sum(float(item.get('total_mats') or 0) for item in history), 2)

    wb = openpyxl.load_workbook(GTR_TEMPLATE_PATH)

    # Remove helper sheets of the template (old INNER sheet with broken formulas + hidden sheet)
    for name in ['INNER', '000']:
        if name in wb.sheetnames:
            del wb[name]

    fill_gtr_sheet(wb['GTR'], total_amount)
    generate_stipend_schedule(history, wb=wb)

    # Open the file on the first sheet (GTR-63)
    wb.active = 0
    for i, sheet in enumerate(wb.worksheets):
        sheet.sheet_view.tabSelected = (i == 0)

    return wb


# ============================================================
# WORD (.docx) - Approval Order / "મંજુરી હુકમ" generation
# ============================================================
WORD_FONT = "Hind Vadodara"
WORD_FONT_SIZE = 11

GUJ_DIGITS = str.maketrans("0123456789", "૦૧૨૩૪૫૬૭૮૯")

def to_gujarati_digits(value):
    return str(value).translate(GUJ_DIGITS)

GUJ_NUMBERS = [
    "", "એક", "બે", "ત્રણ", "ચાર", "પાંચ", "છ", "સાત", "આઠ", "નવ",
    "દસ", "અગિયાર", "બાર", "તેર", "ચૌદ", "પંદર", "સોળ", "સત્તર", "અઢાર", "ઓગણીસ",
    "વીસ", "એકવીસ", "બાવીસ", "ત્રેવીસ", "ચોવીસ", "પચ્ચીસ", "છવ્વીસ", "સત્તાવીસ", "અઠ્ઠાવીસ", "ઓગણત્રીસ",
    "ત્રીસ", "એકત્રીસ", "બત્રીસ", "તેત્રીસ", "ચોત્રીસ", "પાંત્રીસ", "છત્રીસ", "સાડત્રીસ", "આડત્રીસ", "ઓગણચાલીસ",
    "ચાલીસ", "એકતાલીસ", "બેતાલીસ", "ત્રેતાલીસ", "ચુમ્માલીસ", "પિસ્તાલીસ", "છેતાલીસ", "સુડતાલીસ", "અડતાલીસ", "ઓગણપચાસ",
    "પચાસ", "એકાવન", "બાવન", "ત્રેપન", "ચોપન", "પંચાવન", "છપ્પન", "સત્તાવન", "અઠ્ઠાવન", "ઓગણસાઠ",
    "સાઠ", "એકસઠ", "બાસઠ", "ત્રેસઠ", "ચોસઠ", "પાંસઠ", "છાસઠ", "સડસઠ", "અડસઠ", "ઓગણોસિત્તેર",
    "સિત્તેર", "એકોતેર", "બોતેર", "તોતેર", "ચુમોતેર", "પંચોતેર", "છોતેર", "સિત્યોતેર", "ઇઠ્યોતેર", "ઓગણએંસી",
    "એંસી", "એક્યાસી", "બ્યાસી", "ત્યાસી", "ચોર્યાસી", "પંચાસી", "છ્યાસી", "સત્યાસી", "અઠ્યાસી", "નેવ્યાસી",
    "નેવું", "એકાણું", "બાણું", "ત્રાણું", "ચોરાણું", "પંચાણું", "છન્નું", "સત્તાણું", "અઠ્ઠાણું", "નવ્વાણું",
]
GUJ_HUNDREDS = ["", "એકસો", "બસો", "ત્રણસો", "ચારસો", "પાંચસો", "છસો", "સાતસો", "આઠસો", "નવસો"]


def number_to_words_gujarati(num):
    """Amount in Gujarati words (Indian system: કરોડ, લાખ, હજાર, સો)."""
    try:
        num = round(float(num), 2)
    except (TypeError, ValueError):
        return ""
    rupees = int(num)
    paise = int(round((num - rupees) * 100))

    if rupees == 0 and paise == 0:
        return "શૂન્ય પુરા"

    def below_thousand(n):
        parts = []
        if n >= 100:
            parts.append(GUJ_HUNDREDS[n // 100])
            n %= 100
        if n:
            parts.append(GUJ_NUMBERS[n])
        return " ".join(parts)

    parts = []
    crore, rupees = divmod(rupees, 10000000)
    lakh, rupees = divmod(rupees, 100000)
    thousand, rupees = divmod(rupees, 1000)

    if crore:
        parts.append(below_thousand(crore) + " કરોડ")
    if lakh:
        parts.append(GUJ_NUMBERS[lakh] + " લાખ")
    if thousand:
        parts.append(GUJ_NUMBERS[thousand] + " હજાર")
    if rupees:
        parts.append(below_thousand(rupees))

    result = " ".join(parts)
    if paise:
        result = (result + " અને " if result else "") + GUJ_NUMBERS[paise] + " પૈસા"
    return result + " પુરા"


def format_claim_period(cp):
    """'Jan - Mar - 2026' / 'JAN-MAR-2026' -> 'JAN-MAR-26'"""
    text = str(cp or '').strip().upper().replace(' ', '')
    parts = [p for p in text.split('-') if p]
    if parts and parts[-1].isdigit() and len(parts[-1]) == 4:
        parts[-1] = parts[-1][-2:]
    return '-'.join(parts)


def format_amount(value):
    value = float(value or 0)
    if value == int(value):
        return str(int(value))
    return f"{value:.2f}"


def _set_run_font(run, bold=False, underline=False):
    run.font.name = WORD_FONT
    run.font.size = Pt(WORD_FONT_SIZE)
    run.bold = bold
    run.underline = underline
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.insert(0, rfonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rfonts.set(qn(attr), WORD_FONT)
    # complex-script (Gujarati) size + bold
    for tag in ('w:szCs',):
        el = rpr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rpr.append(el)
        el.set(qn('w:val'), str(WORD_FONT_SIZE * 2))
    if bold:
        if rpr.find(qn('w:bCs')) is None:
            rpr.append(OxmlElement('w:bCs'))


def _add_par(doc, text="", bold=False, underline=False, align=None,
             left_indent=None, first_line=None, space_after=0, space_before=0, keep_next=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing = 1.0
    pf.keep_with_next = keep_next
    if left_indent is not None:
        pf.left_indent = left_indent
    if first_line is not None:
        pf.first_line_indent = first_line
    if align is not None:
        p.alignment = align
    if text:
        _set_run_font(p.add_run(text), bold=bold, underline=underline)
    return p


def _set_cell(cell, text, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    _set_run_font(p.add_run(str(text)), bold=bold)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def _set_cell_width(cell, width):
    cell.width = width
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn('w:tcW'))
    if tcW is None:
        tcW = OxmlElement('w:tcW')
        tcPr.append(tcW)
    tcW.set(qn('w:type'), 'dxa')
    tcW.set(qn('w:w'), str(int(width.twips)))


def _row_no_split(row, header=False):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:cantSplit'))
    if header:
        trPr.append(OxmlElement('w:tblHeader'))


def generate_approval_order_docx(history):
    now = datetime.now()
    year_guj = to_gujarati_digits(now.year)
    fy_start = now.year if now.month >= 4 else now.year - 1
    fy_text = to_gujarati_digits(f"{fy_start}-{str(fy_start + 1)[-2:]}")

    total_amount = round(sum(float(item.get('total_mats') or 0) for item in history), 2)

    doc = Document()

    # ---- A4 page, print friendly margins
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    usable_width = section.page_width - section.left_margin - section.right_margin

    # ---- Default font: Hind Vadodara 11 (Latin + Gujarati)
    normal = doc.styles['Normal']
    normal.font.name = WORD_FONT
    normal.font.size = Pt(WORD_FONT_SIZE)
    n_rpr = normal.element.get_or_add_rPr()
    n_fonts = n_rpr.find(qn('w:rFonts'))
    if n_fonts is None:
        n_fonts = OxmlElement('w:rFonts')
        n_rpr.insert(0, n_fonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        n_fonts.set(qn(attr), WORD_FONT)
    n_szcs = OxmlElement('w:szCs')
    n_szcs.set(qn('w:val'), str(WORD_FONT_SIZE * 2))
    n_rpr.append(n_szcs)

    # ---- Header block (right side of the page)
    indent = Cm(9.5)
    _add_par(doc, f"ક્રમાંક:ઔતાસં/સુન/હસબ/MATS/{year_guj}/", left_indent=indent)
    _add_par(doc, "આચાર્યશ્રીની કચેરી,", left_indent=indent)
    _add_par(doc, "ઔદ્યોગિક તાલીમ સંસ્થા,", left_indent=indent)
    _add_par(doc, "બસ સ્ટેન્ડ સામે, સુરેન્દ્રનગર -૩૬૩૦૦૧", left_indent=indent)
    _add_par(doc, "તા.", left_indent=indent, space_after=6)

    # ---- References
    _add_par(doc, "વંચાણે લીધા:", bold=True, underline=True, space_after=2, keep_next=True)
    references = [
        "૧. વ.ક.ગાંધીનગરનો તા.૧૭-૦૧-૨૦૧૯ નો પરિપત્ર ક્રમાંક:રોતાનિ/એટીએસ/છ/૨૦૧૯/૪૮",
        "૨. વ.ક.ના તા.૦૯/૦૮/૨૦૧૯ ના પરિપત્ર ક્રમાંક:રોતાનિ/એટીએસ/છ/૨૦૧૯/૩૩૪",
        "૩. વ.ક.ના તા.૨૩/૦૮/૨૦૧૯ ના પત્ર ક્રમાંક:રોતાનિ/એટીએસ/છ-૧/MATS/૨૦૧૯/૩૮૩",
        "૪. વ.ક.ના તા.૨૯/૦૮/૨૦૧૯ ના પત્ર ક્રમાંક:રોતાનિ/એપ્રે/MATS/છ-૧/૨૦૧૯/૪૩૨",
        "૫. વ.ક.ના તા.૨૧/૦૮/૨૦૧૯ ના પત્ર ક્રમાંક:રોતાનિ/એટીએસ/છ/MATS/૨૦૧૯/૪૫૩",
        "૬. વ.ક.ના તા.૧૨/૦૭/૨૦૨૩ના પત્ર ક્રમાંક:LED/BAW/E-5/11/2023/0155/R-2 SECTION",
        "૭. વ.ક.ના તા.૧૧/૦૮/૨૦૨૩ના પત્ર ક્રમાંક:DET/0175/08/2023",
    ]
    for ref in references:
        _add_par(doc, ref)

    # ---- Order text
    _add_par(doc, "મંજુરી હુકમ:", bold=True, underline=True, space_before=6, space_after=2, keep_next=True)
    _add_par(
        doc,
        "ઉપરોક્ત વંચાણે લીધેલ ઠરાવથી મળેલ સત્તાની રુએ આ સંસ્થાના MATS યોજના અંતર્ગત એકમોને "
        "ચુકવવા પાત્ર નાણાકીય વળતર ની રકમ આ સાથે સામેલ લીસ્ટ મુજબ તેઓના ખાતામાં ચુકવવા આદેશ "
        "કરવામાં આવે છે.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=Cm(1.0), space_after=6)

    # ---- Table (data from database)
    col_widths = [Cm(1.3), Cm(7.0), Cm(3.1), Cm(3.4), Cm(1.9)]
    scale = usable_width / sum(col_widths)
    col_widths = [Emu(int(w * scale)) for w in col_widths]

    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    headers = ["ક્રમ", "એકમનું નામ", "ક્લેઇમ ક્વાર્ટર", "બિલની ચુકવવાપાત્ર રકમ", "રીમાર્ક્સ"]
    for i, h in enumerate(headers):
        _set_cell(table.rows[0].cells[i], h, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _row_no_split(table.rows[0], header=True)

    for idx, item in enumerate(history, start=1):
        cells = table.add_row().cells
        _set_cell(cells[0], idx, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell(cells[1], item.get('company_name', ''))
        _set_cell(cells[2], format_claim_period(item.get('claim_period', '')), align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell(cells[3], format_amount(item.get('total_mats', 0)), align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell(cells[4], "")
        _row_no_split(table.rows[-1])

    total_cells = table.add_row().cells
    _set_cell(total_cells[0], "")
    _set_cell(total_cells[1], "")
    _set_cell(total_cells[2], "TOTAL", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell(total_cells[3], f"{total_amount:.2f}", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell(total_cells[4], "")
    _row_no_split(table.rows[-1])

    for i, col in enumerate(table.columns):
        col.width = col_widths[i]
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            _set_cell_width(cell, col_widths[i])
    # cell padding for readability
    tblPr = table._tbl.tblPr
    mar = OxmlElement('w:tblCellMar')
    for side, w in (('top', 30), ('left', 80), ('bottom', 30), ('right', 80)):
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:w'), str(w))
        el.set(qn('w:type'), 'dxa')
        mar.append(el)
    tblPr.append(mar)

    # ---- Amount in words (Gujarati)
    _add_par(doc, f"(અંકે રૂપિયા {number_to_words_gujarati(total_amount)} )",
             bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=8, space_after=6, keep_next=True)

    # ---- Financial year line
    _add_par(
        doc,
        f"આ અંગેનો ખર્ચ ચાલુ નાણાકીય વર્ષ:{fy_text} માટે ફાળવવામાં આવેલ ગ્રાન્ટમાંથી નીચેના સદર "
        "હેઠળ ઉધારવાનો રહેશે.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=8, keep_next=True)

    # ---- Heads
    heads = [
        ("DEMAND NO:", "057 ATS PLAN"),
        ("MAJOR HEAD:", "2230,LABOUR, EMPLOYMENT AND SKILL DEVELOPMENT"),
        ("SUB MAJOR HEAD:", "03,TRAINIG"),
        ("MINOR HEAD:", "102 APPRENTICESHIP TRAINING"),
        ("SUB HEAD:", "01,EMP-4,National Apprenticeship Training"),
    ]
    for label, value in heads:
        p = _add_par(doc, f"{label}\t{value}", keep_next=True)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(4.2))

    # ---- To
    _add_par(doc, "પ્રતિ,", bold=True, underline=True, space_before=14, keep_next=True)
    _add_par(doc, "જીલ્લા તિજોરી અધિકારીશ્રી,", bold=True, keep_next=True)
    _add_par(doc, "જીલ્લા તિજોરી કચેરી,સુરેન્દ્રનગર", bold=True)

    return doc


# ============================================================
# PDF conversion (LibreOffice) + download routes
# ============================================================
def find_libreoffice():
    import shutil
    for name in ('soffice', 'soffice.exe', 'libreoffice'):
        path = shutil.which(name)
        if path:
            return path
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def convert_to_pdf(file_bytes, extension):
    """Convert xlsx / docx bytes to PDF bytes using LibreOffice (headless)."""
    import tempfile
    import subprocess
    from pathlib import Path

    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError("PDF બનાવવા માટે કમ્પ્યુટરમાં LibreOffice ઇન્સ્ટોલ હોવું જરૂરી છે "
                           "(libreoffice.org). ઇન્સ્ટોલ કર્યા પછી ફરી પ્રયત્ન કરો.")

    with tempfile.TemporaryDirectory() as tmp:
        in_path = os.path.join(tmp, 'document' + extension)
        with open(in_path, 'wb') as f:
            f.write(file_bytes)

        profile_uri = Path(os.path.join(tmp, 'lo_profile')).as_uri()
        cmd = [soffice, f'-env:UserInstallation={profile_uri}', '--headless',
               '--convert-to', 'pdf', '--outdir', tmp, in_path]
        try:
            subprocess.run(cmd, check=True, timeout=180, capture_output=True)
        except subprocess.TimeoutExpired:
            raise RuntimeError("PDF બનાવવામાં વધારે સમય લાગ્યો. ફરી પ્રયત્ન કરો.")
        except subprocess.CalledProcessError:
            raise RuntimeError("PDF બનાવવામાં એરર આવી. ફરી પ્રયત્ન કરો.")

        pdf_path = os.path.join(tmp, 'document.pdf')
        if not os.path.exists(pdf_path):
            raise RuntimeError("PDF ફાઈલ બની શકી નથી. ફરી પ્રયત્ન કરો.")
        with open(pdf_path, 'rb') as f:
            return f.read()


@app.route('/export_excel', methods=['POST'])
def export_excel():
    if 'user' not in session:
        return "ફરીથી લોગિન કરો.", 401

    history = get_selected_bills()
    if not history:
        return "કોઈ રેકોર્ડ પસંદ કર્યો નથી. પહેલા Fetch Bills માં રેકોર્ડ પસંદ કરો.", 400

    if not os.path.exists(GTR_TEMPLATE_PATH):
        return "GTR63_template.xlsx ફાઈલ app.py ની સાથે એ જ ફોલ્ડરમાં મૂકો, પછી ફરી પ્રયત્ન કરો.", 500

    wb = build_workbook(history)
    output = io.BytesIO()
    wb.save(output)
    xlsx_bytes = output.getvalue()

    if request.form.get('file_format') == 'pdf':
        try:
            pdf_bytes = convert_to_pdf(xlsx_bytes, '.xlsx')
        except RuntimeError as e:
            return str(e), 500
        return send_file(io.BytesIO(pdf_bytes), mimetype='application/pdf',
                         as_attachment=True, download_name='Stipend_Bill_Schedule.pdf')

    return send_file(
        io.BytesIO(xlsx_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Stipend_Bill_Schedule.xlsx'
    )


@app.route('/export_word', methods=['POST'])
def export_word():
    if 'user' not in session:
        return "ફરીથી લોગિન કરો.", 401

    history = get_selected_bills()
    if not history:
        return "કોઈ રેકોર્ડ પસંદ કર્યો નથી. પહેલા Fetch Bills માં રેકોર્ડ પસંદ કરો.", 400

    doc = generate_approval_order_docx(history)
    output = io.BytesIO()
    doc.save(output)
    docx_bytes = output.getvalue()

    if request.form.get('file_format') == 'pdf':
        try:
            pdf_bytes = convert_to_pdf(docx_bytes, '.docx')
        except RuntimeError as e:
            return str(e), 500
        return send_file(io.BytesIO(pdf_bytes), mimetype='application/pdf',
                         as_attachment=True, download_name='MATS_Approval_Order.pdf')

    return send_file(
        io.BytesIO(docx_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name='MATS_Approval_Order.docx'
    )


if __name__ == '__main__':
    app.run(debug=True)