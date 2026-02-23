import re
import io
import csv
import logging
import pandas as pd
import pdfplumber
from docx import Document
from flask import Flask, request, render_template_string, send_file, redirect, url_for

logging.basicConfig(level=logging.WARNING)
logger = logging.getLogger(__name__)

# ==========================================
# 1. UI TEMPLATE (HTML + TAILWIND CSS)
# ==========================================
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KAKOS Audit Tool</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Inter', sans-serif; background-color: #f8fafc; }
        .currency { font-family: 'Monaco', monospace; }
        .glass-panel { background: white; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); border-radius: 1rem; }
    </style>
</head>
<body class="text-slate-800">

    <nav class="bg-white border-b border-slate-200 sticky top-0 z-50">
        <div class="max-w-7xl mx-auto px-6 py-4 flex justify-between items-center">
            <a href="/?reset=1" class="flex items-center gap-3 hover:opacity-75 transition-opacity" title="Go to Upload Page">
                <div class="bg-emerald-600 text-white p-2 rounded-lg font-bold">KA</div>
                <div>
                    <h1 class="font-bold text-lg leading-tight">KAKOS AUDIT</h1>
                    <p class="text-xs text-slate-500 font-bold tracking-wider">LOCAL DATA ENGINE</p>
                </div>
            </a>
            {% if filename %}
            <div class="flex items-center gap-4">
                <span class="text-sm text-slate-500 hidden md:block">Active File: <span class="font-mono text-slate-700">{{ filename }}</span></span>
                <a href="/?reset=1" class="text-sm font-bold text-rose-600 hover:text-rose-700 px-3 py-1 bg-rose-50 rounded-lg transition-colors">Reset / Upload New</a>
            </div>
            {% endif %}
        </div>
    </nav>

    <main class="max-w-7xl mx-auto px-6 py-8">
        {% if not filename %}
        <div class="max-w-xl mx-auto mt-20 text-center">
            <div class="glass-panel p-10">
                <div class="w-16 h-16 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mx-auto mb-6">
                    <svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
                </div>
                <h2 class="text-2xl font-bold mb-2">Upload Bank Statement</h2>
                <p class="text-slate-500 mb-8">Upload a bank statement in CSV, Word (.docx), or PDF format.<br>All processing is done locally on your machine.</p>
                
                {% if error %}
                <div class="mb-6 p-4 bg-rose-50 border border-rose-200 rounded-xl text-rose-700 text-sm font-medium">
                    ⚠️ {{ error }}
                </div>
                {% endif %}

                <form action="/" method="post" enctype="multipart/form-data" class="relative">
                    <input type="file" name="file" id="file" class="hidden" accept=".csv, .docx, .pdf" onchange="this.form.submit()">
                    <label for="file" class="block w-full py-4 bg-slate-900 text-white font-bold rounded-xl cursor-pointer hover:bg-slate-800 transition-all shadow-lg hover:shadow-xl">
                        Select CSV, DOCX, or PDF
                    </label>
                </form>
                <p class="mt-4 text-xs text-slate-400">Supported formats: .csv · .docx · .pdf</p>
            </div>
        </div>
        {% else %}
        <div class="space-y-6">
            <div class="glass-panel p-4 flex flex-col md:flex-row justify-between items-center gap-4">
                <form action="/" method="get" class="flex flex-wrap items-center gap-3 w-full md:w-auto">
                    <div class="flex items-center gap-2 bg-slate-50 border border-slate-200 rounded-lg px-3 py-2">
                        <span class="text-xs font-bold text-slate-400 uppercase">From</span>
                        <input type="date" name="start_date" value="{{ filters.start }}" class="bg-transparent text-sm font-semibold focus:outline-none text-slate-700">
                    </div>
                    <div class="flex items-center gap-2 bg-slate-50 border border-slate-200 rounded-lg px-3 py-2">
                        <span class="text-xs font-bold text-slate-400 uppercase">To</span>
                        <input type="date" name="end_date" value="{{ filters.end }}" class="bg-transparent text-sm font-semibold focus:outline-none text-slate-700">
                    </div>
                    <div class="flex items-center gap-2 bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 flex-1 min-w-[160px]">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" class="text-slate-400"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
                        <input type="text" name="search" value="{{ filters.search }}" placeholder="Search description..." class="bg-transparent text-sm font-semibold focus:outline-none text-slate-700 w-full placeholder-slate-300">
                    </div>
                    <button type="submit" class="bg-indigo-600 text-white px-4 py-2 rounded-lg text-sm font-bold hover:bg-indigo-700">Filter</button>
                </form>

                <a href="/export?start_date={{ filters.start }}&end_date={{ filters.end }}&search={{ filters.search }}" class="flex items-center gap-2 bg-emerald-600 text-white px-5 py-2.5 rounded-lg text-sm font-bold hover:bg-emerald-700 shadow-md transition-all">
                    Export Excel
                </a>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                <div class="glass-panel p-5 border-l-4 border-emerald-500">
                    <p class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Total Inflow</p>
                    <p class="text-2xl font-bold text-slate-900 currency">{{ kpis.inflow }}</p>
                </div>
                <div class="glass-panel p-5 border-l-4 border-rose-500">
                    <p class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Total Outflow</p>
                    <p class="text-2xl font-bold text-slate-900 currency">{{ kpis.outflow }}</p>
                </div>
                <div class="glass-panel p-5 border-l-4 border-indigo-500">
                    <p class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Net Movement</p>
                    <p class="text-2xl font-bold currency {{ 'text-emerald-600' if kpis.net_raw > 0 else 'text-rose-600' }}">{{ kpis.net }}</p>
                </div>
                <div class="glass-panel p-5 border-l-4 border-amber-500">
                    <p class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Closing Balance</p>
                    <p class="text-2xl font-bold text-slate-900 currency">{{ kpis.balance }}</p>
                </div>
            </div>

            <div class="glass-panel overflow-hidden">
                <div class="overflow-x-auto">
                    <table class="w-full text-left border-collapse">
                        <thead>
                            <tr class="bg-slate-50 border-b border-slate-200">
                                <th class="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Date</th>
                                <th class="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Description</th>
                                <th class="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Extracted Notes</th>
                                <th class="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider text-right">Debit (GHS)</th>
                                <th class="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider text-right">Credit (GHS)</th>
                                <th class="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider text-right">Balance (GHS)</th>
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-100">
                            {% for row in transactions %}
                            <tr class="hover:bg-slate-50 transition-colors">
                                <td class="px-6 py-3 text-sm font-mono text-slate-600 whitespace-nowrap">{{ row['Booking Date'] }}</td>
                                <td class="px-6 py-3 text-sm text-slate-700 font-medium">{{ row['Description'] }}</td>
                                <td class="px-6 py-3 text-sm text-slate-500 italic">{{ row.get('Extracted Notes', '') }}</td>
                                <td class="px-6 py-3 text-sm font-bold text-right {{ 'text-rose-600' if row['Debit'] != 0 else 'text-slate-200' }}">
                                    {{ "{:,.2f}".format(row['Debit']) if row['Debit'] != 0 else '-' }}
                                </td>
                                <td class="px-6 py-3 text-sm font-bold text-right {{ 'text-emerald-600' if row['Credit'] != 0 else 'text-slate-200' }}">
                                    {{ "{:,.2f}".format(row['Credit']) if row['Credit'] != 0 else '-' }}
                                </td>
                                <td class="px-6 py-3 text-sm font-mono font-bold text-right text-slate-900">
                                    {{ "{:,.2f}".format(row['Balance']) }}
                                </td>
                            </tr>
                            {% endfor %}
                            {% if not transactions %}
                            <tr>
                                <td colspan="6" class="px-6 py-8 text-center text-slate-400 italic">No transactions found for this date range.</td>
                            </tr>
                            {% endif %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        {% endif %}
    </main>
</body>
</html>
"""

# ==========================================
# SHARED UTILITIES
# ==========================================
def clean_money(val):
    """Parse a currency string into a float, handling commas and parenthetical negatives."""
    if not val or str(val).strip() == '':
        return 0.0
    try:
        clean = str(val).replace('"', '').replace(',', '').replace('GH₵', '').strip()
        if clean.startswith('(') and clean.endswith(')'):
            clean = '-' + clean[1:-1]
        return float(clean)
    except ValueError:
        logger.warning("Could not parse monetary value: %r", val)
        return 0.0

EMPTY_DF = lambda: pd.DataFrame(columns=['Booking Date', 'Description', 'Extracted Notes', 'Debit', 'Credit', 'Balance'])

# ==========================================
# 2A. DOCX PARSER ENGINE
# ==========================================
class DocxBankParser:
    def clean_money(self, val):
        return clean_money(val)

    def parse(self, file_stream):
        doc = Document(file_stream)
        data = []
        current_row = None
        extra_desc = []
        
        for table in doc.tables:
            for row in table.rows:
                cells = [c.text.strip().replace('\n', ' ') for c in row.cells]
                if not cells or "CURRENCY :" in cells[0] or "Booking Date" in cells[0]:
                    continue
                    
                booking_date = cells[0]
                reference = cells[1] if len(cells) > 1 else ""
                
                # Handle Starting Balance line
                if "Balance at" in reference:
                    closing_bal = self.clean_money(cells[-1])
                    data.append({
                        'Booking Date': None, 
                        'Description': 'Balance at Period Start', 
                        'Debit': 0.0, 
                        'Credit': 0.0, 
                        'Balance': closing_bal, 
                        'Extracted Notes': ''
                    })
                    continue
                    
                # Is it a main transaction row?
                if re.match(r"^\d{2}\s[A-Z]{3}\s\d{2}$", booking_date):
                    if current_row:
                        clean_extra = []
                        for x in extra_desc:
                            if x and x != current_row['Description'] and x not in clean_extra:
                                clean_extra.append(x)
                        current_row['Extracted Notes'] = " | ".join(clean_extra)
                        data.append(current_row)
                    
                    # Parse dynamic columns based on merged layout
                    if len(cells) == 9:
                        r = {'Booking Date': cells[0], 'Description': cells[4], 
                             'Debit': self.clean_money(cells[6]), 'Credit': self.clean_money(cells[7]), 
                             'Balance': self.clean_money(cells[8]), 'Extracted Notes': ''}
                    elif len(cells) == 7:
                        r = {'Booking Date': cells[0], 'Description': cells[2], 
                             'Debit': self.clean_money(cells[4]), 'Credit': self.clean_money(cells[5]), 
                             'Balance': self.clean_money(cells[6]), 'Extracted Notes': ''}
                    elif len(cells) == 6:
                        desc = cells[2].lower()
                        amt = self.clean_money(cells[4])
                        if "deposit" in desc or "transfer in" in desc or "swift" in desc:
                            r = {'Booking Date': cells[0], 'Description': cells[2], 
                                 'Debit': 0.0, 'Credit': amt, 'Balance': self.clean_money(cells[5]), 'Extracted Notes': ''}
                        else:
                            r = {'Booking Date': cells[0], 'Description': cells[2], 
                                 'Debit': amt, 'Credit': 0.0, 'Balance': self.clean_money(cells[5]), 'Extracted Notes': ''}
                    else:
                        r = {'Booking Date': cells[0], 'Description': cells[2] if len(cells)>2 else "", 
                             'Debit': 0.0, 'Credit': 0.0, 'Balance': self.clean_money(cells[-1]), 'Extracted Notes': ''}
                        
                    current_row = r
                    extra_desc = []
                    
                # Fragmented rows that need to be grouped
                elif current_row and not booking_date:
                    desc = ""
                    if len(cells) == 9: desc = cells[4]
                    elif len(cells) >= 3: desc = cells[2]
                    
                    if desc and ": Chq No -" not in desc:
                        extra_desc.append(desc)

        if current_row:
            clean_extra = []
            for x in extra_desc:
                if x and x != current_row['Description'] and x not in clean_extra:
                    clean_extra.append(x)
            current_row['Extracted Notes'] = " | ".join(clean_extra)
            data.append(current_row)

        if not data:
            return pd.DataFrame(columns=['Booking Date', 'Description', 'Extracted Notes', 'Debit', 'Credit', 'Balance'])
            
        df = pd.DataFrame(data)
        return df

# ==========================================
# 2B. CSV PARSER ENGINE (ORIGINAL)
# ==========================================
class BankParser:
    def __init__(self):
        self.date_pattern = re.compile(r'\b\d{1,2}[\s-](?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[\s-]\d{2,4}\b', re.IGNORECASE)

    def clean_money(self, val):
        return clean_money(val)

    def parse(self, file_content):
        content_str = file_content.decode('utf-8', errors='ignore')
        lines = content_str.splitlines()
        transactions = []
        current_block = []

        for line in lines:
            match = self.date_pattern.search(line)
            if match and match.start() < 5:
                if current_block:
                    transactions.append(self._process_block(current_block))
                current_block = [line]
            else:
                if current_block:
                    current_block.append(line)
        
        if current_block:
            transactions.append(self._process_block(current_block))

        if not transactions:
            df = pd.DataFrame(columns=['Booking Date', 'Description', 'Extracted Notes', 'Debit', 'Credit', 'Balance'])
        else:
            df = pd.DataFrame(transactions)
        return df

    def _process_block(self, lines):
        first_line = lines[0]
        reader = csv.reader([first_line])
        row = list(next(reader))
        
        while row and row[-1].strip() == '':
            row.pop()

        balance = credit = debit = 0.0
        
        if len(row) >= 1: balance = self.clean_money(row[-1])
        if len(row) >= 2: credit = self.clean_money(row[-2])
        if len(row) >= 3: debit = self.clean_money(row[-3])

        full_text = " ".join(lines)
        date_match = self.date_pattern.search(first_line)
        date_str = date_match.group(0) if date_match else ""
        
        desc = full_text.replace(date_str, '', 1) 
        desc = re.sub(r',+', ' ', desc)
        desc = re.sub(r'\s+', ' ', desc).strip()
        
        return {
            'Booking Date': date_str,
            'Description': desc[:200],
            'Extracted Notes': '',
            'Debit': debit,
            'Credit': credit,
            'Balance': balance
        }

# ==========================================
# 2C. PDF PARSER ENGINE
# ==========================================
class PdfBankParser:
    """
    Parses bank statement PDFs exported from Universal Merchant Bank (and similar
    multi-table layouts). Each transaction is its own pdfplumber table with 9 columns:
      [Booking Date, Reference, Acct#, Acct Name, Description, Value Date, Debit, Credit, Balance]
    Continuation rows (notes/cheque info) have an empty first cell and text in col[4].
    """

    DATE_RE = re.compile(r'^\d{2}\s[A-Z]{3}\s\d{2}$')
    FOOTER_KEYS = ('Total Debits', 'Total Credits', 'Closing Balan',
                   'Available Bala', 'Uncleared', 'Booking Date')

    def parse(self, file_stream):
        try:
            records      = []
            extra_notes  = []
            current      = None

            with pdfplumber.open(file_stream) as pdf:
                for page in pdf.pages:
                    for table in page.extract_tables():
                        for row in table:
                            if row is None:
                                continue

                            cells = [str(c).replace('\n', ' ').strip() if c else '' for c in row]
                            if not any(cells):
                                continue

                            col0 = cells[0]

                            # Skip header / footer rows
                            if any(k in col0 for k in self.FOOTER_KEYS):
                                continue

                            # Balance at Period Start
                            if 'Balance at' in ' '.join(cells):
                                if current:
                                    current['Extracted Notes'] = ' | '.join(x for x in extra_notes if x)
                                    records.append(current)
                                    current = None
                                    extra_notes = []
                                records.append({
                                    'Booking Date': None,
                                    'Reference': '',
                                    'Description': 'Balance at Period Start',
                                    'Extracted Notes': '',
                                    'Debit': 0.0, 'Credit': 0.0,
                                    'Balance': clean_money(cells[-1])
                                })
                                continue

                            # Main transaction row — col0 matches date pattern
                            if self.DATE_RE.match(col0):
                                if current:
                                    current['Extracted Notes'] = ' | '.join(x for x in extra_notes if x)
                                    records.append(current)

                                # Col layout: 0=date 1=ref 2=acct# 3=name 4=desc 5=valdate 6=debit 7=credit 8=balance
                                debit  = clean_money(cells[6]) if len(cells) > 6 else 0.0
                                credit = clean_money(cells[7]) if len(cells) > 7 else 0.0
                                bal    = clean_money(cells[8]) if len(cells) > 8 else 0.0
                                desc   = re.sub(r'\s+', ' ', cells[4]).strip() if len(cells) > 4 else ''
                                ref    = re.sub(r'\s+', ' ', cells[1]).strip() if len(cells) > 1 else ''

                                current = {
                                    'Booking Date': col0,
                                    'Reference': ref,
                                    'Description': desc,
                                    'Extracted Notes': '',
                                    'Debit': debit,
                                    'Credit': credit,
                                    'Balance': bal
                                }
                                extra_notes = []

                            # Continuation / notes row — empty col0, notes in col4
                            elif col0 == '' and current and len(cells) > 4:
                                note = re.sub(r'\s+', ' ', cells[4]).strip()
                                skip = (
                                    not note
                                    or note == current['Description']
                                    or ': Chq No' in note
                                    or 'Debit Cheque' in note
                                    or re.match(r'^[\d,]+\.\d{2}$', note)
                                )
                                if not skip:
                                    extra_notes.append(note)

            # Flush last transaction
            if current:
                current['Extracted Notes'] = ' | '.join(x for x in extra_notes if x)
                records.append(current)

            if not records:
                return EMPTY_DF()

            df = pd.DataFrame(records)
            return df

        except Exception as e:
            logger.error("PDF parsing failed: %s", e)
            return EMPTY_DF()

# ==========================================
# 3. FLASK SERVER
# ==========================================
import secrets
app = Flask(__name__)
app.secret_key = secrets.token_hex(32)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20 MB upload limit
DB = {'df': None, 'filename': None}

@app.route('/', methods=['GET', 'POST'])
def index():
    # Handle the Reset functionality
    if request.args.get('reset') == '1':
        DB['df'] = None
        DB['filename'] = None
        return redirect(url_for('index'))

    if request.method == 'POST':
        if 'file' not in request.files: return redirect(request.url)
        file = request.files['file']
        if file.filename == '': return redirect(request.url)
        
        if file:
            filename_lower = file.filename.lower()
            try:
                # Use Docx Engine
                if filename_lower.endswith('.docx'):
                    file_stream = io.BytesIO(file.read())
                    parser = DocxBankParser()
                    df = parser.parse(file_stream)

                # Use PDF Engine
                elif filename_lower.endswith('.pdf'):
                    file_stream = io.BytesIO(file.read())
                    parser = PdfBankParser()
                    df = parser.parse(file_stream)

                # Use CSV Engine
                elif filename_lower.endswith('.csv'):
                    parser = BankParser()
                    df = parser.parse(file.read())

                else:
                    return render_template_string(
                        HTML_TEMPLATE, filename=None,
                        error="Unsupported file type. Please upload a .csv, .docx, or .pdf file."
                    )

                if df.empty:
                    return render_template_string(
                        HTML_TEMPLATE, filename=None,
                        error="No transactions could be extracted from this file. "
                              "Check that it is a valid bank statement."
                    )

                # Ensure proper datetime conversions for all formats
                df['Booking Date'] = pd.to_datetime(df['Booking Date'], format='mixed', errors='coerce')
                df = df.sort_values('Booking Date', na_position='first')

                DB['df'] = df
                DB['filename'] = file.filename
                return redirect(url_for('index'))

            except Exception as e:
                logger.error("Upload processing error: %s", e)
                return render_template_string(
                    HTML_TEMPLATE, filename=None,
                    error=f"Failed to parse file: {e}"
                )

    if DB['df'] is not None:
        df = DB['df'].copy()
        
        # Apply Filters
        start = request.args.get('start_date')
        end = request.args.get('end_date')
        search = request.args.get('search', '').strip()
        if start: df = df[df['Booking Date'] >= pd.to_datetime(start)]
        if end: df = df[df['Booking Date'] <= pd.to_datetime(end)]
        if search:
            mask = (
                df['Description'].str.contains(search, case=False, na=False) |
                df['Extracted Notes'].str.contains(search, case=False, na=False)
            )
            df = df[mask]

        # Calculate KPIs
        kpis = {
            'inflow': f"GH₵ {df['Credit'].sum():,.2f}",
            'outflow': f"GH₵ {df['Debit'].sum():,.2f}",
            'net': f"GH₵ {df['Credit'].sum() - df['Debit'].sum():,.2f}",
            'net_raw': df['Credit'].sum() - df['Debit'].sum(),
            'balance': f"GH₵ {df['Balance'].iloc[-1]:,.2f}" if not df.empty else "0.00"
        }
        
        # Format date safely for HTML rendering, leave missing dates blank (e.g. for Starting Balances)
        df['Booking Date'] = df['Booking Date'].dt.strftime('%d %b %Y').fillna('-')
        
        return render_template_string(
            HTML_TEMPLATE, filename=DB['filename'],
            transactions=df.to_dict('records'), kpis=kpis,
            filters={'start': start or '', 'end': end or '', 'search': search},
            error=None
        )
    
    return render_template_string(HTML_TEMPLATE, filename=None, error=None)

@app.route('/export')
def export():
    if DB['df'] is None: return redirect(url_for('index'))
    
    df = DB['df'].copy()
    start = request.args.get('start_date')
    end = request.args.get('end_date')
    search = request.args.get('search', '').strip()
    if start: df = df[df['Booking Date'] >= pd.to_datetime(start)]
    if end: df = df[df['Booking Date'] <= pd.to_datetime(end)]
    if search:
        mask = (
            df['Description'].str.contains(search, case=False, na=False) |
            df['Extracted Notes'].str.contains(search, case=False, na=False)
        )
        df = df[mask]

    # Pre-format export string dates
    df['Booking Date'] = df['Booking Date'].dt.strftime('%d %b %Y').fillna('-')

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # --- Sheet 1: Transaction Data ---
        df.to_excel(writer, index=False, sheet_name='Audit Data')
        workbook = writer.book
        worksheet = writer.sheets['Audit Data']
        money_fmt = workbook.add_format({'num_format': '"GH₵" #,##0.00', 'align': 'right'})
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1e293b', 'font_color': '#ffffff', 'border': 1})
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, header_fmt)
        worksheet.set_column('D:F', 18, money_fmt)
        worksheet.set_column('A:A', 15)
        worksheet.set_column('B:C', 45)
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

        # --- Sheet 2: Summary ---
        numeric_df = DB['df'].copy()
        if start: numeric_df = numeric_df[numeric_df['Booking Date'] >= pd.to_datetime(start)]
        if end: numeric_df = numeric_df[numeric_df['Booking Date'] <= pd.to_datetime(end)]
        total_inflow = numeric_df['Credit'].sum()
        total_outflow = numeric_df['Debit'].sum()
        net = total_inflow - total_outflow
        closing_bal = numeric_df['Balance'].iloc[-1] if not numeric_df.empty else 0

        summary_ws = workbook.add_worksheet('Summary')
        title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#059669'})
        label_fmt = workbook.add_format({'bold': True, 'bg_color': '#f8fafc', 'border': 1})
        value_fmt = workbook.add_format({'num_format': '"GH₵" #,##0.00', 'border': 1, 'align': 'right'})
        summary_ws.write('A1', 'KAKOS AUDIT — SUMMARY', title_fmt)
        summary_ws.write('A2', f'File: {DB["filename"]}')
        summary_ws.write('A3', f'Filters: {start or "All"} to {end or "All"} | Search: "{search}"' if search else f'Filters: {start or "All"} to {end or "All"}')
        rows = [
            ('Total Inflow (Credits)', total_inflow),
            ('Total Outflow (Debits)', total_outflow),
            ('Net Movement', net),
            ('Closing Balance', closing_bal),
            ('Total Transactions', len(numeric_df)),
        ]
        for i, (label, val) in enumerate(rows):
            r = i + 5
            summary_ws.write(r, 0, label, label_fmt)
            if isinstance(val, float):
                summary_ws.write(r, 1, val, value_fmt)
            else:
                summary_ws.write(r, 1, val, workbook.add_format({'border': 1, 'align': 'right'}))
        summary_ws.set_column('A:A', 30)
        summary_ws.set_column('B:B', 20)

    output.seek(0)
    safe_name = DB['filename'].rsplit('.', 1)[0]
    return send_file(output, download_name=f"Cleaned_{safe_name}.xlsx", as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)