import pandas as pd
import re
from rapidfuzz import process, fuzz
import json
import datetime
from collections import Counter

# --- CONFIGURATION ---
REGISTRY_FILE = 'registry_cleaned.xlsx'
INPUT_FILE = 'Compilate Wrike.xlsx'
OUTPUT_HTML = 'raport_master.html'
OUTPUT_EXCEL = 'extrase_pt_conta.xlsx'

# Sheet mappings: { 'Sheet Name': [List of Column Names to check] }
TARGET_SHEETS = {
    "1  MOBILITATE US (export)": ["NR. CONTRACT"],              # Usually Col J
    "1  REG URBANA US (export)": ["Nr. contract"],              # Usually Col S
    "2  PROIECTE US (export)": ["Nr. contract"],                # Usually Col S
    "2  PROIECTE US (CF - STRATEGII ": ["NR. CONTRACT"]  # Merged column
}
# ---------------------

def clean_text(value):
    """Applies the same regex cleaner used in the first step"""
    if pd.isna(value):
        return ""
    # Handle float values that are whole numbers (remove .0)
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value)
    if text.lower() in ['nan', 'none', '']:
        return ""
    pattern = r'[\s\-/]*\d{2}\.\d{2}\.\d{4}.*'
    return re.sub(pattern, '', text).strip()

def clean_project_name(value):
    """Cleans project name, handling NaN values"""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in ['nan', 'none']:
        return ""
    return text

def clean_numeric_code(value):
    """Cleans numeric codes, removing .0 from whole numbers"""
    if pd.isna(value):
        return ""
    # If it's a float that's actually a whole number, convert to int
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if text.lower() in ['nan', 'none']:
        return ""
    # Handle string that ends with .0
    if text.endswith('.0'):
        return text[:-2]
    return text

def find_project_name_col(df):
    """Attempts to find the column containing the Project Name in the user sheet"""
    # First, check for 'Name' column explicitly (Wrike uses this)
    if 'Name' in df.columns:
        return 'Name'
    # Then check for other keywords, excluding 'valoare' and 'cod' columns
    keywords = ['proiect', 'titlu', 'nume', 'title', 'denumire', 'task']
    exclude = ['valoare', 'cod', 'value', 'code']
    for col in df.columns:
        col_lower = str(col).lower()
        if any(k in col_lower for k in keywords) and not any(e in col_lower for e in exclude):
            return col
    return df.columns[1] if len(df.columns) > 1 else df.columns[0]

def sheet_name_to_filename(sheet_name):
    """Convert sheet name to a valid filename."""
    # Remove common suffixes and clean up
    name = sheet_name.replace(" (export)", "").replace("(export)", "")
    # Replace spaces and special chars with underscores
    name = re.sub(r'[^\w\s-]', '', name)  # Remove special chars except dash
    name = re.sub(r'\s+', '_', name.strip())  # Replace spaces with underscore
    name = name.lower()
    return f"raport_{name}.html"

def generate_html(data, stats, all_exports, missing_from_wrike=None, registry_stats=None,
                  output_file=None, title=None):
    """Generates the HTML file matching your exact requested design."""
    if missing_from_wrike is None:
        missing_from_wrike = []
    if registry_stats is None:
        registry_stats = {'total_rows': 0, 'with_contract': 0, 'unique_contracts': 0, 'no_contract': 0, 'duplicates': 0, 'duplicate_details': {}, 'total_unique': 0, 'found_in_wrike': 0, 'not_in_wrike': 0}
    if output_file is None:
        output_file = OUTPUT_HTML
    if title is None:
        title = "Raport Master - Multiple Exporturi"

    is_single_team = len(all_exports) == 1

    total_valid = sum(s['valid'] for s in stats.values())
    total_dup = sum(s['duplicate'] for s in stats.values())
    total_typo = sum(s['typo'] for s in stats.values())
    total_empty = sum(s['empty'] for s in stats.values())
    total_notfound = sum(s['notfound'] for s in stats.values())
    total_all = total_valid + total_dup + total_typo + total_empty + total_notfound
    
    timestamp = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    
    # Pre-render table rows for the master sheet
    master_prob_rows = ""
    master_valid_rows = ""
    
    # Store JS data for switching tabs
    js_stats = {"-1": {"valid": total_valid, "duplicate": total_dup, "typo": total_typo, "empty": total_empty, "notfound": total_notfound}}
    
    for idx, exp_name in enumerate(all_exports):
        js_stats[str(idx)] = stats[exp_name]

    # HTML TEMPLATE (Truncated for brevity, creates exact structure)
    html = f"""<!DOCTYPE html>
<html lang="ro">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2"></script>
    <style>
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{ font-family: 'Segoe UI', sans-serif; background: linear-gradient(135deg, #0f0c29, #302b63, #24243e); min-height: 100vh; color: #eee; }}
        .container {{ max-width: 1400px; margin: 0 auto; padding: 20px; }}
        header {{ background: rgba(255,255,255,0.1); padding: 20px; border-radius: 12px; margin-bottom: 20px; }}
        h1 {{ color: #a855f7; font-size: 1.8em; }}
        .subtitle {{ color: #888; font-size: 0.9em; margin-top: 5px; }}
        .dashboard {{ display: grid; grid-template-columns: 1fr 300px; gap: 20px; margin-bottom: 20px; }}
        .cards {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px; }}
        .card {{ background: rgba(255,255,255,0.05); padding: 20px; border-radius: 12px; text-align: center; border-left: 4px solid; }}
        .card.valid {{ border-color: #27ae60; }} .card.empty {{ border-color: #f39c12; }}
        .card.notfound {{ border-color: #e74c3c; }} .card.typo {{ border-color: #3498db; }} .card.duplicate {{ border-color: #9b59b6; }}
        .card h2 {{ font-size: 2.5em; margin-bottom: 5px; }}
        .card .pct {{ font-size: 0.9em; color: #888; }}
        .card p {{ color: #aaa; font-size: 0.85em; }}
        .chart-container {{ background: rgba(255,255,255,0.05); padding: 20px; border-radius: 12px; }}
        .export-tabs {{ display: flex; gap: 5px; margin-bottom: 0; flex-wrap: wrap; }}
        .export-tab {{ padding: 12px 20px; background: rgba(255,255,255,0.1); border: none; border-radius: 8px 8px 0 0; color: #aaa; cursor: pointer; }}
        .export-tab.active {{ background: linear-gradient(135deg, #a855f7, #6366f1); color: white; }}
        .export-content {{ display: none; background: rgba(255,255,255,0.05); border-radius: 0 12px 12px 12px; padding: 20px; }}
        .export-content.active {{ display: block; }}
        .inner-tabs {{ display: flex; gap: 5px; margin-bottom: 15px; }}
        .inner-tab {{ padding: 8px 16px; background: rgba(255,255,255,0.1); border: none; border-radius: 6px; color: #aaa; cursor: pointer; }}
        .inner-tab.active {{ background: rgba(168,85,247,0.3); color: #a855f7; }}
        .inner-content {{ display: none; }} .inner-content.active {{ display: block; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 0.85em; table-layout: auto; }}
        td:nth-child(2) {{ min-width: 180px; }}
        th {{ background: rgba(168,85,247,0.9); padding: 10px; text-align: left; cursor: pointer; user-select: none; position: sticky; top: 0; z-index: 10; }}
        th:hover {{ background: rgba(168,85,247,0.4); }}
        th.sort-asc::after {{ content: ' ▲'; font-size: 0.7em; }}
        th.sort-desc::after {{ content: ' ▼'; font-size: 0.7em; }}
        td {{ padding: 8px 10px; border-bottom: 1px solid rgba(255,255,255,0.1); }}
        tr:hover {{ background: rgba(255,255,255,0.05); }}
        .status {{ padding: 3px 8px; border-radius: 12px; font-size: 0.75em; white-space: nowrap; display: inline-block; }}
        .status.valid {{ background: rgba(39,174,96,0.3); color: #2ecc71; }}
        .status.empty {{ background: rgba(243,156,18,0.3); color: #f1c40f; }}
        .status.notfound {{ background: rgba(231,76,60,0.3); color: #e74c3c; }}
        .status.typo {{ background: rgba(52,152,219,0.3); color: #3498db; }}
        .status.duplicate {{ background: rgba(155,89,182,0.3); color: #9b59b6; }}
        .table-wrapper {{ max-height: 500px; overflow: auto; }}
        .registry-summary {{ background: rgba(255,255,255,0.05); padding: 20px; border-radius: 12px; margin-bottom: 20px; }}
        .registry-summary h3 {{ color: #a855f7; margin-bottom: 15px; font-size: 1.1em; }}
        .registry-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 15px; }}
        .registry-item {{ background: rgba(255,255,255,0.03); padding: 15px; border-radius: 8px; text-align: center; }}
        .registry-item .number {{ font-size: 1.8em; font-weight: bold; color: #fff; }}
        .registry-item .label {{ font-size: 0.8em; color: #888; margin-top: 5px; }}
        .registry-item.total {{ border-left: 3px solid #a855f7; }}
        .registry-item.found {{ border-left: 3px solid #27ae60; }}
        .registry-item.notfound {{ border-left: 3px solid #e74c3c; }}
        .registry-item.nocontract {{ border-left: 3px solid #f39c12; }}
        .registry-item.duplicate {{ border-left: 3px solid #e74c3c; background: rgba(231, 76, 60, 0.1); }}
        .registry-item.ok {{ border-left: 3px solid #27ae60; }}
        .registry-item.total-unique {{ border-left: 3px solid #3498db; background: rgba(52, 152, 219, 0.1); }}
        .duplicate-warning {{ margin-top: 15px; background: rgba(231, 76, 60, 0.1); border: 1px solid #e74c3c; border-radius: 8px; overflow: hidden; }}
        .duplicate-header {{ padding: 12px 15px; cursor: pointer; display: flex; justify-content: space-between; align-items: center; color: #e74c3c; font-weight: 500; }}
        .duplicate-header:hover {{ background: rgba(231, 76, 60, 0.15); }}
        .toggle-icon {{ transition: transform 0.3s; }}
        .duplicate-details {{ padding: 15px; background: rgba(0,0,0,0.2); }}
        .duplicate-table {{ width: 100%; border-collapse: collapse; font-size: 0.9em; }}
        .duplicate-table th {{ text-align: left; padding: 8px; border-bottom: 1px solid #444; color: #888; }}
        .duplicate-table td {{ padding: 8px; border-bottom: 1px solid #333; color: #ccc; }}
        .duplicate-table tr:last-child td {{ border-bottom: none; }}
        footer {{ text-align: center; padding: 20px; color: #666; font-size: 0.85em; }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>{title}</h1>
            <div class="subtitle">{'' if is_single_team else f'{len(all_exports)} seturi date | '}Total: {total_all} intrari | Generat: {timestamp}</div>
        </header>
        
        <div class="dashboard">
            <div class="cards">
                <div class="card valid"><h2 id="statValid">{total_valid}</h2><div class="pct" id="pctValid">{round(total_valid/total_all*100,1)}%</div><p>Coduri Valide</p></div>
                <div class="card duplicate"><h2 id="statDuplicate">{total_dup}</h2><div class="pct" id="pctDuplicate">{round(total_dup/total_all*100,1)}%</div><p>Duplicate</p></div>
                <div class="card typo"><h2 id="statTypo">{total_typo}</h2><div class="pct" id="pctTypo">{round(total_typo/total_all*100,1)}%</div><p>Posibile Erori</p></div>
                <div class="card empty"><h2 id="statEmpty">{total_empty}</h2><div class="pct" id="pctEmpty">{round(total_empty/total_all*100,1)}%</div><p>Necompletate</p></div>
                <div class="card notfound"><h2 id="statNotfound">{total_notfound}</h2><div class="pct" id="pctNotfound">{round(total_notfound/total_all*100,1)}%</div><p>Necorelate</p></div>
            </div>
            <div class="chart-container"><canvas id="pieChart" width="250" height="250"></canvas></div>
        </div>
"""
    # Only show registry summary in master report (not for individual team reports)
    if not is_single_team:
        # Build duplicate details HTML if there are duplicates
        duplicate_html = ""
        if registry_stats['duplicates'] > 0:
            duplicate_rows = ""
            for contract_code, occurrences in registry_stats['duplicate_details'].items():
                rows_info = " & ".join([f"Row {o['excel_row']} ({o['project'][:30]}{'...' if len(o['project']) > 30 else ''})" for o in occurrences])
                duplicate_rows += f"<tr><td><strong>{contract_code}</strong></td><td>{rows_info}</td></tr>"

            duplicate_html = f"""
            <div class="duplicate-warning">
                <div class="duplicate-header" onclick="toggleDuplicates()">
                    <span>⚠️ Contracte Duplicate - acelasi numar folosit pentru proiecte diferite</span>
                    <span class="toggle-icon" id="dup-toggle">▼</span>
                </div>
                <div class="duplicate-details" id="duplicate-list" style="display: none;">
                    <table class="duplicate-table">
                        <thead><tr><th>Nr. Contract</th><th>Randuri Excel & Proiecte</th></tr></thead>
                        <tbody>{duplicate_rows}</tbody>
                    </table>
                </div>
            </div>"""

        html += f"""
        <div class="registry-summary">
            <h3>Sumar Registru Contracte</h3>
            <div class="registry-grid">
                <div class="registry-item total">
                    <div class="number">{registry_stats['total_rows']}</div>
                    <div class="label">Total Randuri</div>
                </div>
                <div class="registry-item found">
                    <div class="number">{registry_stats['with_contract']}</div>
                    <div class="label">Cu Nr. Contract</div>
                </div>
                <div class="registry-item nocontract">
                    <div class="number">{registry_stats['no_contract']}</div>
                    <div class="label">Fara Nr. Contract</div>
                </div>
                <div class="registry-item {'duplicate' if registry_stats['duplicates'] > 0 else 'ok'}">
                    <div class="number">{registry_stats['duplicates']}{'⚠️' if registry_stats['duplicates'] > 0 else ''}</div>
                    <div class="label">Contracte Duplicate</div>
                </div>
                <div class="registry-item total-unique">
                    <div class="number">{registry_stats['total_unique']}</div>
                    <div class="label">Total Proiecte Unice</div>
                </div>
            </div>
            {duplicate_html}
        </div>
"""
    # Only show export tabs if there are multiple sheets (master view)
    if not is_single_team:
        html += """        <div class="export-tabs">
            <button class="export-tab active" onclick="showExport(-1, this)">📊 MASTER (Toate)</button>
"""
        for idx, exp_name in enumerate(all_exports):
            html += f'            <button class="export-tab" onclick="showExport({idx}, this)">{exp_name}</button>\n'
        html += '        </div>\n\n'

    # Content loop
    full_dataset = []
    for sheet_name, sheet_data in data.items():
        full_dataset.extend(sheet_data)

    def render_sheet_content(idx, dataset, is_master=False):
        probs = [d for d in dataset if d['status'] != 'VALID']
        valids = [d for d in dataset if d['status'] == 'VALID']

        prefix = 'Master' if is_master else f'Exp{idx}'

        tab_html = f"""        <div id="export{idx}" class="export-content {'active' if is_master else ''}">
            <div class="inner-tabs">
                <button class="inner-tab active" onclick="showInner({idx}, 'prob', this)">Probleme ({len(probs)})</button>
                <button class="inner-tab" onclick="showInner({idx}, 'valid', this)">Valide ({len(valids)})</button>
            </div>
            <div id="export{idx}prob" class="inner-content active">
                <div class="table-wrapper">
                    <table id="table{prefix}Prob">
                        <thead><tr><th onclick="sortTable(this, 0)">Sursa</th><th onclick="sortTable(this, 1)">Status</th><th onclick="sortTable(this, 2)">Linie</th><th onclick="sortTable(this, 3)">NR CONTRACT</th><th onclick="sortTable(this, 4)">Sugestie</th><th onclick="sortTable(this, 5)">Proiect</th><th onclick="sortTable(this, 6)">COD PROIECT</th></tr></thead>
                        <tbody>
"""
        for r in probs:
            status_class = r['status'].lower().replace(' ', '')
            tab_html += f"<tr><td>{r['source']}</td><td><span class='status {status_class}'>{r['status_text']}</span></td><td>{r['line']}</td><td>{r['code']}</td><td>{r['suggest']}</td><td>{r['project']}</td><td>{r['cod_proiect']}</td></tr>\n"
        
        tab_html += f"""                        </tbody>
                    </table>
                </div>
            </div>
            <div id="export{idx}valid" class="inner-content">
                <div class="table-wrapper">
                    <table id="table{prefix}Valid">
                        <thead><tr><th onclick="sortTable(this, 0)">Sursa</th><th onclick="sortTable(this, 1)">Status</th><th onclick="sortTable(this, 2)">Linie</th><th onclick="sortTable(this, 3)">NR CONTRACT</th><th onclick="sortTable(this, 4)">Proiect</th><th onclick="sortTable(this, 5)">COD PROIECT</th></tr></thead>
                        <tbody>
"""
        for r in valids:
            tab_html += f"<tr><td>{r['source']}</td><td><span class='status valid'>VALID</span></td><td>{r['line']}</td><td>{r['code']}</td><td>{r['project']}</td><td>{r['cod_proiect']}</td></tr>\n"

        tab_html += "                        </tbody></table></div></div>\n"
        tab_html += "        </div>\n"
        return tab_html

    # For single-team mode, just render that team's data directly
    # For master mode, render both the combined view and individual sheets
    if is_single_team:
        # Single team - render just the team data (use idx=-1 to make it active)
        html += render_sheet_content(-1, full_dataset, is_master=True)
    else:
        # Render Master (combined view)
        html += render_sheet_content(-1, full_dataset, is_master=True)
        # Render Individual Sheets
        for idx, exp_name in enumerate(all_exports):
            html += render_sheet_content(idx, data[exp_name])

    # Footer and JS
    html += f"""
        <footer>Validare Contracte v4.0 | {len(all_exports)} analize procesate</footer>
    </div>
    <script>
        Chart.register(ChartDataLabels);
        const exportStats = {json.dumps(js_stats)};
        const ctx = document.getElementById('pieChart');
        const pieChart = new Chart(ctx, {{
            type: 'doughnut',
            data: {{
                labels: ['Valide', 'Duplicate', 'Erori', 'Necompletate', 'Necorelate'],
                datasets: [{{ data: [{total_valid}, {total_dup}, {total_typo}, {total_empty}, {total_notfound}], backgroundColor: ['#27ae60', '#9b59b6', '#3498db', '#f39c12', '#e74c3c'], borderWidth: 0 }}]
            }},
            options: {{ responsive: true, plugins: {{ legend: {{ position: 'bottom', labels: {{ color: '#aaa' }} }} }} }}
        }});
        
        function updateDashboard(idx) {{
            const s = exportStats[idx];
            const total = s.valid + s.duplicate + s.typo + s.empty + s.notfound;
            document.getElementById('statValid').textContent = s.valid;
            document.getElementById('statDuplicate').textContent = s.duplicate;
            document.getElementById('statTypo').textContent = s.typo;
            document.getElementById('statEmpty').textContent = s.empty;
            document.getElementById('statNotfound').textContent = s.notfound;
            pieChart.data.datasets[0].data = [s.valid, s.duplicate, s.typo, s.empty, s.notfound];
            pieChart.update();
        }}
        
        function showExport(idx, btn) {{
            document.querySelectorAll('.export-content').forEach(ec => ec.classList.remove('active'));
            document.querySelectorAll('.export-tab').forEach(t => t.classList.remove('active'));
            document.getElementById('export' + idx).classList.add('active');
            btn.classList.add('active');
            updateDashboard(idx);
        }}
        
        function showInner(expIdx, type, btn) {{
            const container = document.getElementById('export' + expIdx);
            container.querySelectorAll('.inner-content').forEach(ic => ic.classList.remove('active'));
            container.querySelectorAll('.inner-tab').forEach(t => t.classList.remove('active'));
            document.getElementById('export' + expIdx + type).classList.add('active');
            btn.classList.add('active');
        }}

        function sortTable(th, colIndex) {{
            const table = th.closest('table');
            const tbody = table.querySelector('tbody');
            const rows = Array.from(tbody.querySelectorAll('tr'));
            const headers = table.querySelectorAll('th');

            // Determine sort direction
            const isAsc = th.classList.contains('sort-asc');
            headers.forEach(h => h.classList.remove('sort-asc', 'sort-desc'));
            th.classList.add(isAsc ? 'sort-desc' : 'sort-asc');

            // Sort rows
            rows.sort((a, b) => {{
                let aVal = a.cells[colIndex].textContent.trim();
                let bVal = b.cells[colIndex].textContent.trim();

                // Try numeric sort for Linie column
                const aNum = parseFloat(aVal);
                const bNum = parseFloat(bVal);
                if (!isNaN(aNum) && !isNaN(bNum)) {{
                    return isAsc ? bNum - aNum : aNum - bNum;
                }}

                // String sort
                return isAsc ? bVal.localeCompare(aVal) : aVal.localeCompare(bVal);
            }});

            // Re-append sorted rows
            rows.forEach(row => tbody.appendChild(row));
        }}

        function toggleDuplicates() {{
            const list = document.getElementById('duplicate-list');
            const icon = document.getElementById('dup-toggle');
            if (list.style.display === 'none') {{
                list.style.display = 'block';
                icon.textContent = '▲';
            }} else {{
                list.style.display = 'none';
                icon.textContent = '▼';
            }}
        }}
    </script>
</body>
</html>
"""
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Report generated successfully: {output_file}")


def export_valid_to_excel(codes_found, registry_full, wrike_names, output_file):
    """Exports valid entries to Excel for accounting."""
    export_data = []
    for code in sorted(codes_found):
        if code in registry_full:
            row = registry_full[code]
            export_data.append({
                'Cod proiect': row['cod_proiect'],
                'Numar Contract': row['numar_contract'],
                'Nume Proiect (Registru)': row['nume_proiect'],
                'Nume Proiect (Wrike)': wrike_names.get(code, ''),
                'Data Contract': row['data_contract']
            })

    if export_data:
        export_df = pd.DataFrame(export_data)
        export_df.to_excel(output_file, index=False)
        print(f"Exported {len(export_data)} valid entries to {output_file}")
    else:
        print("No valid entries to export")


def main():
    print("Loading Registry File...")
    reg_df = pd.read_excel(REGISTRY_FILE)
    
    # Extract Valid Codes and Projects
    reg_col = 'NUMAR CONTRACT'
    proj_col = reg_df.columns[4] # Column E is index 4
    
    valid_codes = {}
    missing_in_registry = []
    no_contract_entries = []  # Registry entries without contract numbers

    # Track contract occurrences for duplicate detection
    contract_occurrences = {}  # {contract_code: [(excel_row, project_name, cod_proiect), ...]}
    total_rows_with_contract = 0

    for index, row in reg_df.iterrows():
        code = clean_text(row[reg_col])
        proj = clean_project_name(row[proj_col])
        cod_proiect = clean_numeric_code(row['COD PROIECT']) if 'COD PROIECT' in reg_df.columns else ""
        excel_row = index + 2  # +1 for header, +1 for 0-based index

        if code:
            total_rows_with_contract += 1
            valid_codes[code] = proj
            # Track for duplicate detection
            if code not in contract_occurrences:
                contract_occurrences[code] = []
            contract_occurrences[code].append({
                'excel_row': excel_row,
                'project': proj,
                'cod_proiect': cod_proiect if cod_proiect else "-"
            })
        else:
            missing_in_registry.append(proj) # Registry has project but no contract code
            # Also track these for the report
            if proj:  # Only if there's a project name
                no_contract_entries.append({
                    "code": "-",
                    "project": proj,
                    "cod_proiect": cod_proiect if cod_proiect else "-"
                })

    # Find duplicate contracts (same contract number used for different projects)
    duplicate_contracts = {code: occurrences for code, occurrences in contract_occurrences.items() if len(occurrences) > 1}
    duplicate_count = len(duplicate_contracts)
    print(f"Found {duplicate_count} duplicate contract numbers affecting {sum(len(v) for v in duplicate_contracts.values())} rows")

    valid_codes_list = list(valid_codes.keys())

    # Build COD PROIECT mapping from the same registry sheet
    cod_proiect_map = {}
    if 'COD PROIECT' in reg_df.columns:
        for index, row in reg_df.iterrows():
            code = clean_text(row[reg_col])
            cod_proiect = clean_numeric_code(row['COD PROIECT'])
            if code and cod_proiect:
                cod_proiect_map[code] = cod_proiect
        print(f"Loaded {len(cod_proiect_map)} COD PROIECT entries")
    else:
        print("Warning: COD PROIECT column not found in registry")

    # Build full registry lookup for export (includes DATA CONTRACT)
    registry_full = {}
    for _, row in reg_df.iterrows():
        code = clean_text(row[reg_col])
        if code:
            registry_full[code] = {
                'cod_proiect': clean_numeric_code(row['COD PROIECT']) if 'COD PROIECT' in reg_df.columns else "",
                'numar_contract': row['NUMAR CONTRACT'],
                'nume_proiect': row['NUME PROIECT'],
                'data_contract': row['DATA CONTRACT']
            }

    print("Loading Wrike Input File...")
    wrike_sheets = pd.read_excel(INPUT_FILE, sheet_name=None)
    
    report_data = {}
    report_stats = {}
    all_exports = []
    codes_found_in_wrike = set()  # Track which registry codes are found in Wrike
    wrike_names = {}  # Track Wrike project names: {contract_code: wrike_project_name}

    for sheet_name, cols_to_check in TARGET_SHEETS.items():
        if sheet_name not in wrike_sheets:
            print(f"Warning: Sheet '{sheet_name}' not found in input file.")
            continue
            
        df = wrike_sheets[sheet_name]
        proj_name_col = find_project_name_col(df)

        for col_name in cols_to_check:
            # Special handling for "Nr. contract.1" to show differently in report
            display_name = sheet_name
            if len(cols_to_check) > 1 and "1" in col_name:
                display_name += " (Col M)"
            else:
                display_name += " (Col L/S/J)"
                
            all_exports.append(display_name)
            report_data[display_name] = []
            stats = {'valid': 0, 'duplicate': 0, 'typo': 0, 'empty': 0, 'notfound': 0}
            
            # Count for duplicates
            if col_name in df.columns:
                cleaned_column = df[col_name].apply(clean_text)
                value_counts = Counter([x for x in cleaned_column if x != ""])
                
                for idx, row in df.iterrows():
                    raw_val = row[col_name]
                    clean_val = clean_text(raw_val)
                    proj_val = clean_project_name(row[proj_name_col]) if proj_name_col in df.columns else ""
                    
                    status = ""
                    status_text = ""
                    suggest = ""
                    cod_proiect_val = "-"

                    # 1. Check for Empty
                    if clean_val == "":
                        status = "empty"
                        status_text = "NR CONTRACT NECOMPLETAT IN WRIKE"
                        stats['empty'] += 1
                        cod_proiect_val = "-"

                    # 2. Check for Exact Match
                    elif clean_val in valid_codes:
                        codes_found_in_wrike.add(clean_val)  # Track this code as found
                        if clean_val not in wrike_names:
                            wrike_names[clean_val] = proj_val  # Store Wrike project name (first occurrence)
                        if value_counts[clean_val] > 1:
                            status = "duplicate"
                            status_text = "COD DUPLICAT"
                            stats['duplicate'] += 1
                            # Keep Wrike name for duplicates
                        else:
                            status = "VALID"
                            status_text = "VALID"
                            stats['valid'] += 1
                            proj_val = valid_codes[clean_val] # Use Registry name for valid
                        cod_proiect_val = cod_proiect_map.get(clean_val, "-")

                    # 3. Check for Typos / Not Found
                    else:
                        match = process.extractOne(clean_val, valid_codes_list, scorer=fuzz.ratio)
                        if match and match[1] >= 85: # 85% similarity threshold
                            codes_found_in_wrike.add(match[0])  # Track suggested code as found
                            if match[0] not in wrike_names:
                                wrike_names[match[0]] = proj_val  # Store Wrike project name (first occurrence)
                            status = "typo"
                            status_text = "POSIBILA EROARE DE TASTARE"
                            suggest = match[0]
                            stats['typo'] += 1
                            cod_proiect_val = cod_proiect_map.get(match[0], "-")
                        else:
                            status = "notfound"
                            status_text = "NR CONTRACT DIN WRIKE NECORELAT CU REGISTRUL"
                            stats['notfound'] += 1
                            cod_proiect_val = "-"

                    report_data[display_name].append({
                        "source": display_name,
                        "line": idx + 2, # +2 for Excel row number (0-idx + header)
                        "code": raw_val if not pd.isna(raw_val) else "",
                        "status": status,
                        "status_text": status_text,
                        "suggest": suggest,
                        "project": proj_val,
                        "cod_proiect": cod_proiect_val
                    })
            else:
                print(f"Warning: Column '{col_name}' not found in sheet '{sheet_name}'.")

            report_stats[display_name] = stats

    # Build list of registry entries not found in any Wrike sheet
    missing_from_wrike = []
    for code, proj_name in valid_codes.items():
        if code not in codes_found_in_wrike:
            missing_from_wrike.append({
                "code": code,
                "project": proj_name,
                "cod_proiect": cod_proiect_map.get(code, "-")
            })

    # Add registry entries that have no contract number
    missing_from_wrike.extend(no_contract_entries)
    print(f"Found {len(missing_from_wrike)} registry entries not in Wrike ({len(no_contract_entries)} without contract number)")

    # Build registry stats for summary
    registry_stats = {
        'total_rows': len(reg_df),
        'with_contract': total_rows_with_contract,
        'unique_contracts': len(valid_codes),
        'no_contract': len(no_contract_entries),
        'duplicates': duplicate_count,
        'duplicate_details': duplicate_contracts,
        'total_unique': len(valid_codes) + len(no_contract_entries),
        'found_in_wrike': len(codes_found_in_wrike),
        'not_in_wrike': len(valid_codes) - len(codes_found_in_wrike)
    }

    # Generate Master Report (all sheets combined)
    generate_html(report_data, report_stats, all_exports, missing_from_wrike, registry_stats,
                  output_file=OUTPUT_HTML, title="Raport Master - Multiple Exporturi")

    # Generate Individual Team Reports (filename derived from sheet name)
    for sheet_name in all_exports:
        if sheet_name in report_data:
            team_data = {sheet_name: report_data[sheet_name]}
            team_stats = {sheet_name: report_stats[sheet_name]}
            team_name = sheet_name.replace(" (export)", "").strip()
            team_output_file = sheet_name_to_filename(sheet_name)
            generate_html(team_data, team_stats, [sheet_name],
                          output_file=team_output_file,
                          title=f"Raport Echipa - {team_name}")

    # Export valid entries to Excel for accounting
    export_valid_to_excel(codes_found_in_wrike, registry_full, wrike_names, OUTPUT_EXCEL)

if __name__ == "__main__":
    main()