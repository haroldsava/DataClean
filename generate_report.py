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

# Sheet mappings: { 'Sheet Name': [List of Column Names to check] }
TARGET_SHEETS = {
    "1  MOBILITATE US (export)": ["NR. CONTRACT"],              # Usually Col J
    "1  REG URBANA US (export)": ["Nr. contract"],              # Usually Col S
    "2  PROIECTE US (export)": ["Nr. contract"],                # Usually Col S
    "2  PROIECTE US (CF - STRATEGII ": ["NR. CONTRACT", "Nr. contract.1"] # Col L and M
}
# ---------------------

def clean_text(value):
    """Applies the same regex cleaner used in the first step"""
    text = str(value)
    if text.lower() in ['nan', 'none', '']:
        return ""
    pattern = r'[\s\-/]*\d{2}\.\d{2}\.\d{4}.*'
    return re.sub(pattern, '', text).strip()

def find_project_name_col(df):
    """Attempts to find the column containing the Project Name in the user sheet"""
    keywords = ['proiect', 'titlu', 'nume', 'title', 'denumire', 'task']
    for col in df.columns:
        if any(k in str(col).lower() for k in keywords):
            return col
    return df.columns[1] if len(df.columns) > 1 else df.columns[0] # Fallback to Col B

def generate_html(data, stats, all_exports):
    """Generates the HTML file matching your exact requested design."""
    
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
    <title>Raport Master - Multiple Exporturi</title>
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
        table {{ width: 100%; border-collapse: collapse; font-size: 0.85em; }}
        th {{ background: rgba(168,85,247,0.2); padding: 10px; text-align: left; }}
        td {{ padding: 8px 10px; border-bottom: 1px solid rgba(255,255,255,0.1); }}
        tr:hover {{ background: rgba(255,255,255,0.05); }}
        .status {{ padding: 3px 8px; border-radius: 12px; font-size: 0.75em; }}
        .status.valid {{ background: rgba(39,174,96,0.3); color: #2ecc71; }}
        .status.empty {{ background: rgba(243,156,18,0.3); color: #f1c40f; }}
        .status.notfound {{ background: rgba(231,76,60,0.3); color: #e74c3c; }}
        .status.typo {{ background: rgba(52,152,219,0.3); color: #3498db; }}
        .status.duplicate {{ background: rgba(155,89,182,0.3); color: #9b59b6; }}
        .table-wrapper {{ max-height: 500px; overflow: auto; }}
        footer {{ text-align: center; padding: 20px; color: #666; font-size: 0.85em; }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>Raport Master - Multiple Exporturi</h1>
            <div class="subtitle">{len(all_exports)} seturi date | Total: {total_all} intrari | Generat: {timestamp}</div>
        </header>
        
        <div class="dashboard">
            <div class="cards">
                <div class="card valid"><h2 id="statValid">{total_valid}</h2><div class="pct" id="pctValid">{round(total_valid/total_all*100,1)}%</div><p>Coduri Valide</p></div>
                <div class="card duplicate"><h2 id="statDuplicate">{total_dup}</h2><div class="pct" id="pctDuplicate">{round(total_dup/total_all*100,1)}%</div><p>Duplicate</p></div>
                <div class="card typo"><h2 id="statTypo">{total_typo}</h2><div class="pct" id="pctTypo">{round(total_typo/total_all*100,1)}%</div><p>Posibile Erori</p></div>
                <div class="card empty"><h2 id="statEmpty">{total_empty}</h2><div class="pct" id="pctEmpty">{round(total_empty/total_all*100,1)}%</div><p>Necompletate</p></div>
                <div class="card notfound"><h2 id="statNotfound">{total_notfound}</h2><div class="pct" id="pctNotfound">{round(total_notfound/total_all*100,1)}%</div><p>Inexistente</p></div>
            </div>
            <div class="chart-container"><canvas id="pieChart" width="250" height="250"></canvas></div>
        </div>
        
        <div class="export-tabs">
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
                        <thead><tr><th>Sursa</th><th>Status</th><th>Linie</th><th>Cod</th><th>Sugestie</th><th>Proiect</th></tr></thead>
                        <tbody>
"""
        for r in probs:
            status_class = r['status'].lower().replace(' ', '')
            tab_html += f"<tr><td>{r['source']}</td><td><span class='status {status_class}'>{r['status_text']}</span></td><td>{r['line']}</td><td>{r['code']}</td><td>{r['suggest']}</td><td>{r['project']}</td></tr>\n"
        
        tab_html += f"""                        </tbody>
                    </table>
                </div>
            </div>
            <div id="export{idx}valid" class="inner-content">
                <div class="table-wrapper">
                    <table id="table{prefix}Valid">
                        <thead><tr><th>Sursa</th><th>Status</th><th>Linie</th><th>Cod</th><th>Proiect</th></tr></thead>
                        <tbody>
"""
        for r in valids:
            tab_html += f"<tr><td>{r['source']}</td><td><span class='status valid'>VALID</span></td><td>{r['line']}</td><td>{r['code']}</td><td>{r['project']}</td></tr>\n"
            
        tab_html += "                        </tbody></table></div></div></div>\n"
        return tab_html

    # Render Master
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
                labels: ['Valide', 'Duplicate', 'Erori', 'Necompletate', 'Inexistente'],
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
    </script>
</body>
</html>
"""
    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Report generated successfully: {OUTPUT_HTML}")


def main():
    print("Loading Registry File...")
    reg_df = pd.read_excel(REGISTRY_FILE)
    
    # Extract Valid Codes and Projects
    reg_col = 'NUMAR CONTRACT'
    proj_col = reg_df.columns[4] # Column E is index 4
    
    valid_codes = {}
    missing_in_registry = []

    for index, row in reg_df.iterrows():
        code = clean_text(row[reg_col])
        proj = str(row[proj_col])
        if code:
            valid_codes[code] = proj
        else:
            missing_in_registry.append(proj) # Registry has project but no contract code
            
    valid_codes_list = list(valid_codes.keys())

    print("Loading Wrike Input File...")
    wrike_sheets = pd.read_excel(INPUT_FILE, sheet_name=None)
    
    report_data = {}
    report_stats = {}
    all_exports = []

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
                    proj_val = str(row[proj_name_col]) if proj_name_col in df.columns else "N/A"
                    
                    status = ""
                    status_text = ""
                    suggest = ""
                    
                    # 1. Check for Empty
                    if clean_val == "":
                        # Is this project listed as missing contract in registry?
                        if any(fuzz.partial_ratio(proj_val, r_proj) > 90 for r_proj in missing_in_registry):
                            status = "empty"
                            status_text = "COD CONTRACT NECOMPLETAT IN REGISTRU"
                        else:
                            status = "empty"
                            status_text = "COD CONTRACT NECOMPLETAT"
                        stats['empty'] += 1
                        
                    # 2. Check for Exact Match
                    elif clean_val in valid_codes:
                        if value_counts[clean_val] > 1:
                            status = "duplicate"
                            status_text = "COD DUPLICAT"
                            stats['duplicate'] += 1
                        else:
                            status = "VALID"
                            status_text = "VALID"
                            stats['valid'] += 1
                        proj_val = valid_codes[clean_val] # Use Registry name if matched
                        
                    # 3. Check for Typos / Not Found
                    else:
                        match = process.extractOne(clean_val, valid_codes_list, scorer=fuzz.ratio)
                        if match and match[1] >= 85: # 85% similarity threshold
                            status = "typo"
                            status_text = "POSIBILA EROARE DE TASTARE"
                            suggest = match[0]
                            stats['typo'] += 1
                        else:
                            status = "notfound"
                            status_text = "COD INEXISTENT IN REGISTRU"
                            stats['notfound'] += 1

                    report_data[display_name].append({
                        "source": display_name,
                        "line": idx + 2, # +2 for Excel row number (0-idx + header)
                        "code": raw_val if not pd.isna(raw_val) else "",
                        "status": status,
                        "status_text": status_text,
                        "suggest": suggest,
                        "project": proj_val
                    })
            else:
                print(f"Warning: Column '{col_name}' not found in sheet '{sheet_name}'.")

            report_stats[display_name] = stats

    generate_html(report_data, report_stats, all_exports)

if __name__ == "__main__":
    main()