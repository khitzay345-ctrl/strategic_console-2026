import pandas as pd
from pathlib import Path
import json
from datetime import datetime

def load_excel_data(file_path):
    """Load all sheets from the Excel file."""
    xls = pd.ExcelFile(file_path)
    return {sheet_name: pd.read_excel(xls, sheet_name=sheet_name) for sheet_name in xls.sheet_names}

def clean_number(val):
    """Convert string numbers to float."""
    if pd.isna(val) or val == "":
        return None
    s = str(val)
    s = ''.join(c for c in s if c.isdigit() or c == '.')
    return float(s) if s else None

def parse_number(value):
    """Parse a number from a string, handling various formats."""
    if value in (None, ''):
        return None
    s = str(value).replace(',', '').replace(' ', '')
    s = re.sub(r"[^\d\.\-]", '', s)
    try:
        return float(s)
    except Exception:
        return None

def fmt_money(value, decimals=2):
    """Format a number as a currency string."""
    num = parse_number(value)
    if num is None:
        return ''
    return f"{num:,.{decimals}f}"

def build_ecom_target_context(sheets):
    """Build context for the ecommerce target section."""
    df = sheets.get('2026 Ecom Target', pd.DataFrame())
    
    if df.empty:
        return {
            "active_tab": "target",
            "target_columns": [],
            "target_rows": [],
            "target_insight": "",
        }
    
    # Clean column names
    df.columns = (
        df.columns.astype(str)
        .str.replace("\u00a0", " ", regex=False)
        .str.replace("\t", " ", regex=False)
        .str.replace("  ", " ", regex=False)
        .str.strip()
    )
    
    # Extract insight text
    insight_text = ""
    first_col = df.columns[0]
    mask = df[first_col].astype(str).str.strip().str.lower().str.contains(r"insight", na=False)
    if mask.any():
        insight_row = df[mask].iloc[0]
        parts = []
        for c in df.columns[1:]:
            clean_val = str(insight_row.get(c, "")).strip()
            if clean_val and clean_val.lower() != 'nan':
                parts.append(clean_val)
        insight_text = " ".join(parts).strip()
    
    # Format rows
    formatted_rows = []
    for _, row in df.iterrows():
        row_out = {}
        for c in df.columns:
            val = row.get(c, "")
            if any(k in c.lower() for k in ["amount", "target", "sales", "moonshot", "fulfillment"]):
                num = parse_number(val)
                row_out[c] = fmt_money(num) if num is not None else ""
            else:
                row_out[c] = val if (val is not None and str(val).strip() != "nan") else ""
        formatted_rows.append(row_out)
    
    return {
        "active_tab": "target",
        "target_columns": list(df.columns),
        "target_rows": formatted_rows,
        "target_insight": insight_text,
    }

def build_ecom_comparison_context(sheets):
    """Build context for the ecommerce comparison section."""
    df = sheets.get('ecom 2024 vs 2025', pd.DataFrame())
    
    if df.empty:
        return {
            "active_tab": "comparison",
            "comp_rows": [],
            "comp_summary": {},
        }
    
    # Clean column names
    df.columns = [str(col).strip() for col in df.columns]
    
    # Process comparison data
    comp_rows = []
    total_2024 = 0
    total_2025 = 0
    count = 0
    
    for _, row in df.iterrows():
        if 'Months' not in row or pd.isna(row['Months']):
            continue
            
        val_2024 = parse_number(row.get('2024', 0)) or 0
        val_2025 = parse_number(row.get('2025', 0)) or 0
        
        # Calculate delta
        delta = val_2025 - val_2024 if val_2024 else 0
        delta_pct = (delta / val_2024 * 100) if val_2024 else 0
        
        # Format values
        fmt_2024 = fmt_money(val_2024, 0) if val_2024 != 0 else "—"
        fmt_2025 = fmt_money(val_2025, 0) if val_2025 != 0 else "—"
        
        # Format delta percentage
        if val_2024 == 0 or val_2025 == 0:
            delta_pct_fmt = "—"
        else:
            delta_pct_fmt = f"{delta_pct:,.1f}%"
        
        comp_rows.append({
            'Months': row['Months'],
            '2024': val_2024,
            '2025': val_2025,
            '2024_fmt': fmt_2024,
            '2025_fmt': fmt_2025,
            'delta': delta,
            'delta_pct': delta_pct,
            'delta_pct_fmt': delta_pct_fmt
        })
        
        # Update totals
        if val_2024 > 0 and val_2025 > 0:
            total_2024 += val_2024
            total_2025 += val_2025
            count += 1
    
    # Calculate averages
    avg_2024 = total_2024 / count if count > 0 else 0
    avg_2025 = total_2025 / count if count > 0 else 0
    
    # Find min and max months
    max_row = max(comp_rows, key=lambda x: x['2025']) if comp_rows else {}
    min_row = min(comp_rows, key=lambda x: x['2025']) if comp_rows else {}
    
    return {
        "active_tab": "comparison",
        "comp_rows": comp_rows,
        "comp_summary": {
            "total_2024": fmt_money(total_2024, 0) if total_2024 > 0 else "—",
            "total_2025": fmt_money(total_2025, 0) if total_2025 > 0 else "—",
            "avg_2024": fmt_money(avg_2024, 0) if avg_2024 > 0 else "—",
            "avg_2025": fmt_money(avg_2025, 0) if avg_2025 > 0 else "—",
            "max_month": max_row.get('Months', '—'),
            "max_value": fmt_money(max_row.get('2025', 0), 0) if max_row else "—",
            "min_month": min_row.get('Months', '—'),
            "min_value": fmt_money(min_row.get('2025', 0), 0) if min_row else "—",
        }
    }

def get_strategy_plan_context(sheets):
    """Build context for the strategy plan section."""
    df = sheets.get('2026 Strategy plan', pd.DataFrame())
    
    if df.empty:
        return {
            "goal_text": "No strategy plan data available.",
            "pillars": {}
        }
    
    # Clean column names
    df.columns = [str(col).strip() for col in df.columns]
    
    # Find the goal text (usually in the first cell)
    goal_text = df.iloc[0][0] if not df.empty and len(df.columns) > 0 else "No goal defined"
    
    # Group by pillar and process entries
    pillars = {}
    for _, row in df.iterrows():
        pillar = row.get('Pillar', '').strip()
        if not pillar or pd.isna(pillar):
            continue
            
        if pillar not in pillars:
            pillars[pillar] = []
            
        # Get photos (assuming comma-separated URLs in a column named 'Photos')
        photos = []
        if 'Photos' in row and pd.notna(row['Photos']):
            photos = [url.strip() for url in str(row['Photos']).split(',') if url.strip()]
        
        entry = {
            'action': row.get('Action', '').strip(),
            'description': row.get('Description', '').strip(),
            'quarter': row.get('Quarter', '').strip(),
            'phase': row.get('Phase', '').strip(),
            'photos': photos
        }
        
        pillars[pillar].append(entry)
    
    return {
        "goal_text": goal_text,
        "pillars": pillars
    }

def get_roadmap_context(sheets):
    """Build context for the roadmap section."""
    df = sheets.get('roadmap', pd.DataFrame())
    
    if df.empty:
        return {
            "quarter_order": [],
            "quarters": {}
        }
    
    # Clean column names
    df.columns = [str(col).strip() for col in df.columns]
    
    # Group by quarter
    quarters = {}
    for _, row in df.iterrows():
        quarter = row.get('Quarter', '').strip()
        if not quarter or pd.isna(quarter):
            continue
            
        if quarter not in quarters:
            quarters[quarter] = []
            
        entry = {
            'Activity_ID': row.get('Activity_ID', '').strip(),
            'Topic': row.get('Topic', '').strip(),
            'Owner': row.get('Owner', '').strip()
        }
        
        quarters[quarter].append(entry)
    
    # Determine quarter order (Q1 2024, Q2 2024, etc.)
    quarter_order = sorted(quarters.keys())
    
    return {
        "quarter_order": quarter_order,
        "quarters": quarters
    }

def extract_styles(template):
    """Extract styles from a template."""
    style_match = re.search(r'<style>([\s\S]*?)</style>', template)
    return style_match.group(1) if style_match else ""

def generate_standalone_html():
    """Generate a standalone HTML file with all the data."""
    # Load all data from Excel
    excel_file = Path(__file__).parent / 'strategic_insight.xlsx'
    sheets = load_excel_data(excel_file)
    
    # Build context for each section
    ecom_target = build_ecom_target_context(sheets)
    ecom_comp = build_ecom_comparison_context(sheets)
    strategy = get_strategy_plan_context(sheets)
    roadmap = get_roadmap_context(sheets)
    
    # Current date for the footer
    current_date = datetime.now().strftime('%Y-%m-%d %H:%M')
    
    # Read templates
    with open('ecom.html', 'r', encoding='utf-8') as f:
        ecom_template = f.read()
    
    with open('strategy_plan.html', 'r', encoding='utf-8') as f:
        strategy_template = f.read()
    
    with open('roadmap.html', 'r', encoding='utf-8') as f:
        roadmap_template = f.read()
    
    # Extract styles
    ecom_styles = extract_styles(ecom_template)
    strategy_styles = extract_styles(strategy_template)
    roadmap_styles = extract_styles(roadmap_template)
    
    # Generate HTML for each section
    ecom_target_html = ecom_template.replace(
        '{% extends "base.html" %}\n{% block content %}',
        ''
    ).replace(
        '{% endblock %}',
        ''
    )
    
    # Strategy Plan
    strategy_plan_html = strategy_template.replace(
        '{% extends "base.html" %}\n{% block content %}',
        ''
    ).replace(
        '{% if not download_mode %}',
        '{% if False %}'  # Disable download button in standalone version
    ).replace(
        '{% endif %}',
        ''
    ).replace(
        '{% endblock %}',
        ''
    )
    
    # Roadmap
    roadmap_html = roadmap_template.replace(
        '{% extends "base.html" %}\n{% block content %}',
        ''
    ).replace(
        '{% endblock %}',
        ''
    )
    
    # E-commerce comparison HTML (simplified for standalone version)
    ecom_comp_html = """
    <div class="dashboard-section">
      <h2 class="section-title">2024 vs 2025 Comparison</h2>
      <div class="table-responsive">
        <table class="table table-hover">
          <thead>
            <tr>
              <th>Months</th>
              <th class="text-end">2024</th>
              <th class="text-end">2025</th>
              <th class="text-end">Δ % (2025 vs 2024)</th>
            </tr>
          </thead>
          <tbody>
    """
    
    # Add comparison rows
    for row in ecom_comp.get('comp_rows', []):
        ecom_comp_html += f"""
            <tr>
              <td>{row['Months']}</td>
              <td class="text-end">{row.get('2024_fmt', '—')}</td>
              <td class="text-end">{row.get('2025_fmt', '—')}</td>
              <td class="text-end">
                <span class="badge bg-{'success' if row.get('delta_pct', 0) >= 0 else 'danger'}">
                  {row.get('delta_pct_fmt', '—')}
                </span>
              </td>
            </tr>
        """
    
    ecom_comp_html += """
          </tbody>
        </table>
      </div>
    </div>
    """
    
    # Combine all sections into a single HTML file
    combined_html = f"""
    <!doctype html>
    <html lang="en">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width,initial-scale=1">
      <title>Strategic Insights Dashboard</title>
      <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
      <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.css" rel="stylesheet">
      <style>
        /* Base styles from base.html */
        :root {{
          --bg-dark: #060d1c;
          --panel-dark: #0c152b;
          --panel-mid: #111c38;
          --primary: #5d6bff;
          --teal: #2CCFBD;
          --text-main: #f8fbff;
          --text-muted: #9ca9c9;
        }}
        
        * {{ box-sizing: border-box; }}
        
        body {{
          margin: 0;
          font-family: "Inter", "Segoe UI", system-ui, sans-serif;
          background: radial-gradient(circle at top, #132146, #060d1c 55%);
          color: var(--text-main);
          min-height: 100vh;
          padding: 20px;
        }}
        
        a {{ text-decoration: none; color: inherit; }}
        
        .dashboard-section {{
          background: var(--panel-dark);
          border-radius: 24px;
          padding: 28px;
          margin-bottom: 30px;
          border: 1px solid rgba(255,255,255,0.07);
          box-shadow: 0 30px 60px rgba(4,6,15,0.65);
        }}
        
        .section-title {{
          font-size: 1.8rem;
          font-weight: 700;
          margin-bottom: 20px;
          color: #fff;
          border-bottom: 2px solid rgba(255,255,255,0.1);
          padding-bottom: 10px;
        }}
        
        /* Import styles from the original templates */
        {ecom_styles}
        {strategy_styles}
        {roadmap_styles}
        
        /* Ensure proper spacing between sections */
        .tab-content {{
          padding: 20px 0;
        }}
        
        .nav-tabs .nav-link {{
          color: var(--text-muted);
        }}
        
        .nav-tabs .nav-link.active {{
          color: #fff;
          background-color: var(--primary);
          border-color: var(--primary);
        }}
        
        .card {{
          background: var(--panel-mid);
          border: 1px solid rgba(255,255,255,0.1);
          color: var(--text-main);
          margin-bottom: 20px;
        }}
        
        .card-header {{
          background: rgba(0,0,0,0.2);
          border-bottom: 1px solid rgba(255,255,255,0.1);
          font-weight: 600;
        }}
        
        .table {{
          color: var(--text-main);
        }}
        
        .table th {{
          border-color: rgba(255,255,255,0.1);
          font-weight: 600;
          text-transform: uppercase;
          font-size: 0.75rem;
          letter-spacing: 0.5px;
          color: var(--text-muted);
        }}
        
        .table td {{
          border-color: rgba(255,255,255,0.05);
          vertical-align: middle;
        }}
        
        .text-muted {{
          color: var(--text-muted) !important;
        }}
        
        .text-primary {{
          color: var(--primary) !important;
        }}
        
        .bg-primary {{
          background-color: var(--primary) !important;
        }}
        
        .btn-primary {{
          background-color: var(--primary);
          border-color: var(--primary);
        }}
        
        .btn-outline-primary {{
          color: var(--primary);
          border-color: var(--primary);
        }}
        
        .btn-outline-primary:hover {{
          background-color: var(--primary);
          color: #fff;
        }}
      </style>
    </head>
    <body>
      <div class="container-fluid">
        <header class="d-flex justify-content-between align-items-center mb-4">
          <h1 class="text-white">Strategic Insights Dashboard</h1>
          <div class="text-muted">Last updated: {current_date}</div>
        </header>
        
        <ul class="nav nav-tabs" id="dashboardTabs" role="tablist">
          <li class="nav-item" role="presentation">
            <button class="nav-link active" id="ecom-tab" data-bs-toggle="tab" data-bs-target="#ecom" type="button" role="tab">
              E-commerce
            </button>
          </li>
          <li class="nav-item" role="presentation">
            <button class="nav-link" id="strategy-tab" data-bs-toggle="tab" data-bs-target="#strategy" type="button" role="tab">
              Strategy Plan
            </button>
          </li>
          <li class="nav-item" role="presentation">
            <button class="nav-link" id="roadmap-tab" data-bs-toggle="tab" data-bs-target="#roadmap" type="button" role="tab">
              Roadmap
            </button>
          </li>
        </ul>
        
        <div class="tab-content" id="dashboardTabsContent">
          <!-- E-commerce Tab -->
          <div class="tab-pane fade show active" id="ecom" role="tabpanel">
            <div class="dashboard-section">
              <h2 class="section-title">E-commerce Performance</h2>
              <ul class="nav nav-tabs" id="ecomTabs" role="tablist">
                <li class="nav-item">
                  <a class="nav-link active" id="target-tab" data-bs-toggle="tab" href="#target" role="tab">2026 Target Plan</a>
                </li>
                <li class="nav-item">
                  <a class="nav-link" id="comparison-tab" data-bs-toggle="tab" href="#comparison" role="tab">2024 vs 2025 Comparison</a>
                </li>
              </ul>
              
              <div class="tab-content mt-4">
                <!-- Target Plan Tab -->
                <div class="tab-pane fade show active" id="target" role="tabpanel">
                  {ecom_target_html}
                </div>
                
                <!-- Comparison Tab -->
                <div class="tab-pane fade" id="comparison" role="tabpanel">
                  {ecom_comp_html}
                </div>
              </div>
            </div>
          </div>
          
          <!-- Strategy Plan Tab -->
          <div class="tab-pane fade" id="strategy" role="tabpanel">
            <div class="dashboard-section">
              {strategy_plan_html}
            </div>
          </div>
          
          <!-- Roadmap Tab -->
          <div class="tab-pane fade" id="roadmap" role="tabpanel">
            <div class="dashboard-section">
              {roadmap_html}
            </div>
          </div>
        </div>
      </div>
      
      <footer class="text-center text-muted mt-5 mb-4">
        <p>Generated on {current_date} | Strategic Insights Dashboard</p>
      </footer>
      
      <!-- Bootstrap JS and dependencies -->
      <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
      
      <script>
        // Initialize tooltips
        document.addEventListener('DOMContentLoaded', function() {{
          var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
          var tooltipList = tooltipTriggerList.map(function (tooltipTriggerEl) {{
            return new bootstrap.Tooltip(tooltipTriggerEl);
          }});
        }});
        
        // Function to open fullscreen image modal
        function openFullScreen(url) {{
          const modal = document.createElement('div');
          modal.style.position = 'fixed';
          modal.style.top = '0';
          modal.style.left = '0';
          modal.style.width = '100%';
          modal.style.height = '100%';
          modal.style.backgroundColor = 'rgba(0, 0, 0, 0.9)';
          modal.style.display = 'flex';
          modal.style.alignItems = 'center';
          modal.style.justifyContent = 'center';
          modal.style.zIndex = '9999';
          modal.onclick = function() {{ document.body.removeChild(modal); }};
          
          const img = document.createElement('img');
          img.src = url;
          img.style.maxWidth = '90%';
          img.style.maxHeight = '90%';
          img.style.borderRadius = '8px';
          
          modal.appendChild(img);
          document.body.appendChild(modal);
        }}
      </script>
    </body>
    </html>
    """
    
    # First, escape any existing curly braces in the content
    def escape_braces(text):
        return text.replace('{', '{{').replace('}', '}}')
    
    ecom_target_html = escape_braces(ecom_target_html)
    ecom_comp_html = escape_braces(ecom_comp_html)
    strategy_plan_html = escape_braces(strategy_plan_html)
    roadmap_html = escape_braces(roadmap_html)
    
    # Format the final HTML with all sections
    html_template = """
    <!doctype html>
    <html lang="en">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width,initial-scale=1">
      <title>Strategic Insights Dashboard</title>
      <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
      <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.css" rel="stylesheet">
      <style>
        /* Base styles */
        :root {{
          --bg-dark: #060d1c;
          --panel-dark: #0c152b;
          --panel-mid: #111c38;
          --primary: #5d6bff;
          --teal: #2CCFBD;
          --text-main: #f8fbff;
          --text-muted: #9ca9c9;
        }}
        
        * {{ box-sizing: border-box; }}
        
        body {{
          margin: 0;
          font-family: "Inter", "Segoe UI", system-ui, sans-serif;
          background: radial-gradient(circle at top, #132146, #060d1c 55%);
          color: var(--text-main);
          min-height: 100vh;
          padding: 20px;
        }}
        
        a {{ text-decoration: none; color: inherit; }}
        
        .dashboard-section {{
          background: var(--panel-dark);
          border-radius: 24px;
          padding: 28px;
          margin-bottom: 30px;
          border: 1px solid rgba(255,255,255,0.07);
          box-shadow: 0 30px 60px rgba(4,6,15,0.65);
        }}
        
        .section-title {{
          font-size: 1.8rem;
          font-weight: 700;
          margin-bottom: 20px;
          color: #fff;
          border-bottom: 2px solid rgba(255,255,255,0.1);
          padding-bottom: 10px;
        }}
        
        /* E-commerce styles */
        .ecom-wrapper {{ max-width: 1200px; margin: 20px auto 28px; }}
        .hero-card {{
          background: radial-gradient(circle at top left,#0b63ff,#04204a 70%);
          color:#fff;
          padding:22px 22px 18px;
          border-radius:22px;
          box-shadow:0 24px 55px rgba(4,32,74,.4);
          margin-bottom:18px;
        }}
        .hero-card h1 {{ margin:0; font-weight:800; letter-spacing:-0.05em; }}
        .hero-card p {{ margin-top:6px;opacity:.88; }}
        .hero-chips{{display:flex;flex-wrap:wrap;gap:.4rem;margin-top:.75rem;}}
        .hero-chip{{padding:0.2rem 0.6rem;border-radius:999px;border:1px solid rgba(191,219,254,0.7);font-size:.72rem;color:#e5e7eb;background:rgba(15,23,42,0.18);}}
        .nav-tabs {{ border:none; margin-bottom:18px; }}
        .nav-tabs .nav-link {{ border:none; border-radius:999px; padding:10px 20px; font-weight:600; color:#c8d4ff; background:rgba(255,255,255,0.06); }}
        .nav-tabs .nav-link.active {{ background:#23d8bb; color:#041024; box-shadow:0 10px 20px rgba(35,216,187,.35); }}
        .table-card {{ background:rgba(12,21,43,0.92); border-radius:24px; padding:22px; box-shadow:0 25px 55px rgba(4,6,15,.55); border:1px solid rgba(255,255,255,0.08); color:#f8fbff; }}
        .e-table {{ width:100%; border-collapse:separate; border-spacing:0; color:#f4f7ff; }}
        .e-table thead th {{ text-align:left; text-transform:uppercase; font-size:.75rem; letter-spacing:1px; padding:14px 12px; color:#d5ddff; border-bottom:2px solid rgba(255,255,255,0.12); background:rgba(255,255,255,0.06); }}
        .e-table tbody td {{ padding:14px 12px; border-bottom:1px solid rgba(255,255,255,0.08); color:#fefefe; font-weight:500; }}
        .e-table tbody tr:last-child td {{ border-bottom:none; }}
        .e-table tbody tr:hover {{ background:rgba(255,255,255,0.06); }}
        .e-table td.text-end {{ text-align:right; color:#fefefe; font-weight:700; }}
        .stat-grid {{ display:grid; grid-template-columns: repeat(auto-fit,minmax(210px,1fr)); gap:18px; margin-bottom:20px; }}
        .stat-card {{
          background:linear-gradient(135deg,#111b3d,#050811);
          border-radius:18px;
          padding:16px;
          box-shadow:0 18px 38px rgba(3,5,12,.7);
          display:flex;
          flex-direction:column;
          gap:2px;
          border:1px solid rgba(255,255,255,0.08);
          color:#eaf1ff;
        }}
        .stat-card span {{ text-transform:uppercase; font-size:.72rem; letter-spacing:.5px; color:#9fb4ff; }}
        .stat-card strong {{ font-size:1.6rem; color:#fff; }}
        .stat-card small {{ font-size:.8rem; color:#b5c2ff; }}
        .delta-pill {{ padding:4px 10px; border-radius:999px; font-size:.75rem; background:rgba(35,216,187,.2); color:#23d8bb; }}
        .bad {{ background:rgba(229,72,72,.2); color:#ffb3b3; }}
        
        /* Strategy plan styles */
        .plan-hero {{
          background: radial-gradient(circle at top,#1d2d4f,#050c1a 65%);
          color:#fff;
          padding:36px;
          border-radius:30px;
          margin-bottom:28px;
          position:relative;
          overflow:hidden;
          display:flex;
          justify-content:space-between;
          gap:18px;
          align-items:center;
        }}
        .plan-hero h1 {{ margin:0; font-weight:800; letter-spacing:-.4px; }}
        .plan-hero p {{ margin-top:6px; color:#cdd7ef; font-size:1.18rem; font-weight:600; }}
        .pillar-card {{
          background:var(--panel-mid);
          border-radius:26px;
          padding:26px;
          margin-bottom:26px;
          border:1px solid rgba(255,255,255,.05);
        }}
        .pillar-header {{ display:flex; flex-wrap:wrap; gap:10px; justify-content:space-between; align-items:center; color:#fff; }}
        .pillar-header h2 {{ margin:0; font-weight:700; }}
        .phase-chip {{
          padding:8px 16px;
          border-radius:999px;
          border:1px solid rgba(255,255,255,.3);
          font-size:.85rem;
          text-transform:uppercase;
          color:#dde4ff;
        }}
        .actions-grid {{ margin-top:18px; display:grid; grid-template-columns: repeat(auto-fit,minmax(320px,1fr)); gap:20px; }}
        .action-panel {{
          background:rgba(255,255,255,.02);
          border:1px solid rgba(255,255,255,.07);
          border-radius:22px;
          padding:18px;
          color:#e4ebff;
          display:flex;
          flex-direction:column;
          gap:14px;
          position:relative;
          overflow:hidden;
        }}
        .action-photo {{
          width:100%;
          aspect-ratio: 4 / 3;
          border-radius:16px;
          background-size:cover;
          background-position:center;
          box-shadow:0 16px 30px rgba(0,0,0,.35);
          cursor:pointer;
        }}
        .meta-row {{ display:flex; gap:10px; flex-wrap:wrap; font-size:.85rem; color:var(--text-muted); }}
        .meta-row span {{ background:rgba(255,255,255,.08); border-radius:999px; padding:4px 10px; }}
        .action-panel h3 {{ margin:0; font-size:1.15rem; color:#fff; }}
        .action-panel p {{ margin:0; font-size:.95rem; color:#cbd4ef; line-height:1.4; }}
        
        /* Roadmap styles */
        .road-bg {{
          position: relative;
          border-radius: 18px;
          overflow: hidden;
          border: 1px solid rgba(255,255,255,.1);
        }}
        .road-bg img {{
          width: 100%;
          object-fit: cover;
          filter: saturate(1.2);
        }}
        .glass-panel {{
          background: rgba(4, 9, 18, .72);
          border: 1px solid rgba(255,255,255,.08);
          border-radius: 18px;
          padding: 24px;
          margin-top: -60px;
          position: relative;
          backdrop-filter: blur(16px);
        }}
        .quarter-card {{
          background: rgba(255,255,255,.04);
          border-radius: 20px;
          padding: 18px;
          margin-bottom: 18px;
          border: 1px solid rgba(255,255,255,.07);
          transition: transform .2s ease;
        }}
        .quarter-card:hover {{ transform: translateY(-6px); }}
        .entry-table {{ width:100%; border-collapse:collapse; margin-top:12px; table-layout:fixed; }}
        .entry-table col.col-activity {{ width: 20%; }}
        .entry-table col.col-topic {{ width: 55%; }}
        .entry-table col.col-owner {{ width: 25%; }}
        .entry-table th, .entry-table td {{ padding:12px 10px; border-bottom:1px solid rgba(255,255,255,.1); }}
        .entry-table th {{ text-transform:uppercase; font-size:.75rem; letter-spacing:.6px; color:#9ab2d3; }}
        .entry-table td {{ color:#e7efff; }}
        .entry-table th:nth-child(3),
        .entry-table td:nth-child(3) {{ text-align:right; }}
        .badge-owner {{ padding:4px 10px; border-radius:999px; background:rgba(255,255,255,.15); font-size:.8rem; }}
        
        /* Common table styles */
        .table {{
          color: var(--text-main);
          width: 100%;
          margin-bottom: 1rem;
          background-color: transparent;
        }}
        .table th {{
          border-color: rgba(255,255,255,0.1);
          font-weight: 600;
          text-transform: uppercase;
          font-size: 0.75rem;
          letter-spacing: 0.5px;
          color: var(--text-muted);
        }}
        .table td {{
          border-color: rgba(255,255,255,0.05);
          vertical-align: middle;
        }}
        .text-end {{ text-align: right !important; }}
        .text-muted {{
          color: var(--text-muted) !important;
        }}
        .badge {{
          font-weight: 600;
          padding: 0.35em 0.65em;
          font-size: 0.75em;
          border-radius: 0.5rem;
        }}
        .bg-success {{
          background-color: #28a745 !important;
        }}
        .bg-danger {{
          background-color: #dc3545 !important;
        }}
      </style>
    </head>
    <body>
      <div class="container-fluid">
        <header class="d-flex justify-content-between align-items-center mb-4">
          <h1 class="text-white">Strategic Insights Dashboard</h1>
          <div class="text-muted">Last updated: {current_date}</div>
        </header>
        
        <ul class="nav nav-tabs" id="dashboardTabs" role="tablist">
          <li class="nav-item" role="presentation">
            <button class="nav-link active" id="ecom-tab" data-bs-toggle="tab" data-bs-target="#ecom" type="button" role="tab">
              <i class="bi bi-cart me-2"></i> E-commerce
            </button>
          </li>
          <li class="nav-item" role="presentation">
            <button class="nav-link" id="strategy-tab" data-bs-toggle="tab" data-bs-target="#strategy" type="button" role="tab">
              <i class="bi bi-easel3 me-2"></i> Strategy Plan
            </button>
          </li>
          <li class="nav-item" role="presentation">
            <button class="nav-link" id="roadmap-tab" data-bs-toggle="tab" data-bs-target="#roadmap" type="button" role="tab">
              <i class="bi bi-map me-2"></i> Roadmap
            </button>
          </li>
        </ul>
        
        <div class="tab-content mt-4" id="dashboardTabsContent">
          <!-- E-commerce Tab -->
          <div class="tab-pane fade show active" id="ecom" role="tabpanel">
            <div class="ecom-wrapper">
              <div class="hero-card">
                <h1>E-commerce Performance</h1>
                <p>Targets for 2026 plus a comparison snapshot of 2024 vs 2025 actuals.</p>
                <div class="hero-chips">
                  <span class="hero-chip">Traffic & Conversion</span>
                  <span class="hero-chip">Average Basket</span>
                  <span class="hero-chip">Year-on-Year Growth</span>
                </div>
              </div>

              <ul class="nav nav-tabs" id="ecomTabs" role="tablist">
                <li class="nav-item">
                  <button class="nav-link active" id="target-tab" data-bs-toggle="tab" data-bs-target="#target" type="button" role="tab">2026 Target Plan</button>
                </li>
                <li class="nav-item">
                  <button class="nav-link" id="comparison-tab" data-bs-toggle="tab" data-bs-target="#comparison" type="button" role="tab">2024 vs 2025 Comparison</button>
                </li>
              </ul>
              
              <div class="tab-content mt-4">
                <!-- Target Plan Tab -->
                <div class="tab-pane fade show active" id="target" role="tabpanel">
                  {ecom_target_html}
                </div>
                
                <!-- Comparison Tab -->
                <div class="tab-pane fade" id="comparison" role="tabpanel">
                  {ecom_comp_html}
                </div>
              </div>
            </div>
          </div>
          
          <!-- Strategy Plan Tab -->
          <div class="tab-pane fade" id="strategy" role="tabpanel">
            <div class="dashboard-section">
              {strategy_plan_html}
            </div>
          </div>
          
          <!-- Roadmap Tab -->
          <div class="tab-pane fade" id="roadmap" role="tabpanel">
            <div class="dashboard-section">
              {roadmap_html}
            </div>
          </div>
        </div>
      </div>
      
      <footer class="text-center text-muted mt-5 mb-4">
        <p>Generated on {current_date} | Strategic Insights Dashboard</p>
      </footer>
      
      <!-- Bootstrap JS and dependencies -->
      <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
      
      <script>
        // Initialize tooltips
        document.addEventListener('DOMContentLoaded', function() {{
          var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
          var tooltipList = tooltipTriggerList.map(function (tooltipTriggerEl) {{
            return new bootstrap.Tooltip(tooltipTriggerEl);
          }});
        }});
        
        // Function to open fullscreen image modal
        function openFullScreen(url) {{
          const modal = document.createElement('div');
          modal.style.position = 'fixed';
          modal.style.top = '0';
          modal.style.left = '0';
          modal.style.width = '100%';
          modal.style.height = '100%';
          modal.style.backgroundColor = 'rgba(0, 0, 0, 0.9)';
          modal.style.display = 'flex';
          modal.style.alignItems = 'center';
          modal.style.justifyContent = 'center';
          modal.style.zIndex = '9999';
          modal.onclick = function() { document.body.removeChild(modal); };
          
          const img = document.createElement('img');
          img.src = url;
          img.style.maxWidth = '90%';
          img.style.borderRadius = '8px';
          
          modal.appendChild(img);
          document.body.appendChild(modal);
        }}
    </script>
    </body>
    </html>
"""
    
    # Format the final HTML with all variables
    try:
        final_html = html_template.format(
            ecom_target_html=ecom_target_html,
            ecom_comp_html=ecom_comp_html,
            strategy_plan_html=strategy_plan_html,
            roadmap_html=roadmap_html,
            current_date=current_date
        )
    except KeyError as e:
        print(f"Error formatting HTML template. Missing key: {e}")
        # Fallback to a simple template if formatting fails
        final_html = f"""
        <!doctype html>
        <html>
        <head>
            <title>Strategic Insights Dashboard</title>
            <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
        </head>
        <body>
            <div class="container mt-4">
                <h1>Strategic Insights Dashboard</h1>
                <p>Last updated: {current_date}</p>
                <div class="card mt-4">
                    <div class="card-header">E-commerce Target</div>
                    <div class="card-body">{ecom_target_html}</div>
                </div>
                <div class="card mt-4">
                    <div class="card-header">E-commerce Comparison</div>
                    <div class="card-body">{ecom_comp_html}</div>
                </div>
                <div class="card mt-4">
                    <div class="card-header">Strategy Plan</div>
                    <div class="card-body">{strategy_plan_html}</div>
                </div>
                <div class="card mt-4">
                    <div class="card-header">Roadmap</div>
                    <div class="card-body">{roadmap_html}</div>
                </div>
            </div>
        </body>
        </html>
        """.format(
            ecom_target_html=ecom_target_html,
            ecom_comp_html=ecom_comp_html,
            strategy_plan_html=strategy_plan_html,
            roadmap_html=roadmap_html,
            current_date=current_date
        )
    
    # Save the standalone HTML file
    output_file = Path(__file__).parent / 'strategic_insights_dashboard.html'
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(final_html)
    
    print(f"Standalone HTML dashboard generated: {output_file}")

if __name__ == "__main__":
    import re  # Import re at the top level
    generate_standalone_html()
