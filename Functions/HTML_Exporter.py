import pandas as pd
from html import escape
import math
import subprocess
from pathlib import Path

def excel_to_html(
    excel_path: str,
    sheet_name: str | int = 0,
    max_rows: int | None = None,
) -> str:
    excel_file = Path(excel_path).expanduser().resolve()
    if not excel_file.exists():
        raise FileNotFoundError(excel_file)

    df = pd.read_excel(excel_file, sheet_name=sheet_name)

    if max_rows:
        df = df.head(max_rows)

    html_table = df.to_html(
        index=False,
        border=0,
        escape=False
    )

    html = f"""
    <html>
    <head>
      <style>
        body {{
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial;
          font-size: 13px;
          color: #222;
        }}
        table {{
          border-collapse: collapse;
          width: 100%;
        }}
        th {{
          background-color: #f2f2f2;
          font-weight: 600;
          border: 1px solid #ccc;
          padding: 6px 8px;
          text-align: left;
        }}
        td {{
          border: 1px solid #ddd;
          padding: 6px 8px;
          white-space: nowrap;
        }}
        tr:nth-child(even) {{
          background-color: #fafafa;
        }}
      </style>
    </head>
    <body>
      {html_table}
    </body>
    </html>
    """

    return html


def excel_to_html_dashboard(
    excel_path: str,
    sheets: list[str] | None = None,
    max_rows_per_sheet: int | None = None,
) -> str:
    """
    Convert Excel dashboard sheets into email-safe HTML.
    - Preserves headers & structure
    - Adds green/red coloring for % columns
    """

    excel_file = Path(excel_path).expanduser().resolve()
    if not excel_file.exists():
        raise FileNotFoundError(excel_file)

    xls = pd.ExcelFile(excel_file)

    if sheets is None:
        sheets = xls.sheet_names

    html_parts = []

    for sheet in sheets:
        df = pd.read_excel(xls, sheet_name=sheet)

        if max_rows_per_sheet:
            df = df.head(max_rows_per_sheet)

        html_parts.append(f"<h3>{escape(sheet)}</h3>")
        html_parts.append("<table>")

        # Header
        html_parts.append("<thead><tr>")
        for col in df.columns:
            html_parts.append(f"<th>{escape(str(col))}</th>")
        html_parts.append("</tr></thead>")

        # Body
        html_parts.append("<tbody>")
        for _, row in df.iterrows():
            html_parts.append("<tr>")
            for col, value in row.items():
                cell_html = format_cell(value, col)
                html_parts.append(cell_html)
            html_parts.append("</tr>")
        html_parts.append("</tbody></table><br>")

    return wrap_email_html("Excel Dashboard", "".join(html_parts))

def format_cell(value, column_name: str) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "<td></td>"

    is_percent = "%" in str(column_name) or "Chng" in str(column_name)

    try:
        num = float(value)
        text = f"{num:.2f}"

        style = ""
        if is_percent:
            style = "color: green;" if num > 0 else "color: red;" if num < 0 else ""
            text = f"{num:.2f}%"

        return f'<td style="{style}">{text}</td>'

    except Exception:
        return f"<td>{escape(str(value))}</td>"


def export_excel_charts(excel_path: str, output_dir: str):
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    applescript = f'''
    tell application "Microsoft Excel"
        open POSIX file "{excel_path}"
        repeat with sh in worksheets of active workbook
            repeat with ch in charts of sh
                set imgPath to POSIX file "{output_dir}/" & name of ch & ".png"
                export ch to imgPath
            end repeat
        end repeat
        close active workbook saving no
    end tell
    '''

    subprocess.run(["osascript", "-e", applescript], check=True)

def embed_chart_images(image_dir: str) -> str:
    html = "<h3>Charts</h3>"
    for img in sorted(Path(image_dir).glob("*.png")):
        html += f'''
        <div style="margin-bottom:20px;">
          <img src="cid:{img.name}" style="max-width:100%;">
        </div>
        '''
    return html

def wrap_email_html(title: str, body_html: str) -> str:
    return f"""
    <html>
    <head>
      <style>
        body {{
          font-family: -apple-system, BlinkMacSystemFont, Arial;
          font-size: 13px;
          color: #222;
        }}
        table {{
          border-collapse: collapse;
          width: 100%;
        }}
        th {{
          background: #f2f2f2;
          border: 1px solid #ccc;
          padding: 6px;
          text-align: left;
        }}
        td {{
          border: 1px solid #ddd;
          padding: 6px;
        }}
        h3 {{
          margin-top: 20px;
        }}
      </style>
    </head>
    <body>
      <h2>{escape(title)}</h2>
      {body_html}
    </body>
    </html>
    """