from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import os
import logging

app = Flask(__name__)
app.logger.setLevel(logging.INFO)


def find_column(df, candidates, substring_fallback=None):
    # try exact matches first (case-insensitive, stripped)
    cols = {c.strip().lower(): c for c in df.columns}
    for cand in candidates:
        if cand.strip().lower() in cols:
            return cols[cand.strip().lower()]
    # fallback: find first column that contains any candidate substring
    for cand in candidates:
        lower_cand = cand.strip().lower()
        for k, orig in cols.items():
            if lower_cand in k:
                return orig
    # optional substring search across all columns
    if substring_fallback:
        for k, orig in cols.items():
            if substring_fallback.lower() in k:
                return orig
    return None


def normalize_card_value(val):
    # convert numeric floats like 12345.0 to '12345', strip whitespace
    try:
        s = str(val).strip()
        # remove trailing .0 that often comes from Excel floats
        if s.endswith('.0'):
            s = s[:-2]
        return s
    except Exception:
        return ''


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        emp_id = request.form.get('empId', '').strip()
        report_type = request.form.get('reportType', '')
        month_input = request.form.get('month', '').strip()

        try:
            df = pd.read_excel('data/attendance.xlsx', sheet_name='PunchReport', engine='openpyxl')
            # normalize column headers
            df.columns = [col.strip() for col in df.columns]
            app.logger.info("Columns read from Excel: %s", df.columns.tolist())

            # find date column and card number column using common variants
            date_col = find_column(df, ['Date', 'Punch Date', 'PunchDate', 'Attendance Date', 'AttendanceDate'], substring_fallback='date')
            card_col = find_column(df, ['Card No', 'CardNumber', 'Card No.', 'Card No', 'Card #', 'Card'], substring_fallback='card')

            if date_col is None:
                return "Could not find a date column in the Excel. Columns available: " + ', '.join(df.columns)

            if card_col is None:
                return "Could not find a Card No column in the Excel. Columns available: " + ', '.join(df.columns)

            # Convert date column
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            if df[date_col].isna().all():
                # All dates failed parse, give diagnostic
                sample_vals = df[date_col].astype(str).head(10).tolist()
                return f"Date column '{date_col}' could not be parsed as dates. Sample values: {sample_vals}"

            # Normalize card column to string forms
            df[card_col] = df[card_col].apply(normalize_card_value)

            # Build Month column in YYYY-MM format
            df['Month'] = df[date_col].dt.strftime('%Y-%m')

            # Try to parse the month input from the form into YYYY-MM
            month_value = None
            if month_input:
                # If user sent YYYY-MM already, accept it
                if pd.Series([month_input]).str.match(r'^\d{4}-\d{2}$').any():
                    month_value = month_input
                else:
                    # try parsing as a date-like string and then format
                    try:
                        parsed = pd.to_datetime(month_input, errors='coerce')
                        if pd.notna(parsed):
                            month_value = parsed.strftime('%Y-%m')
                    except Exception:
                        month_value = None

            # If parsing failed and the form had something, fallback to the raw string
            if not month_value and month_input:
                month_value = month_input

            app.logger.info("Filtering for card '%s' (column '%s') and month '%s'", emp_id, card_col, month_value)

            # Filter safely: compare normalized strings
            if emp_id:
                emp_id_norm = emp_id
            else:
                emp_id_norm = ''

            # Filter rows
            if month_value:
                filtered = df[(df[card_col].astype(str).str.strip() == emp_id_norm) & (df['Month'] == month_value)]
            else:
                filtered = df[df[card_col].astype(str).str.strip() == emp_id_norm]

            if filtered.empty:
                # Provide diagnostics to help find mismatch
                rows_with_card = df[df[card_col].astype(str).str.strip() == emp_id_norm]
                rows_with_month = df[df['Month'] == month_value] if month_value else pd.DataFrame()
                msg_lines = []
                msg_lines.append("No data found for the given Employee ID and Month.")
                msg_lines.append(f"Employee ID searched: '{emp_id_norm}'")
                msg_lines.append(f"Month searched: '{month_value}'")
                msg_lines.append(f"Total rows in sheet: {len(df)}")
                msg_lines.append(f"Rows matching Employee ID only: {len(rows_with_card)}")
                if len(rows_with_card) > 0:
                    msg_lines.append("Sample rows for matching Employee ID (first 5):")
                    msg_lines.append(rows_with_card.head(5).to_string(index=False))
                msg_lines.append(f"Rows matching Month only: {len(rows_with_month)}")
                if len(rows_with_month) > 0:
                    msg_lines.append("Sample rows for matching Month (first 5):")
                    msg_lines.append(rows_with_month.head(5).to_string(index=False))
                msg_lines.append("Available columns: " + ', '.join(df.columns))
                # Also show sample of card column values to see formatting
                sample_cards = df[card_col].dropna().astype(str).head(20).tolist()
                msg_lines.append(f"Sample values from card column '{card_col}': {sample_cards}")
                return "<pre>" + "\n\n".join(msg_lines) + "</pre>"

            # At this point we have filtered rows
            if report_type == 'Punch Report':
                summary = {
                    'Total Days': len(filtered),
                    'PP': (filtered.get('MusterMark') == 'PP').sum() if 'MusterMark' in filtered.columns else 0,
                    'PA': (filtered.get('MusterMark') == 'PA').sum() if 'MusterMark' in filtered.columns else 0,
                    'AP': (filtered.get('MusterMark') == 'AP').sum() if 'MusterMark' in filtered.columns else 0,
                    'AA': (filtered.get('MusterMark') == 'AA').sum() if 'MusterMark' in filtered.columns else 0,
                    'OO': (filtered.get('MusterMark') == 'OO').sum() if 'MusterMark' in filtered.columns else 0,
                    'HH': (filtered.get('MusterMark') == 'HH').sum() if 'MusterMark' in filtered.columns else 0,
                    'Additional': filtered['Additional'].sum() if 'Additional' in filtered.columns else 0
                }
                present_days = summary['PP'] + 0.5 * (summary['PA'] + summary['AP'])
                absent_days = summary['AA'] + 0.5 * (summary['PA'] + summary['AP'])
                summary['Present Days'] = present_days
                summary['Absent Days'] = absent_days
                employee_name = filtered.iloc[0]['Employee Name'].strip() if 'Employee Name' in filtered.columns else ''
                return render_template('report.html', data=filtered.to_dict(orient='records'), emp_id=emp_id, emp_name=employee_name, month=month_value, summary=summary)
            else:
                return render_template('muster.html', data=filtered.to_dict(orient='records'), emp_id=emp_id, month=month_value)
        except Exception as e:
            app.logger.exception("Error loading report")
            return f"Error loading report: {e}"
    return render_template('form.html')


from flask import jsonify
import pandas as pd
import os


@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        # Ensure data folder exists
        os.makedirs('data', exist_ok=True)
        if 'excelFile' not in request.files:
            return jsonify(success=False, message="No file part")
        file = request.files['excelFile']
        if file.filename == '':
            return jsonify(success=False, message="No selected file")
        if file:
            temp_path = os.path.join('data', 'uploaded_temp.xlsx')
            file.save(temp_path)
            try:
                # Example: read uploaded file, update attendance.xlsx, etc.
                # df_uploaded = pd.read_excel(temp_path, engine='openpyxl')
                # ... your update logic ...

                os.remove(temp_path)
                return jsonify(success=True, message="File uploaded and processed!")
            except Exception as e:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                return jsonify(success=False, message=f"Error processing file: {e}")
    return render_template('upload.html')


if __name__ == '__main__':
    # For development only; Render/gunicorn should use a Procfile in production
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
