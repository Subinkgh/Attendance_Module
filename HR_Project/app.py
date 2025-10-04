from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import os

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        emp_id = request.form['empId']
        report_type = request.form['reportType']
        month = request.form['month']
        try:
            df = pd.read_excel('data/attendance.xlsx', sheet_name='PunchReport', engine='openpyxl')
            df.columns = [col.strip() for col in df.columns]
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df['Month'] = df['Date'].dt.strftime('%Y-%m')
            filtered = df[(df['Card No'].astype(str) == emp_id) & (df['Month'] == month)]

            if filtered.empty:
                return "No data found for the given Employee ID and Month."

            if report_type == 'Punch Report':
                summary = {
                    'Total Days': len(filtered),
                    'PP': (filtered['MusterMark'] == 'PP').sum(),
                    'PA': (filtered['MusterMark'] == 'PA').sum(),
                    'AP': (filtered['MusterMark'] == 'AP').sum(),
                    'AA': (filtered['MusterMark'] == 'AA').sum(),
                    'OO': (filtered['MusterMark'] == 'OO').sum(),
                    'HH': (filtered['MusterMark'] == 'HH').sum(),
                    'Additional': filtered['Additional'].sum()
                }
                present_days = summary['PP'] + 0.5 * (summary['PA'] + summary['AP'])
                absent_days = summary['AA'] + 0.5 * (summary['PA'] + summary['AP'])
                summary['Present Days'] = present_days
                summary['Absent Days'] = absent_days
                employee_name = filtered.iloc[0]['Employee Name'].strip() if 'Employee Name' in filtered.columns else ''
                return render_template('report.html', data=filtered.to_dict(orient='records'), emp_id=emp_id, emp_name=employee_name, month=month, summary=summary)
            else:
                return render_template('muster.html', data=filtered.to_dict(orient='records'), emp_id=emp_id, month=month)
        except Exception as e:
            return f"Error loading report: {e}"
    return render_template('form.html')


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
    app.run(debug=True)
