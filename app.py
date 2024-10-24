from flask import Flask, request, render_template, send_file
import pandas as pd
import io

app = Flask(__name__)

attendance_data = None  # Global variable to store attendance data

@app.route('/', methods=['GET', 'POST'])
def index():
    global attendance_data  # Use the global variable

    if request.method == 'POST':
        file = request.files['file']
        if file:
            # Read the Excel file, skipping the first row and using the second row as the header
            df = pd.read_excel(file, skiprows=1)  
            print("DataFrame after loading:\n", df)  # Print the DataFrame

            # Strip whitespace from column names
            df.columns = df.columns.str.strip()

            try:
                # Reference columns directly by their names "Datum" and "Jméno"
                df['Date'] = pd.to_datetime(df['Datum'])  # Use column 'Datum' for dates
                df['User'] = df['Jméno']  # Use column 'Jméno' for users

                # Convert Date column to just the date part
                df['Date'] = df['Date'].dt.date 

                # Count unique days each user was present
                attendance_data = df.groupby('User')['Date'].nunique().reset_index(name='Count')
                print("Attendance DataFrame:\n", attendance_data)  # Print attendance DataFrame

                return render_template('results.html', attendance=attendance_data)

            except KeyError as e:
                return f"KeyError: {str(e)}. Available columns are: {df.columns.tolist()}"

    return render_template('index.html')

@app.route('/export', methods=['GET'])
def export():
    global attendance_data
    if attendance_data is not None:
        # Create a BytesIO buffer
        output = io.BytesIO()
        # Write DataFrame to the buffer as an Excel file
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            attendance_data.to_excel(writer, sheet_name='Attendance', index=False)
        output.seek(0)  # Rewind the buffer

        return send_file(output, as_attachment=True, download_name='attendance_results.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    return "No data available for export."


if __name__ == '__main__':
       app.run(debug=True, host='0.0.0.0', port=8080)
