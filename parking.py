from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import sqlite3
import random
from openpyxl import Workbook
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
import os

app = Flask(__name__)
app.secret_key = "parking_secret_key"
DB = "parking.db"


def init_db():
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS residents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            flat TEXT UNIQUE NOT NULL,
            name TEXT UNIQUE NOT NULL,
            vehicle_number TEXT UNIQUE NOT NULL,
            vehicle_type TEXT NOT NULL,
            slot TEXT
        )
    ''')
    conn.commit()
    conn.close()


def get_connection():
    return sqlite3.connect(DB)


@app.route('/')
def home():
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT flat, name, vehicle_number, vehicle_type, slot FROM residents")
    residents = c.fetchall()
    c.execute("SELECT COUNT(*) FROM residents")
    total_residents = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM residents WHERE slot IS NOT NULL")
    occupied_slots = c.fetchone()[0]
    conn.close()
    return render_template('index.html', residents=residents, total_residents=total_residents,
                           occupied_slots=occupied_slots, available_slots=max(0, 50 - occupied_slots))


@app.route('/add', methods=['POST'])
def add_resident():
    flat = request.form['flat']
    name = request.form['name']
    vehicle_number = request.form['vehicle_number']
    vehicle_type = request.form['vehicle_type']

    try:
        conn = get_connection()
        c = conn.cursor()
        c.execute("INSERT INTO residents (flat, name, vehicle_number, vehicle_type) VALUES (?, ?, ?, ?)",
                  (flat, name, vehicle_number, vehicle_type))
        conn.commit()
        flash('Resident added successfully!', 'success')
    except sqlite3.IntegrityError:
        flash('Duplicate flat, resident, or vehicle number detected!', 'error')
    finally:
        conn.close()

    return redirect(url_for('home'))


@app.route('/allocate', methods=['POST'])
def allocate():
    slots = int(request.form['slots'])
    parking_slots = [f'Slot-{i}' for i in range(1, slots + 1)]
    random.shuffle(parking_slots)

    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT id FROM residents")
    residents = c.fetchall()

    if len(residents) > slots:
        flash('Not enough parking slots!', 'error')
        conn.close()
        return redirect(url_for('home'))

    c.execute("UPDATE residents SET slot = NULL")
    for resident in residents:
        c.execute("UPDATE residents SET slot = ? WHERE id = ?", (parking_slots.pop(), resident[0]))

    conn.commit()
    conn.close()
    flash('Parking allocated successfully!', 'success')
    return redirect(url_for('home'))


@app.route('/export/excel')
def export_excel():
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT flat, name, vehicle_number, vehicle_type, slot FROM residents")
    data = c.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = 'Parking Allocation'
    ws.append(['Flat', 'Resident', 'Vehicle Number', 'Vehicle Type', 'Parking Slot'])

    for row in data:
        ws.append(row)

    filename = 'parking_allocation.xlsx'
    wb.save(filename)
    return send_file(filename, as_attachment=True)


@app.route('/export/pdf')
def export_pdf():
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT flat, name, vehicle_number, vehicle_type, slot FROM residents")
    data = c.fetchall()
    conn.close()

    filename = 'parking_allocation.pdf'
    doc = SimpleDocTemplate(filename)
    styles = getSampleStyleSheet()
    elements = [Paragraph('Parking Allocation Report', styles['Title']), Spacer(1, 20)]

    table_data = [['Flat', 'Resident', 'Vehicle', 'Type', 'Slot']] + list(data)
    table = Table(table_data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))

    elements.append(table)
    doc.build(elements)
    return send_file(filename, as_attachment=True)


if __name__ == '__main__':
    init_db()
    app.run(debug=True)
