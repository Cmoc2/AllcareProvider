# -*- coding: utf-8 -*-
"""
Created on Tue Nov  4 14:51:38 2025

@author: ChristianOrtiz
"""

from flask import Flask, render_template, request, g, redirect, url_for
import sqlite3

app = Flask(__name__)
DATABASE = 'comcare_kaiser_authorizations.db'

def get_db():
    db = getattr(g, '_database', None)
    if db is None:
        db = g._database = sqlite3.connect(DATABASE)
    return db

@app.teardown_appcontext
def close_connection(exception):
    db = getattr(g, '_database', None)
    if db is not None:
        db.close()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/add', methods=['POST'])
def add_authorization():
    data = (
        request.form['patient_name'],
        request.form['mrn_number'],
        request.form['authorization_number'],
        request.form['RN_visit_count'],
        request.form['LVN_visit_count'],
        request.form['PT_visit_count'],
        request.form['OT_visit_count'],
        request.form['ST_visit_count'],
        request.form['MSW_visit_count'],
        request.form['HHA_visit_count'],
        request.form['soc_date'],
        request.form['request_date'],
        request.form['approved_date']
    )
    #Insert into Database
    db = get_db()
    cursor = db.cursor()
    cursor.execute('''
        INSERT INTO authorizations (patient_name, mrn_number, authorization_number, RN_visit_count, LVN_visit_count, PT_visit_count, OT_visit_count, ST_visit_count, MSW_visit_count, HHA_visit_count,soc_date, request_date, approved_date)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', data)
    db.commit()

    return redirect(url_for('summary'),mrn=data[1])

@app.route('/search', methods=['GET'])
def search():
    mrn = request.args.get('mrn')
    db = get_db()
    cursor = db.cursor()
    cursor.execute('SELECT * FROM authorizations WHERE mrn_number = ?', (mrn,))
    rows = cursor.fetchall()
    return render_template('search.html', rows=rows)

@app.route('/summary', methods=['GET'])
def summary():
    mrn = request.args.get('mrn')
    db = get_db()
    cursor = db.cursor()
    
    #Get patient name and MRN
    cursor.execute('SELECT patient_name, mrn_number FROM authorizations WHERE mrn_number = ? LIMIT 1', (mrn,))
    patient_info = cursor.fetchone()  # (patient_name, mrn_number)

    # Get summary grouped by SOC date
    cursor.execute('SELECT soc_date, patient_name, mrn_number, authorization_number, SUM(RN_visit_count), SUM(LVN_visit_count), SUM(PT_visit_count), SUM(OT_visit_count), SUM(ST_visit_count), SUM(MSW_visit_count), SUM(HHA_visit_count) FROM authorizations WHERE mrn_number = ? GROUP BY soc_date', (mrn,))
    rows = cursor.fetchall()
    return render_template('summary.html', rows=rows, patient_info=patient_info)

@app.route('/update/<int:transaction_id>', methods=['GET', 'POST'])
def update_authorization(transaction_id):
    db = get_db()
    cursor = db.cursor()
    if request.method == 'POST':
        approved_date = request.form['approved_date']
        cursor.execute('UPDATE authorizations SET approved_date = ? WHERE transaction_id = ?', (approved_date, transaction_id))
        db.commit()
        
        # Get MRN for redirect
        cursor.execute('SELECT mrn_number FROM authorizations WHERE transaction_id = ?', (transaction_id,))
        result = cursor.fetchone()
        if result:
            mrn = result[0]
            return redirect(url_for('search',mrn=mrn))
        else:
            return "Update Successful, but MRN not found for direct."
    else:
        cursor.execute('SELECT * FROM authorizations WHERE transaction_id = ?', (transaction_id,))
        row = cursor.fetchone()
        return render_template('update.html', row=row)
