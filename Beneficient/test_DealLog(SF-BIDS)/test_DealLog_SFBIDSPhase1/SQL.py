from flask import Flask, render_template, request, redirect, make_response
import pyodbc, datetime
import pandas as pd

app = Flask(__name__)


#connect_db function is used to connect with the database for read and write purposes
def connect_db():
    connection = pyodbc.connect(r'Driver={ODBC Driver 17 for SQL Server};Server=ben-rds-ss-bidsdb-nonprod-east.ckc8jx9o1aev.us-east-1.rds.amazonaws.com;Database=BIDS_DEV;Trusted_Connection=yes',autocommit=True)
    return connection
    #print(connection)


# #index function calls the homepage
# @app.route('/')
# def index():
    con = connect_db()
    cursor = con.cursor()
    sql_query = 'SELECT * FROM Students'
    cursor.execute(sql_query)
    return render_template('index.html')