#################################################################################################################################
#################################################################################################################################
###                                                                                                                           ###
###                                                     DATA VIZ --- App                                                      ###
###                                                                                                                           ###
#################################################################################################################################
#################################################################################################################################


##################################################################################################################################
###    00  -  Define various parameters :                                                                                      ###
##################################################################################################################################

###    00.1  -  Libraries :
from flask import Flask, send_from_directory, render_template, request, jsonify, redirect, url_for
# Flask documents office site: https://flask.palletsprojects.com/en/1.1.x/
from flask_bootstrap import Bootstrap
# BOOTSTRAP site =>  https://getbootstrap.com/docs/5.0/getting-started/introduction/
# FLASK_BOOTSTRAP SITE: https://pythonhosted.org/Flask-Bootstrap/basic-usage.html
from sqlalchemy import create_engine
from pandasql import sqldf
import os
import pandas as pd

pd.options.mode.chained_assignment = None  # avoid SettingWithCopy Warning

from collections import OrderedDict
from flask_sqlalchemy import SQLAlchemy

import datetime
from datetime import timedelta
from datetime import datetime as DT
from dateutil.relativedelta import relativedelta
# json
import json

# download ppt/Excel
try:
    import xlwings as xw
    import win32com.client
    import win32com
    import pythoncom
    download_cash = win32com.__gen_path__
    print(download_cash) # python crash
except: # ImportError
    print("Do not win32 platform")

import numpy as np
# finding the network full path
import ctypes
from ctypes import wintypes



# Instantiate Bootstrap: to Help directly apply ready-made template html page.
app = Flask(__name__)
bootstrap = Bootstrap(app)

###    00.2  -  Parameters :

###         00.2.3   Repository - WARNING TO BE CHANGE in function of the environment !!!
file_path = os.path.abspath(__file__)
# file_path = file_path.replace("\DealerKPIS_backend.py","")

Excel_Path = '<chemin a définir>'
MTData_excel_file = 'DEALER_KPIS.xlsx'


###         00.2.4   Configuration setting
app.config[
    'SQLALCHEMY_DATABASE_URI'] = 'sqlite:///sqlite3.db'  # The database URI that should be used for the connection.
app.config[
    'SQLALCHEMY_TRACK_MODIFICATIONS'] = False  # to disable the modification tracking system and avoid the warning break risk
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = timedelta(
    seconds=1)  # Cache time is set to 1 second, thus css and html reload when you change anything



##################################################################################################################################
###    01  -     Manipulate / Transform DATA                                                                                   ###
##################################################################################################################################

###         00.2.4   Import data from Excel
# DATA MT
#MT_xlsx = pd.ExcelFile(Excel_Path + MTData_excel_file)
#data = pd.read_excel(MT_xlsx, 'DATA')

pysqldf = lambda q: sqldf(q, globals())  # Use sql with Pandas and avoid specifying everytime

###    1.1  -    EXCEL Sheet 'DATA' : SALES DATA MONTHLY / YEARLY
DATA_DEALER = """  select distinct KEY ,
                                UNIT ,
                                TIMEUNIT , 
                                MIN ,
                                MAX , 
 

                                -- for Trading 
                                CAST(N_SALES_B2B_MT_BLIND as int)||' - '||round((N_SALES_B2B_MT_BLIND/N_SALES_ARVALATRADING)*100, 1)||"%" as CONCAT_BLIND_Trading ,
                                CAST(N_SALES_B2B_MT_OPEN as int)||' - '||round((N_SALES_B2B_MT_OPEN/N_SALES_ARVALATRADING)*100, 1)||"%" as CONCAT_OPEN_Trading ,
                                CAST(N_SALES_B2B_MT_BUYNOW as int)||' - '||round((N_SALES_B2B_MT_BUYNOW/N_SALES_ARVALATRADING)*100, 1)||"%" as CONCAT_BUYNOW_Trading ,
                                CAST(N_SALES_B2B_MT_DIRECT as int)||' - '||round((N_SALES_B2B_MT_DIRECT/N_SALES_ARVALATRADING)*100, 1)||"%" as CONCAT_DIRECT_Trading ,                                

                                round(N_BUYERS_B2B_MT/N_CLIENTS_CONNECTED*100,1)||"%" AS BUYER_VS_CONNECTED ,

                                STOCK

                    FROM data as Data_All 

                                       ; """
#DATA_DEALER = pysqldf(DATA_DEALER)


##################################################################################################################################
###    02  -     Create routes (view) for html page and Ajax Jquery                                                                   ###
##################################################################################################################################

###         00.2.6   database and Models:
###         00.2.6.1 Create the database Sqlite (once execute sqlite3.db in the repertory forever)
db = SQLAlchemy()


def init_db(app):
    db.init_app(app)


@app.route('/create_BD/')
def create_DB():
    db.create_all()
    return 'Database created successful'


###   2.0  -     Create engine to connect with Sqlite
engine2 = create_engine("sqlite:///sqlite3.db",
                        encoding='utf-8')  # To final say the SQLAlchemy engine is created with Sqlite3
#DATA_DEALER.to_sql('DATA_DEALER', con=engine2, if_exists='replace', index=False)


###   2.2  -     Show first html home page with overview graphic
@app.route("/")
def DealerKPIS():
    #DATA_DEALER = pd.read_sql("""
    #                            select distinct PORTFOLIOENTITY,
    #                                            DEALER,
    #                                            CAST(N_SALES as int) as N_SALES ,
    #                                            CAST(N_BIDS as int) as N_BIDS
    #
    #                            from DATA_IRIS
    #                            where PORTFOLIOENTITY = {entity}
    #                                ;
    #                            """.format(entity=entity), con=engine2)

    return render_template("DealerKPIS.html",
                           #DATA_DEALER=DATA_DEALER,
                           )




##################################################################################################################################
###    04  -     Run total Project DATAVIZ                                                                                     ###
##################################################################################################################################
import random, threading, webbrowser

if __name__ == "__main__":
    # app.run(host="0.0.0.0", port=8090, debug=True)  # Auto refresh page
    port = 8090
    url = "http://127.0.0.1:{0}".format(port)
    threading.Timer(1.25, lambda: webbrowser.open(url)).start()
    app.run(port=8090, debug=True, use_reloader=False)  # Auto refresh page

    # port=8090, debug=False, use_reloader=False
    # Specific AWS
    # app.run( host="0.0.0.0", port=8090, debug=True)
    # manager.run()
    # In case of coding in Spyder, You may need to copy this link in Google Chrome : http://0.0.0.0:8090/
