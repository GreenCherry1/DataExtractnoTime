import schedule
import time
import pyodbc
import csv
from datetime import datetime,date
from ftplib import FTP_TLS
import ctypes

confirm=ctypes.windll.user32.MessageBoxW(0, "Please confirm adding the daily sync task to the Windows Scheduler", "An action needed", 1)
if confirm==1:
    All=ctypes.windll.user32.MessageBoxW(0, "Contract data for all days?", "An action needed", 4)
    Save=ctypes.windll.user32.MessageBoxW(0, "Save files?", "An action needed", 4)
    
    # for this driver to connect 32-bytes python is nececeary 
    # Connect to the database
    conn = pyodbc.connect(r'Driver=Microsoft Access Driver (*.mdb, *.accdb);DBQ=C:\Shiprite\Shiprite.mdb;')
    cursor = conn.cursor()
    
    today = datetime.now()
    n=f"{today.year}.{today.month}.{today.day}"
    print(All)
    # Extract the data from the table
    if All==6: cursor.execute("SELECT * FROM manifest WHERE Date= ?",n)
    else : cursor.execute("SELECT * FROM manifest")
    rows = cursor.fetchall()
    Cnum = conn.cursor()
    Cnum.execute("SELECT PostNetCDHL_CenterID FROM Setup2 WHERE id= 1")
    roz = Cnum.fetchall()
    # get current date and time
    current_datetime = today.strftime("%m-%d-%y")
    name=f'{roz[0][0]}_Manifest_{current_datetime}.csv'
    path=f'C:\\Shiprite\\'+name

    # convert datetime obj to string
    str_current_datetime = str(current_datetime)
    # Save the data to a CSV file5
    with open(path, "w", encoding="utf-8", newline='') as file:
        writer = csv.writer(file)
        writer.writerow([i[0] for i in cursor.description]) # write headers
        writer.writerows(rows)

    # Close the connection to the database
    conn.close()
    print("Table exported to CSV successfully")

    ftp = FTP_TLS()
    ftp.set_debuglevel(2)
    ftp.connect('ftp.dash-ipostnet.com')
    ftp.login(user='manifest_file_upload@dash-ipostnet.com', passwd='SC{{m@$PZ2;L')
    print(ftp.getwelcome())
    ftp.storbinary('STOR '+name, open(path,'rb'))
    print(ftp.dir())
    ftp.close()






''' 

Ще 1 момент. 
2) додати параметр /savefiles з яким не будемо выдаляти файл після завантаження на FTP, запуск без параметру файл видаляє як є зараз'''
