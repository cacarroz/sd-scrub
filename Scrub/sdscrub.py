
import pandas as pd
import numpy as np

from datetime import datetime
from datetime import timedelta
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog
from tkinter import messagebox
from pytz import timezone

  # Program to filter excel CSV file for Schedule deliveries and pull pertinent columns, write and save a new excel file with the filtered information



root= tk.Tk()

canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'lightsteelblue2', relief = 'raised')
canvas1.pack()

def getCSV ():
    global df

    while True :


    # Chooose what to filter via user input
        filter = simpledialog.askstring(title="Filter selection ", prompt='Filter by AG (Scheduled delivery start date) or  BH (Estimated Arrival Date) ? input 1 or 2 respectively: ')


        if filter == '1':
            x = ('Scheduled delivery start date')
            td = timezone('US/Eastern')
            td = datetime.today() 
            td = td.replace (hour=0,minute=30)
            td= td + timedelta(hours=23) 
            td = td.strftime('%m-%d-%Y %H:%M:%S')
	    

            start_date = '03-01-1996'
            end_date = td

            import_file_path = filedialog.askopenfilename()
            data = pd.read_csv (import_file_path)
            df = pd.DataFrame(data, columns= ['Tracking Id','Ship Option','Scheduled delivery start date',])



            #Change dates to correct format
            df['Scheduled delivery start date'] = pd.to_datetime(df['Scheduled delivery start date'])
           
        #Filter Schedule Delivery column
            contain_values = df[df['Ship Option'].str.contains('SCHEDULED_DELIVERY', na=False)]


            #print (contain_values)
            print(td)


        #Filter dates
            mask = (contain_values[x] > start_date) & (contain_values[x] <= end_date)

            contain_values = contain_values.loc[mask]

        # Create a Pandas Excel writer using XlsxWriter as the engine.
            writer = pd.ExcelWriter('SdscrubAG.xlsx', engine='xlsxwriter')

        # Convert the dataframe to an XlsxWriter Excel object.
            contain_values.to_excel(writer, sheet_name='Sheet1')

        # Close the Pandas Excel writer and output the Excel file.

            writer.save()
            print('Filter succesful saved as Sdscrub')
            messagebox.showinfo("Information","Filter succesful saved as SDscrubAG")


            break

        if filter == '2':
            x = ('Estimated Arrival Date')
            td = timezone('US/Eastern')
            td = datetime.today()
            td = td.replace (hour=0,minute=30)
            td= td + timedelta(hours=23) 
            td = td.strftime('%m-%d-%Y %H:%M:%S')
            


            start_date = '03-01-1996'
            end_date = td

            #
            import_file_path = filedialog.askopenfilename()
            data = pd.read_csv (import_file_path)

            #Select columns on pdf file
            df = pd.DataFrame(data, columns= ['Tracking Id','Ship Option','Estimated Arrival Date',])



            #Change dates to correct format
           
            df['Estimated Arrival Date'] = pd.to_datetime(df['Estimated Arrival Date'])
        #Filter Schedule Delivery column
            contain_values = df[df['Ship Option'].str.contains('SCHEDULED_DELIVERY', na=False)]


            #print (contain_values)
            #print(end_date)
            print(td)

        #Filter dates
            mask = (contain_values[x] > start_date) & (contain_values[x] <= end_date)

            contain_values = contain_values.loc[mask]

        # Create a Pandas Excel writer using XlsxWriter as the engine.
            writer = pd.ExcelWriter('SdscrubBH.xlsx', engine='xlsxwriter')

        # Convert the dataframe to an XlsxWriter Excel object.
            contain_values.to_excel(writer, sheet_name='Sheet1')

        # Close the Pandas Excel writer and output the Excel file.

            writer.save()
            print('Filter succesful saved as Sdscrub')
            messagebox.showinfo("Information","Filter succesful saved as SDscrubBH")
            break

        if filter != '1' or '2' :
            print('Invalid input')
            continue

browseButton_CSV = tk.Button(text="      Import CSV File     ", command=getCSV, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 150, window=browseButton_CSV)





root.mainloop()
