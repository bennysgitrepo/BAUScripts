import pandas as pd
import json
import requests
import csv
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

def main():

    data = requests.get('API-HTTP', auth=('UserID', 'Password'))
    data_parsed = data.json()['Inverters']

    with open("inverterDetail.csv", "w") as csvfile:
        f = csv.writer(csvfile)
        f.writerow(["Inverter: Inverter Number",
                    "Inverter Certificate: Manufacturer/Certificate Holder Account",
                    "Inverter Certificate: Brand Name",
                    "Series", 
                    "Model Number", 
                    "Inverter Certificate: Equipment Category", 
                    "Rated Apparent AC Power (VA per port)", 
                    "No of Phases", 
                    "Inverter Certificate: CEC Approved Date", 
                    "Inverter Certificate: Expiry Date", 
                    "Inverter Certificate: Applicable Standards", 
                    "PQ Mode Cos Power", 
                    "PQ Mode Fixed P Factor", 
                    "Power rate limits compliant with 4777.2", 
                    "PQ Mode Volt / Var", 
                    "PQ Mode Volt / Watt", 
                    "PQ Mode Volt Balance", 
                    "Tested to IEC 62116?", 
                    "Active Anti Island Method"])
        for elem in data_parsed:
            f.writerow([elem["InverterNumber"], 
                        elem["Certificate"]["ManufacturerCertificateHolderAccount"]["Details"]["Name"], 
                        elem["Certificate"]["Details"]["BrandName"], 
                        elem["Details"]["Series__c"],
                        elem["Details"]["Model_Number__c"],
                        elem["Certificate"]["Details"]["EquipmentCategory"],
                        elem["Details"]["Nominal_AC_Power_W__c"],
                        elem["Details"]["No_of_Phases__c"],
                        elem["Certificate"]["Details"]["ApprovalDate"],
                        elem["Certificate"]["Details"]["ExpiryDate"],
                        elem["Certificate"]["Details"]["ApplicableStandards"],
                        elem["Details"]["PQ_Mode_Cos_Power__c"],
                        elem["Details"]["PQ_Mode_Fixed_P_Factor__c"],
                        elem["Details"]["Power_rate_limits_compliant_with_4777_2__c"],
                        elem["Details"]["PQ_Mode_Volt_Var__c"],
                        elem["Details"]["PQ_Mode_Volt_Watt__c"],
                        elem["Details"]["PQ_Mode_Volt_Balance__c"],
                        elem["Details"]["Tested_to_IEC_62116__c"],
                        elem["Details"]["Active_Anti_Island_Method__c"]
                        ])

root= tk.Tk()
root.title('Inverter CSV Downloader')
root.resizable(width=False, height=False)

canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'grey', relief = 'raised')
canvas1.pack()

label1 = tk.Label(root, text='Inverter CSV Downloader', bg = 'darkgrey')
label1.config(font=('helvetica', 20))
canvas1.create_window(150, 60, window=label1)
    
saveAsButton_CSV = tk.Button(text="      Download Inverter CSV     ", command=main, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 130, window=saveAsButton_CSV)

def exitApplication():
    MsgBox = tk.messagebox.askquestion ('Exit Application','Are you sure you want to exit the application',icon = 'warning')
    if MsgBox == 'yes':
       root.destroy()
     
exitButton = tk.Button (root, text='       Exit Application     ',command=exitApplication, bg='brown', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 200, window=exitButton)

root.mainloop()
