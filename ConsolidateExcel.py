import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import concurrent.futures

# Function to open file dialog and select a CSV file
def browseFile():
    filePath = filedialog.askopenfilename(
        title="Select the CSV file",
        filetypes=[("CSV Files", "*.csv"), ("All Files","*.*")], initialdir = os.getcwd()
    )
    if filePath:
        filePathVariable.set(filePath)

def browseDirectory():
    directoryPath = filedialog.askdirectory()
    if directoryPath:
        directoryPathVariable.set(directoryPath)

def checkNumberOfFiles(numberRequired):
    directoryPath = directoryPathVariable.get()
    numberOfFiles = len([i for i in os.listdir(directoryPath) if os.path.isfile(os.path.join(directoryPath, i))])
    response = True
    if numberRequired != numberOfFiles:
        response = messagebox.askyesno("File Count Warning",
            f"The directory has {numberOfFiles} files instead of {numberRequired}.\nDo you want to continue?")
    return response, numberRequired

# Total Only
def prepareTotalByUnitName(workbook):
    workbookTotal = workbook.drop(['F-Description  of Services'], axis=1)
    workbookTotal = workbookTotal.groupby(['A-Unit Name']).sum()
    workbookTotal.to_csv('./Reverse Charges Files/RCMTotalOnly.csv')

# By Service
def prepareTotalByUnitNameAndService(workbook):
    workbookByService = workbook.groupby(['A-Unit Name','F-Description  of Services']).sum()
    workbookByService.to_csv('./Reverse Charges Files/RCMTotalyByService.csv')

# By Service with Subtotal
def prepareTotalByUnitNameAndServiceWithSubtotal(workbook):
    workbookByService = workbook.groupby(['A-Unit Name','F-Description  of Services']).sum().reset_index()

    subtotal = workbookByService.groupby('A-Unit Name').sum()
    subtotal['F-Description  of Services'] = 'Subtotal'
    subtotal = subtotal.reset_index()

    result = pd.concat([workbookByService, subtotal], ignore_index=True)
    result = result.sort_values(by=['A-Unit Name','F-Description  of Services'], key=lambda col: col.map({'Subtotal': 'zzzz'}).fillna(col)).reset_index(drop=True)
    result.to_csv('./Reverse Charges Files/RCMTotalyByServiceWithSubtotal.csv', index=False)

# generating 5% and 18% and Tax Columns
def generate5And18TaxColumns(workbook):
    fiveSlab = []
    eighteenSlab = []

    for rate, taxVal in zip(workbook['H-Rate'], workbook['G-Taxable Value']):
        if rate == 5:
            fiveSlab.append(taxVal)
            eighteenSlab.append(0)
        else:
            fiveSlab.append(0)
            eighteenSlab.append(taxVal)

    workbook = workbook.drop(['H-Rate'], axis=1)
    workbook["5%"] = fiveSlab
    workbook["18%"] = eighteenSlab
    return workbook

def separateExcelWorksheets(pathToFile):
    pathAsList = pathToFile.split("/")
    sheetsToUse = ["Summary", "01 Outward Supply", "02 Reverse Charges", "03 GST-TDS", "04 Inward Supplies (ITC)", "05 Debit & Credit Note"]
    # sheetsToUse = ["Summary", "01 Tax Invoice Outward", "02 Bill Of Supply Outward", "03 Reverse Charges", "04 GST-TDS", "05 Inward Supplies", "06 Debit & Credit Note"]

    for sheet in sheetsToUse:
        try:
            df = pd.read_excel(pathToFile,sheet_name=sheet, skiprows=3)
        except Exception as e:
            messagebox.showerror("Error", f"Problem in processing File. \n{str(e)}")
            return

        df = df.dropna(subset=[df.columns[0]])
        print(f"./Consolidated Files/{pathAsList[-1][:-5]}.{sheet}.csv")
        df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
        df.to_csv(f"./Consolidated Files/{pathAsList[-1][:-5]}.{sheet}.csv", index=False)

def listFilesRecursive(path):
    pool = concurrent.futures.ProcessPoolExecutor(max_workers=os.cpu_count()-1)
    for entry in os.listdir(path):
        fullPath = path + "/" + entry
        print(fullPath)
        if os.path.isdir(fullPath):
            listFilesRecursive(fullPath)
        else:
            if fullPath[-5:] == '.xlsm' or fullPath[-5:] == '.xlsb' or fullPath[-5:] == '.xlxs':
                pool.submit(separateExcelWorksheets, fullPath)
    pool.shutdown(wait=True)

def reverseChargesFile():
    browseFile()
    filePath = filePathVariable.get()
    if not filePath:
        messagebox.showerror("Error", "Please select a file first.")
        return

    try:
        os.makedirs("./Reverse Charges Files", exist_ok=True)

        workbook = pd.read_csv(filePath)
        workbook = workbook.drop(['B-Name of Firm', 'C-Invoice Number Generate By Unit', 'D-Invoice date', 'E-Date of Payment'], axis=1)
        workbook = generate5And18TaxColumns(workbook)

        prepareTotalByUnitName(workbook)
        prepareTotalByUnitNameAndService(workbook)
        prepareTotalByUnitNameAndServiceWithSubtotal(workbook)

        print("Success")
        messagebox.showinfo("Success", "Required CSVs have been generated and are in the Reverse Charges Files folder")

    except Exception as e:
        messagebox.showerror("Error", f"Problem in processing File. \n{str(e)}")

def GSTConsolidation():
    browseFile()
    filePath = filePathVariable.get()
    if not filePath:
        messagebox.showerror("Error", "Please select a file first.")
        return

    try:
        os.makedirs("./GST-TDS Consolidation Files", exist_ok=True)

        dfGST = pd.read_csv(filePath)
        dfGST = dfGST.drop('D-Date of Payment', axis=1)
        dfGST['B-GST No of Supplier'] = dfGST['B-GST No of Supplier'].str.strip()

        columnsToFloat = ['E-Taxable Amount Paid', 'F-TDS-IGST', 'G-TDS-CGST', 'H-TDS-SGST', 'I-Total']
        for columns in columnsToFloat:
            dfGST[columns] = dfGST[columns].replace('NIL', '0', regex=True)
            dfGST[columns] = dfGST[columns].astype('float')

        aggregationFunction = {
            col: ('first' if (col == 'C-Name of Supplier' or col == 'A-Unit Name') else 'sum')
            for col in dfGST.columns if col != 'B-GST No of Supplier'
        }

        dfByGSTAndName = dfGST.groupby(['A-Unit Name','B-GST No of Supplier'], as_index=False).agg(aggregationFunction)
        dfByGSTAndName = dfByGSTAndName.sort_values(by='A-Unit Name')
        dfByGSTAndName.to_csv("./GST-TDS Consolidation Files/GST-TDSandName.csv", index=False)

        dfByGST = dfGST.groupby(['B-GST No of Supplier'], as_index=False).agg(aggregationFunction)
        dfByGST = dfByGST.drop('A-Unit Name', axis=1)
        dfByGST.to_csv("./GST-TDS Consolidation Files/GST-TDSOnly.csv", index=False)

        print("Success")
        messagebox.showinfo("Success", "Required CSVs have been generated and are in the GST Consolidation Files folder")

    except Exception as e:
        messagebox.showerror("Error", f"Problem in processing File. \n{str(e)}")

def inwardInvoiceMatching():
    messagebox.showinfo("Locate File (Excel Worksheet Format)", "Locate the 2B File")
    browseFile()
    filePath2B = filePathVariable.get()
    if not filePath2B:
        messagebox.showerror("Error", "Please select a file first.")
        return

    try:
        df2B = pd.read_excel(filePath2B, sheet_name="B2B", skiprows=4)
        df2B.rename(columns={'Invoice Details': 'Invoice Number', 'Unnamed: 3': 'Invoice Type', 'Unnamed: 4': 'Invoice Date', 'Unnamed: 5': 'Invoice Value', 'Tax Amount': 'Integrated Tax', 'Unnamed: 10': 'Central Tax', 'Unnamed: 11' : 'State/UT Tax', 'Unnamed: 12' : 'Cess'}, inplace=True)
        df2B.drop(index=0, inplace=True)
        df2B['Taxable Value (₹)'] = df2B['Taxable Value (₹)'].astype('float64')
        df2B['Invoice Number'] = df2B['Invoice Number'].astype('string')

    except Exception as e:
        messagebox.showerror("Error", f"Problem in processing 2B File. \n{str(e)}")
        return

    messagebox.showinfo("Locate File (CSV Format)", "Locate the Combined Inward Supplies File")
    browseFile()
    filePathInwardSupply = filePathVariable.get()
    if not filePathInwardSupply:
        messagebox.showerror("Error", "Please select a file first.")
        return

    try:
        csvInwardSupply = pd.read_csv(filePathInwardSupply)
        csvInwardSupply['D-Invoice No.'] = csvInwardSupply['D-Invoice No.'].astype('string')
        csvInwardSupply['G-Taxable Value'] = csvInwardSupply['G-Taxable Value'].astype('float64')

    except Exception as e:
        messagebox.showerror("Error", f"Problem in processing Inward Supply File. \n{str(e)}")
        return

    try:
        os.makedirs("./ITC Files", exist_ok=True)

        dfInvoiceAmount = pd.merge(df2B, csvInwardSupply, left_on=['Invoice Number', 'Taxable Value (₹)'], right_on=['D-Invoice No.','G-Taxable Value'], how='inner')
        dfInvoiceAmount = dfInvoiceAmount[['GSTIN of supplier', 'Trade/Legal name', 'Invoice Number', 'D-Invoice No.', 'Taxable Value (₹)', 'G-Taxable Value', 'A-Unit Name']]
        dfInvoiceAmount.to_csv('./ITC Files/ITCInvoiceAndAmountMatched.csv', index=False)

        dfInvoice = pd.merge(df2B, csvInwardSupply, left_on=['Invoice Number'], right_on=['D-Invoice No.'], how='outer', indicator=True)
        dfInvoice = dfInvoice[['GSTIN of supplier', 'Trade/Legal name', 'Invoice Number', 'D-Invoice No.', 'Taxable Value (₹)', 'G-Taxable Value', 'A-Unit Name', '_merge']]
        dfInvoice[dfInvoice['_merge'] == 'left_only'].to_csv('./ITC Files/ITC2BOnlyInvoice.csv', index=False)
        dfInvoice[dfInvoice['_merge'] == 'right_only'].to_csv('./ITC Files/ITCDivisionOnlyInvoice.csv', index=False)

        InvoiceOnlyArray = dfInvoiceAmount['D-Invoice No.'].to_numpy()
        dfInvoice = dfInvoice.query("`D-Invoice No.` not in @InvoiceOnlyArray")
        dfInvoice = dfInvoice.dropna(subset=['D-Invoice No.'])
        dfInvoice.to_csv("./ITC Files/ITCInvoiceMatchAmountMismatch.csv",index=False)

        print("Success")
        messagebox.showinfo("Success", "Required CSVs have been generated and are in the Invoice Matching Files folder")

    except Exception as e:
        messagebox.showerror("Error", f"Problem in matching data. \n{str(e)}")

def unitConsolidation():
    browseDirectory()
    directoryPath = directoryPathVariable.get()
    if not directoryPath:
        messagebox.showerror("Error", "Please select a file first.")
        return

    if not checkNumberOfFiles(34):
        return

    try:
        os.makedirs("./Consolidated Files", exist_ok=True)

        listFilesRecursive(directoryPath)

        stringsToCombine = ["Summary", "Outward Supply", "Reverse Charges", "GST-TDS", "Inward Supplies (ITC)", "Debit & Credit Note"]
        for string in stringsToCombine:
            # os.system("copy " + string + "\".csv" + " \"Combined " + string + "\".csv")
            print("copy \".\\Consolidated Files\\*" + string + ".csv\" \".\\Consolidated Files\\Combined " + string + ".csv\"")
            os.system("copy \".\\Consolidated Files\\*" + string + ".csv\" \".\\Consolidated Files\\Combined " + string + ".csv\"")

        print("Success")
        messagebox.showinfo("Success", "Required CSVs have been generated and are in the Consolidated Files folder")

    except Exception as e:
        messagebox.showerror("Error", f"Problem in matching data. \n{str(e)}")

def outwardSupplyProcessing():
    browseFile()
    filePath = filePathVariable.get()
    if not filePath:
        messagebox.showerror("Error", "Please select a file first.")
        return

    try:
        os.makedirs("./Outward Supply Files", exist_ok=True)

        dfAll = pd.read_csv(filePath)

        columnsToFloat = ['J-Taxable Value included Mandi & Excluded TCS', 'K-IGST', 'L-CGST', 'M-SGST', 'N-Total Tax']
        for columns in columnsToFloat:
            dfAll[columns] = dfAll[columns].replace('NIL', '0', regex=True)
            dfAll[columns] = dfAll[columns].replace('-', '0', regex=True)
            dfAll[columns] = dfAll[columns].astype('float')

        dfB2B = dfAll.loc[dfAll['B-GSTIN/UIN of Recipient'].str.len() == 15]
        dfB2BTaxable = dfB2B.loc[dfB2B['N-Total Tax']>0]
        dfB2BNil = dfB2B.loc[dfB2B['N-Total Tax']==0]
        dfB2BTaxable.to_csv("./Outward Supply Files/OSB2BTaxable.csv", index=False)
        dfB2BNil.to_csv("./Outward Supply Files/OSB2BNil.csv", index=False)

        dfB2C = dfAll.loc[dfAll['B-GSTIN/UIN of Recipient'].str.len() != 15]
        dfB2CTaxable = dfB2C.loc[dfB2C['N-Total Tax']>0]
        dfB2CNil = dfB2C.loc[dfB2C['N-Total Tax']==0]
        dfB2CTaxable.to_csv("./Outward Supply Files/OSB2CTaxable.csv", index=False)
        dfB2CNil.to_csv("./Outward Supply Files/OSB2CNil.csv", index=False)
        dfB2CNilByGroup = dfB2CNil.groupby('A-UNIT NAME', as_index=False).sum()
        dfB2CNilByGroup.to_csv("./Outward Supply Files/OSB2CNilByGroup.csv", index=False)

        dfNil = pd.concat([dfB2BNil, dfB2CNil], axis=0, ignore_index=True)

        dfB2BTaxable = dfB2BTaxable.drop(['B-GSTIN/UIN of Recipient', 'C-Receiver Name', 'D-Invoice Number', 'E-Item wise Description  of Goods', 'F-Invoice date', 'G-Invoice Value', 'H-HSN Code', 'I- Rate'], axis=1)
        dfB2CTaxable = dfB2CTaxable.drop(['B-GSTIN/UIN of Recipient', 'C-Receiver Name', 'D-Invoice Number', 'E-Item wise Description  of Goods', 'F-Invoice date', 'G-Invoice Value', 'H-HSN Code', 'I- Rate'], axis=1)
        dfNil = dfNil.drop(['B-GSTIN/UIN of Recipient', 'C-Receiver Name', 'D-Invoice Number', 'E-Item wise Description  of Goods', 'F-Invoice date', 'G-Invoice Value', 'H-HSN Code', 'I- Rate'], axis=1)

        dfB2BTaxable = dfB2BTaxable.groupby('A-UNIT NAME', as_index=False).sum()
        dfB2CTaxable = dfB2CTaxable.groupby('A-UNIT NAME', as_index=False).sum()
        dfNil = dfNil.groupby('A-UNIT NAME', as_index=False).sum()

        dfB2BTaxable.rename(columns={'J-Taxable Value included Mandi & Excluded TCS': 'B2B'}, inplace=True)
        dfB2CTaxable.rename(columns={'J-Taxable Value included Mandi & Excluded TCS': 'B2C'}, inplace=True)
        dfNil.rename(columns={'J-Taxable Value included Mandi & Excluded TCS': 'Nil'}, inplace=True)

        dfFinal = pd.concat([dfB2BTaxable, dfB2CTaxable, dfNil]).groupby(['A-UNIT NAME']).sum()
        dfFinal['B2B+B2C Total'] = dfFinal['B2B'] + dfFinal['B2C']
        dfFinal['B2B+B2C+Nil Total']= dfFinal['B2B+B2C Total'] + dfFinal['Nil']

        dfFinal.to_csv('./Outward Supply Files/OSAdvice.csv')

        print("Success")
        messagebox.showinfo("Success", "Required CSVs have been generated and are in the Invoice Matching Files folder")

    except Exception as e:
        messagebox.showerror("Error", f"Problem in matching data. \n{str(e)}")

def outwardSupplyMatching():
    messagebox.showinfo("Locate File (Excel Worksheet Format)", "Locate the UKFDC Software File")
    browseFile()
    UKSoftFilePath = filePathVariable.get()
    if not UKSoftFilePath:
        messagebox.showerror("Error", "Please select a file first.")
        return
    try:
        dfUKSoftFile = pd.read_excel(UKSoftFilePath, sheet_name="Sheet1", skiprows=4)
        dfUKSoftFile['Taxable Value'] = dfUKSoftFile['Total Amount Which Tax will be Calculated'].astype('float64')
        dfUKSoftFile['Invoice No.'] = dfUKSoftFile['Invoice No.'].astype('string')
        columnsToFloat = ['Taxable Value']
        for columns in columnsToFloat:
            dfUKSoftFile[columns] = dfUKSoftFile[columns].replace('NIL', '0', regex=True)
            dfUKSoftFile[columns] = dfUKSoftFile[columns].replace('-', '0', regex=True)
            dfUKSoftFile[columns] = dfUKSoftFile[columns].astype('float')

    except Exception as e:
        messagebox.showerror("Error", f"Problem in processing 2B File. \n{str(e)}")
        return

    messagebox.showinfo("Locate File (CSV Format)", "Locate the Combined Outward Supplies File")
    browseFile()
    filePathOutwardSupply = filePathVariable.get()
    if not filePathOutwardSupply:
        messagebox.showerror("Error", "Please select a file first.")
        return

    try:
        os.makedirs("./Outward Supply Matched Files", exist_ok=True)
        dfOS = pd.read_csv(filePathOutwardSupply)
        columnsToFloat = ['J-Taxable Value included Mandi & Excluded TCS', 'K-IGST', 'L-CGST', 'M-SGST', 'N-Total Tax']
        for columns in columnsToFloat:
            dfOS[columns] = dfOS[columns].replace('NIL', '0', regex=True)
            dfOS[columns] = dfOS[columns].replace('-', '0', regex=True)
            dfOS[columns] = dfOS[columns].astype('float')

        dfInvNoAndValueMatch = pd.merge(dfUKSoftFile, dfOS, left_on=['Invoice No.', 'Taxable Value'], right_on=['D-Invoice Number','J-Taxable Value included Mandi & Excluded TCS'], how='outer', indicator=True)
        dfInvNoAndValueMatch = dfInvNoAndValueMatch[['B-GSTIN/UIN of Recipient', 'C-Receiver Name', 'Invoice No.', 'D-Invoice Number', 'Taxable Value', 'J-Taxable Value included Mandi & Excluded TCS', 'A-UNIT NAME' , 'K-IGST', 'L-CGST', 'M-SGST', 'N-Total Tax', '_merge']]
        dfInvNoAndValueMatch[dfInvNoAndValueMatch['_merge'] == 'left_only'].to_csv("./Outward Supply Matched Files/UKSoftOnly.csv")
        dfInvNoAndValueMatch[dfInvNoAndValueMatch['_merge'] == 'right_only'].to_csv("./Outward Supply Matched Files/CombinedOSOnly.csv")
        dfInvNoAndValueMatch[dfInvNoAndValueMatch['_merge'] == 'both'].to_csv("./Outward Supply Matched Files/CommonOSUKSoft.csv")

        # dfInvoice = pd.merge(df2B, csvInwardSupply, left_on=['Invoice Number'], right_on=['D-Invoice No.'], how='outer', indicator=True)
        # dfInvoice = dfInvoice[['GSTIN of supplier', 'Trade/Legal name', 'Invoice Number', 'D-Invoice No.', 'Taxable Value (₹)', 'G-Taxable Value', 'A-Unit Name', '_merge']]
        # dfInvoice[dfInvoice['_merge'] == 'left_only'].to_csv('./ITC Files/ITC2BOnlyInvoice.csv', index=False)
        # dfInvoice[dfInvoice['_merge'] == 'right_only'].to_csv('./ITC Files/ITCDivisionOnlyInvoice.csv', index=False)

        # InvoiceOnlyArray = dfInvoiceAmount['D-Invoice No.'].to_numpy()
        # dfInvoice = dfInvoice.query("`D-Invoice No.` not in @InvoiceOnlyArray")
        # dfInvoice = dfInvoice.dropna(subset=['D-Invoice No.'])
        # dfInvoice.to_csv("./ITC Files/ITCInvoiceMatchAmountMismatch.csv",index=False)

        print("Success")
        messagebox.showinfo("Success", "Required CSVs have been generated and are in the Invoice Matching Files folder")

    except Exception as e:
        messagebox.showerror("Error", f"Problem in matching data. \n{str(e)}")
def closeApp():
    root.quit()

if __name__ == '__main__':

    root = tk.Tk()
    root.title("GST, Reverse Charge, Inward Invoice, Full Consolidation")
    root.geometry("600x500")

    filePathVariable = tk.StringVar()
    directoryPathVariable = tk.StringVar()

    reverseChargesButton = tk.Button(root, text="Reverse Charges Sheets", command=reverseChargesFile)
    reverseChargesButton.pack(pady=5)

    GSTButton = tk.Button(root, text="GST Consolidation", command=GSTConsolidation)
    GSTButton.pack(pady=5)

    inwardInvoiceButton = tk.Button(root, text="Inward Invoice", command=inwardInvoiceMatching)
    inwardInvoiceButton.pack(pady=5)

    consolidationButton = tk.Button(root, text="Consolidate Excels", command=unitConsolidation)
    consolidationButton.pack(pady=5)

    outwardSupplyButton = tk.Button(root, text="Outward Supply Processing", command=outwardSupplyProcessing)
    outwardSupplyButton.pack(pady=5)

    outwardSupplyButton = tk.Button(root, text="Outward Supply Matching", command=outwardSupplyMatching)
    outwardSupplyButton.pack(pady=5)

    closeButton = tk.Button(root, text="Close", command=closeApp)
    closeButton.pack(pady=5)

    root.mainloop()
