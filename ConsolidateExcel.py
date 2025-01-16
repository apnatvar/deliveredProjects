import os, sys
from argparse import ArgumentParser
import pandas as pd
import concurrent.futures
import logging

logging.basicConfig(level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s')

def loggingForProgram(enableLogging):
    logging.debug('Log Argument(Confirmation): ' + str(enableLogging))
    if ~enableLogging:
        logging.disable(logging.CRITICAL)

def renameFiles(pathToFile):
    pathAsList = pathToFile.split(os.sep)
    pathAsList.pop()
    pathAsList = pathAsList + [pathAsList[-1]+".xlsx"]
    newPathToFile = os.sep.join(pathAsList)
    os.rename(pathToFile,newPathToFile)
    logging.debug('New Name: ', newPathToFile)
    logging.debug('Old Name: ', pathToFile)
    return newPathToFile

def separateExcelWorksheets(pathToFile, sheetNames):
    workbook = pd.ExcelFile(pathToFile)
    pathAsList = pathToFile.split(os.sep)
    for sheet in workbook.sheet_names:
        if sheet in sheetNames:
            df = pd.read_excel(workbook,sheet_name=sheet, skiprows=3)
            df = df.dropna(subset=[df.columns[0]])
            logging.debug("path to file: " + f"{pathAsList[-1][:-5]}.{sheet}.csv")
            logging.debug('dataframe head before dropping unnecessary columns:', df.head())
            df = df.loc[:, ~df.columns.str.startswith('Unnamed')] # deleting unnamed columns are created because of compatibilty issues between certain excel features and pandas features.
            logging.debug('dataframe head after dropping unnecessary columns:\n**Helps check in any important columns were accidentally dropped**', df.head())
            df.to_csv(f"{pathAsList[-1][:-5]}.{sheet}.csv", index=False)

def findFilesToSeparate(args, sheetNames, path='.'):
    pool = concurrent.futures.ThreadPoolExecutor()
    for entry in os.listdir(path):
        fullPath = os.path.join(path, entry)
        logging.debug('Searching in ', fullPath)
        if os.path.isdir(fullPath):
            findFilesToSeparate(args=args, path=fullPath, sheetNames=sheetNames)
        else:
            if fullPath[-5:] == '.xlsx':
                if args.rename:
                    logging.debug('Rename Argument(Confirmation): ' + str(args.rename))
                    fullPath = renameFiles(fullPath)
                logging.debug('file found with path: ',fullPath)
                pool.submit(separateExcelWorksheets, fullPath, sheetNames) # opening and retrieving data from excel sheets takes a lot of time so with multi-threading the program is able to do multiple files concurrently
    pool.shutdown(wait=True)
    return True

def combineSeparatedSheets(sheetNames):
    for sheet in sheetNames:
        os.system("copy *\"" + sheet + "\".csv" + " \"Combined " + sheet + "\".csv")

def compStep(w1, w2, sheetNames): # merges based on requirement submitted
    dfStep = pd.merge(w1, w2, on=sheetNames, how='left', indicator=True)
    dfStepMatched = dfStep[dfStep['_merge'] == 'both']
    dfStepUnmatchedw1 = dfStep[dfStep['_merge'] == 'left_only']
    dfStepMatched = dfStepMatched.drop(columns=['_merge'])
    dfStepUnmatchedw1 = dfStepUnmatchedw1.drop(columns=['_merge'])
    logging.debug('Step ' + str(StepNo) + ' Full Match', dfStep.head())
    logging.debug('Step ' + str(StepNo) + 'Inner Match', dfStepMatched.head())
    logging.debug('Step ' + str(StepNo) + ' Left Only Match', dfStepUnmatchedw1.head())
    dfStepMatched.to_csv("Step" + str(StepNo) + "Matched.csv", index=False)
    dfStepUnmatchedw1.to_csv("Step" + str(StepNo) + "Unmatched.csv", index=False)

def compareSheets(sheetNames):
    pathToWorkbook1 = input("Enter Path/Name of First Workbook\n")
    pathToWorkbook2 = input("Enter Path/Name of Second Workbook\n")
    logging.debug('Path to both workbooks', pathToWorkbook1, pathToWorkbook2)
    workbook1 = pd.ExcelFile(pathToWorkbook1)
    workbook2 = pd.ExcelFile(pathToWorkbook2)
    logging.debug('Dataframe Heads', workbook1.head(), workbook2.head())
    for i in range(3):
        compStep(workbook1, workbook2, sheetNames[0:i],i+1)


parser = ArgumentParser(description="commnad line based tool to extract worksheets from several excel books")
parser.add_argument("--rename", type=bool, help="Choose to rename the files or not to the name of the directory they are currenty in", default=False)
parser.add_argument("--comparison", type=bool, help="Choose to run comparisons between consolidated files", default=False)
parser.add_argument("--log", type=bool, help="Choose to run the program with logging turned on/off", default=False)

args = parser.parse_args()
logging.debug('Rename Argument(Regular Check): ' + str(args.rename))
logging.debug('Comparison Argument(Regular Check): ' + str(args.comparison))
logging.debug('Log Argument(Regular Check): ' + str(args.log))

loggingForProgram(args.log)

sheetNames = ["Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4", "Sheet 5", "Sheet 6"] # sheet names changed before uploading to repo

if args.comparison:
    logging.debug('Comparison Argument(Confirmation): ' + str(args.comparison))
    compareSheets(sheetNames[1:3])
else:
    if findFilesToSeparate(args=args, path='.', sheetNames=sheetNames):
        combineSeparatedSheets(sheetNames)
