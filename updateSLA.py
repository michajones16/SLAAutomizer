import traceback
from datetime import datetime
import shutil
import os
import pandas as pd
from collections import Counter

# Key variables
templatePath = r"C:\Users\wywyguy\Box\Accessibility Shared Folder\Team Member Folders\Wyatt\Programs\SLA Automation\0-Template - YYYY-MM - SLA Report.xlsx"
logPath = r"C:\Users\wywyguy\Box\Accessibility Shared Folder\Team Member Folders\Wyatt\Programs\SLA Automation\SLA Update Program Log.txt"

def log(message):
    with open(logPath, "a", encoding="utf-8") as file:
        timestamp = datetime.now().strftime("%m/%d/%Y %H:%M:%S")
        if message == "BEGIN EXECUTION":
            file.write("\n\n\n")
        file.write(f"{timestamp}: {message}\n")

# Determine which data is which and rename files accordingly
def determineData(paths, now):
    keywords = {
        "prototype": "Prototypes",
        "50%": "50s",
        "psia": "PSIAs",
        "peer": "Peer Verifications"
    }

    fileLabels = {}
    labelCounts = Counter()
    undeterminedFiles = []

    for path in paths:
        if not os.path.exists(path):
            fileLabels[path] = None
            undeterminedFiles.append(path)
            continue

        try:
            df = pd.read_excel(path, engine='openpyxl')
            if 'I' not in df.columns and 8 not in df.columns:
                col_i = df.iloc[:, 8]
            else:
                col_i = df['I'] if 'I' in df.columns else df.iloc[:, 8]
            
            termCounter = Counter()
            for val in col_i.dropna().astype(str):
                for term in keywords:
                    if term.lower() in val.lower():
                        termCounter[term] += 1
            
            if termCounter:
                mostCommonTerm = termCounter.most_common(1)[0][0]
                label = keywords[mostCommonTerm]
                fileLabels[path] = label
                labelCounts[label] += 1
            else:
                fileLabels[path] = None
                undeterminedFiles.append(path)
            
        except Exception as e:
            print(f"Error reading {path}: {e}")
            fileLabels = None
            undeterminedFiles.append(path)

    knownLabels = [label for label in fileLabels.values() if label]
    if len(set(knownLabels)) != len(knownLabels):
        raise ValueError("Duplicate file types detected among existing files.")
    if len(fileLabels) != 4:
        raise ValueError("Total number of files (known + undetermined) must be 4.")

    yearMonth = now.strftime("%Y-%m")
    renamedFiles = {}

    for path, label in fileLabels.items():
        if label:
            newName = f"{yearMonth} - SLA {label}.xlsx"
            renamedFiles[path] = newName
            os.rename(path, os.path.join(os.path.dirname(path), newName))

    return renamedFiles, undeterminedFiles


log("BEGIN EXECUTION")
try:
    # Step 1: Copy the SLA report template
    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
    os.chdir(downloads_path)
    now = datetime.now()
    saveAsName=f"{now:%Y-%m} - SLA Report.xlsx"
    shutil.copy2(templatePath, saveAsName)

    # Step 2: Rename downloaded files according to contents of column I
    paths=[f"All Tasks Report - {now.strftime('%d %b %Y')}.xlsx", f"All Tasks Report - {now.strftime('%d %b %Y')} (1).xlsx", f"All Tasks Report - {now.strftime('%d %b %Y')} (2).xlsx", f"All Tasks Report - {now.strftime('%d %b %Y')} (3).xlsx"]
    determineData(paths, now)

    # Step 3: Clean data and retrieve dates as needed


    # Step 4: Copy relevant data to the monthly report (for the removing 0s step, could we modify the formula to only show up if the fields are inputted? What about fitting the table to the data?)


    # Step 5: Add data to the main report? Formatting?


    # https://byu.instructure.com/courses/1026/pages/sla-update-tutorial

except Exception as e:
    log(f"ERROR: {e}")
    log(traceback.format_exc())