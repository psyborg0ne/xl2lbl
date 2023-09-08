#                     __                      ______              
# .-----.-----.--.--.|  |--.-----.----.-----.|      |.-----.-----.
# |  _  |__ --|  |  ||  _  |  _  |   _|  _  ||  --  ||     |  -__|
# |   __|_____|___  ||_____|_____|__| |___  ||______||__|__|_____|#4689
# |__|        |_____|                 |_____| xl2lbl.py
# 
# Licensed under the terms of the LICENSE file included in this repo.
# Created On   : Fr, September 8th 2023, 12:19:35
# Last Modified: Fr, September 8th 2023, 13:48:53

import pandas as pd
import sys
import os

def readFromFile(_file = 'labels.xlsx', _cols= 'A:E'):
    try:
        df = pd.read_excel(_file, usecols=_cols, engine="openpyxl")
        return df
    except:
        print("Access denied. Please make sure you have saved and closed the excel file before running this script.")
        input("Press any key to exit...")
        sys.exit()

def writeLabels(_dataFrame, _enFile='en-US.txt', _itFile='it-IT.txt'):
    try:
        with open(_enFile, 'w') as lblEnFile, open(_itFile, 'w') as lblItFile:
            lineCount_en = 0
            helpCount_en = 0

            lineCount_it = 0
            helpCount_it = 0

            for row in _dataFrame.itertuples(index=False):
                # English
                line_en = f"{row[0].replace(' ', '')}={row[1]}\r"
                comm_en = f" ;{row[3]}\n"
                print(f"[{_enFile}] <- {line_en}")
                try:
                    lblEnFile.write(line_en)
                    lblEnFile.write(comm_en)
                    lineCount_en += 1

                    # Help labels
                    if(row[4] == 1):
                        print("Adding help label...")
                        line_en_help = f"{row[0].replace(' ', '')}_Help={row[1]}\r"
                        lblEnFile.write(line_en_help)
                        lblEnFile.write(comm_en)
                        helpCount_en += 1

                    print(f'[{_enFile}] - OK')
                except:
                    print(f'[{_enFile}] - FAILED')
                
                # Italian
                line_it = f"{row[0].replace(' ', '')}={row[2]}\r"
                comm_it = f' ;{row[3]}\n'
                print(f"[{_itFile}] <- {line_it}")
                try:
                    lblItFile.write(line_it)
                    lblItFile.write(comm_it)
                    lineCount_it += 1

                    # Help labels
                    if(row[4] == 1):
                        print("Adding help label...")
                        line_it_help = f"{row[0].replace(' ', '')}_Help={row[2]}\r"
                        lblItFile.write(line_it_help)
                        lblItFile.write(comm_it)
                        helpCount_it += 1

                    print(f'[{_itFile}] - OK')
                except:
                    print(f'[{_itFile}] - FAILED')
        print(f"Operation Complete!\n{(len('Summary:')+1) * '-'}\n Summary:\n{(len('Summary:')+1) * '-'}")
        print(f"[{_enFile}] - {lineCount_en} label(s) written. {helpCount_en} help labels written.")
        print(f"[{_itFile}] - {lineCount_it} label(s) written. {helpCount_it} help labels written.")
        print("Thank you, come again!")
    except:
        print("An unknown error occured.")
    finally:
        input("Press any key to exit...")
        sys.exit()

if __name__ == "__main__":
    data = None

    if(len(sys.argv) == 1):
        data = readFromFile()
    elif(len(sys.argv) > 1):   # Passing read file (.xlsx)
        if(os.path.exists(sys.argv[1])):
            data = readFromFile(sys.argv[1].strip())
        else:
            print(f"[ERROR] File {sys.argv[1]} doesn't exist or lies in another place.")
            input("Press any key to exit...")
            sys.exit()

    if(len(sys.argv) == 3):   # Passing english output file (en-US.txt)
        writeLabels(data, sys.argv[2].strip())
    elif(len(sys.argv) == 4):   # Passing english output file (en-US.txt), italian output file (it-IT.txt)
        writeLabels(data, sys.argv[2].strip(), sys.argv[3].strip())
    else:
        writeLabels(data)