'''
@author     Justyn Crook
@version    1.1

JTE takes a text file as input that includes a list of JSON files. It then goes through each
JSON file and exports wanted information into an excel file that includes one or more spreadsheets.
'''

import json
import pandas as pd
import time
from ipwhois import IPWhois

df_cols = []
excel_cols = []

def main():
    global df_cols
    global excel_cols

    filename = input('Enter the text file that contains the location & name of the JSON files: ')
    send_loc = input('Enter the location you would like to save the excel sheet: ')
    xlsx_file = input('Enter the name of the excel sheet you are exporting to: ')
    sheet_num = 1
    sheet_name = 'Sheet'

    sent_loc = send_loc + xlsx_file

    # Open the text file
    with open(filename) as file_locs:
        # Read the .json file names from each line
        for line in file_locs:
            # Remove the \n character from the end of the line
            line = line[:-1]
            # Read data from json file
            with open(line) as json_file:
                data = json.load(json_file)
                
                # Create pandas dataframe & normalize the json data if it is nested
                norm = pd.json_normalize(data["events"], max_level = 1) 
                df = pd.DataFrame(norm)

                # If this is the first file, let user choose which columns are exported
                if sheet_num == 1:
                    print("Please enter each column from the list you would like added to the excel sheet. When finished, type done.")
                    for col in df:
                        df_cols.append(col)
                        print(col)
                
                    # Add an extra line
                    print()
                    
                    # Checks whether the column entered is in the excel columns listed above
                    in_col = ""
                    col_loop = True
                    while col_loop:
                        in_col = input()
                        
                        if "done" in in_col:
                            col_loop = False
                            break
                        elif in_col not in df_cols:
                            print("You entered an invalid column, please try again.")
                            continue
                        else:
                            excel_cols.append(in_col)
                            print("%s added...\n" % in_col)
                    
                    # Unix Time conversion to local date & time
                    print("Is there a UNIX time column you are including? (Y/N)")
                    yes_no = input()
                    unix_time = epoch_conversion(yes_no)

                    #IP Information from WhoIs Lookup which uses ARIN
                    print("Would you like IP information? (Y/N)")
                    ip_ans = input()
                    ip_col = ip_whois(ip_ans)

                print('Reading file: %s\n' % json_file)

                # If the user said they want to look up IP information
                ip_info = []
                if ip_col != None and ip_col in excel_cols:
                    print("Looking up IP information...")
                    last_ip = None
                    last_asn = None
                    # Go through every IP in the JSON file
                    for i in df[ip_col]:
                        # If the last IP is the same as the current, use the previous info.
                        if last_ip == i:
                            ip_info.append(last_asn)
                        # Else look up the IP information
                        else:
                            print(i)
                            ip_who_is = IPWhois(i)
                            ip_data = ip_who_is.lookup_rdap(depth = 1)
                            asn_descrip = ip_data['asn_description']
                            ip_info.append(asn_descrip)
                            last_ip = i
                            last_asn = asn_descrip

                    df['ip_owner'] = ip_info
                    if "ip_owner" not in excel_cols:
                        excel_cols.append("ip_owner")

                # If the user said they want to convert the time
                date_time = []
                if unix_time != None and unix_time in excel_cols:
                    print("Adding date_time column...")
                    # Go through every unix timestamp in the JSON file
                    for epoch in df[unix_time]:
                        # Change real_time depending if you are given milliseconds or seconds
                        real_time = float(epoch) / 1000

                        #Converts the unix seconds into local time
                        local_time = time.ctime(real_time)
                        date_time.append(local_time)
        
                    df['date_time'] = date_time
                    if "date_time" not in excel_cols:
                        excel_cols.append("date_time")
    
                #export dataframe to excel format
                try:
                    if sheet_num == 1:
                        new_sheet = "%s%d" % (sheet_name, sheet_num)
                        df.to_excel(sent_loc, index = False, sheet_name = new_sheet, columns = excel_cols)
                        print("Sheet added...")
                        sheet_num += 1
                    else:
                        new_sheet = "%s%d" % (sheet_name, sheet_num)
                        with pd.ExcelWriter(sent_loc, mode="a") as writer:
                            df.to_excel(writer, index = False, sheet_name = new_sheet, columns = excel_cols)
                        print("Sheet added...")
                        sheet_num += 1
                except e:
                    print("Error...")

    print('Complete!')

def epoch_conversion(ans):
    if 'Y' in ans or 'y' in ans or 'yes' in ans or 'Yes' in ans:
        print("Please enter that column's name: ")
        unix_time = input()
    else:
        unix_time = None
    return unix_time

def ip_whois(ip):
    if 'Y' in ip or 'y' in ip or 'yes' in ip or 'Yes' in ip:
        print("Please enter that column's name: ")
        ip_col = input()
    else:
        ip_col = None
    return ip_col

if __name__ == '__main__':
    main()

