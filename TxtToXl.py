'''
@author     Justyn Crook
@version    1.1

TTE takes a text file as input. It then goes through each line and searches the WhoIs information 
and exports the data into an excel spreadsheet..
'''

import json
import pandas as pd
import time
from ipwhois import IPWhois
import ipwhois

def main():
    df = pd.DataFrame()

    filename = input('Enter the text file location and name: ')
    sent_loc = '../Desktop/TTE.xlsx'

    ip_addr = []
    ip_descrip = []
    ip_country = []

    # Open the text file
    with open(filename) as txt_file:
        for line in txt_file:
            # Remove the \n character from the end of the line
            line = line[:-1]

            print("Looking up IP: " + line)
            
            try:
                ip_who_is = IPWhois(line)
                ip_data = ip_who_is.lookup_rdap(depth = 1)
                asn_descrip = ip_data['asn_description']
                asn_country = ip_data['asn_country_code']

                ip_addr.append(line)
                ip_descrip.append(asn_descrip)
                ip_country.append(asn_country)
            except ValueError:
                ip_addr.append(line)
                ip_descrip.append("IP is not IPv4or IPv6")
                ip_country.append("N/A")
            except ipwhois.exceptions.IPDefinedError:
                ip_addr.append(line)
                ip_descrip.append("IP is ill-defined")
                ip_country.append("N/A")
            except ipwhois.exceptions.WhoisLookupError:
                ip_addr.append(line)
                ip_descrip.append("IP is not IPv4 or IPv6")
                ip_country.append("N/A")
            except ipwhois.exceptions.ASNRegistryError:
                ip_addr.append(line)
                ip_descrip.append("ASN registry lookup failed. Permutations not allowed.")
                ip_country.append("N/A")

        #Add lists to dataframe
        df['IP'] = ip_addr
        df['Description'] = ip_descrip
        df['Country'] = ip_country
        #export dataframe to excel format
        try:
            df.to_excel(sent_loc, index = False)
            print("Sheet added...")
        except e:
            print("Error...")

    print('Complete!')

if __name__ == '__main__':
    main()

