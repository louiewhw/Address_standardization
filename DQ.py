#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Jan 26 12:51:23 2017

This is a Data Quality Tool. 

@author: Louie
"""

import pandas as pd
import numpy as np
import pycountry
import plotly as py
import plotly.graph_objs as go
import re



#============================================================================================

def nFile(data, nFilename):
    '''
    This function is to export Pandas DataFrame into xlsx file
    '''
    writer = pd.ExcelWriter(nFilename + '.xlsx', engine = 'xlsxwriter')  
    data.to_excel(writer, sheet_name = 'sheet1')  
    writer.save()

    
def cleanID(filename):
    '''
    This function is to clean up the unnecessary rows on the excel file.
    Delete rows if DW_Id is not numeric. 
    '''
    oData = pd.read_excel(filename)
    nData = oData[oData.DW_Id.apply(lambda x: isinstance(x, (int)))].set_index('DW_Id') 
    nData = nData.fillna('')
    return nData

def cleanCountry(data, data_country):
    '''
    This function transfer all the Country data into 2 digit code data (example: America -> US)
    
    '''
    
    #Read from Pandas DataFrame into arrays.
    DW_Id = np.array(data.index) #keep track DW_Id
    oData = np.array(data_country,dtype = 'str')
    nData = np.array(oData,dtype = 'str')
    
    #analytical purpose
    caught = 0 
    right = 0
    uncatchable = 0
    uncatchlist = []
    
    
    for i in range(len(oData)):

        try:
            #Checking if the countries is already in 2 digit code
            pycountry.countries.get(alpha_2 = oData[i]).alpha_2
            right += 1
        except:
            try:
                #Transfer all the Country Data into 2 digit code
                nData[i] = pycountry.countries.get(name = oData[i]).alpha_2
                print( oData[i] + ' -> ' + pycountry.countries.get(name = oData[i]).alpha_2 + ' :::: DW_Id = ' + str(DW_Id[i]))
                caught += 1
            except:
                try:
                    #Transfer all the 3 digit code into 2 digit code
                    nData[i] = pycountry.countries.get(alpha_3 = oData[i]).alpha_2
                    print( oData[i] + ' -> ' + pycountry.countries.get(alpha_3 = oData[i]).alpha_2 + ' :::: DW_Id = ' + str(DW_Id[i]))
                    caught += 1
                except:
                    #Danmark exception. Pycountry module spelled 'Denmark'
                    if oData[i] == 'Danmark':
                        nData[i] = 'DK'
                        print(oData[i] + ' -> ' + 'DK' + ' :::: DW_Id = ' + str(DW_Id[i]))
                        caught += 1
                    else:
                        #Record all the uncatched countries
                        uncatchable += 1
                        uncatchlist.append(oData[i])
    
    #analytical purpose
    print('Caught: ' + str(caught))
    print('Right: ' + str(right))
    print('Uncatchable: ' + str(uncatchable))
    print(uncatchlist)
    
    #analytical purpose (Graph)(Optional)
    
    fig = {
    'data': [{'labels': ['Caught', 'Right', 'Uncatchable'],
              'values': [caught, right, uncatchable],
              'type': 'pie'}],
    'layout': {'title': 'Country Data Quality'}
     }
    
    py.offline.plot(fig)
    
    
    return nData

def cleanPhone (data, country, phone):
    '''
    This function is to standardize the phone number.
    US: +1(xxx)xxx-xxxx
    DK:
        
    '''
    #Read from Pandas DataFrame into arrays. *US Only
    DW_Id = np.array(data[data[country] == 'US'].index)
    USData = np.array(data[data[country] == 'US'][phone], dtype = 'str')
    nUSData = np.array(USData, dtype = 'str')
    
    #analytical purpose
    catched = 0
    uncatchable = 0
    uncatchablelist = []
    
    #Looping all the US Phone number
    for i in range(len(USData)):
        #Drop everything except numbers
        nUSData[i] = re.sub("[^0-9]", "", USData[i])
        
        #If phone number contains 11 numbers and the first number is equal to 1, format it.
        if len(nUSData[i]) == 11 and nUSData[i][0] == '1':
            nUSData[i] = ('+' + nUSData[i][0] +'(' + nUSData[i][1:4] + ')' + nUSData[i][4:7] + '-' + nUSData[i][7:11])
            print(USData[i] + ' -> ' + nUSData[i] + ' :::: DW_Id = ' + str(DW_Id[i]))
            catched += 1
        
        #If phone number contains 10 numbers, add +1, and format it
        elif len(nUSData[i]) == 10:
            
            nUSData[i] = ('+1' +'(' + nUSData[i][0:3] + ')' + nUSData[i][3:6] + '-' + nUSData[i][6:10])
            print(USData[i] + ' -> ' + nUSData[i] + ' :::: DW_Id = ' + str(DW_Id[i]))
            catched += 1
        
        ##analytical purpose. Uncatch list.
        else:
            uncatchablelist.append(USData[i])
            uncatchable += 1
    
    #analytical purpose
    print('Total US Phone Number = ' + str(len(nUSData)))
    print('Catched = ' + str(catched))
    print('Uncatchable = ' + str(uncatchable))
    print(uncatchablelist)
    
    return nUSData


def cleanAddress(data, address):
    '''
    This function is to standardize the address. 
    '''
    
    #Read from Pandas DataFrame into arrays.
    DW_Id = np.array(data.index) #keep track DW_Id
    address = np.array(data[address], dtype = 'str')
    
    #Empty arrays
    AddSplit = np.empty(len(address), dtype = list)
    nAddress = []
    
    #List of different case. First element is the desire data format

    street = ['Street','street','st', 'st.', 'str', 'str,']
    avenue = ['Ave', 'ave', 'avenue', 'ave.']
    suite = ['Suite', 'suite', 'suire', 'suite.', 'sutie', 'ste', '#']
    road = ['Road', 'road', 'rd', 'rd.']
    apt = ['APT', 'apt', 'apt.', 'apartment']
    blvd = ['Blvd', 'boulevard', 'blvd.']
    mylist = [street, avenue, suite, road, apt, blvd]
    
    
    for i in range(len(address)):
        #Drop everything except A-Z, a-z, '#', '.', and split it into different element is ther is a space (' ')
        AddSplit[i] = re.sub('[^A-Za-z0-9#. ]',"",address[i]).split(' ')
    
        for j in range(len(AddSplit[i])):
            for k in mylist:
                #if element is equal to desire format, pass.
                if AddSplit[i][j] == k[0]:
                    pass
                #Else if element contains value on list, replace it to the desire data format. 
                elif AddSplit[i][j].lower() in k[1:]:
                    print(AddSplit[i][j] + ' -> ' + k[0] + ' :::: DW_Id = ' + str(DW_Id[i]))
                    AddSplit[i][j] = k[0]

        #Join the splited data into a string
        nAddress.append( ' '.join(AddSplit[i]))
    
    return nAddress

    
    
import googlemaps


gmaps = googlemaps.Client(key = '')

def cleanAddressGoogle(data, address):
    data = np.array(data[address], dtype = 'str')
    ndata = np.array(data)
    
    for i in range(len(data)):
        try:
            ndata[i] = gmaps.geocode(data[i])[0]['formatted_address']
            print(ndata[i])
        except:
            print(ndata[i] + ':::::::::::::ERROR')
    return ndata
        
#========================================================================================

'''
#clean ID

HarCli = cleanID('HarvestClients.xlsx')
HubCus = cleanID('HubSpotCustomers.xlsx')
FinDebt = cleanID('Finance_SOAP_Debtor.xlsx')
FinInv = cleanID('Finance_SOAP_Invoice.xlsx')
Lincense = cleanID('Licenses.xlsx')
'''
'''
#clean country
FinInv['DebtorCountry'] = cleanCountry(FinInv, FinInv.DebtorCountry)
FinDebt['Country'] = cleanCountry(FinDebt, FinDebt.Country)
HubCus['Country'] = cleanCountry(HubCus, HubCus.Country)
Lincense['Country'] = cleanCountry(Lincense, Lincense.Country)
'''
'''
#export file
nFile(FinInv, 'rFinance_SOAP_Invoice')
nFile(FinDebt, 'rFinance_SOAP_Debtor')
nFile(HubCus, 'rHubSpotCustomers')
nFile(HarCli, 'rHarvestClients')
nFile(Lincense, 'rLincense')
'''
'''
#Drop ALL column duplicates (name, country, city, country, etc.)
FinInv = FinInv.drop_duplicates()
FinDebt = FinDebt.drop_duplicates()
HarCli = HarCli.drop_duplicates()
HubCus = HubCus.drop_duplicates()
Lincense = Lincense.drop_duplicates()
'''
'''
#Drop ALL transectional AND unnecessary data
FinDebt = FinDebt[['Name','Address','City','Country','Email','TelephoneAndFaxNumber']]
FinInv = FinInv[['DebtorName', 'DebtorAddress', 'DebtorCity', 'DebtorCountry']]
HarCli = HarCli[['name','details','currency']]
HubCus = HubCus[['Name', 'Street Address', 'Street Address 2', 'City','State/Region','Country','Phone Number','Website URL']]
'''
'''
#Combine to full address
FinDebt['FullAddress'] = FinDebt.Address.astype(str) + ' ' +  FinDebt.City.astype(str) + ' ' + FinDebt.Country.astype(str)
FinInv['FullAddress'] = FinInv.DebtorAddress.astype(str) + ' ' + FinInv.DebtorCity.astype(str) + ' ' + FinInv.DebtorCountry.astype(str)
HubCus['FullAddress'] = HubCus['Street Address'].astype(str) + ' ' + HubCus['Street Address 2'].astype(str) + ' ' + HubCus['City'].astype(str) + ' ' + HubCus['State/Region'].astype(str) + ' ' + HubCus['Country'].astype(str)
'''

'''
#Filter US
FinDebt = FinDebt[FinDebt.Country == 'US']
FinInv = FinInv[FinInv.DebtorCountry == 'US']
HubCus = HubCus[HubCus.Country == 'US']
'''

'''
#Clean Phone
HubCus['Phone Number'][HubCus.Country == 'US'] = cleanPhone (HubCus, 'Country', 'Phone Number')
'''

'''
#Clean Address
HubCus['Street Address'] = cleanAddress(HubCus, 'Street Address')
FinDebt['Address'] = cleanAddress(FinDebt, 'Address')
FinInv['DebtorAddress'] = cleanAddress(FinInv, 'DebtorAddress')
Lincense['Address'] = cleanAddress(Lincense, 'Address')
'''
'''
#Clean Address (google API)
HubCus['FullAddress'] = cleanAddressGoogle(HubCus, 'FullAddress')
FinDebt['FullAddress'] = cleanAddressGoogle(FinDebt, 'FullAddress')
FinInv['FullAddress'] = cleanAddressGoogle(FinInv, 'FullAddress')
'''

'''
#testing
HubCus = cleanID('HubSpotCustomers.xlsx')
HubCus = HubCus[HubCus['Lifecycle Stage'] == 'customer']
HubCus['Country'] = cleanCountry(HubCus, HubCus.Country)
HubCus['Street Address'] = cleanAddress(HubCus, 'Street Address')
HubCus['FullAddress'] = HubCus['Street Address'].astype(str) + ' ' + HubCus['Street Address 2'].astype(str) + ' ' + HubCus['City'].astype(str) + ' ' + HubCus['State/Region'].astype(str) + ' ' + HubCus['Country'].astype(str)
HubCus['FullAddress'] = cleanAddressGoogle(HubCus, 'FullAddress')
nFile(HubCus, 'rHubSpotCustomers')

Lincense = cleanID('Licenses.xlsx')
Lincense['Country'] = cleanCountry(Lincense, Lincense.Country)
Lincense['Address'] = cleanAddress(Lincense, 'Address')
Lincense['FullAddress'] = Lincense['Address'].astype(str) + ' ' + Lincense['City'].astype(str) + ' ' + Lincense['Country'].astype(str)
Lincense['FullAddress'] = cleanAddressGoogle(Lincense, 'FullAddress')
nFile(Lincense, 'rLincense')

FinDebt = cleanID('Finance_SOAP_Debtor.xlsx')
FinDebt['Country'] = cleanCountry(FinDebt, FinDebt.Country)
FinDebt['Address'] = cleanAddress(FinDebt, 'Address')
nFile(FinDebt, 'rFinance_SOAP_Debtor')
FinDebt['FullAddress'] = FinDebt.Address.astype(str) + ' ' +  FinDebt.City.astype(str) + ' ' + FinDebt.Country.astype(str)
FinDebt['FullAddress'] = cleanAddressGoogle(FinDebt, 'FullAddress')
nFile(FinDebt, 'rFinance_SOAP_Debtor')


'''

