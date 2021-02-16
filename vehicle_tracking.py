# -*- coding: utf-8 -*-
"""
Created on Sun Feb 14 23:54:08 2021

@author: KANWAR KELIDE
"""

import pandas as pd
from math import radians, cos, sin, asin, sqrt
import xlsxwriter as xl
import re
from datetime import datetime as dtt
import time
from os import listdir
from os.path import isfile, join
start_time = time.time()

onlyfiles = [f for f in listdir('F:\\Olectra\\Nagpur Data\\') if isfile (join('F:\\Olectra\\Nagpur Data\\',f))]
#print(onlyfiles)
count = len(onlyfiles)
text = [0]
#print (count)
for i in range(0, count):
    print(onlyfiles[i])
    text = onlyfiles[i].split('_Power')
    print(text)
   
unwanted_char = ['_']

final = pd.read_excel('F:\\Olectra\\Nagpur Data\\'+ onlyfiles[0], sheet_name = "Power Analysis")
text = onlyfiles[0].split('_Power')
#print(text)
#text = [sub.replace('_','') for sub in text]

final['Vehicle Number'] = text[0]
for i in range(1, count):
    field = pd.read_excel('F:\\Olectra\\Nagpur Data\\'+ onlyfiles[i], sheet_name = "Power Analysis")
    text = onlyfiles[i].split('_Power')
    #print(text)
    #text = [sub.replace('_','') for sub in text]
    field['Vehicle Number'] = text[0]
    final = final.append(field)
print (final)

busId = final['Vehicle Number']
date = final['Time']
soc = final['SoC']
soc = pd.to_numeric(soc, errors = 'coerce')
power = final['Power(kW)']
odometer = final['Odometer Reading']
speed = final['Speed']
current = final['Total Current']
voltage = final['Total HVoltage']
address = final['Location']

chargingSpan = 0
chargingMinutes = 0
minsPerCharge = 0
added = 0
preRadius = 0
radius = 0
radiusRGI = 0
preRadiusRGI = 0
chargeOnSoC = 0
chargeOffSoC = 0
k = 1
jrny = 1
energy = 0
odoIn = 0
row_count = 0
global j
rc = 1

dc = 0
ac = 0
added = 0
 
def chargingTime(arrival, departure, startIndex, endIndex): # receives arrival/ departure time, index of arrival and departure time;
    #calculates and returns soc added, row index where charging started and stopped
    added = 0
    chargeOff = 0
    chargeOn = 0
    energy = 0
    hourGap = 0
    diff = 0 # stores difference between consecutive rows
    kWh = 0
    for i in range (startIndex, endIndex):
                if (soc.iloc[i+1]>soc.iloc[i] and abs(soc.iloc[i+1] - soc.iloc[i]) == 1): # true if soc increments by 1
                    added =added+1
                    chargeOff = i # index when charging stopped
                    diff = date.iloc[i+1] - date.iloc[i]
                    #print(diff.seconds)
                    hourGap = (1/3600)*diff.seconds # conversion of consecutive row difference from seconds to hours
                    #print(hourGap)
                    energy = hourGap * power.iloc[i] # energy = time difference * power
                    kWh = kWh + energy
                    if added == 1:
                        chargeOn = i # index when charging strated
                    #print (date.iloc[i], soc.iloc[i], soc.iloc[i+1], added)
    if chargeOn != chargeOff:
        added = (soc.iloc[chargeOff] - soc.iloc[chargeOn]) + 1
        #print('Charge on:', soc[chargeOn], 'Charge off', soc[chargeOff])
    return added, chargeOn, chargeOff

def chargingCycle(startIndex, endIndex): # parameters and indices of charge start and stop, return energy added to battery
    diff = 0
    kWh = 0
    energy = 0
    hourGap = 0
    for i in range (startIndex, endIndex):
        #if chargeOn >=  date.iloc[i] and date.iloc[i] <= chargeOff:
            diff = date.iloc[i+1] - date.iloc[i]
            #print(diff.seconds)
            hourGap = (1/3600)*diff.seconds
            #print(hourGap)
            if current.iloc[i] < 0 and current.iloc[i] > -300:
                energy = hourGap * power.iloc[i]
                kWh = kWh + energy
            #print(kWh)
    kWh = -1 * kWh
    #print(kWh)
    return round(kWh,2)

r=1
def printcharge(startIndex, endIndex, r):
    diff = 0
    energy = 0
    hourGap = 0
    j = r
    print (startIndex, endIndex, r)
    for i in range (startIndex, endIndex):
        if current.iloc[i] < 0 and current.iloc[i] > -300:
            diff = date.iloc[i+1] - date.iloc[i]
            hourGap = (1/3600)*diff.seconds
            #print(hourGap)
            energy = hourGap * power.iloc[i]
            #print(energy)
            charge.write(j,0,date.iloc[i])
            charge.write(j,1,busId.iloc[i])
            charge.write(j,2,odometer.iloc[i])
            charge.write(j,3,speed.iloc[i])
            charge.write(j,4,soc.iloc[i])
            charge.write(j,5,current.iloc[i])
            charge.write(j,6,voltage.iloc[i])
            charge.write(j,7,power.iloc[i])
            charge.write(j,8,diff.seconds)
            charge.write(j,9,hourGap)
            charge.write(j,10,energy)
            j = j + 1
    return j  

workbook = xl.Workbook('F:\\Olectra\\Nagpur Data\\NGP In Out and Charging last 3 days.xlsx')
worksheet = workbook.add_worksheet(" In Out")
charge = workbook.add_worksheet("Charging")
rows = workbook.add_worksheet("Row Count")
worksheet.write(0,0,"Depot")
worksheet.write(0,1,"Vehicle Number")
worksheet.write(0,2,"Depot In")
worksheet.write(0,3,"KMs In")
worksheet.write(0,4,"In SoC")
worksheet.write(0,5,"Depot Out")
worksheet.write(0,6,"KMs Out")
worksheet.write(0,7,"Out Soc")
worksheet.write(0,10,"SoC Added")
worksheet.write(0,8,"SoC Charge Start")
worksheet.write(0,9,"SoC Charge Stop")
worksheet.write(0,11,"Charging On")
worksheet.write(0,12,"Charging Off")
worksheet.write(0,13,"Charging Duration(mins)")
worksheet.write(0,14,"Mins. taken per SoC ")
worksheet.write(0,15,"kWh Added")

charge.write(0,0,"Date")
charge.write(0,1,"Vehicle")
charge.write(0,2,"Odometer")
charge.write(0,3,"Speed")
charge.write(0,4,"Charge")
charge.write(0,5,"Current")
charge.write(0,6,"Voltage")
charge.write(0,7,"Power")
charge.write(0,8,"Time diff (secs)")
charge.write(0,9,"Time diff (hrs)")
charge.write(0,10,"Energy (kWh)")

rows.write(0,0,"Vehicle No")
rows.write(0,1,"Row Count")

count = final.shape[0]

for i in range (0, count-1):
    date.iloc[i]= dtt.strptime(date.iloc[i],'%Y-%m-%d %H:%M:%S') # date converted from string to datetime type

for i in range (0, count-1,5):
    if busId.iloc[i] == busId.iloc[i+1]:
        row_count = row_count + 1
       
    else:
        print (busId.iloc[i], row_count + 1)
        rows.write(rc,0,busId.iloc[i])
        rows.write(rc,1,row_count + 1)
        rc = rc + 1
        row_count = 0
        
       
    if str(address.iloc[i]) == "Nagpur Depot" and dc == 0:
        worksheet.write(k,0,'NGP')
        worksheet.write(k,1,busId.iloc[i])
        worksheet.write(k,2,date.iloc[i])
        worksheet.write(k,3,odometer.iloc[i])
        odoIn = odometer.iloc[i]
        worksheet.write(k,4,soc.iloc[i])
        arrival = date.iloc[i]
        startIndex = i
        print('Arrival : ', arrival, soc.iloc[i])
        dc = 1

    if ((address.iloc[i]) != "Nagpur Depot" and dc == 1 and float(odometer.iloc[i]) >= odoIn) or (busId.iloc[i] != busId.iloc[i+1] and dc == 1): #checks whether bus left depot or still at depot and end of data for subjective bus
        worksheet.write(k,5,date.iloc[i])
        worksheet.write(k,6,odometer.iloc[i])
        worksheet.write(k,7,soc.iloc[i])
        departure = date.iloc[i]
        endIndex = i
        print('Departure : ', departure, soc.iloc[i])
        socAdded, chargeOn, chargeOff = chargingTime(arrival, departure, startIndex, endIndex)
        if socAdded > 0: # checks if bus was charged, if yes, then calls energy calculator function
            energy = chargingCycle(chargeOn, chargeOff) # parameters are indices of charge start and charge stop time
            chargingSpan = date.iloc[chargeOff] - date.iloc[chargeOn] # duration of charging in seconds
            chargingMinutes = round(chargingSpan.seconds/60,1) # charging duration in minutes
            minsPerCharge = round(chargingMinutes/socAdded,1)
            chargeOnSoC = soc.iloc[chargeOn-1]  # soc at start of charge
            chargeOffSoC = soc.iloc[chargeOff+1] # soc at end of charge
            #print (timeByCharging)
            r = printcharge(chargeOn, chargeOff, r)
        worksheet.write(k,10,socAdded)
        worksheet.write(k,8,chargeOnSoC)
        worksheet.write(k,9,chargeOffSoC)
        worksheet.write(k,11,date.iloc[chargeOn])
        worksheet.write(k,12,date.iloc[chargeOff])
        worksheet.write(k,13,chargingMinutes)
        worksheet.write(k,14,minsPerCharge)
        worksheet.write(k,15,energy)
        chargingMinutes = 0
        minsPerCharge = 0
        energy = 0
        chargeOff = 0
        chargeOn = 0
        chargeOnSoC = 0
        chargeOffSoC = 0
        k = k + 1
        dc = 0
        #print(type(address.iloc[i]))
    if i == count-2:
        print ("Last Row")
        if (address.iloc[i]) == "Nagpur Depot":
            # checks if bus is at depot and end of rows
            worksheet.write(k,5,date.iloc[i])
            worksheet.write(k,6,odometer.iloc[i])
            worksheet.write(k,7,soc.iloc[i])
            departure = date.iloc[i]
            endIndex = i
            print('Departure : ', departure, soc.iloc[i])
            socAdded, chargeOn, chargeOff = chargingTime(arrival, departure, startIndex, endIndex)
            if socAdded > 0: # checks if bus was charged, if yes, then calls energy calculator function
                energy = chargingCycle(chargeOn, chargeOff) # parameters are indices of charge start and charge stop time
                chargingSpan = date.iloc[chargeOff] - date.iloc[chargeOn] # duration of charging in seconds
                chargingMinutes = round(chargingSpan.seconds/60,1) # charging duration in minutes
                minsPerCharge = round(chargingMinutes/socAdded,1)
                chargeOnSoC = soc.iloc[chargeOn-1] # soc at start of charge
                chargeOffSoC = soc.iloc[chargeOff+1] # soc at end of charge
                #print (timeByCharging)
                r = printcharge(chargeOn, chargeOff, r)
            worksheet.write(k,10,socAdded)
            worksheet.write(k,8,chargeOnSoC)
            worksheet.write(k,9,chargeOffSoC)
            worksheet.write(k,11,date.iloc[chargeOn])
            worksheet.write(k,12,date.iloc[chargeOff])
            worksheet.write(k,13,chargingMinutes)
            worksheet.write(k,14,minsPerCharge)
            worksheet.write(k,15,energy)
            chargingMinutes = 0
            minsPerCharge = 0
            energy = 0
            chargeOff = 0
            chargeOn = 0
            chargeOnSoC = 0
            chargeOffSoC = 0
            k = k + 1
            dc = 0
        print (busId.iloc[i], row_count + 1)
        rows.write(rc,0,busId.iloc[i])
        rows.write(rc,1,row_count + 1)
        rc = rc + 1
   
print("--- %s seconds ---" % (time.time() - start_time))
workbook.close()