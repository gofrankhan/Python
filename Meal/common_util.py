#!/usr/bin/env python

import os
import sys

HOUSE_RENT = 22500
TOTAL_BAZAR = 0
TOTAL_MEAL = 0
TOTAL_PANI_BILL = 0
TOTAL_BUA_BILL_5 = 3200
TOTAL_NET_BILL = 1200
TOTAL_ELECTRICITY_BILL = 0
MEAL_RATE = 0
AYWAN_BUA_BILL = 300

NAMES = ['Avo','Kamal','Gofran','Habibullah','Abdullah']
MEAL_COUNT = []
BAZAR = []
MEAL_COST = []
PAYABLE = []


XML_FOLDER_NAME = 'xml_folder'
XML_FOLDER_PATH = os.getcwd() + os.path.sep + XML_FOLDER_NAME

def take_meal_count_input():
    print("Enter Meal Count:")
    for name in NAMES:
        MEAL_COUNT.append(int(input("%s's Meal %s: " %(name, " "*(11-len(name))))))

def take_bazar_input():
    print("Enter Bazar Amount:")
    for name in NAMES:
        BAZAR.append(int(input("%s's Bazar %s: " %(name, " "*(10-len(name))))))

def calculation_all():
    global TOTAL_MEAL
    for x in MEAL_COUNT:
        TOTAL_MEAL += x

    global TOTAL_BAZAR
    for x in BAZAR:
        TOTAL_BAZAR += x
    
    global MEAL_RATE
    MEAL_RATE =  TOTAL_BAZAR / TOTAL_MEAL
    
    for x in range(0, len(NAMES)):
        MEAL_COST.append(int(MEAL_COUNT[x] * MEAL_RATE))
    
    i = 0
    for name in NAMES:
        PAYABLE.append(int (HOUSE_RENT/5 + TOTAL_NET_BILL/5 + TOTAL_BUA_BILL_5/5 + TOTAL_PANI_BILL/5 + MEAL_COST[i]))
        i = i + 1

def print_summary():
	print('House Rent          = {!r}'.format(HOUSE_RENT)) 
	print('Bua Bill            = {!r}'.format(TOTAL_BUA_BILL_5 + AYWAN_BUA_BILL))
	print('Net Bill            = {!r}'.format(TOTAL_NET_BILL))
	print('Pani Bill           = {!r}'.format(TOTAL_PANI_BILL))
	print('Total Meal          = {!r}'.format(TOTAL_MEAL))
	print('Bazar               = {!r}'.format(TOTAL_BAZAR))
	print('Meal Rate           = {!r}'.format(MEAL_RATE))

def print_result():
    print('-'* 140)
    print("| Name          |  House Rent |   Net Bill |   Bua Bill   |   Pani Bill |   Meal Cost    |   Total Payable |  Total Bazar  |  Settlement     |")
    index = 0
    for name in NAMES:
        print('-'* 140)
        print('| %s    |     %d    |     %d    |     %d      |      %d      |     %d       |      %d       |     %d      |        %d     |' 
        %(NAMES[index] + " "*(10-len(NAMES[index])), HOUSE_RENT/5, TOTAL_NET_BILL/5, TOTAL_BUA_BILL_5/5, TOTAL_PANI_BILL/5, MEAL_COST[index], 
        PAYABLE[index], BAZAR[index], PAYABLE[index]-BAZAR[index]))
        index = index + 1
    print('-'* 140)

def set_arg_variables(rent, bua, net, pani, elect):
    global HOUSE_RENT, TOTAL_BUA_BILL_5, TOTAL_NET_BILL, TOTAL_PANI_BILL
    HOUSE_RENT = rent
    TOTAL_BUA_BILL_5 = bua
    TOTAL_NET_BILL = net
    TOTAL_PANI_BILL = pani
    TOTAL_ELECTRICITY_BILL = elect

