#!/usr/bin/env python

import sys
import os
import argparse

def main():
	parser = argparse.ArgumentParser(description="Meal Calculation System")
	parser.add_argument('--rent', '-r', action='store', dest='rent',
				default=22500,
				type=int,
				help='Store house rent amount which distributed amoung 5.')
	parser.add_argument('--net', '-n', action='store', dest='net',
				default=1200,
				type=int,
				help='Store net bill which distributed amoung 5.')
	parser.add_argument('--bua', '-b', action='store', dest='bua',
				default=3200,
				type=int,
				help='Store bua bill which distributed amoung 5.')

	parser.add_argument('--pani', '-p', action='store', dest='pani',
				default=0,
				type=int,
				help='Store pani bill which distributed amoung 5.')

	parser.add_argument('--version', action='version',
		            version='%(prog)s 1.0')

	names = ['Avo','Kamal','Gofran','Habibullah','Abdullah']
	mealCount = []
	Bazar = []
	print("Enter Meal Count:")
	for name in names:
		mealCount.append(int(input("%s's Meal %s: " %(name, " "*(11-len(name))))))
	print("Enter Bazar Amount:")
	for name in names:
		Bazar.append(int(input("%s's Bazar %s: " %(name, " "*(10-len(name))))))
	total_meal = 0
	for x in mealCount:
		total_meal += x
	total_bazar = 0
	for x in Bazar:
		total_bazar += x

	meal_rate =  total_bazar/total_meal
	mealCost = []
	for x in range(0, len(names)):
		mealCost.append(int(mealCount[x]*meal_rate))

	results = parser.parse_args()
	print('House Rent          = {!r}'.format(results.rent)) 
	print('Bua Bill            = {!r}'.format(results.bua))
	print('Net Bill            = {!r}'.format(results.net))
	print('Pani Bill           = {!r}'.format(results.pani))
	print('Total Meal          = {!r}'.format(total_meal))
	print('Bazar               = {!r}'.format(total_bazar))
	print('Meal Rate           = {!r}'.format(meal_rate))

	payable = []
	i = 0
	for name in names:
		payable.append(int (results.rent/5 + results.net/5 + results.bua/5 + results.pani/5 + mealCost[i]))
		i = i + 1

	print('-'* 140)
	print("| Name          |  House Rent |   Net Bill |   Bua Bill   |   Pani Bill |   Meal Cost   |   Total Payable |  Total Bazar |  Settlement     |")
	index = 0
	for name in names:
		print('-'* 140)
		print('| %s    |     %d    |     %d    |     %d      |      %d      |     %d       |      %d       |     %d      |        %d     |' 
		%(names[index] + " "*(10-len(names[index])), results.rent/5, results.net/5, results.bua/5, results.pani/5, mealCost[index], 
		payable[index], Bazar[index], payable[index]-Bazar[index]))
		index = index + 1
	print('-'* 140)


if __name__=="__main__":
	sys.exit(main())
