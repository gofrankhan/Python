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
	parser.add_argument('-s', action='store',
                    dest='simple_value',
                    help='Store a simple value')

	parser.add_argument('-c', action='store_const',
		            dest='constant_value',
		            const='value-to-store',
		            help='Store a constant value')

	parser.add_argument('-t', action='store_true',
		            default=False,
		            dest='boolean_t',
		            help='Set a switch to true')
	parser.add_argument('-f', action='store_false',
		            default=True,
		            dest='boolean_f',
		            help='Set a switch to false')

	parser.add_argument('-a', action='append',
		            dest='collection',
		            default=[],
		            help='Add repeated values to a list')

	parser.add_argument('-A', action='append_const',
		            dest='const_collection',
		            const='value-1-to-append',
		            default=[],
		            help='Add different values to list')
	parser.add_argument('-B', action='append_const',
		            dest='const_collection',
		            const='value-2-to-append',
		            help='Add different values to list')

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

	results = parser.parse_args()
	print('House Rent          = {!r}'.format(results.rent))
	print('Bua Bill            = {!r}'.format(results.bua))
	print('Net Bill            = {!r}'.format(results.net))
	print('Pani Bill           = {!r}'.format(results.pani))
	print('Total Meal          = {!r}'.format(total_meal))
	print('Bazar               = {!r}'.format(total_bazar))
	print('simple_value        = {!r}'.format(results.simple_value))
	print('constant_value      = {!r}'.format(results.constant_value))
	print('boolean_t           = {!r}'.format(results.boolean_t))
	print('boolean_f           = {!r}'.format(results.boolean_f))
	print('collection          = {!r}'.format(results.collection))
	print('const_collection    = {!r}'.format(results.const_collection))

if __name__=="__main__":
	sys.exit(main())
