#!/usr/bin/env python
from common_util import *
from create_report import *
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
	parser.add_argument('--electricity', '-e', action='store', dest='elect',
				default=0,
				type=int,
				help='Store pani bill which distributed amoung 5.')

	parser.add_argument('--version', action='version',
		            version='%(prog)s 1.0')

	results = parser.parse_args()

	set_arg_variables(results.rent, results.bua, results.net, results.pani, results.elect)
	take_meal_count_input()
	take_bazar_input()
	calculation_all()

	print_summary()
	print_result()

	create_xml_folder()
	create_summary_xml()


if __name__=="__main__":
	sys.exit(main())
  