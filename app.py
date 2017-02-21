#!/usr/bin/env python

import os
import json
import datetime
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import re

from flask import (
	Flask,
	request,
	make_response,
	jsonify
)

app = Flask(__name__)
log = app.logger

def parameters_extractor(params):
	dicts = [params]
	values = []

	while len(dicts):
		d = dicts.pop()

		for value in d.values():
			if isinstance(value, dict):
				dicts.append(value)
			elif isinstance(value, basestring) and len(value) > 0:
				values.append(unicode(value))

	return values

def gsheets_auth():
	print 'auth in progress'
	with open('account.json', 'r') as data_file:
		json_key = json.loads(data_file.read())
	scope = ['https://spreadsheets.google.com/feeds']
	credentials = ServiceAccountCredentials.from_json_keyfile_dict(json_key, scope)
	gc = gspread.authorize(credentials)
	sh = gc.open_by_key('1SKvVzU5CJrlIANTfqfFsrQH34fRnzkHPzFCsYKPIIzw')
	return sh

def sheets_get(spradsheet):
	sales = spradsheet.worksheet("Sales")
	response = spradsheet.worksheet("Response List")
	response_all = response.get_all_values()
	# response_dict = dict(zip(response.col_values(1), [i.split('/') if i.find('/') > -1 else i for i in response.col_values(2)]))
	response_dict = dict(
		zip([i[0] for i in response_all], [i[1].split('/') if i[1].find('/') > -1 else i[1] for i in response_all]))
	return sales, response_dict

sh = gsheets_auth()

@app.route('/webhook', methods=['POST'])
def webhook():
	req = request.get_json(silent=True, force=True)
	date_period_parameter = False
	try:
		action = req.get("result").get('action')
	except AttributeError:
		return "No action, sorry."

	try:
		print req['result']['resolvedQuery']
	except UnicodeEncodeError:
		print req['result']['resolvedQuery'].encode('utf-8')

	if req['result']['parameters'].has_key('date') and len(req['result']['parameters']['date']) > 0:
		try:
			if req['result']['parameters']['date'].has_key('date'):
				req['result']['parameters']['date']['date'] = datetime.datetime.strptime(req['result']['parameters']['date']['date'], '%Y-%m-%d').strftime('%-m/%-d/%Y')
				date_period_parameter = False
			elif req['result']['parameters']['date'].has_key('date-period'):
				req['result']['parameters']['date']['date-period'] = [datetime.datetime.strptime(x, '%Y-%m-%d').strftime('%-m/%-d/%Y') for x in req['result']['parameters']['date']['date-period'].split('/')]
				date_period_parameter = True
		except AttributeError:
			res = {
				"speech": 'wrong parameters',
				"displayText": 'wrong parameters',
			}
			return make_response(jsonify(res))

	if action == 'product.price':
		res = productPrice(req)
	elif action == 'sales.status':
		res = salesStatus(req, False, False, False, date_period_parameter)
	elif action == 'sales.status.paid':
		res = salesStatus(req, True, False, False, date_period_parameter)
	elif action == 'sales.status.booked':
		res = salesStatus(req, False, True, False, date_period_parameter)
	elif action == 'sales.status.returned':
		res = salesStatus(req, False, False, True, date_period_parameter)
	elif action == 'sales.revenue':
		res = salesRevenue(req, date_period_parameter)
	elif action == 'sales.quantity':
		res = salesQuantity(req, date_period_parameter)
	elif action == 'sales.product.quantity':
		res = salesProductQuantity(req, False, False, False, date_period_parameter)
	elif action == 'sales.product.paid.quantity':
		res = salesProductQuantity(req, True, False, False, date_period_parameter)
	elif action == 'sales.product.booked.quantity':
		res = salesProductQuantity(req, False, True, False, date_period_parameter)
	elif action == 'sales.product.returned.quantity':
		res = salesProductQuantity(req, False, False, True, date_period_parameter)
	elif action == 'sales.person':
		res = salesPerson(req, date_period_parameter)
	elif action == 'sales.person.most_money' or action == 'sales.person.least_money':
		res = salesPersonMostLeastMoney(req)
	elif action == 'sales.date':
		res = salesDate(req, date_period_parameter)
	elif action == 'sales.date.most_money' or action == 'sales.date.least_money':
		res = salesDateMostLeastMoney(req)
	elif action == 'sales.product.most_money' or action == 'sales.product.least_money':
		res = salesProductMostLeastMoney(req)
	elif action == 'sales.product.best_selling' or action == 'sales.product.least_selling':
		res = salesProductBestLeastSelling(req)
	elif action == 'sales.status.paid.quantity':
		res = salesStatusQuantity(req, True, False, False, date_period_parameter)
	elif action == 'sales.status.booked.quantity':
		res = salesStatusQuantity(req, False, True, False, date_period_parameter)
	elif action == 'sales.status.returned.quantity':
		res = salesStatusQuantity(req, False, False, True, date_period_parameter)
	else:
		log.error("Unexpeted action.")

	return make_response(jsonify(res))

def productPrice(req):

	sales, response_dict = sheets_get(sh)
	parameters = req['result']['parameters']
	action = req['result']['action']

	product = [i for i in sales.col_values(sales.find('Product').col)[1:] if len(i) > 0]
	ppu_list = [int(float(i.replace('$', '').replace(',', ''))) for i in sales.col_values(sales.find('Price per unit').col)[1:] if len(i) > 0]
	status_dict = dict(zip(product, ppu_list))
	#response = parameters['product'] + ' is $' + str(status_dict[parameters['product']]) + '.00 per unit.'
	#response = re.sub('@\w+', '%s', response_dict[action]) %(parameters['product'], '$' + str(status_dict[parameters['product']]) + '.00')
	response = response_dict[action].replace('@product', parameters['product']).replace('@number', '$' + str(status_dict[parameters['product']]) + '.00')
	contexts = {}

	return {
		"speech": response,
		"displayText": response,
		"contextOut": [contexts]
	}

def salesStatus(req, paid, booked, returned, date_period_parameter):

	sales, response_dict = sheets_get(sh)
	parameters = req['result']['parameters']
	action = req['result']['action']

	try:
		if date_period_parameter == False:
			parameters_list = parameters_extractor(parameters)
			response_list = [x[7] for x in filter(lambda x: x if len(parameters_list) == len(set(x).intersection(set(parameters_list))) else None, sales.get_all_values())]
		else:
			date_period = parameters['date']['date-period']
			time_period_matched_list = filter(lambda x: x if len(x[1]) > 0 and datetime.datetime.strptime(date_period[0], '%m/%d/%Y') <= datetime.datetime.strptime(x[1], '%m/%d/%Y') <= datetime.datetime.strptime(date_period[1], '%m/%d/%Y') else None, sales.get_all_values()[1:])
			del parameters['date']
			parameters_list = parameters_extractor(parameters)
			response_list = [x[7] for x in filter(lambda x: x if len(parameters_list) == len(set(x).intersection(set(parameters_list))) else None, time_period_matched_list)]

		if len(response_list) > 0:
			if paid == False and booked == False and returned == False:
				#response = "It's " + response_list[0].lower() + '.'
				response = re.sub('@\w+', '%s', response_dict[action]) %(response_list[0].lower())
			elif paid == True:
				if response_list[0] == 'Paid':
					#response = "Yes, it's paid."
					response = response_dict[action][0]
				else:
					#response = "No, it's " + response_list[0].lower() + '.'
					response = re.sub('@\w+', '%s', response_dict[action][1]) % (response_list[0].lower())
			elif booked == True:
				if response_list[0] == 'Booked':
					#response = "Yes, it's returned."
					response = response_dict[action][0]
				else:
					#response = "No, it's " + response_list[0].lower() + '.'
					response = re.sub('@\w+', '%s', response_dict[action][1]) % (response_list[0].lower())
			elif returned == True:
				if response_list[0] == 'Returned':
					#response = "Yes, it's paid."
					response = response_dict[action][0]
				else:
					#response = "No, it's " + response_list[0].lower() + '.'
					response = re.sub('@\w+', '%s', response_dict[action][1]) % (response_list[0].lower())
		else:
			response = response_dict['results.not.found']

	except TypeError:
		response = response_dict['code.error']

	contexts = {}

	return {
		"speech": response,
		"displayText": response,
		"contextOut": [contexts]
	}

def salesRevenue(req, date_period_parameter):
	sales, response_dict = sheets_get(sh)

	parameters = req['result']['parameters']
	action = req['result']['action']

	try:
		if date_period_parameter == False:
			parameters_list = parameters_extractor(parameters)
			response_num = reduce(lambda x, y: x + y, [int(float(i[6].replace('$', '').replace(',', ''))) for i in sales.get_all_values() if len(parameters_list) == len(set(i).intersection(set(parameters_list)))])
		else:
			date_period = parameters['date']['date-period']
			time_period_matched_list = filter(lambda x: x if len(x[1]) > 0 and datetime.datetime.strptime(date_period[0], '%m/%d/%Y') <= datetime.datetime.strptime(x[1], '%m/%d/%Y') <= datetime.datetime.strptime(date_period[1], '%m/%d/%Y') else None, sales.get_all_values()[1:])
			del parameters['date']
			parameters_list = parameters_extractor(parameters)
			response_num = reduce(lambda x, y: x + y, [int(float(i[6].replace('$', '').replace(',', ''))) for i in time_period_matched_list if len(parameters_list) == len(set(i).intersection(set(parameters_list)))])
		#response = "Total price is $" + str(response_num) + '.00'
		response = re.sub('@\w+', '%s', response_dict[action]) % ('$' + str(response_num) + '.00')
	except TypeError:
		response = response_dict['results.not.found']

	contexts = {}

	return {
		"speech": response,
		"displayText": response,
		"contextOut": [contexts]
	}

def salesProductQuantity(req, paid, booked, returned, date_period_parameter):
	sales, response_dict = sheets_get(sh)

	parameters = req['result']['parameters']
	action = req['result']['action']

	try:
		if paid == False and booked == False and returned == False:
			if date_period_parameter == False:
				parameters_list = parameters_extractor(parameters)
				response_num = reduce(lambda x, y: x + y, [int(i[4]) for i in sales.get_all_values() if len(parameters_list) == len(set(i).intersection(set(parameters_list)))])
				#response = "Total amount of sold " + req['result']['parameters']['product'] + ' is ' + str(response_num) + '.'
				#response = re.sub('@\w+', '%s', response_dict[action]) % (req['result']['parameters']['product'], response_num)
				response = response_dict[action].replace('@product', req['result']['parameters']['product']).replace('@number', response_num)
			else:
				date_period = parameters['date']['date-period']
				time_period_matched_list = filter(lambda x: x if len(x[1]) > 0 and datetime.datetime.strptime(date_period[0], '%m/%d/%Y') <= datetime.datetime.strptime(x[1], '%m/%d/%Y') <= datetime.datetime.strptime(date_period[1], '%m/%d/%Y') else None, sales.get_all_values()[1:])
				del parameters['date']
				parameters_list = parameters_extractor(parameters)
				response_num = reduce(lambda x, y: x + y, [int(float(i[4].replace('$', '').replace(',', ''))) for i in time_period_matched_list if len(parameters_list) == len(set(i).intersection(set(parameters_list)))])
				#response = "Total amount of sold " + req['result']['parameters']['product'] + ' is ' + str(response_num) + '.'
				#response = re.sub('@\w+', '%s', response_dict[action]) % (req['result']['parameters']['product'], response_num)
				response = response_dict[action].replace('@product', req['result']['parameters']['product']).replace('@number', response_num)
		elif paid == True:
			if date_period_parameter == False:
				parameters_list = parameters_extractor(parameters)
				response_num = reduce(lambda x, y: x + y, [int(i[4]) for i in sales.get_all_values() if len(parameters_list) == len(set(i).intersection(set(parameters_list))) and 'Paid' in i])
				#response = "Total amount of sold " + req['result']['parameters']['product'] + ' is ' + str(response_num) + '.'
				#response = re.sub('@\w+', '%s', response_dict[action]) % (req['result']['parameters']['product'], response_num)
				response = response_dict[action].replace('@product', req['result']['parameters']['product']).replace('@number', response_num)
			else:
				date_period = parameters['date']['date-period']
				time_period_matched_list = filter(lambda x: x if len(x[1]) > 0 and datetime.datetime.strptime(date_period[0], '%m/%d/%Y') <= datetime.datetime.strptime(x[1], '%m/%d/%Y') <= datetime.datetime.strptime(date_period[1], '%m/%d/%Y') else None, sales.get_all_values()[1:])
				del parameters['date']
				parameters_list = parameters_extractor(parameters)
				response_num = reduce(lambda x, y: x + y, [int(float(i[4].replace('$', '').replace(',', ''))) for i in time_period_matched_list if len(parameters_list) == len(set(i).intersection(set(parameters_list))) and 'Paid' in i])
				#response = "Total amount of sold " + req['result']['parameters']['product'] + ' is ' + str(response_num) + '.'
				#response = re.sub('@\w+', '%s', response_dict[action]) % (req['result']['parameters']['product'], response_num)
				response = response_dict[action].replace('@product', req['result']['parameters']['product']).replace('@number', response_num)
		elif booked == True:
			if date_period_parameter == False:
				parameters_list = parameters_extractor(parameters)
				response_num = reduce(lambda x, y: x + y, [int(i[4]) for i in sales.get_all_values() if len(parameters_list) == len(set(i).intersection(set(parameters_list))) and 'Booked' in i])
				#response = "Total amount of booked " + req['result']['parameters']['product'] + ' is ' + str(response_num) + '.'
				#response = re.sub('@\w+', '%s', response_dict[action]) % (req['result']['parameters']['product'], response_num)
				response = response_dict[action].replace('@product', req['result']['parameters']['product']).replace('@number', response_num)
			else:
				date_period = parameters['date']['date-period']
				time_period_matched_list = filter(lambda x: x if len(x[1]) > 0 and datetime.datetime.strptime(date_period[0], '%m/%d/%Y') <= datetime.datetime.strptime(x[1], '%m/%d/%Y') <= datetime.datetime.strptime(date_period[1], '%m/%d/%Y') else None, sales.get_all_values()[1:])
				del parameters['date']
				parameters_list = parameters_extractor(parameters)
				response_num = reduce(lambda x, y: x + y, [int(float(i[4].replace('$', '').replace(',', ''))) for i in time_period_matched_list if len(parameters_list) == len(set(i).intersection(set(parameters_list))) and 'Booked' in i])
				#response = "Total amount of booked " + req['result']['parameters']['product'] + ' is ' + str(response_num) + '.'
				#response = re.sub('@\w+', '%s', response_dict[action]) % (req['result']['parameters']['product'], response_num)
				response = response_dict[action].replace('@product', req['result']['parameters']['product']).replace('@number', response_num)
		elif returned == True:
			if date_period_parameter == False:
				parameters_list = parameters_extractor(parameters)
				response_num = reduce(lambda x, y: x + y, [int(i[4]) for i in sales.get_all_values() if len(parameters_list) == len(set(i).intersection(set(parameters_list))) and 'Returned' in i])
				#response = "Total amount of returned " + req['result']['parameters']['product'] + ' is ' + str(response_num) + '.'
				#response = re.sub('@\w+', '%s', response_dict[action]) % (req['result']['parameters']['product'], response_num)
				response = response_dict[action].replace('@product', req['result']['parameters']['product']).replace('@number', response_num)
			else:
				date_period = parameters['date']['date-period']
				time_period_matched_list = filter(lambda x: x if len(x[1]) > 0 and datetime.datetime.strptime(date_period[0], '%m/%d/%Y') <= datetime.datetime.strptime(x[1], '%m/%d/%Y') <= datetime.datetime.strptime(date_period[1], '%m/%d/%Y') else None, sales.get_all_values()[1:])
				del parameters['date']
				parameters_list = parameters_extractor(parameters)
				response_num = reduce(lambda x, y: x + y, [int(float(i[4].replace('$', '').replace(',', ''))) for i in time_period_matched_list if len(parameters_list) == len(set(i).intersection(set(parameters_list))) and 'Returned' in i])
				#response = "Total amount of returned " + req['result']['parameters']['product'] + ' is ' + str(response_num) + '.'
				#response = re.sub('@\w+', '%s', response_dict[action]) % (req['result']['parameters']['product'], response_num)
				response = response_dict[action].replace('@product', req['result']['parameters']['product']).replace('@number', response_num)
	except TypeError:
		response = response_dict['results.not.found']

	contexts = {}

	return {
		"speech": response,
		"displayText": response,
		"contextOut": [contexts]
	}

def salesPerson(req, date_period_parameter):

	sales, response_dict = sheets_get(sh)

	parameters = req['result']['parameters']
	action = req['result']['action']

	try:
		if date_period_parameter == False:
			parameters_list = parameters_extractor(parameters)
			response_list = [x[2] for x in filter(lambda x: x if len(parameters_list) == len(set(x).intersection(set(parameters_list))) else None, sales.get_all_values())]
		else:
			date_period = parameters['date']['date-period']
			time_period_matched_list = filter(lambda x: x if len(x[1]) > 0 and datetime.datetime.strptime(date_period[0], '%m/%d/%Y') <= datetime.datetime.strptime(x[1], '%m/%d/%Y') <= datetime.datetime.strptime(date_period[1], '%m/%d/%Y') else None, sales.get_all_values()[1:])
			del parameters['date']
			parameters_list = parameters_extractor(parameters)
			response_list = [x[2] for x in filter(lambda x: x if len(parameters_list) == len(set(x).intersection(set(parameters_list))) else None, time_period_matched_list)]
		print date_period_parameter
		#print date_period
		print response_list
		response_list_string = ''
		if len(response_list) > 0:
			if len(response_list) > 1:
				response_list_uniq = list(set(response_list))
				for i in response_list_uniq:
					if response_list_uniq.index(i) != len(response_list_uniq) -2:
						response_list_string += i + ', '
					else:
						response_list_string += i + ' and '

				#response = "It's " + response_list_string[:-2] + '.'
				response = re.sub('@\w+', '%s', response_dict[action]) % (response_list_string[:-2])
			else:
				#response = "It's " + response_list[0] + '.'
				response = re.sub('@\w+', '%s', response_dict[action]) % (response_list[0])
		else:
			response = response_dict['results.not.found']
	except TypeError:
		response = response_dict['results.not.found']

	contexts = {}

	return {
		"speech": response,
		"displayText": response,
		"contextOut": [contexts]
	}

def salesPersonMostLeastMoney(req):

	sales, response_dict = sheets_get(sh)
	parameters = req['result']['parameters']
	action = req['result']['action']

	sales_rep_sale_dict = {}

	if len(parameters['date']) > 0:
		if parameters['date'].has_key('date'):
			for i in [i for i in sales.get_all_values() if len(i[1]) > 0 and i[1] == parameters['date']['date']]:
				num = int(float(i[6].replace('$', '').replace(',', '')))
				if sales_rep_sale_dict.has_key(i[2]) == False:
					sales_rep_sale_dict[i[2]] = num
				else:
					sales_rep_sale_dict[i[2]] += num
		else:
			for i in [i for i in sales.get_all_values()[1:] if len(i[1]) > 0 and datetime.datetime.strptime(parameters['date']['date-period'][0], '%m/%d/%Y') <= datetime.datetime.strptime(i[1], '%m/%d/%Y') <= datetime.datetime.strptime(parameters['date']['date-period'][1], '%m/%d/%Y')]:
				num = int(float(i[6].replace('$', '').replace(',', '')))
				if sales_rep_sale_dict.has_key(i[2]) == False:
					sales_rep_sale_dict[i[2]] = num
				else:
					sales_rep_sale_dict[i[2]] += num
	else:
		for i in [i for i in sales.get_all_values()[1:] if len(i[1]) > 0]:
			num = int(float(i[6].replace('$', '').replace(',', '')))
			if sales_rep_sale_dict.has_key(i[2]) == False:
				sales_rep_sale_dict[i[2]] = num
			else:
				sales_rep_sale_dict[i[2]] += num

	if action == 'sales.person.most_money':
		for i in sales_rep_sale_dict:
			if sales_rep_sale_dict[i] == max(sales_rep_sale_dict.values()):
				max_sales_rep_sale = i
		#response = 'It is ' + max_sales_rep_sale + '.'
		response = re.sub('@\w+', '%s', response_dict[action]) % (max_sales_rep_sale)
	elif action == 'sales.person.least_money':
		for i in sales_rep_sale_dict:
			if sales_rep_sale_dict[i] == min(sales_rep_sale_dict.values()):
				max_sales_rep_sale = i
		#response = 'It is ' + max_sales_rep_sale + '.'
		response = re.sub('@\w+', '%s', response_dict[action]) % (max_sales_rep_sale)

	contexts = {}

	return {
		"speech": response,
		"displayText": response,
		"contextOut": [contexts]
	}

def salesDate(req, date_period_parameter):

	sales, response_dict = sheets_get(sh)

	parameters = req['result']['parameters']
	action = req['action']['action']

	try:
		if date_period_parameter == False:
			parameters_list = parameters_extractor(parameters)
			response_list = [x[1] for x in filter(lambda x: x if len(parameters_list) == len(set(x).intersection(set(parameters_list))) else None, sales.get_all_values())]
		else:
			date_period = parameters['date']['date-period']
			time_period_matched_list = filter(lambda x: x if len(x[1]) > 0 and datetime.datetime.strptime(date_period[0], '%m/%d/%Y') <= datetime.datetime.strptime(x[1], '%m/%d/%Y') <= datetime.datetime.strptime(date_period[1], '%m/%d/%Y') else None, sales.get_all_values()[1:])
			del parameters['date']
			parameters_list = parameters_extractor(parameters)
			response_list = [x[1] for x in filter(lambda x: x if len(parameters_list) == len(set(x).intersection(set(parameters_list))) else None, time_period_matched_list)]

		if response_list and len(response_list) > 0:
			#response = "On " + str(response_list[0]) + '.'
			response = re.sub('@\w+', '%s', response_dict[action]) % (response_list[0])
		else:
			response = response_dict['results.not.found']
	except TypeError:
		response = response_dict['results.not.found']

	contexts = {}

	return {
		"speech": response,
		"displayText": response,
		"contextOut": [contexts]
	}

def salesDateMostLeastMoney(req):

	sales, response_dict = sheets_get(sh)
	action = req['result']['action']
	parameters = req['result']['parameters']

	sales_rep_sale_dict = {}

	for i in [i for i in sales.get_all_values()[1:] if len(i[1]) > 0]:
		num = int(float(i[6].replace('$', '').replace(',', '')))

		if sales_rep_sale_dict.has_key(i[1]) == False:
			sales_rep_sale_dict[i[1]] = num
		else:
			sales_rep_sale_dict[i[1]] += num

	if action == 'sales.date.most_money':
		for i in sales_rep_sale_dict:
			if sales_rep_sale_dict[i] == max(sales_rep_sale_dict.values()):
				max_sales_rep_sale = i
	elif action == 'sales.date.least_money':
		for i in sales_rep_sale_dict:
			if sales_rep_sale_dict[i] == min(sales_rep_sale_dict.values()):
				max_sales_rep_sale = i

	#response = 'On ' + str(max_sales_rep_sale) + '.'
	response = re.sub('@\w+', '%s', response_dict[action]) % (max_sales_rep_sale)
	contexts = {}

	return {
		"speech": response,
		"displayText": response,
		"contextOut": [contexts]
	}

def salesProductMostLeastMoney(req):

	sales, response_dict = sheets_get(sh)
	parameters = req['result']['parameters']
	action = req['result']['action']

	sales_rep_sale_dict = {}

	if len(parameters['date']) > 0:
		if parameters['date'].has_key('date'):
			for i in [i for i in sales.get_all_values() if len(i[1]) > 0 and i[1] == parameters['date']['date']]:
				num = int(float(i[6].replace('$', '').replace(',', '')))
				if sales_rep_sale_dict.has_key(i[3]) == False:
					sales_rep_sale_dict[i[3]] = num
				else:
					sales_rep_sale_dict[i[3]] += num
		else:
			for i in [i for i in sales.get_all_values()[1:] if len(i[1]) > 0 and datetime.datetime.strptime(parameters['date']['date-period'][0], '%m/%d/%Y') <= datetime.datetime.strptime(i[1], '%m/%d/%Y') <= datetime.datetime.strptime(parameters['date']['date-period'][1], '%m/%d/%Y')]:
				num = int(float(i[6].replace('$', '').replace(',', '')))
				if sales_rep_sale_dict.has_key(i[3]) == False:
					sales_rep_sale_dict[i[3]] = num
				else:
					sales_rep_sale_dict[i[3]] += num
	else:
		for i in [i for i in sales.get_all_values()[1:] if len(i[1]) > 0]:
			num = int(float(i[6].replace('$', '').replace(',', '')))
			if sales_rep_sale_dict.has_key(i[3]) == False:
				sales_rep_sale_dict[i[3]] = num
			else:
				sales_rep_sale_dict[i[3]] += num

	if action == 'sales.product.most_money':
		for i in sales_rep_sale_dict:
			if sales_rep_sale_dict[i] == max(sales_rep_sale_dict.values()):
				max_sales_rep_sale = i
				revenue_product = sales_rep_sale_dict[i]
		#response = max_sales_rep_sale + ' has generated the most revenue of $' + str(revenue_product) + '.00.'
		#response = re.sub('@\w+', '%s', response_dict[action]) % (max_sales_rep_sale, '$' + str(revenue_product) + '.00')
		response = response_dict[action].replace('@person', max_sales_rep_sale).replace('@number', '$' + str(revenue_product) + '.00')
	elif action == 'sales.product.least_money':
		for i in sales_rep_sale_dict:
			if sales_rep_sale_dict[i] == min(sales_rep_sale_dict.values()):
				max_sales_rep_sale = i
				revenue_product = sales_rep_sale_dict[i]
		#response = max_sales_rep_sale + ' has generated the least revenue of $' + str(revenue_product) + '.00.'
		#response = re.sub('@\w+', '%s', response_dict[action]) % (max_sales_rep_sale, '$' + str(revenue_product) + '.00')
		response = response_dict[action].replace('@person', max_sales_rep_sale).replace('@number', '$' + str(revenue_product) + '.00')

	contexts = {}

	return {
		"speech": response,
		"displayText": response,
		"contextOut": [contexts]
	}

def salesProductBestLeastSelling(req):

	sales, response_dict = sheets_get(sh)
	parameters = req['result']['parameters']
	action = req['result']['action']

	sales_rep_sale_dict = {}

	if len(parameters['date']) > 0:
		if parameters['date'].has_key('date'):
			for i in [i for i in sales.get_all_values() if len(i[1]) > 0 and i[1] == parameters['date']['date']]:
				num = int(float(i[4].replace('$', '').replace(',', '')))
				if sales_rep_sale_dict.has_key(i[3]) == False:
					sales_rep_sale_dict[i[3]] = num
				else:
					sales_rep_sale_dict[i[3]] += num
		else:
			for i in [i for i in sales.get_all_values()[1:] if len(i[1]) > 0 and datetime.datetime.strptime(parameters['date']['date-period'][0], '%m/%d/%Y') <= datetime.datetime.strptime(i[1], '%m/%d/%Y') <= datetime.datetime.strptime(parameters['date']['date-period'][1], '%m/%d/%Y')]:
				num = int(float(i[4].replace('$', '').replace(',', '')))
				if sales_rep_sale_dict.has_key(i[3]) == False:
					sales_rep_sale_dict[i[3]] = num
				else:
					sales_rep_sale_dict[i[3]] += num
	else:
		for i in [i for i in sales.get_all_values()[1:] if len(i[1]) > 0]:
			num = int(float(i[4].replace('$', '').replace(',', '')))
			if sales_rep_sale_dict.has_key(i[3]) == False:
				sales_rep_sale_dict[i[3]] = num
			else:
				sales_rep_sale_dict[i[3]] += num


	#max_quantity = int(float(sales.cell(sales.find(max(sales_rep_sale_dict)).row, sales.find(max(sales_rep_sale_dict)).col+1).value.replace('$', '').replace(',', '')))
	for i in sales_rep_sale_dict:
		if action == 'sales.product.best_selling':
			if sales_rep_sale_dict[i] == max(sales_rep_sale_dict.values()):
				max_sales_rep_sale = i
		elif action == 'sales.product.least_selling':
			if sales_rep_sale_dict[i] == min(sales_rep_sale_dict.values()):
				max_sales_rep_sale = i

	max_sales_rep_sale_list = sales.findall(max_sales_rep_sale)
	sales_list = sales.findall(str(sales_rep_sale_dict[max_sales_rep_sale]))

	for i in max_sales_rep_sale_list:
		for ii in sales_list:
			if i.row == ii.row:
				product_sale_total = sales.cell(ii.row, ii.col + 2).value

	if sales_list and len(sales_list) > 0:
		#response = "We've sold " + str(sales_rep_sale_dict[max_sales_rep_sale]) + " units of " + str(max_sales_rep_sale) + " for the total of " + str(product_sale_total) + "."
		#response = re.sub('@\w+', '%s', response_dict[action]) % (sales_rep_sale_dict[max_sales_rep_sale], max_sales_rep_sale, product_sale_total)
		response = response_dict[action].replace('@number', sales_rep_sale_dict[max_sales_rep_sale]).replace('@product', max_sales_rep_sale).replace('@sum', product_sale_total)
	else:
		max_sales_rep_sale_val = sales.find(max_sales_rep_sale)
		total_price = int((sales.cell(max_sales_rep_sale_val.row, max_sales_rep_sale_val.col+2).value).replace('$', '').replace('.00', '')) * sales_rep_sale_dict[max_sales_rep_sale]
		#response = "We've sold " + str(sales_rep_sale_dict[max_sales_rep_sale]) + " units of " + str(max_sales_rep_sale) + " for the total of $" + str(total_price) + ".00."
		#response = re.sub('@\w+', '%s', response_dict[action]) % (sales_rep_sale_dict[max_sales_rep_sale], max_sales_rep_sale,"$" + str(total_price) + ".00")
		response = response_dict[action].replace('@number', sales_rep_sale_dict[max_sales_rep_sale]).replace('@product', max_sales_rep_sale).replace('@sum', "$" + str(total_price) + ".00")

	contexts = {}

	return {
		"speech": response,
		"displayText": response,
		"contextOut": [contexts]
	}

def salesStatusQuantity(req, paid, booked, returned, date_period_parameter):

	sales, response_dict = sheets_get(sh)

	parameters = req['result']['parameters']
	action = req['result']['action']

	try:
		if date_period_parameter == False:
			parameters_list = parameters_extractor(parameters)
			response_list = [x[7] for x in filter(lambda x: x if len(parameters_list) == len(set(x).intersection(set(parameters_list))) else None, sales.get_all_values())]
		else:
			date_period = parameters['date']['date-period']
			time_period_matched_list = filter(lambda x: x if len(x[1]) > 0 and datetime.datetime.strptime(date_period[0], '%m/%d/%Y') <= datetime.datetime.strptime(x[1], '%m/%d/%Y') <= datetime.datetime.strptime(date_period[1], '%m/%d/%Y') else None, sales.get_all_values()[1:])
			del parameters['date']
			parameters_list = parameters_extractor(parameters)
			response_list = [x[7] for x in filter(lambda x: x if len(parameters_list) == len(set(x).intersection(set(parameters_list))) else None, time_period_matched_list)]

		if paid == False and booked == False and returned == False:
			if len(response_list) > 0:
				response = "It's " + response_list[0].lower() + '.'
			else:
				response = response_dict['results.not.found']
		elif paid == True:
			quantity = [x[7] for x in filter(lambda x: x if len(parameters_list) == len(set(x).intersection(set(parameters_list))) else None, sales.get_all_values())].count('Paid')
			response = re.sub('@\w+', '%s', response_dict[action]) % (quantity)
		elif booked == True:
			quantity = [x[7] for x in filter(lambda x: x if len(parameters_list) == len(set(x).intersection(set(parameters_list))) else None, sales.get_all_values())].count('Booked')
			response = re.sub('@\w+', '%s', response_dict[action]) % (quantity)
		elif returned == True:
			quantity = [x[7] for x in filter(lambda x: x if len(parameters_list) == len(set(x).intersection(set(parameters_list))) else None, sales.get_all_values())].count('Returned')
			response = re.sub('@\w+', '%s', response_dict[action]) % (quantity)

	except TypeError:
		response = response_dict['results.not.found']

	contexts = {}

	return {
		"speech": response,
		"displayText": response,
		"contextOut": [contexts]
	}

def salesQuantity(req, date_period_parameter):
	sales, response_dict = sheets_get(sh)

	parameters = req['result']['parameters']
	action = req['result']['action']

	if parameters['date'].has_key('date'):
		date_filterd_list = filter(lambda x: x if x[1] == parameters['date']['date'] else None, sales.get_all_values())
	else:
		date_period = parameters['date']['date-period']
		date_filterd_list = filter(lambda x: x if len(x[1]) > 0 and datetime.datetime.strptime(date_period[0], '%m/%d/%Y') <= datetime.datetime.strptime(x[1], '%m/%d/%Y') <= datetime.datetime.strptime(date_period[1], '%m/%d/%Y') else None, sales.get_all_values()[1:])

	# quantity_dict = {i[3]: [i[4],i[6]] for i in date_filterd_list}
	quantity_dict = {}
	for i in date_filterd_list:
		i[3] = [i[4], i[6]]
	if len(quantity_dict) != 0:
		if len(quantity_dict) > 1:
			#strng = 'On '+str(parameters['date']['date'])+' we have sold '
			#strng = 'We have sold '
			strng = ''
			for i in quantity_dict: strng += quantity_dict[i][0] + ' units of ' + i + ', '
			#response = strng[:-2] + '. The total of all sold items is $' + str(reduce(lambda x, y: x + y, [int(float(quantity_dict[i][1].replace('$', '').replace(',', ''))) for i in quantity_dict])) + '.00.'
			#response = re.sub('@\w+', '%s', response_dict[action][0]) % (strng[:-2], '$' + str(reduce(lambda x, y: x + y, [int(float(quantity_dict[i][1].replace('$', '').replace(',', ''))) for i in quantity_dict])) + '.00')
			response = response_dict[action][0].replace('@units', strng[:-2]).replace('@income', '$' + str(reduce(lambda x, y: x + y, [int(float(quantity_dict[i][1].replace('$', '').replace(',', ''))) for i in quantity_dict])) + '.00')
		else:
			#response = 'On '+str(parameters['date']['date'])+' we have sold '+str(quantity_dict.values()[0][0])+' units of '+quantity_dict.keys()[0]+' for the total of '+str(quantity_dict.values()[0][1])+'.'
			#response = 'We have sold '+str(quantity_dict.values()[0][0])+' units of '+quantity_dict.keys()[0]+' for the total of '+str(quantity_dict.values()[0][1])+'.'
			#response = re.sub('@\w+', '%s', response_dict[action][1]) % (quantity_dict.values()[0][0], quantity_dict.keys()[0],quantity_dict.values()[0][1])
			response = response_dict[action][1].replace('@number', quantity_dict.values()[0][0]).replace('@product', quantity_dict.keys()[0]).replace('@income', quantity_dict.values()[0][1])

	else:
		response = "We haven't sold anything."

	contexts = {}

	return {
		"speech": response,
		"displayText": response,
		"contextOut": [contexts]
	}

@app.route('/test', methods=['GET'])
def test():
	return 'salesagent Test is done!'


if __name__ == '__main__':
	port = int(os.getenv('PORT', 5000))

	app.run(
		debug=True,
		port=port,
		host='0.0.0.0'
	)
