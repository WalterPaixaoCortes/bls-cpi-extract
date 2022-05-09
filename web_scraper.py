import requests
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import json
import numpy as np
import datetime
from datetime import timedelta
import holidays
import os
from pyquery import PyQuery as pq
import math
import sys

np.set_printoptions(threshold=sys.maxsize)


path = os.getcwd()

def convert_yyyyq(dates):
	years = dates.values // 10
	months = (dates.values % 10) * 3 - 2
	date = [datetime.datetime(year=years[i], month=months[i], day=1) for i in range(0, len(years))]
	
	date = pd.to_datetime(date)
	return date.date

def get_ncreif():
	curr_y = str(pd.to_datetime("today").year)
	curr_q = str(pd.to_datetime("today").quarter - 2)

	req_sesh_ncreif = requests.session()
		
	# Establish preliminary values
	ncreif_login_url = 'https://www.ncreif.org/login/'
	ncreif_login_values = {'username': 'jeff.stonger@ascentris.com',
			'password': 'Ascentris1!'}

	#Get the RequestVerificationToken
	ncreif_login_page_response = req_sesh_ncreif.get(ncreif_login_url)
	d = pq(ncreif_login_page_response.text)
	ncreif_login_requestverificationtoken = d("form[action=\"/login/\"]").find("input[name=\"__RequestVerificationToken\"]").attr("value")
	ncreif_login_values['__RequestVerificationToken'] = ncreif_login_requestverificationtoken

	#Login to the website
	ncreif_login_response = req_sesh_ncreif.post(ncreif_login_url, data=ncreif_login_values)

	ncreif_data_download_response = req_sesh_ncreif.get("https://www.ncreif.org/globalassets/member-site/products/property/" + curr_y + "q" + curr_q + "/ntbi_vw_" + curr_y + curr_q + ".xlsx")
	wb = open("NCREIF-National-Data.xlsx", "wb")
	wb.write(ncreif_data_download_response.content)
	wb.close()

	ncreif_data_download_response = req_sesh_ncreif.get("https://www.ncreif.org/globalassets/member-site/products/property/" + curr_y + "q" + curr_q + "/trendsreport_" + curr_y + curr_q + ".xlsx")
	wb = open("NCREIF-Trends-Report.xlsx", "wb")
	wb.write(ncreif_data_download_response.content)
	wb.close()

	print("NCREIF Download complete...")

	ncreif = pd.read_excel("NCREIF-National-Data.xlsx", sheet_name="TotalNTBI VW")
	ncreif2 = pd.read_excel("NCREIF-Trends-Report.xlsx", sheet_name="Returns All Props", usecols="B,K:M", header=4)

	ncreif.YYYYQ = convert_yyyyq(ncreif.YYYYQ)
	ncreif2.YYYYQ = convert_yyyyq(ncreif2.YYYYQ)

	n = ncreif2.merge(ncreif)
	return n

def pull_data(data, ncreif):

	us_holidays = holidays.US(years=range(1940, pd.to_datetime("today").year)).keys()

	m = []
	q = []
	d = []
	y = []
	for i, row in data.iterrows():
		if (row.Frequency == "Monthly"):
			m.append(row.Data)
		elif (row.Frequency == "Daily" or row.Frequency == "Weekly"):
			d.append(row.Data)
		elif (row.Frequency == "Quarterly"):
			q.append(row.Data)
		else:
			y.append(row.Data)

	monthly = pd.DataFrame(columns=np.insert(m, 0, "Dates"))
	monthly["Dates"] = pd.date_range(start="01/01/1940", end=pd.to_datetime("today"), freq="MS").date

	weekmask = 'Mon Tue Wed Thu Fri'
	daily = pd.DataFrame(columns=np.insert(d, 0, "Dates"))
	daily["Dates"] = pd.bdate_range(start="01/01/1940", end=pd.to_datetime("today"), freq="C", weekmask=weekmask, holidays=us_holidays).date

	quarterly = pd.DataFrame(columns=np.insert(q, 0, "Dates"))
	quarterly["Dates"] = pd.bdate_range(start="01/01/1940", end=pd.to_datetime("today"), freq="QS").date

	count = 0
	for i, row in data.iterrows():
		
		pulled = False
		"""
		PULL DATA FROM SPECIFIC SOURCE
		"""
		if (row.Source == "FRED"):
			print("Retrieving %s from FRED" % row.Data)

			file = requests.get("https://api.stlouisfed.org/fred/series/observations?series_id={}&api_key={}&file_type={}".format(row.Data,"387a905393cc2ed7a711a24149359f52","json"))

			if file.status_code != 200:
				# This means something went wrong.
				print("ERROR ON %s" % row.Data)
				raise Exception('GET https://api.stlouisfed.org/fred/series/observations {}'.format(file.status_code))

			t = [i['date'] for i in file.json()['observations']]
			dates = pd.to_datetime([i['date'] for i in file.json()['observations']]).date
			values = np.array([i['value'] for i in file.json()['observations']])
			pulled = True			

		elif (row.Source == "NCREIF"):
			print("Retrieving %s from NCREIF" % row.Data)

			values = ncreif[row.Data].to_numpy().astype(str)
			dates = pd.to_datetime(list(ncreif.YYYYQ)).date
			pulled = True

		elif (row.Source == "Nareit"):
			print("Retrieving %s from Nareit" % row.Data)

			nareit_data_download = requests.get("https://www.reit.com/sites/default/files/returns/MonthlyHistoricalReturns.xls")

			wb = open("Nareit.xls", "wb")
			wb.write(nareit_data_download.content)
			wb.close()

			nareit = pd.read_excel("Nareit.xls", sheet_name="Index Data", usecols="A,W:AB", header=7).dropna()
			nareit.columns = ["Date","Total Return","Total Index","Price Return","Price Index","Income Return","Dividend Yield"]
			nareit.Date += timedelta(days=1)
			
			values = nareit[row.Data].to_numpy().astype(str)
			dates = pd.to_datetime(list(nareit.Date)).date
			pulled = True

		if (row.Source == "Yale"):
			print("Retrieving %s from Yale" % row.Data)

			yale_data_download = 0
			while(yale_data_download == 0):
				try:
					yale_data_download = requests.get("http://www.econ.yale.edu/~shiller/data/ie_data.xls")
				except requests.exceptions.ConnectionError:
					print("Error connecting to source, trying again")


			wb = open("Yale.xls", "wb")
			wb.write(yale_data_download.content)
			wb.close()

			yale = pd.read_excel("Yale.xls", sheet_name="Data", usecols="A,M", header=7).dropna()
			
			dates = pd.to_datetime([datetime.datetime.strptime(d, "%Y.%m") for d in yale["Date"].values.astype(str)]).date
			spl = np.round(yale["Date"].values % 1,2)
			mask = spl == 0.1
			dates[mask] = [d.month + 9 for d in dates[mask]]
			values = yale["CAPE"].to_numpy().astype(str)
			pulled = True


		"""
		ADD DATA TO DATAFRAME TO SAVE TO EXCEL LATER
		"""
		if (pulled):
			if (row.Frequency == "Monthly"):
				mask = monthly.Dates.isin(dates).values
				mask2 = np.isin(dates, monthly.Dates.values)
				monthly[row.Data].mask(mask, values[mask2], inplace=True)

			elif (row.Frequency == "Daily" or row.Frequency == "Weekly"):
				mask = daily.Dates.isin(dates).values
				mask2 = np.isin(dates, daily.Dates.values)
				daily[row.Data].mask(mask, values[mask2], inplace=True)

			elif (row.Frequency == "Quarterly"):
				mask = quarterly.Dates.isin(dates).values
				mask2 = np.isin(dates, quarterly.Dates.values)
				quarterly[row.Data].mask(mask, values[mask2], inplace=True)

			count += 1

	return daily, monthly, quarterly, count


def main():

	xls = pd.ExcelFile("DataSummary.xlsx")
	data = pd.read_excel(xls, "Sheet1", usecols="B:F").dropna()
	print(data)

	ncreif = get_ncreif()
	
	daily, monthly, quarterly, count = pull_data(data, ncreif)

	file = "output.xlsx"
	with pd.ExcelWriter(file, engine="openpyxl") as writer:
		daily.to_excel(writer, sheet_name="Daily")
		monthly.to_excel(writer, sheet_name="Monthly")
		quarterly.to_excel(writer, sheet_name="Quarterly")

	print("%d fields populated in output file, '%s'" % (count, file))


if __name__ == '__main__':
	main()
