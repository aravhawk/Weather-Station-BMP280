#!/usr/bin/env/python

# Import all libraries we need!
from smbus import SMBus
from bmp280 import BMP280
import time
import datetime
from datetime import date
from openpyxl import load_workbook

# Initialize the BMP280
bus = SMBus(1)
bmp280 = BMP280(i2c_dev=bus)

# Take first reading and disgard it to avoid garbage first row
temperature = bmp280.get_temperature()
pressure = bmp280.get_pressure()
humidity = bmp280.get_humidity()
time.sleep(1)	

# Load the workbook and select the sheet
wb = load_workbook('/home/pi/Python_Code/weather.xlsx')
sheet = wb['Sheet1']

try:
	while True:
		
		# Read the sensor and get date and time
		temperature = round(bmp280.get_temperature(),1)
		pressure = round(bmp280.get_pressure(),1)
		today = date.today()
		now = datetime.datetime.now().time()

		# Inform the user!
		print('Adding this data to the spreadsheet:')
		print(today)
		print(now)
		print('{}*C {}hPa'.format(temperature, pressure))
	
		# Append data to the spreadsheet
		row = (today, now, temperature, pressure, humidity)
		sheet.append(row)
		
		#Save the workbook
		wb.save('/home/pi/Python_Code/weather.xlsx')

		# Wait for 10 minutes seconds (600 seconds)
		time.sleep(600)

finally:
	# Make sure the workbook is saved!
	wb.save('/home/pi/Python_Code/weather.xlsx')
	
	print('Goodbye!')
