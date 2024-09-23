
#import matplotlib as mpl
#import matplotlib.patches as patches
#from matplotlib.image import BboxImage
import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
#import locale
from datetime import datetime as dt
from datetime import timedelta as td
import pathlib
from pathlib import Path
from pandas.io.formats.style import Styler
#from moneyed import Money
#from moneyed.l10n import format_money
#07808import dataframe_image as dfi
from pandas.plotting import table
import matplotlib.image as mpimg


# Set locale to United States for locale method use.  NOTE: Some functions like locale.currency will change floating point to string type and prevent further math ops.
#locale.setlocale(locale.LC_ALL, '')

#find current working directory to know where to put csv file to be manipulated
#os.getcwd()
#List files and folders in current working directory
#os.listdir()

#Input start date of reporting period from terminal
print('Enter the starting date of the report to be run as MM.DD.YYYY')
ReportStartDate = str(input())

#Method to find and import the raw POS file
def get_file_names_with_strings(str_list):
    full_list = os.listdir("C:\\Users\\satkinson\\Documents\\Igadi\\POS Data Input")
    final_list = [nm for ps in str_list for nm in full_list if ps in nm]

    return final_list
#  make a table that looks like a spreadsheet with column and row headers and export as jpg
#dfi.export(df_sales,'Igadi Sales Report ', {ReportStartDate} {FilenameFormatter} {ReportEndDate},'.jpg')
CurrentPOSdata = get_file_names_with_strings([ReportStartDate])
CurrentPOSdata = str(CurrentPOSdata)[1:-1]
CurrentPOSdata = str(CurrentPOSdata)[1:-1]

#print(ReportEndDate)
#find current working directory to know where to put csv file to be manipulated
#os.getcwd()

os.chdir('C:\\Users\\satkinson\\Documents\\Igadi\\POS Data Input')
#List files and folders in current working directory
#os.getcwd()
#os.listdir
df_sales = pd.read_csv(CurrentPOSdata, index_col=False)
#print(df_sales)

#Calculate the end date of the report period given a 7 day reporting week
ReportEndDate = (dt.strptime(ReportStartDate, '%m.%d.%Y') + td(days=6)).strftime('%m.%d.%Y')
#Calculate previous week file name "Start" and "End"
PreviousWeekStartDate = (dt.strptime(ReportStartDate, '%m.%d.%Y') - td(days=7)).strftime('%m.%d.%Y')
PreviousWeekEndDate = (dt.strptime(ReportStartDate, '%m.%d.%Y') - td(days=1)).strftime('%m.%d.%Y')
#Calculate previous year file name
PreviousYearStartDate = (dt.strptime(ReportStartDate, '%m.%d.%Y') - td(days=364)).strftime('%m.%d.%Y')
PreviousYearEndDate = (dt.strptime(ReportStartDate, '%m.%d.%Y') - td(days=358)).strftime('%m.%d.%Y')
x = (dt.now())
CurrentYear = str(CurrentPOSdata)[-8:-4]
CurrentYearInt = int(CurrentYear)
#print(CurrentYearInt)
PreviousYear = str(CurrentYearInt - 1)
#print(PreviousYear)

#create a new folder for this week's reports
CurrentWeekReportFolder = 'C:\\Users\\satkinson\\Documents\\Igadi\\Sales Reports ' + CurrentYear + '\\Week Of ' + ReportStartDate + ' - ' + ReportEndDate + '\\'
if not os.path.isdir(CurrentWeekReportFolder):
    os.makedirs (CurrentWeekReportFolder)

#START OF GROOMING DATA
#add new colume with Sales - Tax
df_sales["Final Sale"] = df_sales['Final Sale'].round(2)
df_sales["Tax In Dollars"] = df_sales['Tax In Dollars'].round(2)
df_sales ['Sales_Minus_Tax'] = df_sales['Final Sale'].round(2) - df_sales['Tax In Dollars'].round(2)

#add new column with Net Sale after taking out product cost
df_sales ['Net Sale'] = df_sales ['Sales_Minus_Tax'].round(2) - df_sales['Cost'].round(2)

#remove rows with Voided transactions
#main dataframe file is now df_sales2
#First convert any "Sale Was Voided" columns to "Voided" for consistency on input
df_sales2 = df_sales.rename(columns = {'Sale Was Voided':'Voided'}, inplace=True)
df_sales2 = df_sales.drop(df_sales[df_sales.Voided == 'yes'].index)

#Replace NaN entries
df_sales2["Final Sale"].fillna(0, inplace=True)
df_sales2["Tax In Dollars"].fillna(0, inplace=True)
df_sales2["Quantity"].fillna(0, inplace=True)
df_sales2["Cost"].fillna(0, inplace=True)
df_sales2["Weight Sold"].fillna(0, inplace=True)
df_sales2["Location"].fillna(' ', inplace=True)
df_sales2["Category"].fillna(' ', inplace=True)
df_sales2["Brand"].fillna(' ', inplace=True)
df_sales2["Product Name"].fillna('No entry', inplace=True)
df_sales2["Type"].fillna(' ', inplace=True)
df_sales2["Sale Id"].fillna(' ', inplace=True)
df_sales2["Transaction Date"].fillna(' ', inplace=True)
df_sales2["Voided"].fillna(' ', inplace=True)
df_sales2["Strain Name"].fillna(' ', inplace=True)
df_sales2["Customer Type"].fillna(' ', inplace=True)
df_sales2["Transaction Time"].fillna(' ', inplace=True)

# Move all Strain Names into Product Name column if "No entry" in Product Name
df_sales2.loc[df_sales2["Product Name"] == "No entry", "Product Name"] = df_sales2["Strain Name"]
#then delete the Strain Name Column after moving Strain Name into Product Name
del df_sales2["Strain Name"]

#Change Category for any product with a Type that includes "Cartridge" from "Concentrate" to Cartridge
df_sales2.loc[df_sales2["Type"] == "Cartridge", "Category"] = 'Cartridge'
df_sales2.loc[df_sales2["Type"] == "3rd Party - Cartridge", "Category"] = 'Cartridge'
df_sales2.loc[df_sales2["Type"] == "3rd Party - Cartridge - CBD", "Category"] = 'Cartridge'
df_sales2.loc[df_sales2["Type"] == "3rd Party - Cartridge - Hybrid", "Category"] = 'Cartridge'
df_sales2.loc[df_sales2["Type"] == "3rd Party - Cartridge - Indica", "Category"] = 'Cartridge'
df_sales2.loc[df_sales2["Type"] == "3rd Party - Cartridge - Sativa", "Category"] = 'Cartridge'
df_sales2.loc[df_sales2["Type"] == "3rd Party - Cartridge - Hybrid", "Category"] = 'Cartridge'
df_sales2.loc[df_sales2["Type"] == "Promo - Distilled Cartridge", "Category"] = 'Cartridge'
df_sales2.loc[df_sales2["Type"] == "Promo - Live Resin Cartridge", "Category"] = 'Cartridge'
#checkpoint below - Uncomment to data check
#df_sales2.head(40)

# Change Category for any product with a Type that includes "Preroll", "Joint", or "Blunt" from "PackedBud" to "Prerolls"
df_sales2.loc[df_sales2["Type"] == "3rd Party - Joints", "Category"] = 'Prerolls'
df_sales2.loc[df_sales2["Type"] == "blunt", "Category"] = 'Prerolls'
df_sales2.loc[df_sales2["Type"] == "3rd Party - Blunt Pack", "Category"] = 'Prerolls'
df_sales2.loc[df_sales2["Type"] == "3rd Party - Infused Preroll", "Category"] = 'Prerolls'
#checkpoint below - Uncomment to data check
#df_sales2.head(66)

# Change Brand for any product with a Product Name that includes the word "Bloom" to Bloom County
df_sales2.loc[df_sales2["Product Name"] == "Bloom | Joint Hyb - Fried Apples (1g)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Bloom | Joint Hyb - Gary Payton (1g)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Bloom | Joint Hyb - Grape Gasoline (1g)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Bloom | Joint Hyb - LA Kush Cake (1g)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Bloom | Joint Ind - Brain Crasher (1g)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Bloom | Joint Ind - Cake Crasher (1g)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Bloom | Joint Ind - Comatose (1g)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Bloom | Joint Ind - Garlic Crasher (1g)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Bloom | Joint Ind - Kush Crasher (1g)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Bloom | Joint Ind - Modified Grapes (1g)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Bloom | Joint Ind - Punch Mints (1g)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Bloom | Joint Indica - Brain Crasher (1g)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Bloom | Joint Sat - Gelonade (1g)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Bloom | Joint Sat - Golden Goat (1g)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Juicy Fruit - Bloom Co. (3.5g)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Bloom | Joint Sat - Juicy Fruit (1g)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Bloom | Joint Sat - Super Lemon Haze (1g)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Golden Goat - BC (Sativa)", "Brand"] = 'Bloom County'
df_sales2.loc[df_sales2["Product Name"] == "Super Lemon Haze - BC (Sativa)", "Brand"] = 'Bloom County'

#Change Brand for any Brand that includes the word "Bloom Co" to Bloom County
df_sales2.loc[df_sales2["Brand"] == "Bloom Co", "Brand"] = 'Bloom County'
#Change Brand for any Brand that includes the word "BLOOM" to Bloom County
df_sales2.loc[df_sales2["Brand"] == "BLOOM", "Brand"] = 'Bloom County'
#Change Brand for any Brand that includes the word "Bloom" to Bloom County
df_sales2.loc[df_sales2["Brand"] == "Bloom", "Brand"] = 'Bloom County'
#Change Brand for any Brand that includes the word "green dot" to Green Dot
df_sales2.loc[df_sales2["Brand"] == "green dot", "Brand"] = 'Green Dot'
#Change Brand for any Brand that includes the word "Loyalty" to Loyalty Farms
df_sales2.loc[df_sales2["Brand"] == "Loyalty", "Brand"] = 'Loyalty Farms'
#Change Brand for any Brand that includes the word "LOYALTY FARMS" to Loyalty Farms
df_sales2.loc[df_sales2["Brand"] == "LOYALTY FARMS", "Brand"] = 'Loyalty Farms' 
#Change Brand for any Brand that includes the word "loyalty" to Loyalty Farms
df_sales2.loc[df_sales2["Brand"] == "loyalty", "Brand"] = 'Loyalty Farms'  
#Change Brand for any Brand that includes the word "BONSAI" to Bonsai
df_sales2.loc[df_sales2["Brand"] == "BONSAI", "Brand"] = 'Bonsai' 
#Change Brand for any Brand that includes the word "bonsai" to Bonsai
df_sales2.loc[df_sales2["Brand"] == "bonsai", "Brand"] = 'Bonsai'
#Change Brand for any Brand that includes the word "cutting edge" to Cutting Edge
df_sales2.loc[df_sales2["Brand"] == "cutting edge", "Brand"] = 'Cutting Edge'
#Change Brand for any Brand that includes the word "Cutting Edge Cultivation" to Cutting Edge
df_sales2.loc[df_sales2["Brand"] == "Cutting Edge Cultivation", "Brand"] = 'Cutting Edge'
#Change Brand for any Brand that includes the word "DURANGO" to Durango Cannabis
df_sales2.loc[df_sales2["Brand"] == "DURANGO", "Brand"] = 'Durango Cannabis'
#Change Brand for any Brand that includes the word "Durango C C" to Durango Cannabis
df_sales2.loc[df_sales2["Brand"] == "Durango C C", "Brand"] = 'Durango Cannabis'
#Change Brand for any Brand that includes the word "Durango" to Durango Cannabis
df_sales2.loc[df_sales2["Brand"] == "Durango ", "Brand"] = 'Durango Cannabis'
#Change Brand for any Brand that includes the word "MJ Durango" to Durango Cannabis
df_sales2.loc[df_sales2["Brand"] == "MJ Durango", "Brand"] = 'Durango Cannabis'
#Change Brand for any Brand that includes the word "Cutting Edge Cultivation" to Cutting Edge
df_sales2.loc[df_sales2["Brand"] == "Cutting Edge Cultivation", "Brand"] = 'Cutting Edge'
#Change Brand for any Brand that includes the word "Honest" to Honest Marijuana
df_sales2.loc[df_sales2["Brand"] == "Honest", "Brand"] = 'Honest Marijuana'
#Change Brand for any Brand that includes the word "Honest Blunt" to Honest Marijuana
df_sales2.loc[df_sales2["Brand"] == "Honest Blunt", "Brand"] = 'Honest Marijuana'
#Change Brand for any Brand that includes the word "Honest Blunts" to Honest Marijuana Company
df_sales2.loc[df_sales2["Brand"] == "Honest Marijuna Company", "Brand"] = 'Honest Marijuana'
#Change Brand for any Brand that includes the word "HOST" to Host Cannabis
df_sales2.loc[df_sales2["Brand"] == "HOST", "Brand"] = 'Host Cannabis'
#Change Brand for any Brand that includes the word "Host" to Host Cannabis
df_sales2.loc[df_sales2["Brand"] == "Host", "Brand"] = 'Host Cannabis'
#Change Brand for any Brand that includes the word "Host Cannabis Co." to Host Cannabis
df_sales2.loc[df_sales2["Brand"] == "Host Cannabis Co.", "Brand"] = 'Host Cannabis'
#Change Brand for any Brand that includes the word "KIND LOVE" to Kind Love
df_sales2.loc[df_sales2["Brand"] == "KIND LOVE", "Brand"] = 'Kind Love'
#Change Brand for any Brand that includes the word "KindLove" to Kind Love
df_sales2.loc[df_sales2["Brand"] == "KindLove", "Brand"] = 'Kind Love'
#Change Brand for any Brand that includes the word "Lttle Smokey's" to Lil' Smokey's
df_sales2.loc[df_sales2["Brand"] == "Lttle Smokey's", "Brand"] = "Lil' Smokey's"
#Change Brand for any Brand that includes the word "Smokeys" to Lil' Smokey's
df_sales2.loc[df_sales2["Brand"] == "Smokeys", "Brand"] = "Lil' Smokey's"
#Change Brand for any Brand that includes the word "Smokeys" to Lil' Smokey's
df_sales2.loc[df_sales2["Brand"] == "Smokeys", "Brand"] = "Lil' Smokey's"
#Change Brand for any Brand that includes the word "NATTY REMS" to Natty Rems
df_sales2.loc[df_sales2["Brand"] == "NATTY REMS", "Brand"] = 'Natty Rems'
#Change Brand for any Brand that includes the word "SHIFT" to Shift
df_sales2.loc[df_sales2["Brand"] == "SHIFT", "Brand"] = 'Shift'
#Change Brand for any Brand that includes the word "Mary Jane's" to Mary Jane's Medicinals
df_sales2.loc[df_sales2["Brand"] == "Mary Jane's", "Brand"] = "Mary Jane's Medicinals"

# Change Brand for any product with a Product Name that includes the words "Apple Fritter - Clarity (Hybrid)" or "Purple Punch - CK (Indica)" 
#           or "Apple Fritter - BC (Hybrid)" or "Purple Punch - Clarity (Indica)" or "Purple Punch - BC (Indica)" "Gas Cream - CK (Hybrid)" 
#           to Clarity
df_sales2.loc[df_sales2["Product Name"] == "Apple Fritter - Clarity (Hybrid)", "Brand"] = 'Clarity'
df_sales2.loc[df_sales2["Product Name"] == "Purple Punch - CK (Indica)", "Brand"] = 'Clarity'
df_sales2.loc[df_sales2["Product Name"] == "Apple Fritter - BC (Hybrid)", "Brand"] = 'Clarity'
df_sales2.loc[df_sales2["Product Name"] == "Purple Punch - Clarity (Indica)", "Brand"] = 'Clarity'
df_sales2.loc[df_sales2["Product Name"] == "Purple Punch - BC (Indica)", "Brand"] = 'Clarity'
df_sales2.loc[df_sales2["Product Name"] == "Gas Cream - CK (Hybrid)", "Brand"] = 'Clarity'

# Change Brand for any product with a Product Name that includes the words "DCC | Joint Hyb - Cherry Stompers (1g)" or "DCC | Joint Hyb - Cindy 99 x Afghani (1g)" 
#           or "DCC | Joint Ind - Star Roses (1g)" or "DCC | Joint Ind - Wingsuit (1g)" or "DCC | Joint Sat - Cherry Plantains (1g)" or
#           "DCC | Joint Sat - Huckleberry (1g)" or "DCC | Joint Sat - Matanuska Thunder F**k (1g)" or "DCC | Joint Sat - Matanuska Thunder Fuck (1g)" or "DCC | Joint Sat - Reba (1g)"
#           or "Shift | Joint Hyb - Jet Fuel x Gelato (1g)" or "DCC | Joint Sat - Star Crush (1g)" to Durango Cannabis
df_sales2.loc[df_sales2["Product Name"] == "DCC | Joint Hyb - Cherry Stompers (1g)", "Brand"] = 'Durango Cannabis'
df_sales2.loc[df_sales2["Product Name"] == "DCC | Joint Hyb - Cindy 99 x Afghani (1g)", "Brand"] = 'Durango Cannabis'
df_sales2.loc[df_sales2["Product Name"] == "DCC | Joint Ind - Star Roses (1g)", "Brand"] = 'Durango Cannabis'
df_sales2.loc[df_sales2["Product Name"] == "DCC | Joint Ind - Wingsuit (1g)", "Brand"] = 'Durango Cannabis'
df_sales2.loc[df_sales2["Product Name"] == "DCC | Joint Sat - Cherry Plantains (1g)", "Brand"] = 'Durango Cannabis'
df_sales2.loc[df_sales2["Product Name"] == "DCC | Joint Sat - Huckleberry (1g)", "Brand"] = 'Durango Cannabis'
df_sales2.loc[df_sales2["Product Name"] == "DCC | Joint Sat - Matanuska Thunder F**k (1g)", "Brand"] = 'Durango Cannabis'
df_sales2.loc[df_sales2["Product Name"] == "DCC | Joint Sat - Matanuska Thunder Fuck (1g)", "Brand"] = 'Durango Cannabis'
df_sales2.loc[df_sales2["Product Name"] == "DCC | Joint Sat - Reba (1g)", "Brand"] = 'Durango Cannabis'
df_sales2.loc[df_sales2["Product Name"] == "Shift | Joint Hyb - Jet Fuel x Gelato (1g)", "Brand"] = 'Durango Cannabis'
df_sales2.loc[df_sales2["Product Name"] == "DCC | Joint Sat - Star Crush (1g)", "Brand"] = 'Durango Cannabis'

# Change Brand for any product with a Product Name that includes the words "HM | Honest Blunts 2 Pack - Indica (1.6g)" or "Honest | Blunt 2 Pack - Indica" or
#            "Honest | Blunt 2 Pack - Sativa" or "Honest | Blunt 6 Pack -  Sativa" or "HM | Honest Blunts 6 Pack - Indica (4.8g)" or
#              "Honest | Blunt 6 Pack - Indica"  or "HM | Honest Blunts 6 Pack - Indica (4.8g)" or "Honest | Blunt 6 Pack - Hybrid" to Honest Marijuana
df_sales2.loc[df_sales2["Product Name"] == "HM | Honest Blunts 2 Pack - Indica (1.6g)", "Brand"] = 'Honest Marijuana'
df_sales2.loc[df_sales2["Product Name"] == "Honest | Blunt 2 Pack - Indica", "Brand"] = 'Honest Marijuana'
df_sales2.loc[df_sales2["Product Name"] == "Honest | Blunt 2 Pack - Sativa", "Brand"] = 'Honest Marijuana'
df_sales2.loc[df_sales2["Product Name"] == "Honest | Blunt 6 Pack -  Sativa", "Brand"] = 'Honest Marijuana'
df_sales2.loc[df_sales2["Product Name"] == "Honest | Blunt 6 Pack - Indica", "Brand"] = 'Honest Marijuana'
df_sales2.loc[df_sales2["Product Name"] == "Honest | Blunt 6 Pack - Sativa", "Brand"] = 'Honest Marijuana'
df_sales2.loc[df_sales2["Product Name"] == "HM | Honest Blunts 6 Pack - Indica (4.8g)", "Brand"] = 'Honest Marijuana'
df_sales2.loc[df_sales2["Product Name"] == "Honest | Blunt 6 Pack - Hybrid", "Brand"] = 'Honest Marijuana'

# Change Brand for any product with a Product Name that includes the words "Tropicanna Banana - Host (Hybrid)" or "White Runtz - Host (Hybrid)" 
#           or "Cherry Slimeade - Host (Hybrid)" or "Gary P - Host (Hybrid)" or "Gelato Cake - Host (Indica)" or "Gorilla Pops - Host (Hybrid)" "Quattro Kush - Host (Indica)" 
#           or "Root Beer Slushie - Host (Sativa)" or "White Truffle - Host (Indica)" to Host Cannabis Co.
df_sales2.loc[df_sales2["Product Name"] == "Tropicanna Banana - Host (Hybrid)", "Brand"] = 'Host Cannabis Co.'
df_sales2.loc[df_sales2["Product Name"] == "White Runtz - Host (Hybrid)", "Brand"] = 'Host Cannabis Co.'
df_sales2.loc[df_sales2["Product Name"] == "Cherry Slimeade - Host (Hybrid)", "Brand"] = 'Host Cannabis Co.'
df_sales2.loc[df_sales2["Product Name"] == "Deluxe Sugarcane - Host (Indica)", "Brand"] = 'Host Cannabis Co.'
df_sales2.loc[df_sales2["Product Name"] == "Gary P - Host (Hybrid)", "Brand"] = 'Host Cannabis Co.'
df_sales2.loc[df_sales2["Product Name"] == "Gelato Cake - Host (Indica)", "Brand"] = 'Host Cannabis Co.'
df_sales2.loc[df_sales2["Product Name"] == "Gorilla Pops - Host (Hybrid)", "Brand"] = 'Host Cannabis Co.'
df_sales2.loc[df_sales2["Product Name"] == "Quattro Kush - Host (Indica)", "Brand"] = 'Host Cannabis Co.'
df_sales2.loc[df_sales2["Product Name"] == "Root Beer Slushie - Host (Sativa)", "Brand"] = 'Host Cannabis Co.'
df_sales2.loc[df_sales2["Product Name"] == "White Truffle - Host (Indica)", "Brand"] = 'Host Cannabis Co.'

# Change Brand for any product with a Product Name that includes the words "Cherry Slimeade - LF (Hybrid)" or "Strawberry Fritter - LF (Hybrid)"  
#           to Loyalty Farms
df_sales2.loc[df_sales2["Product Name"] == "Cherry Slimeade - LF (Hybrid)", "Brand"] = 'Loyalty Farms'
df_sales2.loc[df_sales2["Product Name"] == "Strawberry Fritter - LF (Hybrid)", "Brand"] = 'Loyalty Farms'

# Change Brand for any product with a Product Name that includes the words "3rd Party - Bath (CBD)" or "3rd Party - Topical (CBD)"  
#           or "3rd Party - Transdermal" or "3rd Party - Transdermal (CBD)"  to Mary Jane's Medicinals
df_sales2.loc[df_sales2["Product Name"] == "3rd Party - Bath (CBD)", "Brand"] = "Mary Jane's Medicinals"
df_sales2.loc[df_sales2["Product Name"] == "3rd Party - Topical (CBD)", "Brand"] = "Mary Jane's Medicinals"
df_sales2.loc[df_sales2["Product Name"] == "3rd Party - Transdermal", "Brand"] = "Mary Jane's Medicinals"
df_sales2.loc[df_sales2["Product Name"] == "3rd Party - Transdermal (CBD)", "Brand"] = "Mary Jane's Medicinals"

# Change Brand for any product with a Product Name that includes the words "Alien Haze - NR (Sativa)" or "Cherry Kush Mints - NR (Indica)"  
#           or "Golden Goat - BC (Sativa)" or "Golden Goat - NR (Sativa)" to Natty Rems
df_sales2.loc[df_sales2["Product Name"] == "Alien Haze - NR (Sativa)", "Brand"] = "Natty Rems"
df_sales2.loc[df_sales2["Product Name"] == "Cherry Kush Mints - NR (Indica)", "Brand"] = "Natty Rems"
df_sales2.loc[df_sales2["Product Name"] == "Golden Goat - BC (Sativa)", "Brand"] = "Natty Rems"
df_sales2.loc[df_sales2["Product Name"] == "Golden Goat - NR (Sativa)", "Brand"] = "Natty Rems"

# Change Brand for any product with a Product Name that includes the words "Bubble Bath - RG (Indica)" or "Cake Mix - RG (Indica)"  
#           or "Gary x Jealousy - RG (Hybrid)" or "Grapple Pie - RG (Hybrid)" or "Oreoz - RG (Hybrid)" or "Pink Grapes - RG (Sativa)" 
#           or "Punch Breath - RG (Hybrid)" or "V-Lone Runtz - RG (Indica)" to Red Garden
df_sales2.loc[df_sales2["Product Name"] == "Bubble Bath - RG (Indica)", "Brand"] = "Red Garden"
df_sales2.loc[df_sales2["Product Name"] == "Cake Mix - RG (Indica)", "Brand"] = "Red Garden"
df_sales2.loc[df_sales2["Product Name"] == "Gary x Jealousy - RG (Hybrid)", "Brand"] = "Red Garden"
df_sales2.loc[df_sales2["Product Name"] == "Grapple Pie - RG (Hybrid)", "Brand"] = "Red Garden"
df_sales2.loc[df_sales2["Product Name"] == "Oreoz - RG (Hybrid)", "Brand"] = "Red Garden"
df_sales2.loc[df_sales2["Product Name"] == "Pink Grapes - RG (Sativa)", "Brand"] = "Red Garden"
df_sales2.loc[df_sales2["Product Name"] == "Punch Breath - RG (Hybrid)", "Brand"] = "Red Garden"
df_sales2.loc[df_sales2["Product Name"] == "V-Lone Runtz - RG (Indica)", "Brand"] = "Red Garden"

# Change Type to PPEOJ if Category=PackedBud and Type=Top Shelf
df_sales2.loc[(df_sales2["Category"] == "PackedBud") & (df_sales2["Type"] == "Top Shelf"), "Type"] = 'PPEOJ'
# Change Category to NonEdible if Category=Topical
df_sales2.loc[(df_sales2["Category"] == "Topical"), "Category"] = 'NonEdible'

# RECONCIATION OF BRAND NAMES
#Change Brand for any Brand that includes the word "Bloom Co" to Bloom County
df_sales2.loc[df_sales2["Brand"] == "Bloom Co", "Brand"] = 'Bloom County'
#Change Brand for any Brand that includes the word "BLOOM" to Bloom County
df_sales2.loc[df_sales2["Brand"] == "BLOOM", "Brand"] = 'Bloom County'
#Change Brand for any Brand that includes the word "Bloom" to Bloom County
df_sales2.loc[df_sales2["Brand"] == "Bloom", "Brand"] = 'Bloom County'
#Change Brand for any Brand that includes the word "green dot" to Green Dot
df_sales2.loc[df_sales2["Brand"] == "green dot", "Brand"] = 'Green Dot'
#Change Brand for any Brand that includes the word "Loyalty" to Loyalty Farms
df_sales2.loc[df_sales2["Brand"] == "Loyalty", "Brand"] = 'Loyalty Farms'
#Change Brand for any Brand that includes the word "LOYALTY FARMS" to Loyalty Farms
df_sales2.loc[df_sales2["Brand"] == "LOYALTY FARMS", "Brand"] = 'Loyalty Farms' 
#Change Brand for any Brand that includes the word "loyalty" to Loyalty Farms
df_sales2.loc[df_sales2["Brand"] == "loyalty", "Brand"] = 'Loyalty Farms'  
#Change Brand for any Brand that includes the word "BONSAI" to Bonsai
df_sales2.loc[df_sales2["Brand"] == "BONSAI", "Brand"] = 'Bonsai' 
#Change Brand for any Brand that includes the word "bonsai" to Bonsai
df_sales2.loc[df_sales2["Brand"] == "bonsai", "Brand"] = 'Bonsai'
#Change Brand for any Brand that includes the word "cutting edge" to Cutting Edge
df_sales2.loc[df_sales2["Brand"] == "cutting edge", "Brand"] = 'Cutting Edge'
#Change Brand for any Brand that includes the word "Cutting Edge Cultivation" to Cutting Edge
df_sales2.loc[df_sales2["Brand"] == "Cutting Edge Cultivation", "Brand"] = 'Cutting Edge'
#Change Brand for any Brand that includes the word "DURANGO" to Durango Cannabis
df_sales2.loc[df_sales2["Brand"] == "DURANGO", "Brand"] = 'Durango Cannabis'
#Change Brand for any Brand that includes the word "Durango C C" to Durango Cannabis
df_sales2.loc[df_sales2["Brand"] == "Durango C C", "Brand"] = 'Durango Cannabis'
#Change Brand for any Brand that includes the word "Durango" to Durango Cannabis
df_sales2.loc[df_sales2["Brand"] == "Durango", "Brand"] = 'Durango Cannabis'
#Change Brand for any Brand that includes the word "MJ Durango" to Durango Cannabis
df_sales2.loc[df_sales2["Brand"] == "MJ Durango", "Brand"] = 'Durango Cannabis'
#Change Brand for any Brand that includes the word "Cutting Edge Cultivation" to Cutting Edge
df_sales2.loc[df_sales2["Brand"] == "Cutting Edge Cultivation", "Brand"] = 'Cutting Edge'
#Change Brand for any Brand that includes the word "Honest" to Honest Marijuana
df_sales2.loc[df_sales2["Brand"] == "Honest", "Brand"] = 'Honest Marijuana'
#Change Brand for any Brand that includes the word "Honest Blunt" to Honest Marijuana
df_sales2.loc[df_sales2["Brand"] == "Honest Blunt", "Brand"] = 'Honest Marijuana'
#Change Brand for any Brand that includes the word "Honest Blunts" to Honest Marijuana Company
df_sales2.loc[df_sales2["Brand"] == "Honest Marijuna Company", "Brand"] = 'Honest Marijuana'
#Change Brand for any Brand that includes the word "HOST" to Host Cannabis
df_sales2.loc[df_sales2["Brand"] == "HOST", "Brand"] = 'Host Cannabis'
#Change Brand for any Brand that includes the word "Host" to Host Cannabis
df_sales2.loc[df_sales2["Brand"] == "Host", "Brand"] = 'Host Cannabis'
#Change Brand for any Brand that includes the word "Host Cannabis Co." to Host Cannabis
df_sales2.loc[df_sales2["Brand"] == "Host Cannabis Co.", "Brand"] = 'Host Cannabis'
#Change Brand for any Brand that includes the word "KIND LOVE" to Kind Love
df_sales2.loc[df_sales2["Brand"] == "KIND LOVE", "Brand"] = 'Kind Love'
#Change Brand for any Brand that includes the word "KindLove" to Kind Love
df_sales2.loc[df_sales2["Brand"] == "KindLove", "Brand"] = 'Kind Love'
#Change Brand for any Brand that includes the word "Lttle Smokey's" to Lil' Smokey's
df_sales2.loc[df_sales2["Brand"] == "Lttle Smokey's", "Brand"] = "Lil' Smokey's"
#Change Brand for any Brand that includes the word "Smokeys" to Lil' Smokey's
df_sales2.loc[df_sales2["Brand"] == "Smokeys", "Brand"] = "Lil' Smokey's"
#Change Brand for any Brand that includes the word "Smokeys" to Lil' Smokey's
df_sales2.loc[df_sales2["Brand"] == "Smokey's", "Brand"] = "Lil' Smokey's"
#Change Brand for any Brand that includes the word "Smokeys" to Lil' Smokey's
df_sales2.loc[df_sales2["Brand"] == "Lil' Smonkey's", "Brand"] = "Lil' Smokey's"
#Change Brand for any Brand that includes the word "NATTY REMS" to Natty Rems
df_sales2.loc[df_sales2["Brand"] == "NATTY REMS", "Brand"] = 'Natty Rems'
#Change Brand for any Brand that includes the word "SHIFT" to Shift
df_sales2.loc[df_sales2["Brand"] == "SHIFT", "Brand"] = 'Shift'

#remove rows with "Totals" and "Weight Sold" in Location column to clean up dataframe
#main dataframe file is now df_sales4
df_sales3 = df_sales2.drop(df_sales2[df_sales2.Location == 'Totals:'].index)
df_sales4 = df_sales3.drop(df_sales3[df_sales3.Location == 'Weight Sold'].index)

#Drop last rows with irrelevant data
df_sales4 = df_sales4.drop(df_sales4[df_sales4.Location == 'Totals:'].index)
df_sales4 = df_sales4.drop(df_sales4[df_sales4.Location == 'Weight Sold'].index)
df_sales4 = df_sales4.drop(df_sales4[df_sales4.Quantity == 0].index )

#Change date Transaction Date format from DD/MM/YY to DD/MM
df_sales4.style.format({'Transaction Date':'{:%m}'})
#print(df_sales4['Transaction Date'])

# Change Tranaction Time to 24 hour format
df_sales4['Transaction Time'] = pd.to_datetime(df_sales4['Transaction Time'], format='%I:%M:%S %p').dt.strftime('%H:%M:%S')
df_sales4['Transaction Time'] = pd.to_datetime(df_sales4['Transaction Time'], format='%H:%M:%S').dt.strftime('%H:%M')
# print(df_sales4.head(5))

GroomedData = df_sales4
filename = 'Groomed Data Week of '
FilenameFormatter = ' - '
GroomedData.to_excel(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.xlsx')
#END OF GROOMING DATA

print(GroomedData.head(10))