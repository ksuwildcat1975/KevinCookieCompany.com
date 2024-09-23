import matplotlib as mpl
import matplotlib.patches as patches
from matplotlib.image import BboxImage
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
import dataframe_image as dfi
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
df_sales ['Sales_Minus_Tax'] = df_sales['Final Sale'] - df_sales['Tax In Dollars']

#add new column with Net Sale after taking out product cost
df_sales ['Net Sale'] = df_sales ['Sales_Minus_Tax'] - df_sales['Cost']

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

GroomedData = df_sales4
filename = 'Groomed Data Week of '
FilenameFormatter = ' - '
GroomedData.to_excel(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.xlsx')
#END OF GROOMING DATA


#START OF SALES REPORT
#Block plt.show() from displaying in console unless debugging
plt.show(block=False)
# Create a series for unique transactions for each location
NumberOfTransactions = df_sales4.groupby('Location')['Sale Id'].nunique()


#Count number of unique transactions and generate average per transaction at each store
#First create a series "AverageTransaction"
AverageTransaction = (df_sales4.groupby('Location')['Final Sale'].sum()) / (df_sales4.groupby('Location')['Sale Id'].nunique())

#Convert the series into a dataframe and round to two decimal points
df_sales6 = pd.DataFrame(AverageTransaction).round(2)
#Add the name to the column
df_sales6.columns = ['Avg. Transaction']

#Calculate WoW for Avg. 
#Find the current path name for the preivous week's sales data
FilenameFormatter = ' - '
PreviousWeekDataPath = 'C:\\Users\\satkinson\\Documents\\Igadi\\Sales Reports ' + CurrentYear + '\\Week of ' + PreviousWeekStartDate + ' - ' + PreviousWeekEndDate
os.chdir(PreviousWeekDataPath)
filename = ('Average Ticket and # of Transactions Week of ' + PreviousWeekStartDate + FilenameFormatter + PreviousWeekEndDate + '.xlsx')
PreviousWeekAvgTicket = (os.path.abspath(filename))
#print(PreviousWeekAvgTicket)
#import Previous Sale Report Excel file and normalize
PreviousWeekAvgReport = pd.read_excel(PreviousWeekAvgTicket)
PreviousWeekAvgReport = PreviousWeekAvgReport.set_index('Location')
#Add the WoW percentages for Average Transactions
df_sales6['Previous Week Avg. Transaction'] = PreviousWeekAvgReport['Avg. Transaction']
df_sales6['WoW Sales Delta'] = df_sales6['Avg. Transaction'] - df_sales6['Previous Week Avg. Transaction']
df_sales6.loc[:, 'Avg. WoW'] = (df_sales6['WoW Sales Delta'] / df_sales6['Avg. Transaction']).round(2)

#Calculate YoY for Avg. 
#Find the current path name for the preivous year's sales data
PreviousYearDataPath = 'C:\\Users\\satkinson\\Documents\\Igadi\\Sales Reports ' + PreviousYear + '\\Week of ' + PreviousYearStartDate + ' - ' + PreviousYearEndDate
os.chdir(PreviousYearDataPath)
filename = ('Average Ticket and # of Transactions Week of ' + PreviousYearStartDate + FilenameFormatter + PreviousYearEndDate + '.xlsx')
PreviousYearAvgTicket = (os.path.abspath(filename))
#import Previous Year Sales Report Excel file and normalize indecies
PreviousYearAvgReport = pd.read_excel(PreviousYearAvgTicket)
PreviousYearAvgReport = PreviousYearAvgReport.set_index('Location')

#Add the YoY percentages for Average Transactions
df_sales6['Previous Year Avg. Transaction'] = PreviousYearAvgReport['Avg. Transaction']
df_sales6['YoY Sales Delta'] = df_sales6['Avg. Transaction'] - df_sales6['Previous Year Avg. Transaction']
df_sales6.loc[:, 'Avg. YoY'] = (df_sales6['YoY Sales Delta'] / df_sales6['Avg. Transaction']).round(2)
#Account for Locations added in 2023 (i.e. Lyons) by replacing NaN's
df_sales6["Previous Year Avg. Transaction"].fillna(0, inplace=True)
df_sales6["YoY Sales Delta"].fillna(0, inplace=True)
df_sales6["Avg. YoY"].fillna(0, inplace=True)
#df_sales6.head(10)


#Add the current # of Transcations and WoW percentages
df_sales6['# Transactions'] = NumberOfTransactions
df_sales6['Previous Week Transactions'] = PreviousWeekAvgReport['# Transactions']
df_sales6['Transaction # Delta'] = df_sales6['# Transactions'] - df_sales6['Previous Week Transactions']
df_sales6.loc[:, 'Trans. WoW'] = (df_sales6['Transaction # Delta'] / df_sales6['Previous Week Transactions']).round(2)

#Add the YoY percentages for Number of Transactions
df_sales6['Previous Year Transactions'] = PreviousYearAvgReport['# Transactions']
df_sales6['YoY Transaction Delta'] = df_sales6['# Transactions'] - df_sales6['Previous Year Transactions']
df_sales6.loc[:, 'Transactions YoY'] = (df_sales6['YoY Transaction Delta'] / df_sales6['# Transactions']).round(2)
#Account for Locations added in 2023 (i.e. Lyons) by replacing NaN's with zeros
df_sales6["Previous Year Transactions"].fillna(0, inplace=True)
df_sales6["YoY Transaction Delta"].fillna(0, inplace=True)
df_sales6["Transactions YoY"].fillna(0, inplace=True)

# Calculate total average ticket and total number of transactions for all locations
TotalAvgTicket = df_sales6['Avg. Transaction'].mean()
TotalAvgTicket = round(TotalAvgTicket, 2)
PreviousWeekTotalAvgTicket = df_sales6['Previous Week Avg. Transaction'].mean()
PreviousWeekTotalAvgTicket = round(PreviousWeekTotalAvgTicket, 2)
AvgTotalTicketDelta = TotalAvgTicket - PreviousWeekTotalAvgTicket
AvgTotalTicketDelta = round(AvgTotalTicketDelta, 2)
AvgTotalTicketWoW = AvgTotalTicketDelta / PreviousWeekTotalAvgTicket
AvgTotalTicketWoW = round(AvgTotalTicketWoW, 2)
PreviousYearTotalAvgTicket = df_sales6['Previous Year Avg. Transaction'].mean()
PreviousYearTotalAvgTicket = round(PreviousYearTotalAvgTicket, 2)
PreviousYearTotalAvgTicketDelta = TotalAvgTicket - PreviousYearTotalAvgTicket
PreviousYearTotalAvgTicketDelta = round(PreviousYearTotalAvgTicketDelta, 2)
AvgTotalTicketYoY = PreviousYearTotalAvgTicketDelta / TotalAvgTicket
AvgTotalTicketYoY = round(AvgTotalTicketYoY, 2)


TotalTransactions = df_sales6['# Transactions'].sum()
TotalTransactions.round()
PreviousWeekTotalTransactions = df_sales6['Previous Week Transactions'].sum()
PreviousWeekTotalTransactions.round(0)
TotalTransactionDelta = TotalTransactions - PreviousWeekTotalTransactions
TranactionsWoW = TotalTransactionDelta / PreviousWeekTotalTransactions
TranactionsWoW = TranactionsWoW.round(2)
PreviousYearTotalTransactions = df_sales6['Previous Year Transactions'].sum()
PreviousYearTotalTransactions.round(0)
TotalTransactionYoYDelta = TotalTransactions - PreviousYearTotalTransactions
TranactionsYoY = TotalTransactionYoYDelta / PreviousYearTotalTransactions
TranactionsYoY = TranactionsYoY.round(2)

#Add a row for the total average ticket and total number of transactions for all locations
df_sales6.loc['Grand Total'] = [TotalAvgTicket, PreviousWeekTotalAvgTicket, AvgTotalTicketDelta, AvgTotalTicketWoW, PreviousYearTotalAvgTicket, PreviousYearTotalAvgTicketDelta, 
                          AvgTotalTicketYoY, TotalTransactions, PreviousWeekTotalTransactions, TotalTransactionDelta, TranactionsWoW, PreviousYearTotalTransactions, TotalTransactionYoYDelta, TranactionsYoY]

AvgAndNumberTransactions = df_sales6
filename = 'Average Ticket and # of Transactions Week of '
#Format and output Excel file for future use
AvgAndNumberTransactions.to_excel(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.xlsx')

#FORMAT AND OUTPUT TABLE TO IMAGE
df_sales6 = df_sales6.reset_index(drop=False)
#print(df_sales6)    
# df_sales6.head(10)
#Simple .table method to export dataframe to image
#define figure and axes
fig, ax = plt.subplots()
# hide the axes
ax.patch.set_visible(True)
fig.patch.set_visible(True)
ax.axis('off')
ax.axis('tight')
ColumnsTest = ['Location','Avg. Ticket','Avg. Ticket WoW','Avg. Ticket YoY', '# Transactions', 'Transactions WoW', 'Transactions YoY' ]
# RowsTest = ['Granby', 'Lafayette', 'Lyons', 'Northglenn']
TableTest = df_sales6[['Location', 'Avg. Transaction','Avg. WoW','Avg. YoY', '# Transactions', 'Trans. WoW', 'Transactions YoY' ]]

TableTest = TableTest.reset_index(drop=True)
TableTest["Avg. Transaction"] = TableTest['Avg. Transaction'].map('${:,.2f}'.format)
TableTest["Avg. WoW"] = TableTest['Avg. WoW'].map('{:.0%}'.format)
TableTest["Avg. YoY"] = TableTest['Avg. YoY'].map('{:.0%}'.format)
TableTest["Trans. WoW"] = TableTest['Trans. WoW'].map('{:.0%}'.format)
TableTest["Transactions YoY"] = TableTest['Transactions YoY'].map('{:.0%}'.format)
#print(TableTest)    
ccolors = plt.cm.Greens(np.full(len(ColumnsTest), 0.4))

#create table
table = ax.table(cellText=TableTest.values, colLabels=ColumnsTest, loc='center', cellLoc='center', colColours=ccolors, edges='closed')
#table.set_fontsize(20)
table.scale(1, 3)

table.auto_set_column_width(col=list(range(len(TableTest.columns))))
table.auto_set_font_size(True)
#Alternate the row colors
# for i in range(0, len(TableTest), 2):
#     for j in range(len(ColumnsTest)):
#         table[(i,j)].set_facecolor("lightgrey")
for i in range(0, len(TableTest), 2):
    for j in range(len(ColumnsTest)):
        table[(i+1,j)].set_facecolor("royalblue")
        table[(i+1,j)].set_text_props(weight='bold', color='w')

for i in range(0, len(TableTest), 2):
    for j in range(len(ColumnsTest)):
        table[(i,j)].set_facecolor("cornflowerblue")
        table[(i,j)].set_text_props(weight='bold', color='w')

#Reset column labels row background color and font color
for j in range(len(ColumnsTest)):
    table[(0,j)].set_facecolor("midnightblue")
    table[(0,j)].set_text_props(weight='bold', color='w')




#table.set_linewidth(0.5)  # set cell border width
bbox=[0, 0, 1, 1]

#display table
fig.tight_layout()

#save table
filename = 'Average Ticket and # of Transactions Week of '
# fig.savefig('Igadi Average Ticket Number of Transactions.png', bbox_inches='tight', dpi=300)
fig.savefig(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.png', bbox_inches='tight', dpi=300)


# #Create a dataframe for the top 10 products by sales
Top5BrandsByCategory = df_sales4.groupby('Category')['Final Sale'].sum().nlargest(5)
#Convert the series into a dataframe and round to two decimal points
df_sales7 = pd.DataFrame(Top5BrandsByCategory).round(2)
#Add the name to the column
df_sales7.columns = ['Total Sales']
#Export the dataframe as Excel to preserve for further Excel operations as needed or re-import
filename = 'Top Brands by Category Week of '
df_sales7.to_excel(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.xlsx')


## Pivot table functions to sum Sales by Category and Type by Store
# Start with generating base pivot table for first Sales by Location report
df_sales5 = df_sales4.pivot_table(index=["Location"], columns=["Transaction Date"], values=["Final Sale"], aggfunc="sum")
#print(df_sales5)
# Calculate the daily sales totals by store location and date. 'Location' is the index
sum = df_sales5.sum()
#print(sum)
#Give the new row a name
sum.name = 'Daily Totals'
#Transpose the rows and columns (Daily Totals needs to be a row rather than a column generated by the sum function) and append to dataframe.  NOTE: append will be deprecated
# in a future version and this code will need to be replaced by concat.  Using append here is the right tool for now.
Daily_Sales_by_Location = df_sales5.append(sum.transpose())
#Add a column "Grand Total" which shows weekly total sales by store location calculated by row.  Using axis 1 sums across rows rather than down columns.
Daily_Sales_by_Location.loc[:,"Grand Total"] = Daily_Sales_by_Location.sum(axis=1)
#print(Daily_Sales_by_Location)
# Set the filepath and filename to print file with date of the report added to the filename
filename = 'Sales Report By Location Week of '
Daily_Sales_by_Location.to_excel(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.xlsx')

#CREATING FORMATTED TABLE FOR SALES BY LOCATION
#Format the table for Sales by Location
#Convert pivot table to dataframe
#Daily_Sales_by_Location = pd.DataFrame(Daily_Sales_by_Location)
#select only the values for the columns we want to display to create a new dataframe
Daily_Sales_by_Location = Daily_Sales_by_Location.iloc[:,[0,1,2,3,4,5,6,7]]
#Reset the index to make Location a column rather than the index
Daily_Sales_by_Location = Daily_Sales_by_Location.reset_index(drop=False)
#Make Transaction Date the column headers
#LATER WORK  TRY TO USE ACTUATL DATE AS VARIBLE TO INSERT COLUMN NAME RATHER THAN DAY OF WEEK i.e. Date1 = Daily_Sales_by_Location.columns[1]

Daily_Sales_by_Location.columns = ['Location', "Monday",'Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday','Grand Total']
DailySalesByLocationTable = Daily_Sales_by_Location

#Format float numbers for sales to currency
DailySalesByLocationTable["Grand Total"] = DailySalesByLocationTable['Grand Total'].map('${:,.0f}'.format)
DailySalesByLocationTable["Monday"] = DailySalesByLocationTable["Monday"].map('${:,.0f}'.format)
DailySalesByLocationTable["Tuesday"] = DailySalesByLocationTable['Tuesday'].map('${:,.0f}'.format)
DailySalesByLocationTable["Wednesday"] = DailySalesByLocationTable['Wednesday'].map('${:,.0f}'.format)
DailySalesByLocationTable["Thursday"] = DailySalesByLocationTable['Thursday'].map('${:,.0f}'.format)
DailySalesByLocationTable["Friday"] = DailySalesByLocationTable['Friday'].map('${:,.0f}'.format)
DailySalesByLocationTable["Saturday"] = DailySalesByLocationTable['Saturday'].map('${:,.0f}'.format)
DailySalesByLocationTable["Sunday"] = DailySalesByLocationTable['Sunday'].map('${:,.0f}'.format)

#Simple .table method to export dataframe to image
#define figure and axes
fig, ax = plt.subplots()
# hide the axes
ax.patch.set_visible(True)
fig.patch.set_visible(True)
ax.axis('off')
ax.axis('tight')
plt.ioff()
Date1 = Daily_Sales_by_Location.columns[1]
Columns2 = ['Location', Date1,'Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday','Grand Total']
Rows2 = ['Central City', 'Granby', 'Idaho Springs', 'Lafayette', 'Louisville', 'Lyons', 'Nederland', 'Northglenn', 'Tabernash', 'Grand Total']  
ccolors = plt.cm.Greens(np.full(len(Columns2), 0.4))#create table
table = ax.table(cellText=DailySalesByLocationTable.values, colLabels=Columns2, loc='center', cellLoc='center', colColours=ccolors, edges='closed')
#table.set_fontsize(20)
#Scale the table to show all columns
table.scale(1, 3)


table.auto_set_column_width(col=list(range(len(DailySalesByLocationTable.columns))))
#Auto set row height
# table.auto_set_row_height(row=list(range(len(DailySalesByLocationTable.rows))))
table.auto_set_font_size(True)
#Alternate the row colors
# for i in range(0, len(DailySalesByLocationTable), 2):
#     for j in range(len(Columns2)):
#         table[(i+1,j)].set_facecolor("lightgrey")

for i in range(0, len(DailySalesByLocationTable), 2):
    for j in range(len(Columns2)):
        table[(i+1,j)].set_facecolor("royalblue")
        table[(i+1,j)].set_text_props(weight='bold', color='w')

for i in range(0, len(TableTest), 2):
    for j in range(len(Columns2)):
        table[(i,j)].set_facecolor("cornflowerblue")
        table[(i,j)].set_text_props(weight='bold', color='w')

#Reset column labels row background color and font color
for j in range(len(Columns2)):
    table[(0,j)].set_facecolor("midnightblue")
    table[(0,j)].set_text_props(weight='bold', color='w')


#table.set_linewidth(0.5)  # set cell border width
bbox=[0, 0, 1, 1]
#bbox=[0, 0, 10, 5]

#display table
fig.tight_layout()

#save table
filename = 'Sales Report By Location Week of '
fig.savefig(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.png', bbox_inches='tight', dpi=300)

# #Create a line graph for Sales by Location by Day of Week by Location LATER WORK ON THIS
# plt.plot(Daily_Sales_by_Location['Central City'], label='Central City')
# plt.plot(Daily_Sales_by_Location['Granby'], label='Granby')
# # plt.plot(Daily_Sales_by_Location['Wednesday'], label='Wednesday')

# # plt.plot(Daily_Sales_by_Location['Thursday'], label='Thursday')
# # plt.plot(Daily_Sales_by_Location['Friday'], label='Friday')
# # plt.plot(Daily_Sales_by_Location['Saturday'], label='Saturday')
# # plt.plot(Daily_Sales_by_Location['Sunday'], label='Sunday')
# # plt.plot(Daily_Sales_by_Location['Grand Total'], label='Grand Total')
# plt.legend()
# plt.title('Sales by Location by Day of Week')
# plt.xlabel('Location')
# plt.ylabel('Sales')
# plt.xticks(rotation='vertical')
# plt.tight_layout()
# filename = 'Sales by Location by Day of Week Week of '
# plt.savefig(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.png', dpi=300)
# plt.show()
# #END OF SALES BY LOCATION REPORT


#PERCETAGE OF SALES BY CATEGORY TABLE 1
#Calculate the total sales for the week
TotalSales = df_sales4['Final Sale'].sum()
#Calculate the per cent of sales by category
CalculatePerCentByCategory = (df_sales4.groupby('Category')['Final Sale'].sum())/TotalSales
#Convert the series into a dataframe and round to two decimal points
df_sales8 = pd.DataFrame(CalculatePerCentByCategory).round(2)
#Reset the index to make Category a column rather than the index
df_sales8 = df_sales8.reset_index(drop=False)
#Add the name to the first two columns
df_sales8.columns = ['Category', '% Of Sales']

# #Retrieve the previous week's sales data
PreviousWeekDataPath = 'C:\\Users\\satkinson\\Documents\\Igadi\\Sales Reports ' + CurrentYear + '\\Week of ' + PreviousWeekStartDate + ' - ' + PreviousWeekEndDate
os.chdir(PreviousWeekDataPath)
filename = ('Sales Percent by Category Week of ' + PreviousWeekStartDate + FilenameFormatter + PreviousWeekEndDate + '.xlsx')
#import Previous Sale Report Excel file and normalize
PreviousWeekSalesPercent = pd.read_excel(filename)
#Add the WoW percentages from the previous week by adding a column to the dataframe
df_sales8['Previous Week'] = PreviousWeekSalesPercent['% Of Sales'] 

#Retrieve the previous year's sales data
PreviousYearDataPath = 'C:\\Users\\satkinson\\Documents\\Igadi\\Sales Reports ' + PreviousYear + '\\Week of ' + PreviousYearStartDate + ' - ' + PreviousYearEndDate
os.chdir(PreviousYearDataPath)
filename = ('Sales Percent by Category Week of ' + PreviousYearStartDate + FilenameFormatter + PreviousYearEndDate + '.xlsx')
# import Previous Year Sales Report Excel file and add column for Previous Year Percent of Sales
PreviousYearSalesPercent = pd.read_excel(filename)
df_sales8['Previous Year'] = PreviousYearSalesPercent['% Of Sales']
#print(df_sales8)
PerCentOfSalesByCategory = df_sales8
#Export the dataframe as Excel to preserve for further Excel operations as needed or re-import
filename = 'Sales Percent by Category Week of '
PerCentOfSalesByCategory.to_excel(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.xlsx')

#FORMAT AND OUTPUT CATEGORY % OF SALES TABLE TO IMAGE
#Simple .table method to export dataframe to image
PerCentOfSalesByCategoryTable = PerCentOfSalesByCategory
#Format float numbers for sales to percentage
PerCentOfSalesByCategoryTable["% Of Sales"] = PerCentOfSalesByCategoryTable["% Of Sales"].map('{:,.0%}'.format)
PerCentOfSalesByCategoryTable["Previous Week"] = PerCentOfSalesByCategoryTable["Previous Week"].map('{:,.0%}'.format)
PerCentOfSalesByCategoryTable["Previous Year"] = PerCentOfSalesByCategoryTable["Previous Year"].map('{:,.0%}'.format)

#define figure and axes
fig, ax = plt.subplots()
# hide the axes
ax.patch.set_visible(True)
fig.patch.set_visible(True)
ax.axis('off')
ax.axis('tight')
Columns3 = ['Category','% of Sales','Previous Week', 'Previous Year']
#Rows2 = ['Central City', 'Granby', 'Idaho Springs', 'Lafayette', 'Louisville', 'Lyons', 'Nederland', 'Northglenn', 'Tabernash', 'Grand Total']  
ccolors = plt.cm.Greens(np.full(len(Columns3), 0.4))
#create table
table = ax.table(cellText=PerCentOfSalesByCategoryTable.values, colLabels=Columns3, loc='center', cellLoc='center', colColours=ccolors, edges='closed')
table.scale(1, 3)

table.auto_set_column_width(col=list(range(len(PerCentOfSalesByCategoryTable.columns))))
#Auto set row height
# table.auto_set_row_height(row=list(range(len(DailySalesByLocationTable.rows))))
table.auto_set_font_size(True)
#Alternate the row colors
# for i in range(0, len(PerCentOfSalesByCategoryTable), 2):
#     for j in range(len(Columns3)):
#         table[(i+1,j)].set_facecolor("lightgrey")

for i in range(0, len(PerCentOfSalesByCategoryTable), 2):
    for j in range(len(Columns3)):
        table[(i+1,j)].set_facecolor("mediumseagreen")
        table[(i+1,j)].set_text_props(weight='bold', color='w')

for i in range(0, len(PerCentOfSalesByCategoryTable), 2):
    for j in range(len(Columns3)):
        table[(i,j)].set_facecolor("limegreen")
        table[(i,j)].set_text_props(weight='bold', color='w')

#Reset column labels row background color and font color
for j in range(len(Columns3)):
    table[(0,j)].set_facecolor("darkgreen")
    table[(0,j)].set_text_props(weight='bold', color='w')
#table.set_linewidth(0.5)  # set cell border width
#bbox=[0, 0, 1, 1]
bbox=[0, 0, 10, 5]

#display table
fig.tight_layout()

#save table
filename = 'Category % of Sales Week of '
# fig.savefig('Igadi Average Ticket Number of Transactions.png', bbox_inches='tight', dpi=300)
fig.savefig(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.png', dpi=800)


#BAR GRAPH FOR SALES BY CATEGORY VERSUS LOCATION TABLE 2
#Left below in commented code for now as reference for use of margins in pivot_table function
#Table2_test = df_sales4.pivot_table(index=["Location"], columns=["Category"], values=["Final Sale"], aggfunc='sum', margins=True, margins_name='Total')

df_sales9 = df_sales4.pivot_table(index="Location", columns="Category", values="Final Sale", aggfunc='sum')
df_sales10 = df_sales9.transpose()
# filename = 'Sales by Category vs Location Pivot Table Raw Week of '
# df_sales10.to_excel(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.xlsx')
CalculatePerCentByLocationandCategory = (df_sales10/ df_sales10.sum() * 100).round(2)
filename = 'Sales by Category vs Location Calculations Week '
CalculatePerCentByLocationandCategory.to_excel(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.xlsx')
CalculatePerCentByLocationandCategory2 = CalculatePerCentByLocationandCategory.transpose()

#Make the bar graph for Table 2
ax = CalculatePerCentByLocationandCategory2.plot(kind='bar', stacked=True, title = 'Sales by Category vs Location', figsize=(15,10))
#Add the values to each section of the bars in the bar chart.  Suppress labels for any value less than "threshold" for readability
threshold=2
ax.bar_label(ax.containers[0], fmt=lambda x: f'{x:0.0f}%' if x > threshold else '', label_type='center')
ax.bar_label(ax.containers[1], fmt=lambda x: f'{x:0.0f}%' if x > threshold else '', label_type='center')
ax.bar_label(ax.containers[2], fmt=lambda x: f'{x:0.0f}%' if x > threshold else '', label_type='center')
ax.bar_label(ax.containers[3], fmt=lambda x: f'{x:0.0f}%' if x > threshold else '', label_type='center')
ax.bar_label(ax.containers[4], fmt=lambda x: f'{x:0.0f}%' if x > threshold else '', label_type='center')
ax.bar_label(ax.containers[5], fmt=lambda x: f'{x:0.0f}%' if x > threshold else '', label_type='center')
ax.bar_label(ax.containers[6], fmt=lambda x: f'{x:0.0f}%' if x > threshold else '', label_type='center')
ax.bar_label(ax.containers[7], fmt=lambda x: f'{x:0.0f}%' if x > threshold else '', label_type='center')

#Move the legend and rotate the bar label text from vertical to horizontal
plt.xticks(rotation='horizontal')
plt.legend(bbox_to_anchor=(0.5, 1.1), loc='upper center', fancybox=True, shadow=True, ncol=8)
plt.tight_layout()
#histogram = np.histogram(CalculatePerCentByLocationandCategory2, bins=8)
# save plot
#ReportWeekString = str(ReportWeek)
filename = ('Sales by Category vs Location Week ')
plt.savefig(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.png')

#END OF BAR GRAPH FOR SALES BY CATEGORY VERSUS LOCATION TABLE 2

#SIX TABLES FOR TOP 5 BRANDS BY CATEGORY

#Define a function to set alternate row colors for the tables
def set_row_colors():
    table[(1,0)].set_facecolor('Lightgray')
    table[(1,1)].set_facecolor('Lightgray')
    table[(3,0)].set_facecolor('Lightgray')
    table[(3,1)].set_facecolor('Lightgray')

#Create a dataframe for the top 5 brands by Category "Edible"
Top5BrandsByEdible = df_sales4.groupby(['Category','Brand'])['Final Sale'].sum()
Top5BrandsByEdible = Top5BrandsByEdible.loc['Edible'].nlargest(5)
#Convert the series into a dataframe and round to two decimal points
df_sales11 = pd.DataFrame(Top5BrandsByEdible).round(2)
#Reset the index to make Category a column rather than the index
df_sales11 = df_sales11.reset_index(drop=False)
#Add the name to the first two columns
df_sales11.columns = ['Brand', 'Total Sales']
#Export the dataframe as Excel to preserve for further Excel operations as needed or re-import
filename = 'Top Brands by Edible Week of '
df_sales11.to_excel(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.xlsx')

#FORMAT AND OUTPUT TOP EDIBLE TABLE TO IMAGE
#Simple .table method to export dataframe to image
Top5BrandsByEdibleTable = df_sales11
#Format float numbers for sales to currency
Top5BrandsByEdibleTable["Total Sales"] = Top5BrandsByEdibleTable['Total Sales'].map('${:,.0f}'.format)
#Replace all 0 values with blank space
Top5BrandsByEdibleTable = Top5BrandsByEdibleTable.replace("$0", '')
#define figure and axes
fig, ax = plt.subplots()
# hide the axes
ax.patch.set_visible(True)
fig.patch.set_visible(True)
ax.axis('off')
ax.axis('tight')
Columns4 = ['Brand','Total Sales']
# ccolors = plt.cm.Greens(np.full(len(Columns4), 0.4))
#create table
table = ax.table(cellText=Top5BrandsByEdibleTable.values, colLabels=Columns4, loc='center', cellLoc='center', edges='closed')
table.scale(1, 3)

table.auto_set_column_width(col=list(range(len(Top5BrandsByEdibleTable.columns))))
#Auto set row height
# table.auto_set_row_height(row=list(range(len(DailySalesByLocationTable.rows))))
table.auto_set_font_size(True)
#Alternate the row colors
# for i in range(0, len(Top5BrandsByEdibleTable), 2):
#     for j in range(len(Columns4)):
#         table[(i+1,j)].set_facecolor("lightgrey")

for i in range(0, len(Top5BrandsByEdibleTable), 2):
    for j in range(len(Columns4)):
        table[(i+1,j)].set_facecolor("mediumseagreen")
        table[(i+1,j)].set_text_props(weight='bold', color='w')

for i in range(0, len(Top5BrandsByEdibleTable), 2):
    for j in range(len(Columns4)):
        table[(i,j)].set_facecolor("limegreen")
        table[(i,j)].set_text_props(weight='bold', color='w')

#Reset column labels row background color and font color
for j in range(len(Columns4)):
    table[(0,j)].set_facecolor("darkgreen")
    table[(0,j)].set_text_props(weight='bold', color='w')

table.header_row = None

bbox=[0, 0, 1, 1]
#display table
fig.tight_layout()

#save table
filename = 'Top Brands by Edible Week of '
# fig.savefig('Igadi Average Ticket Number of Transactions.png', bbox_inches='tight', dpi=300)
fig.savefig(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.png', dpi=300)

#Create a dataframe for the top 5 brands by Category "PackedBud"
Top5BrandsByPackedBud = df_sales4.groupby(['Category','Brand'])['Final Sale'].sum()
Top5BrandsByPackedBud = Top5BrandsByPackedBud.loc['PackedBud'].nlargest(5)
#Convert the series into a dataframe and round to two decimal points
df_sales12 = pd.DataFrame(Top5BrandsByPackedBud).round(2)
#Reset the index to make Category a column rather than the index
df_sales12 = df_sales12.reset_index(drop=False)
#Add the name to the first two columns
df_sales12.columns = ['Brand', 'Total Sales']
#Export the dataframe as Excel to preserve for further Excel operations as needed or re-import
filename = 'Top Brands By Packed Bud Week of '
df_sales12.to_excel(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.xlsx')

#FORMAT AND OUTPUT TOP PACKED BUD TABLE TO IMAGE
#Simple .table method to export dataframe to image
Top5BrandsByPackedBudTable = df_sales12
#Format float numbers for sales to currency
Top5BrandsByPackedBudTable["Total Sales"] = Top5BrandsByPackedBudTable['Total Sales'].map('${:,.0f}'.format)
#Replace all 0 values with blank space
Top5BrandsByPackedBudTable = Top5BrandsByPackedBudTable.replace("$0", '')
#define figure and axes
fig, ax = plt.subplots()
# hide the axes
ax.patch.set_visible(True)
fig.patch.set_visible(True)
ax.axis('off')
ax.axis('tight')
Columns5 = ['Brand','Total Sales']
# ccolors = plt.cm.Greens(np.full(len(Columns5), 0.4))
#create table
table = ax.table(cellText=Top5BrandsByPackedBudTable.values, colLabels=Columns5, loc='center', cellLoc='center', edges='closed')
table.scale(1, 3)

table.auto_set_column_width(col=list(range(len(Top5BrandsByPackedBudTable.columns))))
#Auto set row height
# table.auto_set_row_height(row=list(range(len(DailySalesByLocationTable.rows))))
table.auto_set_font_size(True)
#Alternate the row colors
# for i in range(0, len(Top5BrandsByPackedBubTable), 2):
#     for j in range(len(Columns5)):
#         table[(i+1,j)].set_facecolor("lightgrey")

for i in range(0, len(Top5BrandsByPackedBudTable), 2):
    for j in range(len(Columns5)):
        table[(i+1,j)].set_facecolor("mediumseagreen")
        table[(i+1,j)].set_text_props(weight='bold', color='w')

for i in range(0, len(Top5BrandsByPackedBudTable), 2):
    for j in range(len(Columns5)):
        table[(i,j)].set_facecolor("limegreen")
        table[(i,j)].set_text_props(weight='bold', color='w')

#Reset column labels row background color and font color
for j in range(len(Columns5)):
    table[(0,j)].set_facecolor("darkgreen")
    table[(0,j)].set_text_props(weight='bold', color='w')

table.header_row = None
bbox=[0, 0, 1, 1]

#display table
fig.tight_layout()

#save table
filename = 'Top Brands by Packed Bud Week of '
# fig.savefig('Igadi Average Ticket Number of Transactions.png', bbox_inches='tight', dpi=300)
fig.savefig(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.png', dpi=300)



#Create a dataframe for the top 5 brands by Category "BulkBud"
Top5BrandsByBulkBud = df_sales4.groupby(['Category','Brand'])['Final Sale'].sum()
Top5BrandsByBulkBud = Top5BrandsByBulkBud.loc['BulkBud'].nlargest(5)
#Convert the series into a dataframe and round to two decimal points
df_sales13 = pd.DataFrame(Top5BrandsByBulkBud).round(2)
#Reset the index to make Category a column rather than the index
df_sales13 = df_sales13.reset_index(drop=False)
#Add the name to the first two columns
df_sales13.columns = ['Brand', 'Total Sales']
#Export the dataframe as Excel to preserve for further Excel operations as needed or re-import
filename = 'Top Brands by Bulk Bud Week of '
df_sales13.to_excel(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.xlsx')

#FORMAT AND OUTPUT TOP BULKBUD TABLE TO IMAGE
#Simple .table method to export dataframe to image
Top5BrandsByBulkBudTable = df_sales13
#Format float numbers for sales to currency
Top5BrandsByBulkBudTable["Total Sales"] = Top5BrandsByBulkBudTable['Total Sales'].map('${:,.0f}'.format)
#Replace all 0 values with blank space
Top5BrandsByBulkBudTable = Top5BrandsByBulkBudTable.replace("$0", '')
#define figure and axes
fig, ax = plt.subplots()
# hide the axes
ax.patch.set_visible(True)
fig.patch.set_visible(True)
ax.axis('off')
ax.axis('tight')
Columns6 = ['Brand','Total Sales']
# ccolors = plt.cm.Greens(np.full(len(Columns6), 0.4))
#create table
table = ax.table(cellText=Top5BrandsByBulkBudTable.values, colLabels=Columns6, loc='center', cellLoc='center', edges='closed')
table.scale(1, 3)

table.auto_set_column_width(col=list(range(len(Top5BrandsByBulkBudTable.columns))))
#Auto set row height
# table.auto_set_row_height(row=list(range(len(DailySalesByLocationTable.rows))))
table.auto_set_font_size(True)
#Alternate the row colors
# for i in range(0, len(Top5BrandsByBulkBudTable), 2):
#     for j in range(len(Columns6)):
#         table[(i+1,j)].set_facecolor("lightgrey")

for i in range(0, len(Top5BrandsByBulkBudTable), 2):
    for j in range(len(Columns6)):
        table[(i+1,j)].set_facecolor("mediumseagreen")
        table[(i+1,j)].set_text_props(weight='bold', color='w')

for i in range(0, len(Top5BrandsByBulkBudTable), 2):
    for j in range(len(Columns6)):
        table[(i,j)].set_facecolor("limegreen")
        table[(i,j)].set_text_props(weight='bold', color='w')

#Reset column labels row background color and font color
for j in range(len(Columns6)):
    table[(0,j)].set_facecolor("darkgreen")
    table[(0,j)].set_text_props(weight='bold', color='w')
    
table.header_row = None
bbox=[0, 0, 1, 1]

#display table
fig.tight_layout()

#save table
filename = 'Top Brands by Bulk Bud Week of '
# fig.savefig('Igadi Average Ticket Number of Transactions.png', bbox_inches='tight', dpi=300)
fig.savefig(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.png', dpi=300)


#Create a dataframe for the top 5 brands by Category "Cartridge"
Top5BrandsByCartridge = df_sales4.groupby(['Category','Brand'])['Final Sale'].sum()
Top5BrandsByCartridge = Top5BrandsByCartridge.loc['Cartridge'].nlargest(5)

#Convert the series into a dataframe and round to two decimal points
df_sales14 = pd.DataFrame(Top5BrandsByCartridge).round(2)
#Reset the index to make Category a column rather than the index
df_sales14 = df_sales14.reset_index(drop=False)
#Add the name to the first two columns
df_sales14.columns = ['Brand', 'Total Sales']
#If number of rows is less than 5, add rows with 0 values to make 5 rows
if len(df_sales14) < 5:
    padding = pd.DataFrame({'Brand': [''] * (5 - len(df_sales14)), 'Total Sales': [0] * (5 - len(df_sales14))})
    df_sales14 = df_sales14.append(padding, ignore_index=True)
# print(df_sales14)

#Export the dataframe as Excel to preserve for further Excel operations as needed or re-import
filename = 'Top Brands by Cartridge Week of '
df_sales14.to_excel(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.xlsx')

#FORMAT AND OUTPUT TOP CARTRIDGE TABLE TO IMAGE
#Simple .table method to export dataframe to image
Top5BrandsByCartridgeTable = df_sales14
#Format float numbers for sales to currency
Top5BrandsByCartridgeTable["Total Sales"] = Top5BrandsByCartridgeTable['Total Sales'].map('${:,.0f}'.format)
#Replace all 0 values with blank space
Top5BrandsByCartridgeTable = Top5BrandsByCartridgeTable.replace("$0", '')
#define figure and axes
fig, ax = plt.subplots()
# hide the axes
ax.patch.set_visible(True)
fig.patch.set_visible(True)
ax.axis('off')
ax.axis('tight')
Columns7 = ['Brand','Total Sales']
# ccolors = plt.cm.Greens(np.full(len(Columns7), 0.4))
#create table
table = ax.table(cellText=Top5BrandsByCartridgeTable.values, colLabels=Columns7, loc='center', cellLoc='center', edges='closed')
table.scale(1, 3)

table.auto_set_column_width(col=list(range(len(Top5BrandsByCartridgeTable.columns))))
#Auto set row height
# table.auto_set_row_height(row=list(range(len(DailySalesByLocationTable.rows))))
table.auto_set_font_size(True)
#Alternate the row colors
# for i in range(0, len(Top5BrandsByCartridgeTable), 2):
#     for j in range(len(Columns7)):
#         table[(i+1,j)].set_facecolor("lightgrey")

for i in range(0, len(Top5BrandsByCartridgeTable), 2):
    for j in range(len(Columns7)):
        table[(i+1,j)].set_facecolor("mediumseagreen")
        table[(i+1,j)].set_text_props(weight='bold', color='w')

for i in range(0, len(Top5BrandsByCartridgeTable), 2):
    for j in range(len(Columns7)):
        table[(i,j)].set_facecolor("limegreen")
        table[(i,j)].set_text_props(weight='bold', color='w')

#Reset column labels row background color and font color
for j in range(len(Columns7)):
    table[(0,j)].set_facecolor("darkgreen")
    table[(0,j)].set_text_props(weight='bold', color='w')

table.header_row = None

bbox=[0, 0, 1, 1]

#display table
fig.tight_layout()

#save table
filename = 'Top Brands by Cartridge Week of '
# fig.savefig('Igadi Average Ticket Number of Transactions.png', bbox_inches='tight', dpi=300)
fig.savefig(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.png', dpi=300)



#Create a dataframe for the top 5 brands by Category "Concentrate"
Top5BrandsByConcentrate = df_sales4.groupby(['Category','Brand'])['Final Sale'].sum()
Top5BrandsByConcentrate = Top5BrandsByConcentrate.loc['Concentrate'].nlargest(5)
#Convert the series into a dataframe and round to two decimal points
df_sales15 = pd.DataFrame(Top5BrandsByConcentrate).round(2)
#Reset the index to make Category a column rather than the index
df_sales15 = df_sales15.reset_index(drop=False)
#Add the name to the first two columns
df_sales15.columns = ['Brand', 'Total Sales']
#Export the dataframe as Excel to preserve for further Excel operations as needed or re-import
filename = 'Top Brands by Concentrate Week of '
df_sales15.to_excel(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.xlsx')

#FORMAT AND OUTPUT TOP CONCENTRATE TABLE TO IMAGE
#Simple .table method to export dataframe to image
Top5BrandsByConcentrateTable = df_sales15
#Format float numbers for sales to currency
Top5BrandsByConcentrateTable["Total Sales"] = Top5BrandsByConcentrateTable['Total Sales'].map('${:,.0f}'.format)
#Replace all 0 values with blank space
Top5BrandsByConcentrateTable = Top5BrandsByConcentrateTable.replace("$0", '')
#define figure and axes
fig, ax = plt.subplots()
# hide the axes
ax.patch.set_visible(True)
fig.patch.set_visible(True)
ax.axis('off')
ax.axis('tight')
Columns8 = ['Brand','Total Sales']
# ccolors = plt.cm.Greens(np.full(len(Columns8), 0.4))
#create table
table = ax.table(cellText=Top5BrandsByConcentrateTable.values, colLabels=Columns8, loc='center', cellLoc='center', edges='closed')
table.scale(1, 3)

table.auto_set_column_width(col=list(range(len(Top5BrandsByConcentrateTable.columns))))
#Auto set row height
# table.auto_set_row_height(row=list(range(len(DailySalesByLocationTable.rows))))
table.auto_set_font_size(True)
#Alternate the row colors
# for i in range(0, len(Top5BrandsByConcentrateTable), 2):
#     for j in range(len(Columns8)):
#         table[(i+1,j)].set_facecolor("lightgrey")

for i in range(0, len(Top5BrandsByConcentrateTable), 2):
    for j in range(len(Columns8)):
        table[(i+1,j)].set_facecolor("mediumseagreen")
        table[(i+1,j)].set_text_props(weight='bold', color='w')

for i in range(0, len(Top5BrandsByConcentrateTable), 2):
    for j in range(len(Columns8)):
        table[(i,j)].set_facecolor("limegreen")
        table[(i,j)].set_text_props(weight='bold', color='w')

#Reset column labels row background color and font color
for j in range(len(Columns8)):
    table[(0,j)].set_facecolor("darkgreen")
    table[(0,j)].set_text_props(weight='bold', color='w')

table.header_row = None

bbox=[0, 0, 1, 1]

#display table
fig.tight_layout()

#save table
filename = 'Top Brands by Concentrate Week of '
# fig.savefig('Igadi Average Ticket Number of Transactions.png', bbox_inches='tight', dpi=300)
fig.savefig(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.png', dpi=300)


#Create a dataframe for the top 5 brands by Category "Prerolls"
Top5BrandsByPrerolls = df_sales4.groupby(['Category','Brand'])['Final Sale'].sum()
Top5BrandsByPrerolls = Top5BrandsByPrerolls.loc['Prerolls'].nlargest(5)
#Convert the series into a dataframe and round to two decimal points
df_sales16 = pd.DataFrame(Top5BrandsByPrerolls).round(2)
#Reset the index to make Category a column rather than the index
df_sales16 = df_sales16.reset_index(drop=False)
#Add the name to the first two columns
df_sales16.columns = ['Brand', 'Total Sales']
#Export the dataframe as Excel to preserve for further Excel operations as needed or re-import
filename = 'Top Brands by Prerolls Week of '
df_sales16.to_excel(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.xlsx')

#FORMAT AND OUTPUT TOP PREROLLS TABLE TO IMAGE
#Simple .table method to export dataframe to image
Top5BrandsByPrerollsTable = df_sales16
#Format float numbers for sales to currency
Top5BrandsByPrerollsTable["Total Sales"] = Top5BrandsByPrerollsTable['Total Sales'].map('${:,.0f}'.format)
#Replace all 0 values with blank space
Top5BrandsByPrerollsTable = Top5BrandsByPrerollsTable.replace("$0", '')
#define figure and axes
fig, ax = plt.subplots()
# hide the axes
ax.patch.set_visible(True)
fig.patch.set_visible(True)
ax.axis('off')
ax.axis('tight')
Columns9 = ['Brand','Total Sales']
ccolors = plt.cm.Greens(np.full(len(Columns9), 0.4))
#create table
table = ax.table(cellText=Top5BrandsByPrerollsTable.values, loc='center', colLabels=Columns9, cellLoc='center', edges='closed')

table.scale(1, 3)

table.auto_set_column_width(col=list(range(len(Top5BrandsByPrerollsTable.columns))))

#Auto set row height
#table.auto_set_row_height(row=list(range(len(Top5BrandsByPrerollsTable.rows))))
table.auto_set_font_size(True)
#Alternate the row colors
# for i in range(0, len(Top5BrandsByPrerollsTable), 2):
#     for j in range(len(Columns9)):
#         table[(i+1,j)].set_facecolor("lightgrey")

for i in range(0, len(Top5BrandsByPrerollsTable), 2):
    for j in range(len(Columns9)):
        table[(i+1,j)].set_facecolor("mediumseagreen")
        table[(i+1,j)].set_text_props(weight='bold', color='w')

for i in range(0, len(Top5BrandsByPrerollsTable), 2):
    for j in range(len(Columns9)):
        table[(i,j)].set_facecolor("limegreen")
        table[(i,j)].set_text_props(weight='bold', color='w')

#Reset column labels row background color and font color
for j in range(len(Columns9)):
    table[(0,j)].set_facecolor("darkgreen")
    table[(0,j)].set_text_props(weight='bold', color='w')

table.header_row = None
bbox=[0, 0, 1, 1]

#display table
fig.tight_layout()

plt.show()
#save table
filename = 'Top Brands by Prerolls Week of '
# fig.savefig('Igadi Average Ticket Number of Transactions.png', bbox_inches='tight', dpi=300)
fig.savefig(f'{CurrentWeekReportFolder}{filename}{ReportStartDate}{FilenameFormatter}{ReportEndDate}.png', dpi=300)

#Create a dataframe for the top 5 brands by Category "Topicals"





## ?:how to use concat to add a row to a dataframe?
#df_sales5 = pd.concat([df_sales4, sum], axis=1)

# ?: display the dataframe as a table
#df_sales5.style.format({'Final Sale':'{:.2f}'})










