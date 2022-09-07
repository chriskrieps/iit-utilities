from typing import Iterable
import pandas as pd
import xlsxwriter
from functools import reduce

# Load excel data from each sheet into a dictionary of dataframes
dict_df = pd.read_excel('.\Data\OG Meter Data Summary.xlsx', sheet_name=None)

# Create an excel writer and create a target file for updated data
Excelwriter = pd.ExcelWriter(".\Data\Clean Meter Data.xlsx", engine="xlsxwriter")

# Create empty dictionary to hold consolidated building dataframes
cons_dict_df = {}

# Building Dictionary
build_dict = {
    'TN':'Tech North',
    'TC':'Tech Central',
    'TS':'Tech South',
    'AM':'Alumni Hall',
    'ASA':'ASA',
    'ASP':'ASP',
    'CH':'Carman Hall',
    'DTD':'DTD',
    'RE':'John Rettaliata Engineering Center',
    'FS':'Facilities Building',
    'FH':'Farr Hall',
    'GL':'Galvin Library',
    'HP':'Heating Plant',
    'HH':'Herman Hall',
    'INC':'Incubator',
    'IT':'IIT Tower',
    'KP':'Kappa',
    'KH':'Keating Hall',
    'PS':'Pritzker Science Center',
    'LSR':'Life Science Research',
    'MH':'Machinery Hall',
    'MB':'Main Building',
    'MTB':'Metals Building',
    'MSV':'MSV',
    'EH':'MSV East Hall',
    'FO':'MSV Fowler Hall',
    'GH':'MSV Graduate Hall',
    'LH':'MSV Lewis Hall',
    'MC':'MTTC',
    'PH':'Perlstein Hall',
    'PKP':'PKP',
    'PKS':'PKS',
    'SH':'Siegel Hall',
    'SPE':'SPE',
    'SVN':'SSV North Building',
    'SVM':'SSV Middle Building',
    'SVS':'SSV South Building',
    'SB':'Stuart Building',
    'TBC':'TBC',
    'VA1':'Vandercook 1',
    'VA2':'Vandercook 2',
    'WH':'Wishnick Hall',
    'BH':'Bailey Hall',
    'CR':'Crown Hall',
    'CU':'Cunningham Hall',
    'GU':'Gunsaulus Hall',
    'SS':'South Substation'
}

# Building Dictionary
building_areas = {
    'TN': 41576,
    'TC': 80660,
    'TS': 82554,
    'AM': 39542,
    'ASA':15616,
    'ASP':12402,
    'CH': 69559,
    'DTD': 18360,
    'RE': 133990,
    'FS': 20363,
    'FH': 24894,
    'GL': 92978,
    'HP': 16298,
    'HH': 111135,
    'INC': 69841,
    'IT': 392894,
    'KP': 15616,
    'KH': 53163,
    'PS': 123454,
    'LSR': 106758,
    'MH': 30399,
    'MB': 63155,
    'MTB': 68398,
    'MSV': 87647,
    'EH': 30886,
    'FO': 24062,
    'GH': 30453,
    'LH': 32318,
    'MC': 93667,
    'PH': 102517,
    'PKP':12876,
    'PKS':12876,
    'SH': 63711,
    'SPE': 15616,
    'SVN': 36500,
    'SVM': 36500,
    'SVS': 36500,
    'SB': 83906,
    'TBC':140788,
    'VA1': 33317,
    'VA2': 27319,
    'WH': 62913,
    'BH': 69559,
    'CR': 53901,
    'CU': 70269,
    'GU': 82898,
    'SS': 0
}

# Iterate over each sheet name within dict_df
for sht in dict_df:

    # Get the data frame associated with the sheet name
    df = dict_df[sht]

    # Check if there's a consumption column, if not then skip
    if 'CONSUMPTION' not in df.columns:
        continue

    # Get the building abbreviation from the sheet name
    bld_name = sht.split(" ")[0]
    #print('Building Name: '+bld_name)

    # Get the utility type (elec, gas, hw, etc)
    util_type = sht.split(" ")[1].lower()

    if util_type.__contains__("elec"):
        util_type = "Elec"
    elif util_type.__contains__("gas"):
        util_type = "Gas"
    elif util_type.__contains__("chw"):
        util_type = "CHW"
    elif util_type.__contains__("hw"):
        util_type = "HW"
    elif util_type.__contains__("dw"):
        util_type = "DW"
    elif util_type.__contains__("steam"):
        util_type = "Steam"
    else:
        util_type = util_type

    #print("Utility Type: "+util_type)

    # Only grab the column data between DATE and CONSUMPTION
    df = df.loc[:,'DATE':'CONSUMPTION']

    # Remove missing values from dataframe
    df = df[df['METER READING'] != 'Missing']

    # Get any outliers from the consumption column
    q1 = df['CONSUMPTION'].quantile(0.25)
    q3 = df['CONSUMPTION'].quantile(0.75)

    # If quartiles don't exist, remove this sheet from the dictionary and move to next sheet to check
    if q1 == 'nan' or q3 == 'nan':
        dict_df.pop(sht)
        continue

    # Get range of non-outliers
    iqr = q3 - q1

    # Strip the outliers out of our data
    df = df[~((df['CONSUMPTION'] < (q1 - 1.5 * iqr)) |(df['CONSUMPTION'] > (q3 + 1.5 * iqr)))]

    # Get Average consumption for each month
    df = df.groupby(pd.PeriodIndex(df['DATE'], freq="M"))['CONSUMPTION'].mean().reset_index()

    # Multiply the avg consumption by number of days in month
    df["CONSUMPTION"] = df["CONSUMPTION"].mul(df["DATE"].dt.daysinmonth)

    # Rename the consumption column to be include utility type name
    df.rename(columns={'CONSUMPTION': util_type+' Consumption'}, inplace=True)

    # Overwrite the dataframe dictionary with our cleaned up data
    dict_df[sht] = df

    # Check if the building abbrev already exists in the consolidated dataframe
    # If already exists, merge the data
    if bld_name in cons_dict_df:
        if (util_type+' Consumption_7') in cons_dict_df[bld_name].columns:
            cons_dict_df[bld_name] = cons_dict_df[bld_name].merge(df, on='DATE', how='left', suffixes=("_9","_0"))
        elif (util_type+' Consumption_5') in cons_dict_df[bld_name].columns:
            cons_dict_df[bld_name] = cons_dict_df[bld_name].merge(df, on='DATE', how='left', suffixes=("_7","_8"))
        elif (util_type+' Consumption_3') in cons_dict_df[bld_name].columns:
            cons_dict_df[bld_name] = cons_dict_df[bld_name].merge(df, on='DATE', how='left', suffixes=("_5","_6"))
        elif (util_type+' Consumption_1') in cons_dict_df[bld_name].columns:
            cons_dict_df[bld_name] = cons_dict_df[bld_name].merge(df, on='DATE', how='left', suffixes=("_3","_4"))
        else:
            cons_dict_df[bld_name] = cons_dict_df[bld_name].merge(df, on='DATE', how='left', suffixes=("_1","_2"))

    # Else, add new data frame
    else:
        cons_dict_df[bld_name] = df

    # Add new column to each dataframe for building names
    df.insert(loc=1,column='Building', value=build_dict[bld_name])
    #cons_dict_df[bld_name].insert(loc=1,column='Building', value=build_dict[bld_name])

    # Add each dataframe to the excel Sheet
    df.to_excel(Excelwriter, sheet_name=sht,index=False)

# Save the final Excel file
Excelwriter.save()

#
# Time to create a new excel file which consolidates all of our building data into a single sheet for each building
#

# Create an excel writer and create a target file for consolidated data
Excelwriter2 = pd.ExcelWriter(".\Data\Consolidated Meter Data.xlsx", engine="xlsxwriter")

# Write consolidated dataframe dictionary to an excel file
for sht in cons_dict_df:

    # Get dataframe from dictionary based on sheet name
    cons_df = cons_dict_df[sht]

    # For each column convert to energy
    for column in cons_df:

        print('Column: '+column)

        ener_df = pd.DataFrame()

        # Check if we are looking at the DATE column

        # convert units to kBTU/h depending on source and assumptions listed
        util_type = column.split(" ")[0]
        match util_type:
            case "Elec":
                # Assume kW, 1 kW = 3.412 kBTU/h
                energy = cons_df[column]*3.412
            case "Gas":
                # Assume therms, 1 therm = 100,000 BTU/h = 100 kBTU/h
                energy = cons_df[column]*100
            case "CHW":
                # Assume BTU meter, 1000 BTU = 1 kBTU
                energy = cons_df[column]/1000
            case "HW":
                # Assume BTU meter, 1000 BTU = 1 kBTU
                energy = cons_df[column]/1000
            case "DW":
                # Assume gallons, what do I do with that?
                energy = cons_df[column]
            case "Steam":
                # Assume lbs/hr, estimate 15 psig steam at 912 latent heat of vaporization
                energy = cons_df[column]*912
            case "DATE":
                continue
            case "Building":
                continue

        print('energy: ')
        print(energy)
        ener_df['Total Energy'] = energy

        if 'Total Energy' in cons_df:

        # Add this energy to the Total Energy column
            cons_df['Total Energy'] = cons_df['Total Energy'] + ener_df['Total Energy']

        else: 
            cons_df['Total Energy'] = ener_df['Total Energy']

    
    cons_df["EUI"] = cons_df['Total Energy']/building_areas[sht]

    # Save dataframe to excel doc
    cons_df.to_excel(Excelwriter2, sheet_name=sht, index=False)

# Create a single combined dataframe with all building data
df_merged = pd.concat(cons_dict_df)

#reduce(lambda  left,right: pd.merge(left,right,on=['DATE'], how='outer'), cons_dict_df)

# Create a single combined sheet the combined dataframe data
df_merged.to_excel(Excelwriter2, sheet_name='Combined',index=False,)

# Save the excel file
Excelwriter2.save()

# Pull out the TN Elec sheet as a single dataframe to test
df = cons_dict_df['MTB']

# Print the dataframe
print(df)
