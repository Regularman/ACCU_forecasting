import collections
from main import ERA
from main import (interpolate_cost, interpolate_volume, display, ERA)
import os
from pyxlsb import open_workbook
import re
import pandas as pd


## Function that preprocesses excel into a datastore of ERA objects
def pre_process(sheet, sector_name, total_num_ERA):
    data = collections.defaultdict(dict)

    num_ERA = 0

    ## Counts the number of ERA from the spreadsheet
    for row in sheet.rows():
      if (row[1].v == "Outputs"):
          break    
      if (row[3].v is not None):
        if (type(row[3].v) == float and int(row[3].v) >= num_ERA):
            num_ERA = int(row[3].v)
    print(f"{sheet.name} {num_ERA}")
    
    ## Starts at 2022 (col 15) and ends at 2050 (col 43) in the template
    start = 15
    end = 43
    sub_sector_emission = 0
    for i, row in enumerate(sheet.rows(), start=1):  # start=1 for 1-based row indexing
      if i == 30:
        if (row[7].v is not None):
          sub_sector_emission = row[7].v
          print(f"The subsector emission is {sub_sector_emission}")
          break
      if i == 31:
        if (row[7].v is not None):
          sub_sector_emission = row[7].v
          print(f"The subsector emission is {sub_sector_emission}")
          break
    for row in sheet.rows():
        if (type(row[3].v) == float):
            for yr in range(int(start), int(end)+1): 
                key = total_num_ERA + row[3].v
                data[key][yr] = ERA(row[3].v, sector_name, sheet.name, row[5].v, yr+2007)
                data[key][yr].set_subsector_volume(sub_sector_emission)
        if (row[1].v == "Outputs"):
            break
    
    is_emission_cost = False
    is_emission_volume = False
    is_emission_relative_volume = False
    for row in sheet.rows():
        if (row[1].v == "Outputs"):
          break
        if (row[3].v == "Emissions Abatement Potential - Abatement Cost"):
          is_emission_cost = True
        if (row[3].v == "Emissions Abatement Potential - Absolute (tpa-CO2e)"):
          is_emission_volume = True
          is_emission_cost = False
        if (row[3].v == "Emissions Abatement Potential - Relative (%% of Emitting Activity)"):
          is_emission_volume = False
          is_emission_cost = False
          is_emission_relative_volume = True
        if (type(row[3].v) == float and is_emission_cost):
          for yr in range(int(start), int(end)+1): 
            if row[yr].v is not None:
              key = total_num_ERA + row[3].v
              data[key][yr].set_abatement_cost(row[yr].v)
            else: 
              key = total_num_ERA + row[3].v
              data[key][yr].set_abatement_cost(-1)
        if (type(row[3].v) == float and is_emission_volume):
          for yr in range(int(start), int(end)+1): 
            if row[yr].v is not None:
              key = total_num_ERA + row[3].v
              data[key][yr].set_volume(row[yr].v)
            else: 
              key = total_num_ERA + row[3].v
              data[key][yr].set_volume(-1)

    ret_data = []
    for outer_key, outer in data.items():
        for inner_key, inner in outer.items():
            ret_data.append(inner)
    return (num_ERA, ret_data)
## Function that exports all interpolated data into a new spreadsheet
def export(data, num_ERA):

    ## Converted into dicts for optimisation
    data_dict = list_to_dict(data)
    ## Need to create the dataframe from data here
    cost_header = ["Sector", "Emitting Activity", "ERA"]
    for year in range(2022, 2051):
        cost_header.append(year)
    volume_header = ["Sector", "Emitting Activity", "ERA", "Sector Percentage"]
    for year in range(2022, 2051):
      volume_header.append(year)

    IDs = []
    for ID in range(1, num_ERA+1):
        IDs.append(ID)

    volume_data = []
    for ERA_id in range(1, num_ERA + 1):

        volume = data_dict[ERA_id][2022].volume
        
        ERA_data = []
        ERA_data.append(data_dict[ERA_id][2022].sector_name)        
        ERA_data.append(data_dict[ERA_id][2022].sub_sector_name)
        ERA_data.append(data_dict[ERA_id][2022].action)
        ERA_data.append(volume)

        for year in range (2022, 2051):
            ERA_data.append(float(data_dict[ERA_id][year].volume) * float(data_dict[ERA_id][year].subsector_volume))

            
        volume_data.append(ERA_data)
    volume_df = pd.DataFrame(volume_data, index=IDs, columns=volume_header)  # doctest: +SKIP

    cost_data = []
    for ERA_id in range(1, num_ERA + 1):
        ERA_data = []
        ERA_data.append(data_dict[ERA_id][2022].sector_name)        
        ERA_data.append(data_dict[ERA_id][2022].sub_sector_name)
        ERA_data.append(data_dict[ERA_id][2022].action)

        for year in range (2022, 2051):
            ERA_data.append(data_dict[ERA_id][year].abatement_cost)
        cost_data.append(ERA_data)
    abatement_cost_df = pd.DataFrame(cost_data, index=IDs, columns=cost_header)  # doctest: +SKIP
    with pd.ExcelWriter(
        "./output/output.xlsx", 
        mode="w", 
        engine="openpyxl") as writer:
            volume_df.to_excel(writer, sheet_name="volume")  # doctest: +SKIP
            abatement_cost_df.to_excel(writer, sheet_name="abatement_cost")
    print("Exported data")

def list_to_dict(data):
    # Takes the dictionary and returns it in an array              
    ret_data = collections.defaultdict(dict)
    for id, item in enumerate(data):
        ret_data[int(id/29)+1][item.year] = item
    return ret_data

datastore = []
files = os.listdir("./compile")
input_files = []
for file in files:
  if re.match("^[0-9]+.", file):
      input_files.append(file)
total_num_ERA = 0
for file in input_files:
    # datastore.append(pre_process(f"./compile/{file}"))
    workbook = open_workbook(f"./compile/{file}")
    datasheets = []
    for sheet in workbook.sheets:
      if re.match("^[0-9]+.", sheet):
          datasheets.append(sheet)

    for sheet in datasheets:
      num_ERA, data = pre_process(workbook.get_sheet(sheet), file, total_num_ERA)
      data = interpolate_volume(data)
      data = interpolate_cost(data)
      total_num_ERA += num_ERA

      for data_point in data:
         datastore.append(data_point)

display(datastore)
print(len(datastore))
print(total_num_ERA)
export(datastore, total_num_ERA)