#! usr/bin/env python3 

import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from pprint import pprint
import random
from pyxlsb import open_workbook as open_xlsb
import os
import CLI
import collections

## Emission Reducing Activity Class Object
## volume is the CO2 in tonnes that can be abated
## potential is the % of CO2 volume that can be technically abated through ERA
class ERA():
    def __init__(self, id, sector_name, sub_sector_name, action, year):
        self.id = id
        self.sector_name = sector_name
        self.sub_sector_name = sub_sector_name
        self.action = action
        self.volume = float(-2)
        self.year = int(year)
        self.abatement_cost = float(-2)

    ## Function for debugging, just prints out the ERA object 
    def print_entity(self):
        print(f"id: {self.id}")  
        print(f"Sector: {self.sector_name}")
        print(f"Activity Name: {self.sub_sector_name}")
        print(f"Action Name: {self.action}")
        print(f"Year: {self.year}")
        print(f"Volume: {self.volume}")
        print(f"Abatement_cost: {self.abatement_cost}")
        print("------")

    def set_volume(self, volume):
        self.volume = volume
    def set_abatement_cost(self, abatement_cost):
        self.abatement_cost = abatement_cost


## Function for debuggin to see what data is selected
def display(data):
    for act in data:
        print("--------")
        act.print_entity()

## When user input a year and emitting activity to analyse, this function returns
## a filtered array from the datastore
def data_select(year, sector, data):
    target = []
    for act in data:
        if (int(act.year) == int(year) and act.sector_name.lower() == sector.lower()):
            target.append(act)
    return target


## Function that preprocesses excel into a datastore of ERA objects
def pre_process(sheet):
    data = collections.defaultdict(dict)
    num_ERA = 0

    ## Counts the number of ERA from the spreadsheet
    for row in sheet.rows():
        if (type(row[0].v) == float and int(row[0].v) >= num_ERA):
            num_ERA = int(row[0].v)

    print(f"The num_ERA = {num_ERA}")
    ## Starts at 2022 (col 5) and ends at 2050 (col 33) in the template
    start = 5 
    end = 33

    for row in sheet.rows():
        if (type(row[0].v) == float):
            for yr in range(int(start), int(end)+1): 
                data[row[0].v][yr] = ERA(row[0].v, row[1].v, row[2].v, row[3].v, yr+2017)
        if (row[0].v == num_ERA):
            break

    for row in sheet.rows():
        if (type(row[0].v) == float):
            for yr in range(int(start), int(end)+1): 
                if (row[4].v == "Cost"):
                    if (row[yr].v == '-'):
                       data[row[0].v][yr].set_abatement_cost(-1)
                    else:
                        data[row[0].v][yr].set_abatement_cost(row[yr].v)
                elif (row[4].v == "Volume"):
                    if (row[yr].v == '-'):
                        data[row[0].v][yr].set_volume(-1)   
                    else:
                        data[row[0].v][yr].set_volume(row[yr].v)   

    # Takes the dictionary and returns it in an array              
    ret_data = []
    for outer_key, inner_dict in data.items():
        for inner_key, value in inner_dict.items():
            ret_data.append(value)
    return ret_data

# function that does polynomial interpolation (this can be swapped out for log)
def poly_function(target, year):
    x = []
    y = []
    for point in target:
        x.append(point[0])
        y.append(point[1])
    poly = np.polyfit(x,y,2)
    x_prime = year - 2022
    new_vol = 0
    for i in range(0, len(poly)):
        new_vol += poly[i]*(x_prime**(len(poly)-1-i))
    return new_vol

## Check number of existing data for that particular activity 
def scope(id, year, data, check_volume):
    ## Variable for the number of available data points for each year
    available = 0
    target = []
    if (check_volume == True):
        for act in data:
            if (int(act.id) == int(id) and int(act.volume) != int(-1)):
                available += 1
                target.append((act.year, act.volume))
        if available == 0:
            print("Insufficient data inputted, check that all relevant cells are filled")
            exit()
        if (available == 1):    
            new_vol = target[0][1] 
            return new_vol
        # get all data points and interpolate between the given data points
        else:
            new_vol = poly_function(target, year)
            return new_vol
    else:
        for act in data:
            if (int(act.id) == int(id) and int(act.abatement_cost) != int(-1)):
                available += 1
                target.append((act.year, act.abatement_cost))
        if (available == 1):    
            new_cost = target[0][1] 
            return new_cost
        # get all data points and interpolate
        else:
            new_cost = poly_function(target, year)
            return new_cost
                
# Apply fitting algorithm to given data: Uses polynomial interpolation
def interpolate_volume(data):
    target = []
    for act in data:
        if (int(act.volume) == int(-1)):
            # Check number of existing data for that year
            new_vol = scope(act.id, act.year, data, check_volume=True)
            new_ERA = ERA(act.id, act.sector_name, act.sub_sector_name, act.action, act.year+2017)
            new_ERA.set_abatement_cost(act.abatement_cost)
            new_ERA.set_volume(new_vol)
            target.append(new_ERA)
        else:
            target.append(act)
    return target

def interpolate_cost(data):
    target = []
    for act in data:
        if (int(act.abatement_cost) == int(-1)):
            # Check number of existing data for that year
            new_cost = scope(act.id, act.year, data, check_volume=False)
            new_ERA = ERA(act.id, act.sector_name, act.sub_sector_name, act.action, act.year+2017)
            new_ERA.set_abatement_cost(new_cost)
            new_ERA.set_volume(act.volume)
            target.append(new_ERA)
        else:
            target.append(act)
    return target


def main(): 

    ## Code for processing excel into datastore (blackbox)
    dir_list = os.listdir("./input")
    if len(dir_list) > 1:
        print("There are more than one file in the input folder")
        return
    wb = open_xlsb(f"./input/{dir_list[0]}")
    sheet = wb.get_sheet(1)
    data = pre_process(sheet)
    
    display(data)
    ## interpolating
    data = interpolate_volume(data)
    data = interpolate_cost(data)
    ## Inputs
    CLI.intro()
    # sector, year =  CLI.inputs(data)
    sector = "Industry"
    year = 2030

    ## Selects the data based on the filter given
    plot_data = data_select(year, sector, data)
    ## sort the data based on abatement cost to form a cost ranking
    plot_data = sorted(plot_data, key=lambda x: x.abatement_cost)

    ## Splits the data into abatement cost and volume
    abatement_cost = []
    volume = [] 
    for dp in plot_data:
        abatement_cost.append(dp.abatement_cost)
        volume.append(dp.volume)
    num_target_ERA = len(volume)

    CLI.debug(abatement_cost, volume)

    track_volume = 0
    print("calculating", end='', flush=True)
    for i in range(0, int(num_target_ERA)):
        bar_range = range(track_volume, track_volume + int(volume[i]/1000))
        track_volume += int(volume[i]/1000)
        plt.bar(bar_range, abatement_cost[i], color=(random.random(), random.random(), random.random()), align='edge', width=1)
        print(".", end ='', flush=True)

    plt.show()    


if __name__=="__main__":
    main()