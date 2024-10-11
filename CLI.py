def intro():
  print("#######################################################################")
  print("\tAURECON MACC GENERATOR")
  print("\twritten by Melvin Chan")
  print("\ton python 3.12")
  print("\tFeel free to ask for inquire if you need more data work done!")
  print("#######################################################################x")

def inputs(data):
  ## Variables (Variable to change)
  sector = input("Enter Sector: ")
  sector_set = set()
  for act in data: 
      sector_set.add(act.sector_name)
  if sector not in sector_set:
      print("Sector not in data, the available sectors are:")
      for sector in sector_set:
          print("\t"+sector)
      return

  year = int(input("Enter year: "))
  if int(year) not in range(2022, 2051):
      print("incorrect year, only projecting between 2022-2050 inclusive")
      return
  return (sector, year)

def debug(abatement_cost, volume):
    print("This is the abatement cost")
    print(abatement_cost)
    print("This is the volume")
    print(volume)