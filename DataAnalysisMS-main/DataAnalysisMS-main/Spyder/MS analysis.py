import pandas as pd
import matplotlib.pyplot as plt
import openpyxl

#reading data from the text file
workbook = "MSdata.asc"
df = pd.read_table(workbook, skiprows = 6)

test_column = df.iloc[0]

a = df["Ion Curent [A]"].min()
print(a)
#user chooses one of the measurable m/z values and the cycle of measurement
masses = ["2", "14", "15", "18", "20", "22", "26", "28", "29", "30", "31", "32", "37", "39", "40", "41", "42", "43", "45", "46", "48", "55", "56", "57", "58", "59", "60", "61", "62", "63", "64", "67", "68", "70", "72", "74"]
while True:
    a = str(input("Choose the m/z ratio: "))
    if a not in masses:
        print("This value was not measured. Please choose again: ")
    else:
        break

b = input("Choose the cycle (number or \"all\"): ")

#pandas takes a given interval of measurements based on user cycle selection
if b == "all":
    cycle_value = df
elif b == 1:
    b = int(b)
    cycle_value = df.loc[1:1850, :]
else:
    b = int(b)
    start_point = 1850 + (b - 1)*460
    end_point = start_point + 460
    cycle_value = df.loc[start_point:end_point, :]

#pandas reads the columns assigned to the m/z, which the user selected and plots the data
if a in masses:
    index_a = masses.index(a)
    print(index_a)
    if index_a != 0:
        nx = ".{}".format(index_a)
    else:
        nx = ""
        
value = cycle_value[["Time{}".format(nx),"Ion Current [A]{}".format(nx)]]
ax = value.plot(x = "Time{}".format(nx), y = "Ion Current [A]{}".format(nx))

print(cycle_value)

time_list = df['Time'].to_list()
stripped_time_list = [i.split(" ") for i in time_list] 

PM_list = list()
AM_list = list()

for element in stripped_time_list:
    if element[-1] == "PM":      
        PM_list.append(element[1].split(":"))
        PM_list.append("PM")
        
    elif element[-1] == "AM":
        AM_list.append(element[1].split(":"))
        AM_list.append("PM")


plt.savefig("MSdata.png", dpi = 1000)
plt.show()


#the graph is output to an Excel file
wb = openpyxl.Workbook()
ws = wb.worksheets[0]
img = openpyxl.drawing.image.Image("MSdata.png")
img.anchor = "A1"
ws.add_image(img)
wb.save("MSdata.xlsx")
   



