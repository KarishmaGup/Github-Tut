import pandas as pd
import matplotlib.pyplot as plt
import openpyxl


#reading data from the text file and making 2 separate dataframes
workbook = "test3_1.txt"
df = pd.read_table(workbook, skiprows = 69, decimal=",")
df2 = pd.read_table(workbook, skiprows = 13, sep = "delimiter", engine = "python")
measured_quantities = ["time/s", "control/V", "Ewe/V", "<I>/mA", "cycle number", "P/W", "(Q-Qo)/C"]
time_list = list()

#making a list of the absolute time of measuring start
for column in df2.columns:
    time_list.append(column)


time_list = [i.split(" ")[-1] for i in time_list] 
stripped_time = [i.split(":") for i in time_list]

EC_AM_list = list()
EC_PM_list = list()

if int(stripped_time[0][0]) > 12:
    stripped_time[0][0] = str(int(stripped_time[0][0]) - 12)
    # PM_sum = ":".join(stripped_time[0])  #alternative for having the time in one chunk               
    EC_PM_list.append(stripped_time[0])
    EC_PM_list.append("PM")
    print(EC_PM_list)

else:
    # AM_sum = ":".join(stripped_time[0])    #alternative for having the time in one chunk        
    EC_AM_list.append(stripped_time[0])
    EC_AM_list.append("AM")


#user selects the x and y axes:
while True:
    x_axis = str(input("Choose x axis from the following: time/s, control/V, Ewe/V, <I>/mA, cycle number, P/W: "))
    if x_axis not in measured_quantities:
        print("\nInvalid input")
    else:
        break

while True:
    y_axis = str(input("Choose y axis from the following: time/s, control/V, Ewe/V, <I>/mA, cycle number, P/W: "))
    if y_axis not in measured_quantities:
        print("\nInvalid input")
    else:
        break
 
#the user inputs desired cycle or cycle range
while True:
    try:    
        cycle_range = str(input("Select the first cycle or \"all\": "))
        cycle_range_list = [x.strip() for x in cycle_range.split(',')]
        
        if cycle_range == "all":
            cycle_value = df
            break
    
        else:
            sp = int(cycle_range_list[0])
            ep = int(cycle_range_list[-1])

            if ep >= sp:
                cycle_value = df[(df["cycle number"]>= sp) & (df["cycle number"]<= ep)]
                break
            
    except ValueError:
        print("\nInvalid input")
        

#the graph is plotted and output into Excel
value = cycle_value[[x_axis, y_axis]]  
ax = value.plot(x = x_axis, y = y_axis)
print(value)

plt.savefig("ECdata.png", dpi = 1000)
plt.show()

wb = openpyxl.Workbook()
ws = wb.worksheets[0]
img = openpyxl.drawing.image.Image("ECdata.png")
img.anchor = "A1"
ws.add_image(img)
wb.save("ECdata.xlsx")

print(df)
