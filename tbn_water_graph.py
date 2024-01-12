import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import datetime as datetime


def wrangle(filepath):
   # read csv file into dataframe
   df = pd.read_excel(filepath)

   # drop columns "Unnamed: 1", "Unnamed: 2", "Unnamed: 3", "Unnamed: 4" and "index = 0"
   df.drop(columns = ["Unnamed: 1", "Unnamed: 2", "Unnamed: 3", "Unnamed: 4"], inplace = True)
   df.drop(index = [0, 1, 2, 3], inplace = True)

   # reset index and drop old index
   df.reset_index(inplace = True)
   df.drop(columns = "index", inplace = True)

   # assign new column  
   new_header = df.iloc[0]
   df = pd.DataFrame(df.values[1:], columns = new_header).set_index("ENGINE NO")

   # select from dataframe only engines and their tbn values
   df = df.iloc[::2] 

   return df  

df = wrangle(r"C:\Users\HP\Desktop\Engine Lube Oil Analysis Report - 2023.xlsx")
df.head()

# selecct only gas engines from aksa1 and aksa2
ph1_gas_dgs = df[7:14]
ph3_gas_dgs = df[17:19]
all_gas_dgs = pd.concat([ph1_gas_dgs, ph3_gas_dgs])

# select only cells with values and ignore "NaN" values in the last column and put it in a DataFrame
gas_nonna_values = all_gas_dgs.iloc[:, -1][all_gas_dgs.iloc[:, -1].notna()].astype(float)
b = pd.DataFrame(gas_nonna_values)

# Create a new columns named "gas_new_oil_tbn" and reset index
b["gas_new_oil_tbn"] = (b.iloc[:, -1]/b.iloc[:, -1]) * 12
b = b.reset_index()
b.head()

gas_date = b.columns[-2].strftime("%d-%b-%y")

# code out hfo engines
ph1_hfo_dgs = df[:7]
ph3_hfo_dgs1 = df[14:17]
ph3_hfo_dgs2 = df[19:22]
all_hfo_dgs = pd.concat([ph1_hfo_dgs, ph3_hfo_dgs1, ph3_hfo_dgs2])

# select only cells with values and ignore "NaN" values in the last column and put it in a DataFrame
hfo_nonna_values = (all_hfo_dgs.iloc[:, -1][all_hfo_dgs.iloc[:, -1].notna()]).astype(float)
a = pd.DataFrame(hfo_nonna_values)

# Create a new columns named "hfo_new_oil_tbn" and reset index
a["hfo_new_oil_tbn"] = (a.iloc[:, -1]/a.iloc[:, -1]) * 30
a = a.reset_index()

hfo_date = a.columns[-2].strftime("%d-%b-%y")

# GRAPH FOR GAS ENGINES
fig, ax = plt.subplots(figsize=(10, 4.5))
# set deterioration limit
deteriorating_limit = 12 * 0.5
plt.axhline(
    deteriorating_limit, linestyle = "--",
    color = "red", label = "Oil Deterioration Limit")

# Set the width of each bar group and the positions of bars
bar_width = 0.4

# Create a bar graph of the "gas_new_oil_tbn"
plt.bar(
    b["ENGINE NO"], b["gas_new_oil_tbn"],
    bar_width, label= " New Oil TBN",
    align='center', color = "orange", edgecolor="orange")    

# Create a bar graph of the last column of TBN data
plt.bar(
    b["ENGINE NO"], b.iloc[:, -2],
        bar_width, label= f"Used Oil TBN [{gas_date}]",
    align='edge', edgecolor='C2', linewidth=1, color = "C2"
        )

# annotate tbn values of used oil to the top of the bar chart
for i, value in enumerate(b.iloc[:, -2]):
    plt.text(i + bar_width/10, value/1.1, str(value), color = "white", weight = "bold", fontsize = 8,
            )
    
# Add labels, title, and legend
plt.xlabel("Engine Number")
plt.ylabel('TBN Value [mgKOH/g]')
plt.title("Distribution of Lube Oil TBN For Online Gas Powered Engines", weight = "bold")
plt.legend(loc = "lower left", fontsize = 8);

# GRAPH FOR HFO ENGINES
fig, ax = plt.subplots(figsize=(10.5, 4.5))

deteriorating_limit = 30 * 0.5
plt.axhline(deteriorating_limit, linestyle = "--", color = "b", label = "Oil Deterioration Limit")

# Set the width of each bar group and the positions of bars
bar_width = 0.4

# Create the bar graph with double bars
plt.bar(
    a["ENGINE NO"], a["hfo_new_oil_tbn"],
    bar_width, label= " New Oil TBN",
    align='center', color = "orange"
)

# Create a bar graph of the last column of TBN data
plt.bar(
    a["ENGINE NO"], a.iloc[:, -2],
        bar_width, label= f"Used Oil TBN [{hfo_date}]",
    align='edge', edgecolor="brown",
    linewidth=1, color = "brown"
        )

# annotate tbn values of used oil to the top of the bar chart
for i, value in enumerate(a.iloc[:, -2]):
    plt.text(i + bar_width/15, value/1.1,
             str(value), color = "white", weight = "bold", fontsize = 6.8)
    
# Add labels, title, and legend
plt.xlabel("Engine Number")
plt.ylabel('TBN Value [mgKOH/g]')
plt.title("Distribution of Lube Oil TBN For Online HFO Powered Engines", weight = "bold")
plt.legend(loc = "lower left", fontsize = 8);

def wrangle(filepath):
    # read csv file into dataframe
    df1 = pd.read_excel(filepath)
    
    # drop columns "Unnamed: 1", "Unnamed: 2", "Unnamed: 3", "Unnamed: 4" and "index = 0"
    df1.drop(columns = ["Unnamed: 1", "Unnamed: 2", "Unnamed: 3", "Unnamed: 4"], inplace = True)
    df1.drop(index = [0, 1, 2, 3], inplace = True)
    
    # reset index and drop old index
    df1.reset_index(inplace = True)
    df1.drop(columns = "index", inplace = True)
    
#     # assign new column  
    new_header = df1.iloc[0]
    df1 = pd.DataFrame(df1.values[1:], columns = new_header)
  
    # select from dataframe only engines and their %water values
    df1 = df1.iloc[1::2] 
   
    # add DG numbers for dataframe
    df1["DG NO"] = df.index
    
    df1 = df1.set_index("DG NO").drop(columns = "ENGINE NO")
    water_non_df_non_na_values = (df1.iloc[:, -1][df1.iloc[:, -1].notna()]).astype(float)
    df1 = pd.DataFrame(water_non_df_non_na_values)
    return df1


df1 = wrangle(r"C:\Users\HP\Desktop\Engine Lube Oil Analysis Report - 2023.xlsx")
df1.head()


water_date = df1.columns[-1].strftime("%d-%b-%y")


# Extract last column of df1 above
last_water_results = df1.iloc[:, -1]
last_water_results.head()


def wrangle(filepath):
    df = pd.read_excel(filepath, sheet_name = "Viscosity")
    
    # drop ENGINE LUBE OIL DAILY VISCOSITY ANALYSIS REPORT FOR AKSA ENERGY GHANA - 2023 column
    df.drop(
            columns = "ENGINE LUBE OIL DAILY VISCOSITY ANALYSIS REPORT FOR AKSA ENERGY GHANA - 2023",
        inplace = True
    )
    
    # rename index "1" in column "Unnamed: 0" to DG_No
    new_label = "DG_No"
    df.at[1, "Unnamed: 0"] = new_label  
    
    # drop irrelevant index namely "0, 2" and drop columns Unnamed: 1", "Unnamed: 3", "Unnamed: 4"
    df.drop(index = [0, 2], inplace = True)
    df.drop(columns = ["Unnamed: 1", "Unnamed: 3", "Unnamed: 4"], inplace = True)
    
    # Assign columns with date as new header
    new_header = df.iloc[0]
    df = pd.DataFrame(df.values[1:], columns = new_header)
    
    # set DG_No as index
    df.set_index("DG_No", inplace = True)  
    
    viscosity_non_df_non_na_values = (df.iloc[:, -1][df.iloc[:, -1].notna()]).astype(float)
    df = pd.DataFrame(viscosity_non_df_non_na_values)
    
    return df  


df = wrangle(r"C:\Users\HP\Desktop\Engine Lube Oil Analysis Report - 2023.xlsx")
df.tail()


hfo_date = df.columns[-1].strftime("%d-%b-%y")


# Extract last column of df above
last_vis_results = df.iloc[:, -1]
last_vis_results.head()


# Create subplots ax and ax2
fig, (ax, ax2) = plt.subplots(figsize = (9.0, 5.8), ncols=2, sharey=True, width_ratios = [1, 1.8])

# set lower and upper limits for viscosity grapgh
upper_limit = 130*1.4
lower_limit = 130 - 130*0.25
plt.axvline(upper_limit, linestyle = "--", color = "r", label = "Upper Viscosity Limit")
plt.axvline(lower_limit, linestyle = "--", color = "b", label = "Lower Viscosity Limit")

# invert x-axis of %water graph so that bars can face left
ax.invert_xaxis()

# set shared y-axis to the right side of %water graph
ax.yaxis.tick_right()

# set upper limit for %water data
upper_limit = 0.15
plt.ax=ax.axvline(
    upper_limit, linestyle = "--", color = "brown", label = "Water Content in Oil Limit"
)
    
# plot "water in oil" graph
last_water_results.plot(
    kind='barh',
    ax=ax,
    color = "orange",
    xlabel = "Water in Oil [%]",
    label = f"%Water in Oil [{water_date}]",
)

# place viscosity values on top of bars for %water graph
for index, value in enumerate(last_water_results):
    plt.ax=ax.text(
        value, index, str(value), ha='right',
        va='center', color = "magenta", weight = "bold", fontsize = 6)
    
# plot "viscosity" graph
last_vis_results.plot(
    kind='barh', ax=ax2,
    color = "g",
    xlabel = "Viscosity @ 40deg [cSt]",
    label = f"Used Oil Viscosity [{hfo_date}]"
)

# place viscosity values on top of bars for viscosity graph
for index, value in enumerate(last_vis_results):
    plt.text(
        value, index, str(value), ha='right',
        va='center', color = "white", weight = "bold", fontsize = 7.0)
    
# show legends for both graphs 
plt.ax=ax2.legend(loc = "upper left", fontsize = 7)
plt.ax=ax.legend(loc = "upper left", fontsize = 7)

# set title for the whole figure, and position title just above figure by setting y = 0.92
fig.suptitle(
    "Distribution of Lube Oil Viscosity and %Water Values For Online Engine ",
    weight = "bold",
    fontsize = 10, y = 0.92
)

# remove "DG NO" label on the y axis
ax.set_ylabel("")

plt.show()