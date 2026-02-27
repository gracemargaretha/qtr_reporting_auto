import pandas as pd

#base year to calculate MLF and Curt FY
base_year = 2021

#store workbook
wb_IT = "Q1 26 Model update issues tracker Baringa RC.xlsx"

#open worksheet
ws_gll_list = pd.read_excel(wb_IT,sheet_name="GLL List")

#filter generators included quarterly report
genList = ws_gll_list[ws_gll_list["Included in reporting?"] == "Included"].copy()

#keep only relevant columns
cols_keep = ["Name", "BusNum", "GenID", "Tech", "Location", "COD"]

genList = genList[cols_keep].copy()

#append Name to add SF/WF/BESS
genList["Name"] = genList.apply(
    lambda row: (
        row["Name"] + " SF"
        if str(row["Tech"]).lower() == "solar"
        else row["Name"] + " WF"
        if "wind" in str(row["Tech"]).lower()
        else row["Name"]
    ),
    axis=1
)

#sort genList based on Location and alphabetically

genList = genList.sort_values(["Location","Name"])

name_list = genList["Name"].tolist()
tech_list = genList["Tech"].tolist()
state_list = genList["Location"].tolist()

#create unique IDs for generators
genList["UniqueID"] = genList["BusNum"].astype(int).astype(str) + "_" + genList["GenID"].astype(int).astype(str)

