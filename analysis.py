import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Reading the database
data = pd.read_csv(r'C:\Users\rodri\Desktop\CloudWalk\logistics-case-v3.csv')

# Renaming columns
data = data.rename(columns={'Created At': 'Creation', 'In Transit To Local Distribution At': 'To_Local','Local Distribution At': 'At_Local',
                        'In Transit To Deliver At': 'To_Deliver','Delivered At': 'At_Deliver', 'Delivery Estimate Date': 'Deliver_ET',
                        'Delivery Addresses → City': 'Deliver_City', 'Delivery Addresses → State': 'Deliver_State'})
# Formating datas
data["Creation"] = data["Creation"].astype(str).str[:-4]
data["Creation"] = data["Creation"].str.replace('T', ' ')
data["To_Local"] = data["To_Local"].str.replace('T', ' ')
data["At_Local"] = data["At_Local"].str.replace('T', ' ')
data["To_Deliver"] = data["To_Deliver"].str.replace('T', ' ')
data["At_Deliver"] = data["At_Deliver"].str.replace('T', ' ')

# Calculating days
data.insert(9, "Days_To_Local", 0)
data.Days_To_Local = (((pd.to_datetime(data.At_Local) - pd.to_datetime(data.To_Local)).dt.days))
data.insert(10, "Days_To_Deliver", 0)
data.Days_To_Deliver = (((pd.to_datetime(data.At_Deliver) - pd.to_datetime(data.To_Deliver)).dt.days))
data.insert(11, "Real_Time", 0)
data.Real_Time = ((pd.to_datetime(data.At_Deliver) - pd.to_datetime(data.Creation)).dt.days)
data.insert(12, "Estimated_Time", 0)
data.Estimated_Time = (((pd.to_datetime(data.Deliver_ET) - pd.to_datetime(data.Creation)).dt.days))
data.insert(13, "Diff", 0)
data.Diff = data.Real_Time - data.Estimated_Time
data = data.drop(data.columns[[1,2,3,4,5,6]], axis = 1)

data_AC = data.query('Deliver_State == "AC"')
data_AL = data.query('Deliver_State == "AL"')
data_AM = data.query('Deliver_State == "AM"')
data_AP = data.query('Deliver_State == "AP"')
data_BA = data.query('Deliver_State == "BA"')
data_CE = data.query('Deliver_State == "CE"')
data_DF = data.query('Deliver_State == "DF"')
data_ES = data.query('Deliver_State == "ES"')
data_GO = data.query('Deliver_State == "GO"')
data_MA = data.query('Deliver_State == "MA"')
data_MG = data.query('Deliver_State == "MG"')
data_MS = data.query('Deliver_State == "MS"')
data_MT = data.query('Deliver_State == "MT"')
data_PA = data.query('Deliver_State == "PA"')
data_PB = data.query('Deliver_State == "PB"')
data_PE = data.query('Deliver_State == "PE"')
data_PI = data.query('Deliver_State == "PI"')
data_PR = data.query('Deliver_State == "PR"')
data_RJ = data.query('Deliver_State == "RJ"')
data_RN = data.query('Deliver_State == "RN"')
data_RO = data.query('Deliver_State == "RO"')
data_RR = data.query('Deliver_State == "RR"')
data_RS = data.query('Deliver_State == "RS"')
data_SC = data.query('Deliver_State == "SC"')
data_SE = data.query('Deliver_State == "SE"')
data_SP = data.query('Deliver_State == "SP"')
data_TO = data.query('Deliver_State == "TO"')

writer = pd.ExcelWriter(r'C:\Users\rodri\Desktop\CloudWalk\CloudWalk.xlsx', engine='xlsxwriter')
data_AC.to_excel(writer, sheet_name = 'AC')
data_AL.to_excel(writer, sheet_name = 'AL')
data_AM.to_excel(writer, sheet_name = 'AM')
data_AP.to_excel(writer, sheet_name = 'AP')
data_BA.to_excel(writer, sheet_name = 'BA')
data_CE.to_excel(writer, sheet_name = 'CE')
data_DF.to_excel(writer, sheet_name = 'DF')
data_ES.to_excel(writer, sheet_name = 'ES')
data_GO.to_excel(writer, sheet_name = 'GO')
data_MA.to_excel(writer, sheet_name = 'MA')
data_MG.to_excel(writer, sheet_name = 'MG')
data_MS.to_excel(writer, sheet_name = 'MS')
data_MT.to_excel(writer, sheet_name = 'MT')
data_PA.to_excel(writer, sheet_name = 'PA')
data_PB.to_excel(writer, sheet_name = 'PB')
data_PE.to_excel(writer, sheet_name = 'PE')
data_PI.to_excel(writer, sheet_name = 'PI')
data_PR.to_excel(writer, sheet_name = 'PR')
data_RJ.to_excel(writer, sheet_name = 'RJ')
data_RN.to_excel(writer, sheet_name = 'RN')
data_RO.to_excel(writer, sheet_name = 'RO')
data_RR.to_excel(writer, sheet_name = 'RR')
data_RS.to_excel(writer, sheet_name = 'RS')
data_SE.to_excel(writer, sheet_name = 'SE')
data_SC.to_excel(writer, sheet_name = 'SC')
data_SP.to_excel(writer, sheet_name = 'SP')
data_TO.to_excel(writer, sheet_name = 'TO')

writer.save()
writer.close()