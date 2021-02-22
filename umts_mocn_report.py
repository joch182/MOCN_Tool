from tkinter import Menu
from tkinter import *
from tkinter.ttk import * 
from tkinter import filedialog, Menu
import glob
from tkinter import messagebox, ttk
import numpy as np
import pandas as pd
import jinja2
import xlsxwriter

'''
First create the functionality for the buttons of the GUI
'''

def input_umts_daily_report_button():
    # Ask for CSV data of template UMTS daily report
    global input_umts_daily_report_data
    input_umts_daily_report_data = filedialog.askopenfilename()

def input_umts_daily_report1_button():
    # Ask for CSV data of template UMTS daily report1
    global input_umts_daily_report1_data
    input_umts_daily_report1_data = filedialog.askopenfilename()

def input_umts_mocn_button():
    # Ask for CSV data of template UMTS MOCN
    global input_umts_mocn_data
    input_umts_mocn_data = filedialog.askopenfilename()

def output_button():
    global output_data
    output_data = filedialog.askdirectory()

# UMTS raw counters template download
def download_umts_raw_counters_template():
    global umts_template
    umts_template = filedialog.askdirectory()
    jinja_env = jinja2.Environment(loader=jinja2.FileSystemLoader('assets'))
    template = jinja_env.get_template('UMTS_KPI_Template.xml')
    render = template.render()
    with open(umts_template+"/UMTS_KPI_Template.xml", "w") as file:
        file.write(render)
    messagebox.showinfo(message="UMTS U2000 templates saved at: "+umts_template, title="Finished")  
    umts_template = ""
    return

# UMTS MOCN Report generation button handler
def generate_umts_mocn_delivery_report_button():
    global input_umts_daily_report_data
    global input_umts_daily_report1_data
    global input_umts_mocn_data
    global output_data
    if input_umts_daily_report_data != "" and input_umts_daily_report1_data != "" and input_umts_mocn_data != "" and output_data != "":
        filename1 = glob.glob(input_umts_daily_report_data)
        filename2 = glob.glob(input_umts_daily_report1_data)
        filename3 = glob.glob(input_umts_mocn_data)
        # Perform the calcuation and generate the excel report
        UMTS_MOCN_REPORT(filename1, filename2, filename3)
        messagebox.showinfo(message="File saved at: "+output_data, title="Finished")
        input_umts_daily_report_data = ""
        input_umts_daily_report1_data = ""
        input_umts_mocn_data = ""
        output_data = ""
    else:
        messagebox.showinfo(message="Please input correct information", title="Error")

'''
Following perform the calculations required
'''
def UMTS_MOCN_REPORT(umts_daily, umts_daily2, umts_mocn):
    umts_daily_df = pd.read_csv(umts_daily[0],skiprows=7)
    umts_daily2_df = pd.read_csv(umts_daily2[0],skiprows=7)
    umts_mocn_df = pd.read_csv(umts_mocn[0],skiprows=7)
    # Format the cellname and drop useless columns
    umts_daily_df = format_cellname_sitename(umts_daily_df)
    umts_daily2_df = format_cellname_sitename(umts_daily2_df)
    umts_mocn_df = format_cellname_sitename(umts_mocn_df)
    # Set the index only to prepare the initial columns time, site, cellname
    umts_daily_df = set_index(umts_daily_df)
    umts_daily2_df = set_index(umts_daily2_df)
    umts_mocn_df = set_index(umts_mocn_df)
    # Merge the 2 DF from daily data
    umts_daily_join_df = pd.merge(umts_daily_df, umts_daily2_df, left_index=True, right_index=True)
    # Clear memory for useless df
    del umts_daily_df
    del umts_daily2_df
    # Remove the index
    umts_daily_join_df = umts_daily_join_df.reset_index()
    umts_mocn_df = umts_mocn_df.reset_index()

    # Required to aggregate the same column with count and sum
    umts_daily_join_df['VS.Cell.UnavailTime.Sys (s) count'] = umts_daily_join_df['VS.Cell.UnavailTime.Sys (s)']

    # Generate the data for each sheet
    df_day_att_tef = calculate_umts_kpi(umts_daily_join_df, ["Date"])
    df_cell_att_tef = calculate_umts_kpi(umts_daily_join_df, ["Date","Site","Cell","CellID"])
    df_hour_att_tef = calculate_umts_kpi(umts_daily_join_df, ["Date","Hour"])
    df_cellhour_att_tef = calculate_umts_kpi(umts_daily_join_df, ["Date","Hour", "Cell", "CellID"])
    del umts_daily_join_df
    df_day_att = calculate_mocn(umts_mocn_df, ["Date"], 'ATT')
    df_day_tef = calculate_mocn(umts_mocn_df, ["Date"], 'TEF')
    df_cell_att = calculate_mocn(umts_mocn_df, ["Date","Site","Cell","CellID"], 'ATT')
    df_cell_tef = calculate_mocn(umts_mocn_df, ["Date","Site","Cell","CellID"], 'TEF')
    df_hour_att = calculate_mocn(umts_mocn_df, ["Date","Hour"], 'ATT')
    df_hour_tef = calculate_mocn(umts_mocn_df, ["Date","Hour"], 'TEF')
    df_cellhour_att = calculate_mocn(umts_mocn_df, ["Date","Hour", "Cell", "CellID"], 'ATT')
    df_cellhour_tef = calculate_mocn(umts_mocn_df, ["Date","Hour", "Cell", "CellID"], 'TEF')
    del umts_mocn_df

    # Initiate the excel write process
    global writer
    global workbook
    writer = pd.ExcelWriter(output_data+'/results.xlsx',engine='xlsxwriter')
    workbook=writer.book

    # Generate the summary sheet
    generate_summary_sheet(df_day_att_tef, df_day_att, df_day_tef)
    
    # Generate sheets for raw data
    generate_excel_sheets('Day(AT&T+TLF)', df_day_att_tef)
    del df_day_att_tef
    generate_excel_sheets('Day(AT&T)', df_day_att)
    del df_day_att
    generate_excel_sheets('Day(TLF)', df_day_tef)
    del df_day_tef
    generate_excel_sheets('Cell(AT&T+TLF)', df_cell_att_tef)
    del df_cell_att_tef
    generate_excel_sheets('Cell(AT&T)', df_cell_att)
    del df_cell_att
    generate_excel_sheets('Cell(TLF)', df_cell_tef)
    del df_cell_tef
    generate_excel_sheets('Hour(AT&T+TLF)', df_hour_att_tef)
    del df_hour_att_tef
    generate_excel_sheets('Hour(AT&T)', df_hour_att)
    del df_hour_att
    generate_excel_sheets('Hour(TLF)', df_hour_tef)
    del df_hour_tef
    generate_excel_sheets('Cell Hour(AT&T+TLF)', df_cellhour_att_tef)
    del df_cellhour_att_tef
    generate_excel_sheets('Cell Hour(AT&T)', df_cellhour_att)
    del df_cellhour_att
    generate_excel_sheets('Cell Hour(TLF)', df_cellhour_tef)
    del df_cellhour_tef

    # Save the excel file
    writer.save()
    return

# Format for time, site, cell name and cellid
def format_cellname_sitename(df):
    if df.columns.tolist()[3] == "BSC6910UCell":
        df['CellName'] = df['BSC6910UCell'].str[6:]
        df = df.drop(columns=['Period', 'NE Name', 'BSC6910UCell'])
    else:
        df['CellName'] = df['BSC6900UCell'].str[6:]
        df = df.drop(columns=['Period', 'NE Name', 'BSC6900UCell'])

    df['CellFilter1'] = df['CellName'].str.split(expand=True)[0]
    df['CellID1'] = df['CellName'].str.split(expand=True)[1]

    df['CellFilter2'] = df['CellFilter1'].str.split(pat="=", expand=True)[0]
    df['CellID2'] = df['CellID1'].str.split(pat="=", expand=True)[1]

    df['Cell'] = df['CellFilter2'].str.split(pat=",", expand=True)[0]
    df['CellID'] = df['CellID2'].str.split(pat=",", expand=True)[0]

    df['Site'] = df['Cell'].str.split(pat="_", expand=True)[0]
    del df['CellFilter1']
    del df['CellFilter2']
    del df['CellID1']
    del df['CellID2']
    del df['CellName']

    df['Start_Time'] = pd.to_datetime(df['Start Time']).dt.strftime('%Y-%m-%d %H:00')
    df['Date'] = pd.to_datetime(df['Start Time']).dt.strftime('%Y-%m-%d')
    df['Hour'] = pd.to_datetime(df['Start Time']).dt.strftime('%H:00')
    del df['Start Time']

    return df

# Define multi index for start time, site, cellname and cellid
def set_index(df):
    df = df.set_index(['Start_Time', 'Date', 'Hour', 'Site', 'Cell', 'CellID'])
    return df

# Function to calculate MOCN KPI
def calculate_mocn(df, groupby_list, network):
    df_grouped = df.groupby(by=groupby_list).sum()
    if network == 'ATT':
        df_grouped['Traffic DL Volume (GB)'] =  round((df_grouped['VS.PSLoad.DLThruput.MOCN.PLMN0 (byte)'] + df_grouped['VS.PSLoad.DLThruput.MOCN.PLMN1 (byte)'])/(8*1024*1024*1024),2)
        df_grouped['Traffic UL Volume (GB)'] =  round((df_grouped['VS.PSLoad.ULThruput.MOCN.PLMN0 (byte)'] + df_grouped['VS.PSLoad.ULThruput.MOCN.PLMN1 (byte)'])/(8*1024*1024*1024),2)
        df_grouped['Voice Erlangs'] =  round((df_grouped['VS.CS.Erlang.Equiv.MOCN.PLMN0 (Erl)'] + df_grouped['VS.CS.Erlang.Equiv.MOCN.PLMN1 (Erl)'])*(30/60),2)
        df_grouped['Total MOCN Traffic(GB)'] =  df_grouped['Traffic DL Volume (GB)'] + df_grouped['Traffic UL Volume (GB)']
        df_grouped.drop(df_grouped.columns.difference(['Traffic DL Volume (GB)','Traffic UL Volume (GB)','Voice Erlangs','Total MOCN Traffic(GB)']), 1, inplace=True)
    elif network == 'TEF':
        df_grouped['Traffic DL Volume (GB)'] =  round((df_grouped['VS.PSLoad.DLThruput.MOCN.PLMN2 (byte)'])/(8*1024*1024*1024),2)
        df_grouped['Traffic UL Volume (GB)'] =  round((df_grouped['VS.PSLoad.ULThruput.MOCN.PLMN2 (byte)'])/(8*1024*1024*1024),2)
        df_grouped['Voice Erlangs'] =  round((df_grouped['VS.CS.Erlang.Equiv.MOCN.PLMN2 (Erl)'])*(30/60),2)
        df_grouped['Total MOCN Traffic(GB)'] =  df_grouped['Traffic DL Volume (GB)'] + df_grouped['Traffic UL Volume (GB)']
        df_grouped.drop(df_grouped.columns.difference(['Traffic DL Volume (GB)','Traffic UL Volume (GB)','Voice Erlangs','Total MOCN Traffic(GB)']), 1, inplace=True)
    return df_grouped

# Function to calculate UMTS KPI not MOCN
def calculate_umts_kpi(df, groupby_list):
    df = df.replace('NIL', 0)
    df['VS.HSDPA.MeanChThroughput (kbit/s)'] = pd.to_numeric(df['VS.HSDPA.MeanChThroughput (kbit/s)'])
    df['VS.HSUPA.MeanChThroughput (kbit/s)'] = pd.to_numeric(df['VS.HSUPA.MeanChThroughput (kbit/s)'])
    df_grouped = df.groupby(by=groupby_list).agg({
        'VS.AMR.Erlang.BestCell (None)': 'sum',
        'VS.VP.Erlang.BestCell (None)': 'sum',
        'VS.HSDPA.MeanChThroughput.TotalBytes (byte)': 'sum',
        'VS.SRNCIubBytesPSFACH.Tx (byte)': 'sum',
        'VS.SRNCIubBytesPSEFACH.Tx (byte)': 'sum',
        'VS.PS.Bkg.DL.8.Traffic (bit)': 'sum',
        'VS.PS.Bkg.DL.16.Traffic (bit)': 'sum',
        'VS.PS.Bkg.DL.32.Traffic (bit)': 'sum',
        'VS.PS.Bkg.DL.64.Traffic (bit)': 'sum',
        'VS.PS.Bkg.DL.128.Traffic (bit)': 'sum',
        'VS.PS.Bkg.DL.144.Traffic (bit)': 'sum',
        'VS.PS.Bkg.DL.256.Traffic (bit)': 'sum',
        'VS.PS.Bkg.DL.384.Traffic (bit)': 'sum',
        'VS.PS.Int.DL.8.Traffic (bit)': 'sum',
        'VS.PS.Int.DL.16.Traffic (bit)': 'sum',
        'VS.PS.Int.DL.32.Traffic (bit)': 'sum',
        'VS.PS.Int.DL.64.Traffic (bit)': 'sum',
        'VS.PS.Int.DL.128.Traffic (bit)': 'sum',
        'VS.PS.Int.DL.144.Traffic (bit)': 'sum',
        'VS.PS.Int.DL.256.Traffic (bit)': 'sum',
        'VS.PS.Int.DL.384.Traffic (bit)': 'sum',
        'VS.PS.Str.DL.8.Traffic (bit)': 'sum',
        'VS.PS.Str.DL.16.Traffic (bit)': 'sum',
        'VS.PS.Str.DL.32.Traffic (bit)': 'sum',
        'VS.PS.Str.DL.64.Traffic (bit)': 'sum',
        'VS.PS.Str.DL.128.Traffic (bit)': 'sum',
        'VS.PS.Str.DL.144.Traffic (bit)': 'sum',
        'VS.PS.Str.DL.256.Traffic (bit)': 'sum',
        'VS.PS.Str.DL.384.Traffic (bit)': 'sum',
        'VS.PS.Conv.DL.Traffic (bit)': 'sum',
        'VS.DcchSRB.Dl.Traffic (bit)': 'sum',
        'VS.HSUPA.MeanChThroughput.TotalBytes (byte)': 'sum',
        'VS.SRNCIubBytesPSRACH.Rx (byte)': 'sum',
        'VS.SRNCIubBytesPSERACH.Rx (byte)': 'sum',
        'VS.PS.Bkg.UL.8.Traffic (bit)': 'sum',
        'VS.PS.Bkg.UL.16.Traffic (bit)': 'sum',
        'VS.PS.Bkg.UL.32.Traffic (bit)': 'sum',
        'VS.PS.Bkg.UL.64.Traffic (bit)': 'sum',
        'VS.PS.Bkg.UL.128.Traffic (bit)': 'sum',
        'VS.PS.Bkg.UL.144.Traffic (bit)': 'sum',
        'VS.PS.Bkg.UL.256.Traffic (bit)': 'sum',
        'VS.PS.Bkg.UL.384.Traffic (bit)': 'sum',
        'VS.PS.Int.UL.8.Traffic (bit)': 'sum',
        'VS.PS.Int.UL.16.Traffic (bit)': 'sum',
        'VS.PS.Int.UL.32.Traffic (bit)': 'sum',
        'VS.PS.Int.UL.64.Traffic (bit)': 'sum',
        'VS.PS.Int.UL.128.Traffic (bit)': 'sum',
        'VS.PS.Int.UL.144.Traffic (bit)': 'sum',
        'VS.PS.Int.UL.256.Traffic (bit)': 'sum',
        'VS.PS.Int.UL.384.Traffic (bit)': 'sum',
        'VS.PS.Str.UL.8.Traffic (bit)': 'sum',
        'VS.PS.Str.UL.16.Traffic (bit)': 'sum',
        'VS.PS.Str.UL.32.Traffic (bit)': 'sum',
        'VS.PS.Str.UL.64.Traffic (bit)': 'sum',
        'VS.PS.Str.UL.128.Traffic (bit)': 'sum',
        'VS.PS.Conv.UL.Traffic (bit)': 'sum',
        'VS.DcchSRB.Ul.Traffic (bit)': 'sum',
        'RRC.SuccConnEstab.OrgConvCall (None)': 'sum',
        'RRC.SuccConnEstab.TmConvCall (None)': 'sum',
        'RRC.SuccConnEstab.EmgCall (None)': 'sum',
        'VS.SuccCellUpdt.OrgConvCall.PCH (None)': 'sum',
        'VS.SuccCellUpdt.EmgCall.PCH (None)': 'sum',
        'VS.SuccCellUpdt.TmConvCall.PCH (None)': 'sum',
        'RRC.AttConnEstab.OrgConvCall (None)': 'sum',
        'RRC.AttConnEstab.TmConvCall (None)': 'sum',
        'RRC.AttConnEstab.EmgCall (None)': 'sum',
        'VS.AttCellUpdt.OrgConvCall.PCH (None)': 'sum',
        'VS.AttCellUpdt.TmConvCall.PCH (None)': 'sum',
        'VS.AttCellUpdt.EmgCall.PCH (None)': 'sum',
        'VS.RAB.SuccEstabCS.Conv (None)': 'sum',
        'VS.RAB.SuccEstabCS.Str (None)': 'sum',
        'VS.RAB.AttEstabCS.Conv (None)': 'sum',
        'VS.RAB.AttEstabCS.Str (None)': 'sum',
        'RRC.SuccConnEstab.OrgBkgCall (None)': 'sum',
        'RRC.SuccConnEstab.TmBkgCall (None)': 'sum',
        'RRC.SuccConnEstab.OrgInterCall (None)': 'sum',
        'RRC.SuccConnEstab.TmItrCall (None)': 'sum',
        'RRC.SuccConnEstab.OrgStrCall (None)': 'sum',
        'RRC.SuccConnEstab.TmStrCall (None)': 'sum',
        'RRC.SuccConnEstab.OrgHhPrSig (None)': 'sum',
        'RRC.SuccConnEstab.TmHhPrSig (None)': 'sum',
        'RRC.SuccConnEstab.OrgLwPrSig (None)': 'sum',
        'RRC.SuccConnEstab.TmLwPrSig (None)': 'sum',
        'RRC.SuccConnEstab.OrgSubCall (None)': 'sum',
        'RRC.SuccConnEstab.Unkown (None)': 'sum',
        'RRC.SuccConnEstab.CallReEst (None)': 'sum',
        'VS.SuccCellUpdt.PageRsp (None)': 'sum',
        'VS.SuccCellUpdt.ULDataTrans (None)': 'sum',
        'RRC.AttConnEstab.OrgBkgCall (None)': 'sum',
        'RRC.AttConnEstab.TmBkgCall (None)': 'sum',
        'RRC.AttConnEstab.OrgInterCall (None)': 'sum',
        'RRC.AttConnEstab.TmInterCall (None)': 'sum',
        'RRC.AttConnEstab.OrgStrCall (None)': 'sum',
        'RRC.AttConnEstab.TmStrCall (None)': 'sum',
        'RRC.AttConnEstab.OrgHhPrSig (None)': 'sum',
        'RRC.AttConnEstab.TmHhPrSig (None)': 'sum',
        'RRC.AttConnEstab.OrgLwPrSig (None)': 'sum',
        'RRC.AttConnEstab.TmLwPrSig (None)': 'sum',
        'RRC.AttConnEstab.OrgSubCall (None)': 'sum',
        'RRC.AttConnEstab.Unknown (None)': 'sum',
        'RRC.AttConnEstab.CallReEst (None)': 'sum',
        'VS.AttCellUpdt.PageRsp (None)': 'sum',
        'VS.AttCellUpdt.ULDataTrans (None)': 'sum',
        'VS.RAB.SuccEstabPS.Conv (None)': 'sum',
        'VS.RAB.SuccEstabPS.Str (None)': 'sum',
        'VS.RAB.SuccEstabPS.Int (None)': 'sum',
        'VS.RAB.SuccEstabPS.Bkg (None)': 'sum',
        'VS.DCCC.Succ.F2D.AfterP2F (None)': 'sum',
        'VS.RAB.AttEstabPS.Conv (None)': 'sum',
        'VS.RAB.AttEstabPS.Str (None)': 'sum',
        'VS.RAB.AttEstabPS.Int (None)': 'sum',
        'VS.RAB.AttEstabPS.Bkg (None)': 'sum',
        'VS.DCCC.Att.F2D.AfterP2F (None)': 'sum',
        'VS.HSDPA.MeanChThroughput (kbit/s)': 'mean',
        'VS.HSUPA.MeanChThroughput (kbit/s)': 'mean',
        'VS.CellDCHUEs (None)': 'mean',
        'VS.Cell.UnavailTime.Sys (s)' : 'sum',
        'VS.Cell.UnavailTime.Sys (s) count': 'count'
    })
    df_grouped['Voice Erlangs'] =  round((df_grouped['VS.AMR.Erlang.BestCell (None)'] + df_grouped['VS.VP.Erlang.BestCell (None)'])*(30/60),2)

    df_grouped['Traffic DL Volume (GB)'] = round(((df_grouped['VS.HSDPA.MeanChThroughput.TotalBytes (byte)'] + df_grouped['VS.SRNCIubBytesPSFACH.Tx (byte)'] + df_grouped['VS.SRNCIubBytesPSEFACH.Tx (byte)'])/(1024*1024*1024)) + ((df_grouped['VS.PS.Bkg.DL.8.Traffic (bit)']+ df_grouped['VS.PS.Bkg.DL.16.Traffic (bit)'] + df_grouped['VS.PS.Bkg.DL.32.Traffic (bit)'] + df_grouped['VS.PS.Bkg.DL.64.Traffic (bit)'] + df_grouped['VS.PS.Bkg.DL.128.Traffic (bit)'] + df_grouped['VS.PS.Bkg.DL.144.Traffic (bit)'] + df_grouped['VS.PS.Bkg.DL.256.Traffic (bit)'] + df_grouped['VS.PS.Bkg.DL.384.Traffic (bit)']+
    df_grouped['VS.PS.Int.DL.8.Traffic (bit)'] + df_grouped['VS.PS.Int.DL.16.Traffic (bit)'] + df_grouped['VS.PS.Int.DL.32.Traffic (bit)'] + df_grouped['VS.PS.Int.DL.64.Traffic (bit)'] +
    df_grouped['VS.PS.Int.DL.128.Traffic (bit)'] + df_grouped['VS.PS.Int.DL.144.Traffic (bit)'] + df_grouped['VS.PS.Int.DL.256.Traffic (bit)'] + df_grouped['VS.PS.Int.DL.384.Traffic (bit)'] + df_grouped['VS.PS.Str.DL.8.Traffic (bit)'] + df_grouped['VS.PS.Str.DL.16.Traffic (bit)'] + df_grouped['VS.PS.Str.DL.32.Traffic (bit)'] + df_grouped['VS.PS.Str.DL.64.Traffic (bit)'] + df_grouped['VS.PS.Str.DL.128.Traffic (bit)'] + df_grouped['VS.PS.Str.DL.144.Traffic (bit)'] + df_grouped['VS.PS.Str.DL.256.Traffic (bit)'] + df_grouped['VS.PS.Str.DL.384.Traffic (bit)'] + df_grouped['VS.PS.Conv.DL.Traffic (bit)']+df_grouped['VS.DcchSRB.Dl.Traffic (bit)'])/(8*1024*1024*1024)),2)
    
    df_grouped['Traffic UL Volume (GB)'] = round(((df_grouped['VS.HSUPA.MeanChThroughput.TotalBytes (byte)'] + df_grouped['VS.SRNCIubBytesPSRACH.Rx (byte)'] + df_grouped['VS.SRNCIubBytesPSERACH.Rx (byte)'])/(1024*1024*1024)) + ((df_grouped['VS.PS.Bkg.UL.8.Traffic (bit)']+ df_grouped['VS.PS.Bkg.UL.16.Traffic (bit)'] + df_grouped['VS.PS.Bkg.UL.32.Traffic (bit)'] + df_grouped['VS.PS.Bkg.UL.64.Traffic (bit)'] + df_grouped['VS.PS.Bkg.UL.128.Traffic (bit)'] + df_grouped['VS.PS.Bkg.UL.144.Traffic (bit)'] + df_grouped['VS.PS.Bkg.UL.256.Traffic (bit)'] + df_grouped['VS.PS.Bkg.UL.384.Traffic (bit)']+
    df_grouped['VS.PS.Int.UL.8.Traffic (bit)'] + df_grouped['VS.PS.Int.UL.16.Traffic (bit)'] + df_grouped['VS.PS.Int.UL.32.Traffic (bit)'] + df_grouped['VS.PS.Int.UL.64.Traffic (bit)'] +
    df_grouped['VS.PS.Int.UL.128.Traffic (bit)'] + df_grouped['VS.PS.Int.UL.144.Traffic (bit)'] + df_grouped['VS.PS.Int.UL.256.Traffic (bit)'] + df_grouped['VS.PS.Int.UL.384.Traffic (bit)'] + df_grouped['VS.PS.Str.UL.8.Traffic (bit)'] + df_grouped['VS.PS.Str.UL.16.Traffic (bit)'] + df_grouped['VS.PS.Str.UL.32.Traffic (bit)'] + df_grouped['VS.PS.Str.UL.64.Traffic (bit)'] + df_grouped['VS.PS.Str.UL.128.Traffic (bit)'] + df_grouped['VS.PS.Conv.UL.Traffic (bit)']+df_grouped['VS.DcchSRB.Ul.Traffic (bit)'])/(8*1024*1024*1024)),2)

    df_grouped['Total MOCN Traffic(GB)'] = df_grouped['Traffic UL Volume (GB)'] + df_grouped['Traffic DL Volume (GB)']

    df_grouped['Accessibility CS (%)'] = round(100 * ((df_grouped['RRC.SuccConnEstab.OrgConvCall (None)'] + df_grouped['RRC.SuccConnEstab.TmConvCall (None)'] + df_grouped['RRC.SuccConnEstab.EmgCall (None)'] + df_grouped['VS.SuccCellUpdt.OrgConvCall.PCH (None)'] + df_grouped['VS.SuccCellUpdt.EmgCall.PCH (None)'] + df_grouped['VS.SuccCellUpdt.TmConvCall.PCH (None)'])/(df_grouped['RRC.AttConnEstab.OrgConvCall (None)']+df_grouped['RRC.AttConnEstab.TmConvCall (None)']+df_grouped['RRC.AttConnEstab.EmgCall (None)']+df_grouped['VS.AttCellUpdt.OrgConvCall.PCH (None)']+df_grouped['VS.AttCellUpdt.TmConvCall.PCH (None)']+df_grouped['VS.AttCellUpdt.EmgCall.PCH (None)'])) * ((df_grouped['VS.RAB.SuccEstabCS.Conv (None)'] + df_grouped['VS.RAB.SuccEstabCS.Str (None)'])/(df_grouped['VS.RAB.AttEstabCS.Conv (None)'] + df_grouped['VS.RAB.AttEstabCS.Str (None)'])),2)

    df_grouped['Accessibility PS (%)'] = round(100 * (((df_grouped['RRC.SuccConnEstab.OrgBkgCall (None)'] + df_grouped['RRC.SuccConnEstab.TmBkgCall (None)'] + df_grouped['RRC.SuccConnEstab.OrgInterCall (None)'] + df_grouped['RRC.SuccConnEstab.TmItrCall (None)'] + df_grouped['RRC.SuccConnEstab.OrgStrCall (None)'] + df_grouped['RRC.SuccConnEstab.TmStrCall (None)'] + df_grouped['RRC.SuccConnEstab.OrgHhPrSig (None)'] + df_grouped['RRC.SuccConnEstab.TmHhPrSig (None)'] + df_grouped['RRC.SuccConnEstab.OrgLwPrSig (None)'] + df_grouped['RRC.SuccConnEstab.TmLwPrSig (None)'] + df_grouped['RRC.SuccConnEstab.OrgSubCall (None)'] + df_grouped['RRC.SuccConnEstab.Unkown (None)'] + df_grouped['RRC.SuccConnEstab.CallReEst (None)'] + df_grouped['VS.SuccCellUpdt.PageRsp (None)'] + df_grouped['VS.SuccCellUpdt.ULDataTrans (None)'])-(df_grouped['VS.SuccCellUpdt.OrgConvCall.PCH (None)'] + df_grouped['VS.SuccCellUpdt.EmgCall.PCH (None)'] + df_grouped['VS.SuccCellUpdt.TmConvCall.PCH (None)']))/(((df_grouped['RRC.AttConnEstab.OrgBkgCall (None)'] + df_grouped['RRC.AttConnEstab.TmBkgCall (None)'] + df_grouped['RRC.AttConnEstab.OrgInterCall (None)'] + df_grouped['RRC.AttConnEstab.TmInterCall (None)'] + df_grouped['RRC.AttConnEstab.OrgStrCall (None)'] + df_grouped['RRC.AttConnEstab.TmStrCall (None)'] + df_grouped['RRC.AttConnEstab.OrgHhPrSig (None)'] + df_grouped['RRC.AttConnEstab.TmHhPrSig (None)'] + df_grouped['RRC.AttConnEstab.OrgLwPrSig (None)'] + df_grouped['RRC.AttConnEstab.TmLwPrSig (None)'] + df_grouped['RRC.AttConnEstab.OrgSubCall (None)'] + df_grouped['RRC.AttConnEstab.Unknown (None)'] + df_grouped['RRC.AttConnEstab.CallReEst (None)'] + df_grouped['VS.AttCellUpdt.PageRsp (None)'] + df_grouped['VS.AttCellUpdt.ULDataTrans (None)'])-(df_grouped['VS.AttCellUpdt.OrgConvCall.PCH (None)'] + df_grouped['VS.AttCellUpdt.TmConvCall.PCH (None)'] + df_grouped['VS.AttCellUpdt.EmgCall.PCH (None)'])))) * ((df_grouped['VS.RAB.SuccEstabPS.Conv (None)'] + df_grouped['VS.RAB.SuccEstabPS.Str (None)'] + df_grouped['VS.RAB.SuccEstabPS.Int (None)'] + df_grouped['VS.RAB.SuccEstabPS.Bkg (None)'] + df_grouped['VS.DCCC.Succ.F2D.AfterP2F (None)'])/(df_grouped['VS.RAB.AttEstabPS.Conv (None)'] + df_grouped['VS.RAB.AttEstabPS.Str (None)'] + df_grouped['VS.RAB.AttEstabPS.Int (None)'] + df_grouped['VS.RAB.AttEstabPS.Bkg (None)'] + df_grouped['VS.DCCC.Att.F2D.AfterP2F (None)'])),2)
    
    df_grouped['Average DL Throughput(Mbps)'] = round(df_grouped['VS.HSDPA.MeanChThroughput (kbit/s)'], 2)

    df_grouped['Average UL Throughput(Mbps)'] = round(df_grouped['VS.HSUPA.MeanChThroughput (kbit/s)'], 2)
    
    df_grouped['Availability'] = round(100 - (df_grouped['VS.Cell.UnavailTime.Sys (s)']/(df_grouped['VS.Cell.UnavailTime.Sys (s) count']*3600)),2)

    df_grouped['PS Sim Users'] = round((df_grouped['VS.CellDCHUEs (None)'] * 30/60),2)

    df_grouped.drop(df_grouped.columns.difference(['Traffic DL Volume (GB)', 'Traffic UL Volume (GB)', 'Voice Erlangs','Total MOCN Traffic(GB)', 'Accessibility CS (%)', 'Accessibility PS (%)', 'Average DL Throughput(Mbps)', 'Average UL Throughput(Mbps)', 'Availability', 'PS Sim Users']), 1, inplace=True)
    return df_grouped

# Generate each excel sheet with the respective dataframe (not including summary sheet)
def generate_excel_sheets(name, df):
    worksheet=workbook.add_worksheet(name)
    worksheet.set_column(0, 0, 11)

    if name == 'Day(AT&T+TLF)':
        worksheet.set_column(1, 1, 12)
        worksheet.set_column(2, 4, 21)
        worksheet.set_column(5, 6, 17)
        worksheet.set_column(7, 8, 28)
        worksheet.set_column(9, 10, 12)
    elif name == 'Day(AT&T)' or name == 'Day(TLF)':
        worksheet.set_column(1, 2, 21)
        worksheet.set_column(3, 3, 12)
        worksheet.set_column(4, 4, 21)
    elif name == 'Cell(AT&T+TLF)':
        worksheet.set_column(1, 1, 17)
        worksheet.set_column(2, 2, 21)
        worksheet.set_column(3, 3, 5)
        worksheet.set_column(4, 4, 12)
        worksheet.set_column(5, 7, 21)
        worksheet.set_column(8, 9, 17)
        worksheet.set_column(10, 11, 28)
        worksheet.set_column(12, 13, 12)
    elif name == 'Cell(AT&T)' or name == 'Cell(TLF)':
        worksheet.set_column(1, 1, 17)
        worksheet.set_column(2, 2, 21)
        worksheet.set_column(3, 3, 5)
        worksheet.set_column(4, 5, 21)
        worksheet.set_column(6, 6, 12)
        worksheet.set_column(7, 7, 21)
    elif name == 'Hour(AT&T+TLF)':
        worksheet.set_column(1, 1, 5)
        worksheet.set_column(2, 2, 12)
        worksheet.set_column(3, 5, 21)
        worksheet.set_column(6, 7, 17)
        worksheet.set_column(8, 9, 28)
        worksheet.set_column(10, 11, 12)
    elif name == 'Hour(AT&T)' or name == 'Hour(TLF)':
        worksheet.set_column(1, 1, 5)
        worksheet.set_column(2, 3, 21)
        worksheet.set_column(4, 4, 12)
        worksheet.set_column(5, 5, 21)
    elif name == 'Cell Hour(AT&T+TLF)':
        worksheet.set_column(1, 1, 5)
        worksheet.set_column(2, 2, 21)
        worksheet.set_column(3, 3, 5)
        worksheet.set_column(4, 4, 12)
        worksheet.set_column(5, 7, 21)
        worksheet.set_column(8, 9, 17)
        worksheet.set_column(10, 11, 28)
        worksheet.set_column(12, 13, 12)
    elif name == 'Cell Hour(AT&T)' or name == 'Cell Hour(TLF)':
        worksheet.set_column(1, 1, 5)
        worksheet.set_column(2, 2, 21)
        worksheet.set_column(3, 3, 5)
        worksheet.set_column(4, 5, 21)
        worksheet.set_column(6, 6, 12)
        worksheet.set_column(7, 7, 21)

    writer.sheets[name] = worksheet
    df.to_excel(writer,sheet_name=name,startrow=0 , startcol=0, merge_cells=False)
    return

def generate_summary_sheet(df_both, df_att, df_tef):
    worksheet=workbook.add_worksheet('Summary')
    worksheet.set_column(0, 0, 11)
    worksheet.set_column(1, 1, 28)
    worksheet.set_column(2, 3, 11)
    worksheet.set_column(4, 4, 15)
    worksheet.set_column(6, 6, 11)
    worksheet.set_column(7, 7, 17)

    bold_index = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'white'})

    KPI_format = workbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#92D050'})

    subtitle_green_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#92D050'})

    subtitle_green_format.set_text_wrap()

    subtitle_yellow_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#FFFF00'})

    subtitle_yellow_format.set_text_wrap()

    MOCN_KPI_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#FCD5B4'})

    merge__titles_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#00B0F0'})

    merge__titles_format.set_text_wrap()

    merge_index_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'white'})

    KPI_PLMN_data_format = workbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#FFFF00'})

    traffic_percentage_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'white',
     'num_format': '0.00%',
    'font_color': '#00B050'})

    mocn_percentage_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#FCD5B4',
    'num_format': '0.00%',
    'font_color': '#00B050'})

    worksheet.merge_range('A1:A2', 'Type', merge__titles_format)
    worksheet.merge_range('B1:B2', 'KPI', merge__titles_format)
    worksheet.write('C1', 'Before', merge__titles_format)
    worksheet.merge_range('D1:G1', 'After MOCN '+str(df_both.index[-1]), merge__titles_format)
    worksheet.merge_range('A3:A7', 'Traffic', merge_index_format)
    worksheet.merge_range('A8:A9', 'Accessibility', merge_index_format)
    worksheet.merge_range('A10:A11', 'Integrity', merge_index_format)
    worksheet.write('A12', 'Availability', bold_index)
    worksheet.write('B3', 'PS Sim Users', KPI_format)
    worksheet.write('B4', 'Voice Erlangs', KPI_format)
    worksheet.write('B5', 'Traffic DL Volume (GB)', KPI_format)
    worksheet.write('B6', 'Traffic UL Volume (GB)', KPI_format)
    worksheet.write('B7', 'Total MOCN Traffic(GB)', MOCN_KPI_format)
    worksheet.write('B8', 'Accessibility CS (%)', KPI_format)
    worksheet.write('B9', 'Accessibility PS (%)', KPI_format)
    worksheet.write('B10', 'Throughput DL Average (Mbps)', KPI_format)
    worksheet.write('B11', 'Throughput UL Average (Mbps)', KPI_format)
    worksheet.write('B12', 'Availability', KPI_format)
    worksheet.write('C2', 'Before MOCN AT&T', subtitle_green_format)
    worksheet.write('D2', 'After MOCN AT&T+TLF', subtitle_green_format)
    worksheet.write('E2', 'MOCN AT&T PLMN(0 and 1)', subtitle_yellow_format)
    worksheet.write('F2', 'MOCN TLF PLMN (2)', subtitle_yellow_format)
    worksheet.write('G2', 'TLF/(TLF+AT&T)', merge__titles_format)

    worksheet.write('C3', df_both.iloc[0]['PS Sim Users'], KPI_format)
    worksheet.write('C4', df_both.iloc[0]['Voice Erlangs'], KPI_format)
    worksheet.write('C5', df_both.iloc[0]['Traffic DL Volume (GB)'], KPI_format)
    worksheet.write('C6', df_both.iloc[0]['Traffic UL Volume (GB)'], KPI_format)
    worksheet.write('C7', df_both.iloc[0]['Total MOCN Traffic(GB)'], MOCN_KPI_format)
    worksheet.write('C8', df_both.iloc[0]['Accessibility CS (%)'], KPI_format)
    worksheet.write('C9', df_both.iloc[0]['Accessibility PS (%)'], KPI_format)
    worksheet.write('C10', df_both.iloc[0]['Average DL Throughput(Mbps)'], KPI_format)
    worksheet.write('C11', df_both.iloc[0]['Average UL Throughput(Mbps)'], KPI_format)
    worksheet.write('C12', df_both.iloc[0]['Availability'], KPI_format)

    worksheet.write('D3', df_both.iloc[-1]['PS Sim Users'], KPI_format)
    worksheet.write('D4', df_both.iloc[-1]['Voice Erlangs'], KPI_format)
    worksheet.write('D5', df_both.iloc[-1]['Traffic DL Volume (GB)'], KPI_format)
    worksheet.write('D6', df_both.iloc[-1]['Traffic UL Volume (GB)'], KPI_format)
    worksheet.write('D7', df_both.iloc[-1]['Total MOCN Traffic(GB)'], MOCN_KPI_format)
    worksheet.write('D8', df_both.iloc[-1]['Accessibility CS (%)'], KPI_format)
    worksheet.write('D9', df_both.iloc[-1]['Accessibility PS (%)'], KPI_format)
    worksheet.write('D10', df_both.iloc[-1]['Average DL Throughput(Mbps)'], KPI_format)
    worksheet.write('D11', df_both.iloc[-1]['Average UL Throughput(Mbps)'], KPI_format)
    worksheet.write('D12', df_both.iloc[-1]['Availability'], KPI_format)

    worksheet.write('E3', '', KPI_PLMN_data_format)
    worksheet.write('E4', df_att.iloc[-1]['Voice Erlangs'], KPI_PLMN_data_format)
    worksheet.write('E5', df_att.iloc[-1]['Traffic DL Volume (GB)'], KPI_PLMN_data_format)
    worksheet.write('E6', df_att.iloc[-1]['Traffic UL Volume (GB)'], KPI_PLMN_data_format)
    worksheet.write('E7', df_att.iloc[-1]['Total MOCN Traffic(GB)'], MOCN_KPI_format)
    worksheet.write('E8', '', KPI_PLMN_data_format)
    worksheet.write('E9', '', KPI_PLMN_data_format)
    worksheet.write('E10', '', KPI_PLMN_data_format)
    worksheet.write('E11', '', KPI_PLMN_data_format)
    worksheet.write('E12', '', KPI_PLMN_data_format)

    worksheet.write('F3', '', KPI_PLMN_data_format)
    worksheet.write('F4', df_tef.iloc[-1]['Voice Erlangs'], KPI_PLMN_data_format)
    worksheet.write('F5', df_tef.iloc[-1]['Traffic DL Volume (GB)'], KPI_PLMN_data_format)
    worksheet.write('F6', df_tef.iloc[-1]['Traffic UL Volume (GB)'], KPI_PLMN_data_format)
    worksheet.write('F7', df_tef.iloc[-1]['Total MOCN Traffic(GB)'], MOCN_KPI_format)
    worksheet.write('F8', '', KPI_PLMN_data_format)
    worksheet.write('F9', '', KPI_PLMN_data_format)
    worksheet.write('F10', '', KPI_PLMN_data_format)
    worksheet.write('F11', '', KPI_PLMN_data_format)
    worksheet.write('F12', '', KPI_PLMN_data_format)

    worksheet.write('G3', '', traffic_percentage_format)
    worksheet.write('G4', (df_tef.iloc[-1]['Voice Erlangs'] / (df_tef.iloc[-1]['Voice Erlangs'] + df_att.iloc[-1]['Voice Erlangs'])), traffic_percentage_format)
    worksheet.write('G5', (df_tef.iloc[-1]['Traffic DL Volume (GB)'] / (df_tef.iloc[-1]['Traffic DL Volume (GB)'] + df_att.iloc[-1]['Traffic DL Volume (GB)'])), traffic_percentage_format)
    worksheet.write('G6', (df_tef.iloc[-1]['Traffic UL Volume (GB)'] / (df_tef.iloc[-1]['Traffic UL Volume (GB)'] + df_att.iloc[-1]['Traffic UL Volume (GB)'])), traffic_percentage_format)
    worksheet.write('G7', (df_tef.iloc[-1]['Total MOCN Traffic(GB)'] / (df_tef.iloc[-1]['Total MOCN Traffic(GB)'] + df_att.iloc[-1]['Total MOCN Traffic(GB)'])), mocn_percentage_format)

    worksheet.write('G8', '', traffic_percentage_format)
    worksheet.write('G9', '', traffic_percentage_format)
    worksheet.write('G10', '', traffic_percentage_format)
    worksheet.write('G11', '', traffic_percentage_format)
    worksheet.write('G12', '', traffic_percentage_format)

    writer.sheets['Summary'] = worksheet
    return