from tkinter import Menu
from tkinter import *
from tkinter.ttk import * 
from tkinter import filedialog, Menu
import glob
from tkinter import messagebox, ttk
import numpy as np
from numpy.lib.function_base import average
import pandas as pd
import jinja2
import xlsxwriter

'''
First create the functionality for the buttons of the GUI
'''

def input_lte_daily_report_button():
    # Ask for CSV data of template lte daily report
    global input_lte_daily_report_data
    input_lte_daily_report_data = filedialog.askopenfilename()

def input_lte_daily_report1_button():
    # Ask for CSV data of template lte daily report1
    global input_lte_daily_report1_data
    input_lte_daily_report1_data = filedialog.askopenfilename()

def input_lte_mocn_button():
    # Ask for CSV data of template UMTS MOCN
    global input_lte_mocn_data
    input_lte_mocn_data = filedialog.askopenfilename()

def input_lte_nbiot_mocn_report_button():
    # Ask for CSV data of template lte nbiot mocn report
    global input_lte_nbiot_mocn_report_data
    input_lte_nbiot_mocn_report_data = filedialog.askopenfilename()

def input_lte_nbiot_report1_button():
    # Ask for CSV data of template lte nbiot report1
    global input_lte_nbiot_report1_data
    input_lte_nbiot_report1_data = filedialog.askopenfilename()

def input_lte_emtc_button():
    # Ask for CSV data of template lte emtc
    global input_lte_emtc_data
    input_lte_emtc_data = filedialog.askopenfilename()

def output_button():
    global output_data
    output_data = filedialog.askdirectory()

# LTE raw counters template download
def download_lte_raw_counters_template():
    global lte_template
    lte_template = filedialog.askdirectory()
    jinja_env = jinja2.Environment(loader=jinja2.FileSystemLoader('assets'))
    template = jinja_env.get_template('LTE_KPI_Template.xml')
    render = template.render()
    with open(lte_template+"/LTE_KPI_Template.xml", "w") as file:
        file.write(render)
    messagebox.showinfo(message="LTE U2000 templates saved at: "+lte_template, title="Finished")  
    lte_template = ""
    return

# LTE MOCN Report generation button handler
def generate_lte_mocn_delivery_report_button():
    global input_lte_daily_report_data
    global input_lte_daily_report1_data
    global input_lte_mocn_data
    global input_lte_nbiot_mocn_report_data
    global input_lte_nbiot_report1_data
    global input_lte_emtc_data
    global output_data
    if input_lte_daily_report_data != "" and input_lte_daily_report1_data != "" and input_lte_mocn_data != "" and output_data != "" and input_lte_emtc_data != "":
        filename1 = glob.glob(input_lte_daily_report_data)
        filename2 = glob.glob(input_lte_daily_report1_data)
        filename3 = glob.glob(input_lte_mocn_data)
        filename4 = glob.glob(input_lte_emtc_data)
        # Perform the calcuation and generate the excel report
        LTE_MOCN_REPORT(filename1, filename2, filename3, filename4)
        messagebox.showinfo(message="File saved at: "+output_data, title="Finished")
        input_lte_daily_report_data = ""
        input_lte_daily_report1_data = ""
        input_lte_mocn_data = ""
        input_lte_emtc_data = ""
        output_data = ""
    else:
        messagebox.showinfo(message="Please input all required files:", title="Error")

'''
Following perform the calculations required
'''
def LTE_MOCN_REPORT(lte_daily, lte_daily1, lte_mocn, lte_emtc):
    lte_daily_df = pd.read_csv(lte_daily[0],skiprows=7)
    lte_daily2_df = pd.read_csv(lte_daily1[0],skiprows=7)
    lte_mocn_df = pd.read_csv(lte_mocn[0],skiprows=7)
    lte_emtc_df = pd.read_csv(lte_emtc[0],skiprows=7)
    # Format the cellname and drop useless columns
    lte_daily_df = format_cells(lte_daily_df)
    lte_daily2_df = format_cells(lte_daily2_df)
    lte_mocn_df = format_cells(lte_mocn_df)
    lte_emtc_df = format_cells(lte_emtc_df)
    # Set the index only to prepare the initial columns time, site, cellname
    lte_daily_df = set_index(lte_daily_df)
    lte_daily2_df = set_index(lte_daily2_df)
    lte_emtc_df = set_index(lte_emtc_df)
    lte_mocn_df = set_index(lte_mocn_df)
    # Merge the 2 DF from daily data
    lte_cell_join_temp_df = pd.merge(lte_daily_df, lte_daily2_df, left_index=True, right_index=True)
    lte_cell_join_df = pd.merge(lte_cell_join_temp_df, lte_emtc_df, left_index=True, right_index=True)
    # Clear memory for useless df
    del lte_cell_join_temp_df
    del lte_daily_df
    del lte_daily2_df
    del lte_emtc_df
    # Remove the index
    lte_cell_join_df = lte_cell_join_df.reset_index()
    lte_mocn_df = lte_mocn_df.reset_index()
    # Required to aggregate the same column with count and sum
    lte_cell_join_df['L.Cell.Avail.Dur (s) count'] = lte_cell_join_df['L.Cell.Avail.Dur (s)']

    # Get the cell config integrity values
    cell_config_integrity = get_cell_config_integrity(lte_cell_join_df)

    # Generate the data for each sheet
    df_day_att_tef = calculate_lte_kpi(lte_cell_join_df, ["Date"])
    df_cell_att_tef = calculate_lte_kpi(lte_cell_join_df, ["Date","NE Name","CellName"])
    df_hour_att_tef = calculate_lte_kpi(lte_cell_join_df, ["Date","Hour"])
    df_cellhour_att_tef = calculate_lte_kpi(lte_cell_join_df, ["Date","Hour", "CellName"])
    
    df_day_att = calculate_lte_mocn(lte_mocn_df, ["Date"], 'ATT', lte_cell_join_df)
    df_day_tef = calculate_lte_mocn(lte_mocn_df, ["Date"], 'TEF', lte_cell_join_df)
    df_cell_att = calculate_lte_mocn(lte_mocn_df, ["Date","NE Name","CellName"], 'ATT', lte_cell_join_df)
    df_cell_tef = calculate_lte_mocn(lte_mocn_df, ["Date","NE Name","CellName"], 'TEF', lte_cell_join_df)
    df_hour_att = calculate_lte_mocn(lte_mocn_df, ["Date","Hour"], 'ATT', lte_cell_join_df)
    df_hour_tef = calculate_lte_mocn(lte_mocn_df, ["Date","Hour"], 'TEF', lte_cell_join_df)
    df_cellhour_att = calculate_lte_mocn(lte_mocn_df, ["Date","Hour", "CellName"], 'ATT', lte_cell_join_df)
    df_cellhour_tef = calculate_lte_mocn(lte_mocn_df, ["Date","Hour", "CellName"], 'TEF', lte_cell_join_df)
    del lte_cell_join_df
    del lte_mocn_df

    # Initiate the excel write process
    global writer
    global workbook
    writer = pd.ExcelWriter(output_data+'/results.xlsx',engine='xlsxwriter')
    workbook=writer.book

    # Generate the summary sheet
    generate_summary_sheet(df_day_att_tef, df_day_att, df_day_tef, cell_config_integrity)

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

# Format for time and cell name
def format_cells(df):

    df = df.drop(columns=['Period'])
    df['CellFilter1'] = df['Cell'].str.split(expand=True)[7]
    df['CellFilter2'] = df['CellFilter1'].str.split(pat="=", expand=True)[1]
    df['CellName'] = df['CellFilter2'].str.split(pat=",", expand=True)[0]
    df = df.replace('NIL', 0)

    del df['CellFilter1']
    del df['CellFilter2']
    del df['Cell']

    df['Start_Time'] = pd.to_datetime(df['Start Time']).dt.strftime('%Y-%m-%d %H:00')
    df['Date'] = pd.to_datetime(df['Start Time']).dt.strftime('%Y-%m-%d')
    df['Hour'] = pd.to_datetime(df['Start Time']).dt.strftime('%H:00')
    del df['Start Time']

    return df

# Define multi index for start time, site, cellname and cellid
def set_index(df):
    df = df.set_index(['Start_Time', 'Date', 'Hour', 'NE Name', 'CellName'])
    return df

# Function to calculate MOCN KPI
def calculate_lte_mocn(df, groupby_list, network, df_notmocn):
    if network == 'ATT':
        df = df.loc[(df['CnOperator'] == 'CN Operator ID=0') | (df['CnOperator'] == 'CN Operator ID=1')]
    else:
        df = df.loc[(df['CnOperator'] == 'CN Operator ID=2')]

    df['L.RBUsedOwn.DL.PLMN (Max)'] = df['L.RBUsedOwn.DL.PLMN (None)']
    df['L.RBUsedOwn.UL.PLMN (Max)'] = df['L.RBUsedOwn.UL.PLMN (None)']

    df_grouped = df.groupby(by=groupby_list).agg({
        'L.Traffic.User.Avg.PLMN (None)': 'mean',
        'L.Thrp.bits.DL.PLMN (bit)': 'sum',
        'L.Thrp.bits.UL.PLMN (bit)': 'sum',
        'L.Thrp.bits.DL.LastTTI.PLMN (bit)': 'sum',
        'L.Thrp.Time.DL.RmvLastTTI.PLMN (ms)': 'sum',
        'L.Thrp.bits.UE.UL.LastTTI.PLMN (bit)': 'sum',
        'L.Thrp.Time.UE.UL.RmvLastTTI.PLMN (ms)': 'sum',
        'L.RBUsedOwn.DL.PLMN (None)': 'sum',
        'L.E-RAB.SuccEst.PLMN (None)': 'sum',
        'L.E-RAB.AttEst.PLMN (None)': 'sum',
        'L.E-RAB.AbnormRel.PLMN (None)': 'sum',
        'L.E-RAB.AbnormRel.MME.PLMN (None)': 'sum',
        'L.E-RAB.NormRel.PLMN (None)': 'sum',
        'L.IRATHO.E2W.ExecSuccOut.PLMN (None)': 'sum',
        'L.IRATHO.E2G.ExecSuccOut.PLMN (None)': 'sum',
        'L.HHO.IntraeNB.IntraFreq.ExecSuccOut.PLMN (None)': 'sum',
        'L.HHO.IntereNB.IntraFreq.ExecSuccOut.PLMN (None)': 'sum',
        'L.HHO.IntraeNB.IntraFreq.ExecAttOut.PLMN (None)': 'sum',
        'L.HHO.IntereNB.IntraFreq.ExecAttOut.PLMN (None)': 'sum',
        'L.CSFB.PrepSucc.PLMN (None)': 'sum',
        'L.CSFB.PrepAtt.PLMN (None)': 'sum',
        'L.E-RAB.SuccEst.PLMN.QCI.1 (None)': 'sum',
        'L.E-RAB.AttEst.PLMN.QCI.1 (None)': 'sum',
        'L.E-RAB.AbnormRel.PLMN.QCI.1 (None)': 'sum',
        'L.E-RAB.AbnormRel.MME.VoIP.PLMN (None)': 'sum',
        'L.E-RAB.NormRel.PLMN.QCI.1 (None)': 'sum',
        'L.E-RAB.SessionTime.HighPrecision.PLMN.QCI1 (100 ms)': 'sum',
        'L.Thrp.bits.DL.PLMN.QCI.8 (bit)': 'sum',
        'L.Thrp.bits.UL.PLMN.QCI.8 (bit)': 'sum',
        'L.RBUsedOwn.UL.PLMN (None)': 'sum',
        'L.HHO.IntraeNB.IntraFreq.PrepAttOut.PLMN (None)': 'sum',
        'L.HHO.IntereNB.IntraFreq.PrepAttOut.PLMN (None)': 'sum',
        'L.HHO.IntraeNB.InterFreq.ExecSuccOut.PLMN (None)': 'sum',
        'L.HHO.IntereNB.InterFreq.ExecSuccOut.PLMN (None)': 'sum',
        'L.HHO.IntraeNB.InterFreq.PrepAttOut.PLMN (None)': 'sum',
        'L.HHO.IntereNB.InterFreq.PrepAttOut.PLMN (None)': 'sum',
        'L.HHO.X2.IntraFreq.ExecSuccOut.PLMN (None)': 'sum',
        'L.HHO.X2.InterFreq.ExecSuccOut.PLMN (None)': 'sum',
        'L.HHO.X2.IntraFreq.PrepAttOut.PLMN (None)': 'sum',
        'L.HHO.X2.InterFreq.PrepAttOut.PLMN (None)': 'sum',
        'L.RRCRedirection.E2W.PLMN (None)': 'sum',
        'L.RRCRedirection.E2W.CSFB.PLMN (None)': 'sum',
        'L.RBUsedOwn.DL.PLMN (Max)': 'max',
        'L.RBUsedOwn.UL.PLMN (Max)': 'max'
    })

    df_notmocn['L.ChMeas.PRB.DL.Avail (Max)'] = df_notmocn['L.ChMeas.PRB.DL.Avail (None)']
    df_notmocn['L.ChMeas.PRB.UL.Avail (Max)'] = df_notmocn['L.ChMeas.PRB.UL.Avail (None)']
    df_notmocn['L.Cell.Avail.Dur (Count)'] = df_notmocn['L.Cell.Avail.Dur (s)']

    df_grouped_notmocn = df_notmocn.groupby(by=groupby_list).agg({
        'L.UECNTX.AbnormRel (None)': 'sum',
        'L.UECNTX.NormRel (None)': 'sum',
        'L.UECNTX.Rel.MME (None)': 'sum',
        'L.ChMeas.PRB.DL.Avail (None)': 'sum',
        'L.ChMeas.PRB.UL.Avail (None)': 'sum',
        'L.ChMeas.PRB.DL.Avail (Max)': 'max',
        'L.ChMeas.PRB.UL.Avail (Max)': 'max',
        'L.Cell.Avail.Dur (s)': 'sum',
        'L.Cell.Unavail.Dur.EnergySaving (s)': 'sum',
        'L.Cell.Avail.Dur (Count)': 'count'
    })
    df_grouped_notmocn.drop(df_grouped_notmocn.columns.difference(['L.UECNTX.AbnormRel (None)', 'L.UECNTX.NormRel (None)', 'L.UECNTX.Rel.MME (None)', 'L.ChMeas.PRB.DL.Avail (None)', 'L.ChMeas.PRB.UL.Avail (None)', 'L.ChMeas.PRB.DL.Avail (Max)', 'L.ChMeas.PRB.UL.Avail (Max)', 'L.Cell.Avail.Dur (s)', 'L.Cell.Unavail.Dur.EnergySaving (s)', 'L.Cell.Avail.Dur (Count)']), 1, inplace=True)
    df_grouped = pd.merge(df_grouped, df_grouped_notmocn, left_index=True, right_index=True)

    df_grouped['VoLTE Traffic(Erls)'] =  round((df_grouped['L.E-RAB.SessionTime.HighPrecision.PLMN.QCI1 (100 ms)'])/(10*3600),2)
    df_grouped['Traffic DL Volume (GB)'] = round((df_grouped['L.Thrp.bits.DL.PLMN (bit)'])/(8*1024*1024*1024),2)
    df_grouped['Traffic UL Volume (GB)'] = round((df_grouped['L.Thrp.bits.UL.PLMN (bit)'])/(8*1024*1024*1024),2)
    df_grouped['WBB traffic(GB)'] = round((df_grouped['L.Thrp.bits.DL.PLMN.QCI.8 (bit)'] + df_grouped['L.Thrp.bits.UL.PLMN.QCI.8 (bit)'])/(8*1024*1024*1024),2)
    # df_grouped['eMTC traffic(GB)']
    # df_iot['NB-IoT traffic(RABs)'] = round((df_iot['L.NB.UECNTX.NormRel.PLMN (None)'] + df_iot['L.NB.UECNTX.AbnormRel.PLMN)']),2)
    df_grouped['Total MOCN traffic (GB)'] = round((df_grouped['Traffic DL Volume (GB)'] + df_grouped['Traffic UL Volume (GB)']),2)
    df_grouped['Accessibility PS (%)'] = round((100 * (df_grouped['L.E-RAB.SuccEst.PLMN (None)'])/(df_grouped['L.E-RAB.AttEst.PLMN (None)'])),2)
    df_grouped['Accessibility Volte (%)'] = round((100 * (df_grouped['L.E-RAB.SuccEst.PLMN.QCI.1 (None)'])/(df_grouped['L.E-RAB.AttEst.PLMN.QCI.1 (None)'])),2)
    df_grouped['Retainability PS (%)'] = round(100 - (100 * (df_grouped['L.E-RAB.AbnormRel.PLMN (None)'])/(df_grouped['L.E-RAB.AbnormRel.PLMN (None)'] + df_grouped['L.E-RAB.NormRel.PLMN (None)'] + df_grouped['L.IRATHO.E2W.ExecSuccOut.PLMN (None)'])),2)

    df_grouped['Retainability PS (%) (incl# MME)'] = round(100 - (100 * (df_grouped['L.E-RAB.AbnormRel.PLMN (None)'] + df_grouped['L.E-RAB.AbnormRel.MME.PLMN (None)'])/(df_grouped['L.E-RAB.AbnormRel.PLMN (None)'] + df_grouped['L.E-RAB.NormRel.PLMN (None)'] + df_grouped['L.IRATHO.E2W.ExecSuccOut.PLMN (None)'])),2)
    df_grouped['VoLTE Ret (%)'] = round(100 - (100 * (df_grouped['L.E-RAB.AbnormRel.PLMN.QCI.1 (None)'])/(df_grouped['L.E-RAB.AbnormRel.PLMN.QCI.1 (None)'] + df_grouped['L.E-RAB.NormRel.PLMN.QCI.1 (None)'])),2)
    df_grouped['VoLTE Ret (%) (incl# MME)'] = round(100 - (100 * (df_grouped['L.E-RAB.AbnormRel.PLMN.QCI.1 (None)'] + df_grouped['L.E-RAB.AbnormRel.MME.VoIP.PLMN (None)'])/(df_grouped['L.E-RAB.AbnormRel.PLMN.QCI.1 (None)'] + df_grouped['L.E-RAB.NormRel.PLMN.QCI.1 (None)'])),2)
    df_grouped['Average DL Throughput(Mbps)'] = round((df_grouped['L.Thrp.bits.DL.PLMN (bit)'] - df_grouped['L.Thrp.bits.DL.LastTTI.PLMN (bit)'])/(df_grouped['L.Thrp.Time.DL.RmvLastTTI.PLMN (ms)'] * 1000),2)
    df_grouped['Average UL Throughput(Mbps)'] = round((df_grouped['L.Thrp.bits.UL.PLMN (bit)'] - df_grouped['L.Thrp.bits.UE.UL.LastTTI.PLMN (bit)'])/(df_grouped['L.Thrp.Time.UE.UL.RmvLastTTI.PLMN (ms)'] * 1000),2)
    df_grouped['HO Intra Success Rate (%)'] = round((100 * (df_grouped['L.HHO.IntraeNB.IntraFreq.ExecSuccOut.PLMN (None)'] + df_grouped['L.HHO.IntereNB.IntraFreq.ExecSuccOut.PLMN (None)'])/(df_grouped['L.HHO.IntraeNB.IntraFreq.PrepAttOut.PLMN (None)'] + df_grouped['L.HHO.IntereNB.IntraFreq.PrepAttOut.PLMN (None)'])),2)
    df_grouped['HO Inter Success Rate (%)'] = round((100 * (df_grouped['L.HHO.IntraeNB.InterFreq.ExecSuccOut.PLMN (None)'] + df_grouped['L.HHO.IntereNB.InterFreq.ExecSuccOut.PLMN (None)'])/(df_grouped['L.HHO.IntraeNB.InterFreq.PrepAttOut.PLMN (None)'] + df_grouped['L.HHO.IntereNB.InterFreq.PrepAttOut.PLMN (None)'])),2)
    df_grouped['HO X2 Success Rate (%)'] = round((100 * (df_grouped['L.HHO.X2.IntraFreq.ExecSuccOut.PLMN (None)'] + df_grouped['L.HHO.X2.InterFreq.ExecSuccOut.PLMN (None)'])/(df_grouped['L.HHO.X2.IntraFreq.PrepAttOut.PLMN (None)'] + df_grouped['L.HHO.X2.InterFreq.PrepAttOut.PLMN (None)'])),2)
    df_grouped['HO S1 Success Rate (%)'] = round((100 * (df_grouped['L.HHO.IntereNB.IntraFreq.ExecSuccOut.PLMN (None)'] - df_grouped['L.HHO.X2.IntraFreq.ExecSuccOut.PLMN (None)'] + df_grouped['L.HHO.IntereNB.InterFreq.ExecSuccOut.PLMN (None)'] - df_grouped['L.HHO.X2.InterFreq.ExecSuccOut.PLMN (None)'])/(df_grouped['L.HHO.IntereNB.IntraFreq.PrepAttOut.PLMN (None)'] - df_grouped['L.HHO.X2.IntraFreq.PrepAttOut.PLMN (None)'] + df_grouped['L.HHO.IntereNB.InterFreq.PrepAttOut.PLMN (None)'] - df_grouped['L.HHO.X2.InterFreq.PrepAttOut.PLMN (None)'])),2)
    df_grouped['CSFB Success Rate(%)'] = round(100 * (df_grouped['L.CSFB.PrepSucc.PLMN (None)'])/(df_grouped['L.CSFB.PrepAtt.PLMN (None)']),2)

    df_grouped['Redirection Rate (%)'] = round((100 * (df_grouped['L.RRCRedirection.E2W.PLMN (None)'] - df_grouped['L.RRCRedirection.E2W.CSFB.PLMN (None)'])/(df_grouped['L.UECNTX.AbnormRel (None)'] + df_grouped['L.UECNTX.NormRel (None)'] - df_grouped['L.UECNTX.Rel.MME (None)'])),2)
    df_grouped['Availability'] = round((100 * (df_grouped['L.Cell.Avail.Dur (s)'] + df_grouped['L.Cell.Unavail.Dur.EnergySaving (s)'])/(df_grouped['L.Cell.Avail.Dur (Count)']*3600)),2)
    df_grouped['PS Sim Users'] = round(df_grouped['L.Traffic.User.Avg.PLMN (None)'],2)
    df_grouped['Average DL PRB Usage(%)'] = round(100 * (df_grouped['L.RBUsedOwn.DL.PLMN (None)'])/(df_grouped['L.ChMeas.PRB.DL.Avail (None)']),2)
    df_grouped['Max DL PRB Usage(%)'] = round(100 * (df_grouped['L.RBUsedOwn.DL.PLMN (Max)'])/(df_grouped['L.RBUsedOwn.UL.PLMN (Max)']))
    df_grouped['Average UL PRB Usage(%)'] = round(100 * (df_grouped['L.RBUsedOwn.UL.PLMN (None)'])/(df_grouped['L.ChMeas.PRB.UL.Avail (None)']),2)
    df_grouped['Max UL PRB Usage(%)'] = round(100 * (df_grouped['L.RBUsedOwn.UL.PLMN (Max)'])/(df_grouped['L.ChMeas.PRB.UL.Avail (Max)']),2)

    df_grouped.drop(df_grouped.columns.difference(['VoLTE Traffic(Erls)','Traffic DL Volume (GB)','Traffic UL Volume (GB)','WBB traffic(GB)', 'Total MOCN traffic (GB)', 'Accessibility PS (%)', 'Accessibility Volte (%)', 'Retainability PS (%)', 'Retainability PS (%) (incl# MME)', 'VoLTE Ret (%)', 'VoLTE Ret (%) (incl# MME)', 'Average DL Throughput(Mbps)', 'Average UL Throughput(Mbps)', 'HO Intra Success Rate (%)', 'HO Inter Success Rate (%)', 'HO X2 Success Rate (%)', 'HO S1 Success Rate (%)', 'CSFB Success Rate(%)', 'Redirection Rate (%)', 'Availability', 'PS Sim Users', 'Average DL PRB Usage(%)', 'Max DL PRB Usage(%)', 'Average UL PRB Usage(%)', 'Max UL PRB Usage(%)']), 1, inplace=True)
    return df_grouped

# Function to calculate UMTS KPI not MOCN
def calculate_lte_kpi(df, groupby_list):
    df = df.replace('NIL', 0)
    df['L.ChMeas.PRB.DL.Avail (Max)'] = df['L.ChMeas.PRB.DL.Avail (None)']
    df['L.ChMeas.PRB.UL.Avail (Max)'] = df['L.ChMeas.PRB.UL.Avail (None)']
    df['L.ChMeas.PRB.DL.DrbUsed.Avg (Max)'] = df['L.ChMeas.PRB.DL.DrbUsed.Avg (None)']
    df['L.ChMeas.PRB.UL.DrbUsed.Avg (Max)'] = df['L.ChMeas.PRB.UL.DrbUsed.Avg (None)']
    df['L.Cell.Avail.Dur (Count)'] = df['L.Cell.Avail.Dur (s)']

    df_grouped = df.groupby(by=groupby_list).agg({
        'L.Traffic.User.Avg (None)': 'mean',
        'L.Thrp.bits.DL (bit)': 'sum',
        'L.Thrp.bits.DL.SRB (bit)': 'sum',
        'L.Thrp.bits.UL.SRB (bit)': 'sum',
        'L.Thrp.bits.UL (bit)': 'sum',
        'L.Thrp.bits.DL.LastTTI (bit)': 'sum',
        'L.Thrp.Time.DL.RmvLastTTI (ms)': 'sum',
        'L.Thrp.Time.UE.UL.RmvLastTTI (ms)': 'sum',
        'L.Thrp.bits.UE.UL.LastTTI (bit)': 'sum',
        'L.ChMeas.PRB.DL.DrbUsed.Avg (None)': 'sum',
        'L.ChMeas.PRB.DL.Avail (None)': 'sum',
        'L.RRC.ConnReq.Succ (None)': 'sum',
        'L.RRC.ConnReq.Succ.MoSig (None)': 'sum',
        'L.RRC.ConnReq.Att (None)': 'sum',
        'L.RRC.ConnReq.Att.MoSig (None)': 'sum',
        'L.S1Sig.ConnEst.Succ (None)': 'sum',
        'L.S1Sig.ConnEst.Att (None)': 'sum',
        'L.E-RAB.SuccEst (None)': 'sum',
        'L.E-RAB.AttEst (None)': 'sum',
        'L.E-RAB.AbnormRel (None)': 'sum',
        'L.E-RAB.AbnormRel.MME (None)': 'sum',
        'L.E-RAB.NormRel.IRatHOOut (None)': 'sum',
        'L.E-RAB.NormRel (None)': 'sum',
        'L.HHO.IntraeNB.IntraFreq.ExecSuccOut (None)': 'sum',
        'L.HHO.IntereNB.IntraFreq.ExecSuccOut (None)': 'sum',
        'L.HHO.IntraeNB.IntraFreq.ExecAttOut (None)': 'sum',
        'L.HHO.IntereNB.IntraFreq.ExecAttOut (None)': 'sum',
        'L.CSFB.PrepSucc (None)': 'sum',
        'L.CSFB.PrepAtt (None)': 'sum',
        'L.E-RAB.SuccEst.QCI.1 (None)': 'sum',
        'L.E-RAB.AttEst.QCI.1 (None)': 'sum',
        'L.E-RAB.FailEst.X2AP.VoIP (None)': 'sum',
        'L.E-RAB.AbnormRel.QCI.1 (None)': 'sum',
        'L.E-RAB.AbnormRel.MME.VoIP (None)': 'sum',
        'L.E-RAB.NormRel.QCI.1 (None)': 'sum',
        'L.E-RAB.NormRel.IRatHOOut.QCI.1 (None)': 'sum',
        'L.E-RAB.SessionTime.HighPrecision.QCI1 (100 ms)': 'sum',
        'L.Cell.Avail.Dur (s)': 'sum',
        'L.Cell.Unavail.Dur.EnergySaving (s)': 'sum',
        'L.Thrp.bits.DL.QCI.8 (bit)': 'sum',
        'L.Thrp.bits.UL.QCI.8 (bit)': 'sum',
        'L.ChMeas.PRB.UL.DrbUsed.Avg (None)': 'sum',
        'L.ChMeas.PRB.UL.Avail (None)': 'sum',
        'L.HHO.IntraeNB.IntraFreq.PrepAttOut (None)': 'sum',
        'L.HHO.IntereNB.IntraFreq.PrepAttOut (None)': 'sum',
        'L.HHO.IntraeNB.InterFreq.ExecSuccOut (None)': 'sum',
        'L.HHO.IntereNB.InterFreq.ExecSuccOut (None)': 'sum',
        'L.HHO.IntraeNB.InterFreq.PrepAttOut (None)': 'sum',
        'L.HHO.IntereNB.InterFreq.PrepAttOut (None)': 'sum',
        'L.RRCRedirection.E2W (None)': 'sum',
        'L.RRCRedirection.E2W.CSFB (None)': 'sum',
        'L.UECNTX.AbnormRel (None)': 'sum',
        'L.UECNTX.NormRel (None)': 'sum',
        'L.UECNTX.Rel.MME (None)': 'sum',
        'L.HHO.X2.IntraFreq.ExecSuccOut (None)': 'sum',
        'L.HHO.X2.InterFreq.ExecSuccOut (None)': 'sum',
        'L.HHO.X2.IntraFreq.PrepAttOut (None)': 'sum',
        'L.HHO.X2.InterFreq.PrepAttOut (None)': 'sum',
        'L.Thrp.eMTC.bits.DL (bit)': 'sum',
        'L.Thrp.eMTC.bits.UL (bit)': 'sum',
        'L.ChMeas.PRB.DL.Avail (Max)': 'max',
        'L.ChMeas.PRB.UL.Avail (Max)': 'max',
        'L.ChMeas.PRB.DL.DrbUsed.Avg (Max)': 'max',
        'L.ChMeas.PRB.UL.DrbUsed.Avg (Max)': 'max',
        'L.Cell.Avail.Dur (Count)': 'count'
    })

    df_grouped['VoLTE Traffic(Erls)'] =  round((df_grouped['L.E-RAB.SessionTime.HighPrecision.QCI1 (100 ms)'])/(10*3600),2)
    df_grouped['Traffic DL Volume (GB)'] = round((df_grouped['L.Thrp.bits.DL (bit)'])/(8*1024*1024*1024),2)
    df_grouped['Traffic UL Volume (GB)'] = round((df_grouped['L.Thrp.bits.UL (bit)'])/(8*1024*1024*1024),2)
    df_grouped['WBB traffic(GB)'] = round((df_grouped['L.Thrp.bits.DL.QCI.8 (bit)'] + df_grouped['L.Thrp.bits.UL.QCI.8 (bit)'])/(8*1024*1024*1024),2)
    # df_grouped['eMTC traffic(GB)']
    # df_iot['NB-IoT traffic(RABs)'] = round((df_iot['L.NB.UECNTX.NormRel.PLMN (None)'] + df_iot['L.NB.UECNTX.AbnormRel.PLMN)']),2)
    df_grouped['Total MOCN traffic (GB)'] = round((df_grouped['Traffic DL Volume (GB)'] + df_grouped['Traffic UL Volume (GB)']),2)
    df_grouped['Accessibility PS (%)'] = round((100 * ((df_grouped['L.RRC.ConnReq.Succ (None)'] - df_grouped['L.RRC.ConnReq.Succ.MoSig (None)'])/(df_grouped['L.RRC.ConnReq.Att (None)']-df_grouped['L.RRC.ConnReq.Att.MoSig (None)'])) * ((df_grouped['L.S1Sig.ConnEst.Succ (None)'])/(df_grouped['L.S1Sig.ConnEst.Att (None)'])) *((df_grouped['L.E-RAB.SuccEst (None)'])/(df_grouped['L.E-RAB.AttEst (None)']))),2)
    df_grouped['Accessibility Volte (%)'] = round((100 * (df_grouped['L.E-RAB.SuccEst.QCI.1 (None)'])/(df_grouped['L.E-RAB.AttEst.QCI.1 (None)'] - df_grouped['L.E-RAB.FailEst.X2AP.VoIP (None)'])),2)
    df_grouped['Retainability PS (%)'] = round(100 - (100 * (df_grouped['L.E-RAB.AbnormRel (None)'])/(df_grouped['L.E-RAB.AbnormRel (None)'] + df_grouped['L.E-RAB.NormRel (None)'] + df_grouped['L.E-RAB.NormRel.IRatHOOut (None)'])),2)
    df_grouped['Retainability PS (%) (incl# MME)'] = round(100 - (100 * (df_grouped['L.E-RAB.AbnormRel (None)'] + df_grouped['L.E-RAB.AbnormRel.MME (None)'])/(df_grouped['L.E-RAB.AbnormRel (None)'] + df_grouped['L.E-RAB.NormRel (None)'] + df_grouped['L.E-RAB.NormRel.IRatHOOut (None)'])),2)
    df_grouped['VoLTE Ret (%)'] = round(100 - (100 * (df_grouped['L.E-RAB.AbnormRel.QCI.1 (None)'])/(df_grouped['L.E-RAB.AbnormRel.QCI.1 (None)'] + df_grouped['L.E-RAB.NormRel.QCI.1 (None)'] + df_grouped['L.E-RAB.NormRel.IRatHOOut.QCI.1 (None)'])),2)
    df_grouped['VoLTE Ret (%) (incl# MME)'] = round(100 - (100 * (df_grouped['L.E-RAB.AbnormRel.QCI.1 (None)'] + df_grouped['L.E-RAB.AbnormRel.MME.VoIP (None)'])/(df_grouped['L.E-RAB.AbnormRel.QCI.1 (None)'] + df_grouped['L.E-RAB.NormRel.QCI.1 (None)'] + df_grouped['L.E-RAB.NormRel.IRatHOOut.QCI.1 (None)'])),2)
    df_grouped['Average DL Throughput(Mbps)'] = round((df_grouped['L.Thrp.bits.DL (bit)'] - df_grouped['L.Thrp.bits.DL.LastTTI (bit)'])/(df_grouped['L.Thrp.Time.DL.RmvLastTTI (ms)'] * 1000),2)
    df_grouped['Average UL Throughput(Mbps)'] = round((df_grouped['L.Thrp.bits.UL (bit)'] - df_grouped['L.Thrp.bits.UE.UL.LastTTI (bit)'])/(df_grouped['L.Thrp.Time.UE.UL.RmvLastTTI (ms)'] * 1000),2)

    df_grouped['HO Intra Success Rate (%)'] = round((100 * (df_grouped['L.HHO.IntraeNB.IntraFreq.ExecSuccOut (None)'] + df_grouped['L.HHO.IntereNB.IntraFreq.ExecSuccOut (None)'])/(df_grouped['L.HHO.IntraeNB.IntraFreq.PrepAttOut (None)'] + df_grouped['L.HHO.IntereNB.IntraFreq.PrepAttOut (None)'])),2)
    df_grouped['HO Inter Success Rate (%)'] = round((100 * (df_grouped['L.HHO.IntraeNB.InterFreq.ExecSuccOut (None)'] + df_grouped['L.HHO.IntereNB.InterFreq.ExecSuccOut (None)'])/(df_grouped['L.HHO.IntraeNB.InterFreq.PrepAttOut (None)'] + df_grouped['L.HHO.IntereNB.InterFreq.PrepAttOut (None)'])),2)
    df_grouped['HO X2 Success Rate (%)'] = round((100 * (df_grouped['L.HHO.X2.IntraFreq.ExecSuccOut (None)'] + df_grouped['L.HHO.X2.InterFreq.ExecSuccOut (None)'])/(df_grouped['L.HHO.X2.IntraFreq.PrepAttOut (None)'] + df_grouped['L.HHO.X2.InterFreq.PrepAttOut (None)'])),2)
    df_grouped['HO S1 Success Rate (%)'] = round((100 * (df_grouped['L.HHO.IntereNB.IntraFreq.ExecSuccOut (None)'] - df_grouped['L.HHO.X2.IntraFreq.ExecSuccOut (None)'] + df_grouped['L.HHO.IntereNB.InterFreq.ExecSuccOut (None)'] - df_grouped['L.HHO.X2.InterFreq.ExecSuccOut (None)'])/(df_grouped['L.HHO.IntereNB.IntraFreq.PrepAttOut (None)'] - df_grouped['L.HHO.X2.IntraFreq.PrepAttOut (None)'] + df_grouped['L.HHO.IntereNB.InterFreq.PrepAttOut (None)'] - df_grouped['L.HHO.X2.InterFreq.PrepAttOut (None)'])),2)
    df_grouped['CSFB Success Rate(%)'] = round(100 * (df_grouped['L.CSFB.PrepSucc (None)'])/(df_grouped['L.CSFB.PrepAtt (None)']),2)
    df_grouped['Redirection Rate (%)'] = round((100 * (df_grouped['L.RRCRedirection.E2W (None)'] - df_grouped['L.RRCRedirection.E2W.CSFB (None)'])/(df_grouped['L.UECNTX.AbnormRel (None)'] + df_grouped['L.UECNTX.NormRel (None)'] - df_grouped['L.UECNTX.Rel.MME (None)'])),2)
    df_grouped['Availability'] = round((100 * (df_grouped['L.Cell.Avail.Dur (s)'] + df_grouped['L.Cell.Unavail.Dur.EnergySaving (s)'])/(df_grouped['L.Cell.Avail.Dur (Count)']*3600)),2)
    df_grouped['PS Sim Users'] = round(df_grouped['L.Traffic.User.Avg (None)'],2)
    df_grouped['Average DL PRB Usage(%)'] = round(100 * (df_grouped['L.ChMeas.PRB.DL.DrbUsed.Avg (None)'])/(df_grouped['L.ChMeas.PRB.DL.Avail (None)']),2)
    df_grouped['Max DL PRB Usage(%)'] = round(100 * (df_grouped['L.ChMeas.PRB.DL.DrbUsed.Avg (Max)'])/(df_grouped['L.ChMeas.PRB.DL.Avail (Max)']))
    df_grouped['Average UL PRB Usage(%)'] = round(100 * (df_grouped['L.ChMeas.PRB.UL.DrbUsed.Avg (None)'])/(df_grouped['L.ChMeas.PRB.UL.Avail (None)']),2)
    df_grouped['Max UL PRB Usage(%)'] = round(100 * (df_grouped['L.ChMeas.PRB.UL.DrbUsed.Avg (Max)'])/(df_grouped['L.ChMeas.PRB.UL.Avail (Max)']),2)

    df_grouped.drop(df_grouped.columns.difference(['VoLTE Traffic(Erls)','Traffic DL Volume (GB)','Traffic UL Volume (GB)','WBB traffic(GB)', 'Total MOCN traffic (GB)', 'Accessibility PS (%)', 'Accessibility Volte (%)', 'Retainability PS (%)', 'Retainability PS (%) (incl# MME)', 'VoLTE Ret (%)', 'VoLTE Ret (%) (incl# MME)', 'Average DL Throughput(Mbps)', 'Average UL Throughput(Mbps)', 'HO Intra Success Rate (%)', 'HO Inter Success Rate (%)', 'HO X2 Success Rate (%)', 'HO S1 Success Rate (%)', 'CSFB Success Rate(%)', 'Redirection Rate (%)', 'Availability', 'PS Sim Users', 'Average DL PRB Usage(%)', 'Max DL PRB Usage(%)', 'Average UL PRB Usage(%)', 'Max UL PRB Usage(%)']), 1, inplace=True)
    
    return df_grouped

# Generate each excel sheet with the respective dataframe (not including summary sheet)
def generate_excel_sheets(name, df):
    worksheet=workbook.add_worksheet(name)
    worksheet.set_column(0, 0, 11)
    if name == 'Day(AT&T+TLF)' or name == 'Day(AT&T)' or name == 'Day(TLF)':
        worksheet.set_column(1, 1, 17)
        worksheet.set_column(2, 8, 20)
        worksheet.set_column(9, 9, 29)
        worksheet.set_column(10, 10, 12)
        worksheet.set_column(11, 11, 23)
        worksheet.set_column(12, 13, 28)
        worksheet.set_column(14, 19, 21)
        worksheet.set_column(20, 20, 11)
        worksheet.set_column(21, 21, 16)
        worksheet.set_column(22, 25, 23)
    elif name == 'Cell(AT&T+TLF)' or name == 'Cell(AT&T)' or name == 'Cell(TLF)':
        worksheet.set_column(1, 1, 15)
        worksheet.set_column(2, 2, 22)
        worksheet.set_column(3, 3, 19)
        worksheet.set_column(4, 10, 20)
        worksheet.set_column(11, 11, 29)
        worksheet.set_column(12, 12, 13)
        worksheet.set_column(13, 13, 24)
        worksheet.set_column(14, 15, 28)
        worksheet.set_column(16, 17, 23)
        worksheet.set_column(18, 21, 21)
        worksheet.set_column(22, 23, 12)
        worksheet.set_column(24, 24, 23)
        worksheet.set_column(25, 25, 20)
        worksheet.set_column(26, 26, 23)
        worksheet.set_column(27, 27, 20)
    elif name == 'Hour(AT&T+TLF)' or name == 'Hour(AT&T)' or name == 'Hour(TLF)':
        worksheet.set_column(1, 1, 5)
        worksheet.set_column(2, 2, 17)
        worksheet.set_column(3, 9, 20)
        worksheet.set_column(10, 10, 29)
        worksheet.set_column(11, 11, 12)
        worksheet.set_column(12, 12, 23)
        worksheet.set_column(13, 14, 28)
        worksheet.set_column(15, 20, 21)
        worksheet.set_column(21, 21, 11)
        worksheet.set_column(22, 22, 16)
        worksheet.set_column(23, 24, 23)
        worksheet.set_column(25, 26, 23)
    elif name == 'Cell Hour(AT&T+TLF)' or name == 'Cell Hour(AT&T)' or name == 'Cell Hour(TLF)':
        worksheet.set_column(1, 1, 5)
        worksheet.set_column(2, 2, 20)
        worksheet.set_column(3, 3, 17)
        worksheet.set_column(4, 10, 20)
        worksheet.set_column(11, 11, 29)
        worksheet.set_column(12, 12, 12)
        worksheet.set_column(13, 13, 23)
        worksheet.set_column(14, 15, 28)
        worksheet.set_column(16, 21, 21)
        worksheet.set_column(22, 22, 11)
        worksheet.set_column(23, 23, 16)
        worksheet.set_column(24, 25, 23)
        worksheet.set_column(26, 26, 23)
        worksheet.set_column(27, 27, 20)
    writer.sheets[name] = worksheet
    df.to_excel(writer,sheet_name=name,startrow=0 , startcol=0, merge_cells=False)
    return

def generate_summary_sheet(df_both, df_att, df_tef, cell_config_integrity):
    worksheet=workbook.add_worksheet('Summary')
    worksheet.set_column(0, 0, 11)
    worksheet.set_column(1, 1, 31)
    worksheet.set_column(2, 3, 12)
    worksheet.set_column(4, 4, 15)
    worksheet.set_column(5, 6, 11)

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
    worksheet.merge_range('A3:A8', 'Traffic', merge_index_format)
    worksheet.merge_range('A9:A10', 'Accessibility', merge_index_format)
    worksheet.merge_range('A11:A14', 'Retainability', merge_index_format)
    worksheet.merge_range('A15:A21', 'Integrity', merge_index_format)
    worksheet.merge_range('A22:A26', 'Mobility', merge_index_format)
    worksheet.write('A27', 'Availability', merge_index_format)
    worksheet.write('B3', 'PS Sim Users', KPI_format)
    worksheet.write('B4', 'VoLTE Traffic(Erls)', KPI_format)
    worksheet.write('B5', 'Traffic DL Volume (GB)', KPI_format)
    worksheet.write('B6', 'Traffic UL Volume (GB)', KPI_format)
    worksheet.write('B7', 'WBB traffic (GB)', KPI_format)
    worksheet.write('B8', 'Total MOCN traffic (GB)', MOCN_KPI_format)
    worksheet.write('B9', 'Accessibility PS (%)', KPI_format)
    worksheet.write('B10', 'Accessibility VoLTE (%)', KPI_format)
    worksheet.write('B11', 'Retainability PS (%)', KPI_format)
    worksheet.write('B12', 'Retainability PS (%) (incl. MME)', KPI_format)
    worksheet.write('B13', 'VoLTE Ret (%)', KPI_format)
    worksheet.write('B14', 'VoLTE Ret (%) (incl. MME)', KPI_format)
    worksheet.write('B15', 'Throughput DL Average (Mbps)', KPI_format)
    worksheet.write('B16', 'Throughput UL Average (Mbps)', KPI_format)
    worksheet.write('B17', 'PRB Usage DL Average (%)', KPI_format)
    worksheet.write('B18', 'PRB Usage DL Max (%)', KPI_format)
    worksheet.write('B19', 'PRB Usage UL Average (%)', KPI_format)
    worksheet.write('B20', 'PRB Usage UL Max (%)', KPI_format)
    worksheet.write('B21', 'Cell Configuration Integrity', KPI_format)
    worksheet.write('B22', 'HO Intra Success Rate (%)', KPI_format)
    worksheet.write('B23', 'HO Inter Success Rate (%)', KPI_format)
    worksheet.write('B24', 'HO S1 Success Rate (%)', KPI_format)
    worksheet.write('B25', 'CSFB Success Rate (%)', KPI_format)
    worksheet.write('B26', 'Redirection Rate (%)', KPI_format)
    worksheet.write('B27', 'Availability', KPI_format)
    worksheet.write('C2', 'Before MOCN AT&T', subtitle_green_format)
    worksheet.write('D2', 'After MOCN AT&T+TLF', subtitle_green_format)
    worksheet.write('E2', 'MOCN AT&T PLMN(0 and 1)', subtitle_yellow_format)
    worksheet.write('F2', 'MOCN TLF PLMN (2)', subtitle_yellow_format)
    worksheet.write('G2', 'TLF/(TLF+AT&T)', merge__titles_format)

    worksheet.write('C3', df_both.iloc[0]['PS Sim Users'], KPI_format)
    worksheet.write('C4', df_both.iloc[0]['VoLTE Traffic(Erls)'], KPI_format)
    worksheet.write('C5', df_both.iloc[0]['Traffic DL Volume (GB)'], KPI_format)
    worksheet.write('C6', df_both.iloc[0]['Traffic UL Volume (GB)'], KPI_format)
    worksheet.write('C7', df_both.iloc[0]['WBB traffic(GB)'], KPI_format)
    worksheet.write('C8', df_both.iloc[0]['Total MOCN traffic (GB)'], MOCN_KPI_format)
    worksheet.write('C9', df_both.iloc[0]['Accessibility PS (%)'], KPI_format)
    worksheet.write('C10', df_both.iloc[0]['Accessibility Volte (%)'], KPI_format)
    worksheet.write('C11', df_both.iloc[0]['Retainability PS (%)'], KPI_format)
    worksheet.write('C12', df_both.iloc[0]['Retainability PS (%) (incl# MME)'], KPI_format)
    worksheet.write('C13', df_both.iloc[0]['VoLTE Ret (%)'], KPI_format)
    worksheet.write('C14', df_both.iloc[0]['VoLTE Ret (%) (incl# MME)'], KPI_format)
    worksheet.write('C15', df_both.iloc[0]['Average DL Throughput(Mbps)'], KPI_format)
    worksheet.write('C16', df_both.iloc[0]['Average UL Throughput(Mbps)'], KPI_format)
    worksheet.write('C17', df_both.iloc[0]['Average DL PRB Usage(%)'], KPI_format)
    worksheet.write('C18', df_both.iloc[0]['Max DL PRB Usage(%)'], KPI_format)
    worksheet.write('C19', df_both.iloc[0]['Average UL PRB Usage(%)'], KPI_format)
    worksheet.write('C20', df_both.iloc[0]['Max UL PRB Usage(%)'], KPI_format)
    worksheet.write('C21', cell_config_integrity[0], KPI_format)
    worksheet.write('C22', df_both.iloc[0]['HO Intra Success Rate (%)'], KPI_format)
    worksheet.write('C23', df_both.iloc[0]['HO Inter Success Rate (%)'], KPI_format)
    worksheet.write('C24', df_both.iloc[0]['HO S1 Success Rate (%)'], KPI_format)
    worksheet.write('C25', df_both.iloc[0]['CSFB Success Rate(%)'], KPI_format)
    worksheet.write('C26', df_both.iloc[0]['Redirection Rate (%)'], KPI_format)
    worksheet.write('C27', df_both.iloc[0]['Availability'], KPI_format)

    worksheet.write('D3', df_both.iloc[-1]['PS Sim Users'], KPI_format)
    worksheet.write('D4', df_both.iloc[-1]['VoLTE Traffic(Erls)'], KPI_format)
    worksheet.write('D5', df_both.iloc[-1]['Traffic DL Volume (GB)'], KPI_format)
    worksheet.write('D6', df_both.iloc[-1]['Traffic UL Volume (GB)'], KPI_format)
    worksheet.write('D7', df_both.iloc[-1]['WBB traffic(GB)'], KPI_format)
    worksheet.write('D8', df_both.iloc[-1]['Total MOCN traffic (GB)'], MOCN_KPI_format)
    worksheet.write('D9', df_both.iloc[-1]['Accessibility PS (%)'], KPI_format)
    worksheet.write('D10', df_both.iloc[-1]['Accessibility Volte (%)'], KPI_format)
    worksheet.write('D11', df_both.iloc[-1]['Retainability PS (%)'], KPI_format)
    worksheet.write('D12', df_both.iloc[-1]['Retainability PS (%) (incl# MME)'], KPI_format)
    worksheet.write('D13', df_both.iloc[-1]['VoLTE Ret (%)'], KPI_format)
    worksheet.write('D14', df_both.iloc[-1]['VoLTE Ret (%) (incl# MME)'], KPI_format)
    worksheet.write('D15', df_both.iloc[-1]['Average DL Throughput(Mbps)'], KPI_format)
    worksheet.write('D16', df_both.iloc[-1]['Average UL Throughput(Mbps)'], KPI_format)
    worksheet.write('D17', df_both.iloc[-1]['Average DL PRB Usage(%)'], KPI_format)
    worksheet.write('D18', df_both.iloc[-1]['Max DL PRB Usage(%)'], KPI_format)
    worksheet.write('D19', df_both.iloc[-1]['Average UL PRB Usage(%)'], KPI_format)
    worksheet.write('D20', df_both.iloc[-1]['Max UL PRB Usage(%)'], KPI_format)
    worksheet.write('D21', cell_config_integrity[1], KPI_format)
    worksheet.write('D22', df_both.iloc[-1]['HO Intra Success Rate (%)'], KPI_format)
    worksheet.write('D23', df_both.iloc[-1]['HO Inter Success Rate (%)'], KPI_format)
    worksheet.write('D24', df_both.iloc[-1]['HO S1 Success Rate (%)'], KPI_format)
    worksheet.write('D25', df_both.iloc[-1]['CSFB Success Rate(%)'], KPI_format)
    worksheet.write('D26', df_both.iloc[-1]['Redirection Rate (%)'], KPI_format)
    worksheet.write('D27', df_both.iloc[-1]['Availability'], KPI_format)

    worksheet.write('E3', df_att.iloc[-1]['PS Sim Users'], KPI_PLMN_data_format)
    worksheet.write('E4', df_att.iloc[-1]['VoLTE Traffic(Erls)'], KPI_PLMN_data_format)
    worksheet.write('E5', df_att.iloc[-1]['Traffic DL Volume (GB)'], KPI_PLMN_data_format)
    worksheet.write('E6', df_att.iloc[-1]['Traffic UL Volume (GB)'], KPI_PLMN_data_format)
    worksheet.write('E7', df_att.iloc[-1]['WBB traffic(GB)'], KPI_PLMN_data_format)
    worksheet.write('E8', df_att.iloc[-1]['Total MOCN traffic (GB)'], MOCN_KPI_format)
    worksheet.write('E9', df_att.iloc[-1]['Accessibility PS (%)'], KPI_PLMN_data_format)
    worksheet.write('E10', df_att.iloc[-1]['Accessibility Volte (%)'], KPI_PLMN_data_format)
    worksheet.write('E11', df_att.iloc[-1]['Retainability PS (%)'], KPI_PLMN_data_format)
    worksheet.write('E12', df_att.iloc[-1]['Retainability PS (%) (incl# MME)'], KPI_PLMN_data_format)
    worksheet.write('E13', df_att.iloc[-1]['VoLTE Ret (%)'], KPI_PLMN_data_format)
    worksheet.write('E14', df_att.iloc[-1]['VoLTE Ret (%) (incl# MME)'], KPI_PLMN_data_format)
    worksheet.write('E15', df_att.iloc[-1]['Average DL Throughput(Mbps)'], KPI_PLMN_data_format)
    worksheet.write('E16', df_att.iloc[-1]['Average UL Throughput(Mbps)'], KPI_PLMN_data_format)
    worksheet.write('E17', df_att.iloc[-1]['Average DL PRB Usage(%)'], KPI_PLMN_data_format)
    worksheet.write('E18', df_att.iloc[-1]['Max DL PRB Usage(%)'], KPI_PLMN_data_format)
    worksheet.write('E19', df_att.iloc[-1]['Average UL PRB Usage(%)'], KPI_PLMN_data_format)
    worksheet.write('E20', df_att.iloc[-1]['Max UL PRB Usage(%)'], KPI_PLMN_data_format)
    worksheet.write('E21', cell_config_integrity[1], KPI_PLMN_data_format)
    worksheet.write('E22', df_att.iloc[-1]['HO Intra Success Rate (%)'], KPI_PLMN_data_format)
    worksheet.write('E23', df_att.iloc[-1]['HO Inter Success Rate (%)'], KPI_PLMN_data_format)
    worksheet.write('E24', df_att.iloc[-1]['HO S1 Success Rate (%)'], KPI_PLMN_data_format)
    worksheet.write('E25', df_att.iloc[-1]['CSFB Success Rate(%)'], KPI_PLMN_data_format)
    worksheet.write('E26', df_att.iloc[-1]['Redirection Rate (%)'], KPI_PLMN_data_format)
    worksheet.write('E27', '', KPI_PLMN_data_format)

    worksheet.write('F3', df_tef.iloc[-1]['PS Sim Users'], KPI_PLMN_data_format)
    worksheet.write('F4', df_tef.iloc[-1]['VoLTE Traffic(Erls)'], KPI_PLMN_data_format)
    worksheet.write('F5', df_tef.iloc[-1]['Traffic DL Volume (GB)'], KPI_PLMN_data_format)
    worksheet.write('F6', df_tef.iloc[-1]['Traffic UL Volume (GB)'], KPI_PLMN_data_format)
    worksheet.write('F7', df_tef.iloc[-1]['WBB traffic(GB)'], KPI_PLMN_data_format)
    worksheet.write('F8', df_tef.iloc[-1]['Total MOCN traffic (GB)'], MOCN_KPI_format)
    worksheet.write('F9', df_tef.iloc[-1]['Accessibility PS (%)'], KPI_PLMN_data_format)
    worksheet.write('F10', df_tef.iloc[-1]['Accessibility Volte (%)'], KPI_PLMN_data_format)
    worksheet.write('F11', df_tef.iloc[-1]['Retainability PS (%)'], KPI_PLMN_data_format)
    worksheet.write('F12', df_tef.iloc[-1]['Retainability PS (%) (incl# MME)'], KPI_PLMN_data_format)
    worksheet.write('F13', df_tef.iloc[-1]['VoLTE Ret (%)'], KPI_PLMN_data_format)
    worksheet.write('F14', df_tef.iloc[-1]['VoLTE Ret (%) (incl# MME)'], KPI_PLMN_data_format)
    worksheet.write('F15', df_tef.iloc[-1]['Average DL Throughput(Mbps)'], KPI_PLMN_data_format)
    worksheet.write('F16', df_tef.iloc[-1]['Average UL Throughput(Mbps)'], KPI_PLMN_data_format)
    worksheet.write('F17', df_tef.iloc[-1]['Average DL PRB Usage(%)'], KPI_PLMN_data_format)
    worksheet.write('F18', df_tef.iloc[-1]['Max DL PRB Usage(%)'], KPI_PLMN_data_format)
    worksheet.write('F19', df_tef.iloc[-1]['Average UL PRB Usage(%)'], KPI_PLMN_data_format)
    worksheet.write('F20', df_tef.iloc[-1]['Max UL PRB Usage(%)'], KPI_PLMN_data_format)
    worksheet.write('F21', cell_config_integrity[1], KPI_PLMN_data_format)
    worksheet.write('F22', df_tef.iloc[-1]['HO Intra Success Rate (%)'], KPI_PLMN_data_format)
    worksheet.write('F23', df_tef.iloc[-1]['HO Inter Success Rate (%)'], KPI_PLMN_data_format)
    worksheet.write('F24', df_tef.iloc[-1]['HO S1 Success Rate (%)'], KPI_PLMN_data_format)
    worksheet.write('F25', df_tef.iloc[-1]['CSFB Success Rate(%)'], KPI_PLMN_data_format)
    worksheet.write('F26', df_tef.iloc[-1]['Redirection Rate (%)'], KPI_PLMN_data_format)
    worksheet.write('F27', '', KPI_PLMN_data_format)

    worksheet.write('G3', (df_tef.iloc[-1]['PS Sim Users'] / (df_tef.iloc[-1]['PS Sim Users'] + df_att.iloc[-1]['PS Sim Users'])), traffic_percentage_format)
    worksheet.write('G4', (df_tef.iloc[-1]['VoLTE Traffic(Erls)'] / (df_tef.iloc[-1]['VoLTE Traffic(Erls)'] + df_att.iloc[-1]['VoLTE Traffic(Erls)'])), traffic_percentage_format)
    worksheet.write('G5', (df_tef.iloc[-1]['Traffic DL Volume (GB)'] / (df_tef.iloc[-1]['Traffic DL Volume (GB)'] + df_att.iloc[-1]['Traffic DL Volume (GB)'])), traffic_percentage_format)
    worksheet.write('G6', (df_tef.iloc[-1]['Traffic UL Volume (GB)'] / (df_tef.iloc[-1]['Traffic UL Volume (GB)'] + df_att.iloc[-1]['Traffic UL Volume (GB)'])), traffic_percentage_format)
    worksheet.write('G7', (df_tef.iloc[-1]['WBB traffic(GB)'] / (df_tef.iloc[-1]['WBB traffic(GB)'] + df_att.iloc[-1]['WBB traffic(GB)'])), traffic_percentage_format)
    worksheet.write('G8', (df_tef.iloc[-1]['Total MOCN traffic (GB)'] / (df_tef.iloc[-1]['Total MOCN traffic (GB)'] + df_att.iloc[-1]['Total MOCN traffic (GB)'])), mocn_percentage_format)

    worksheet.write('G9', '', traffic_percentage_format)
    worksheet.write('G10', '', traffic_percentage_format)
    worksheet.write('G11', '', traffic_percentage_format)
    worksheet.write('G12', '', traffic_percentage_format)
    worksheet.write('G13', '', traffic_percentage_format)
    worksheet.write('G14', '', traffic_percentage_format)
    worksheet.write('G15', '', traffic_percentage_format)
    worksheet.write('G16', '', traffic_percentage_format)
    worksheet.write('G17', '', traffic_percentage_format)
    worksheet.write('G18', '', traffic_percentage_format)
    worksheet.write('G19', '', traffic_percentage_format)
    worksheet.write('G20', '', traffic_percentage_format)
    worksheet.write('G21', '', traffic_percentage_format)
    worksheet.write('G22', '', traffic_percentage_format)
    worksheet.write('G23', '', traffic_percentage_format)
    worksheet.write('G24', '', traffic_percentage_format)
    worksheet.write('G25', '', traffic_percentage_format)
    worksheet.write('G26', '', traffic_percentage_format)
    worksheet.write('G27', '', traffic_percentage_format)

    writer.sheets['Summary'] = worksheet
    return

def get_cell_config_integrity(df):
    list_cell_integrity = []
    date_list = df['Start_Time'].unique().tolist()

    before_df = df.loc[df['Start_Time'] == date_list[0]]
    after_df = df.loc[df['Start_Time'] == date_list[-1]]
    list_cell_integrity = [before_df['CellName'].nunique(), after_df['CellName'].nunique()]
    del date_list
    del before_df
    del after_df

    return list_cell_integrity