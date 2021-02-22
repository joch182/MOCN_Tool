from tkinter import Menu
from tkinter import *
from tkinter.ttk import * 
from tkinter import messagebox, ttk
import lte_mocn_report, umts_mocn_report

input_data = ""
output_data = ""
lte_template = ""

root = Tk()
root.title("AT&T KPI TOOL")
root.geometry('580x220')
menubar = Menu(root)
tab_parent = ttk.Notebook(root)

tab1 = ttk.Frame(tab_parent)
tab_parent.add(tab1, text="UMTS MOCN")

tab2 = ttk.Frame(tab_parent)
tab_parent.add(tab2, text="LTE MOCN")

tab_parent.pack(expand=1, fill='both')

# Tab UMTS MOCN
GETUMTSTemplate_Button = Button(tab1, text="Get U2000 Templates", style='Send.TButton', command=umts_mocn_report.download_umts_raw_counters_template)
GETUMTSTemplate_Button.grid(row = 0, column = 0, pady = 20, padx = 10)

UMTS_daily_report_input_Button = Button(tab1, text="Input UMTS daily report (csv)", command=umts_mocn_report.input_umts_daily_report_button)
UMTS_daily_report_input_Button.grid(row = 1, column = 0, padx = 10) 

UMTS_daily_report1_input_Button = Button(tab1, text="Input UMTS daily report1 (csv)", command=umts_mocn_report.input_umts_daily_report1_button)
UMTS_daily_report1_input_Button.grid(row = 1, column = 1, padx = 10) 

UMTS_MOCN_data_Button = Button(tab1, text="Input UMTS MOCN (csv)", command=umts_mocn_report.input_umts_mocn_button)
UMTS_MOCN_data_Button.grid(row = 1, column = 2, padx = 10)

UMTS_MOCN_OutputDir_Button = Button(tab1,text="Output Directory", command=umts_mocn_report.output_button)
UMTS_MOCN_OutputDir_Button.grid(row = 2, column = 1, pady = 20, padx = 10)

UMTS_MOCN_Analyze_Button = Button(tab1, text="UMTS MOCN Report", style='Send.TButton', command=umts_mocn_report.generate_umts_mocn_delivery_report_button)
UMTS_MOCN_Analyze_Button.grid(row = 3, column = 0, pady = 20, padx = 10)

# Tab UMTS MOCN
GETLTETemplate_Button = Button(tab2, text="Get U2000 Templates", style='Send.TButton', command=lte_mocn_report.download_lte_raw_counters_template)
GETLTETemplate_Button.grid(row = 0, column = 0, pady = 10, padx = 10)

LTE_daily_report_input_Button = Button(tab2, text="Input LTE daily report (csv)", command=lte_mocn_report.input_lte_daily_report_button)
LTE_daily_report_input_Button.grid(row = 1, column = 0, padx = 10) 

UMTS_daily_report1_input_Button = Button(tab2, text="Input LTE daily report1 (csv)", command=lte_mocn_report.input_lte_daily_report1_button)
UMTS_daily_report1_input_Button.grid(row = 1, column = 1, padx = 10) 

LTE_MOCN_data_Button = Button(tab2, text="Input LTE MOCN (csv)", command=lte_mocn_report.input_lte_mocn_button)
LTE_MOCN_data_Button.grid(row = 1, column = 2, padx = 10)

# LTE_daily_report_input_Button = Button(tab2, text="Input LTE NBIOT MOCN report (csv)", command=lte_mocn_report.input_lte_nbiot_mocn_report_button)
# LTE_daily_report_input_Button.grid(row = 2, column = 0, pady = 10,  padx = 10) 

# LTE_nbiot_report1_input_Button = Button(tab2, text="Input LTE NBIOT report (csv)", command=lte_mocn_report.input_lte_nbiot_report1_button)
# LTE_nbiot_report1_input_Button.grid(row = 2, column = 1, pady = 10,  padx = 10) 

LTE_emtc_Button = Button(tab2, text="Input LTE eMTC report (csv)", command=lte_mocn_report.input_lte_emtc_button)
LTE_emtc_Button.grid(row = 2, column = 2, pady = 10, padx = 10)

LTE_MOCN_OutputDir_Button = Button(tab2,text="Output Directory", command=lte_mocn_report.output_button)
LTE_MOCN_OutputDir_Button.grid(row = 3, column = 1, pady = 10, padx = 10)

LTE_MOCN_Analyze_Button = Button(tab2, text="LTE MOCN Report", style='Send.TButton', command=lte_mocn_report.generate_lte_mocn_delivery_report_button)
LTE_MOCN_Analyze_Button.grid(row = 4, column = 0, pady = 10, padx = 10)

root.config(menu=menubar)
root.mainloop()