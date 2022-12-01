# Python Inbuilt Functions
import re
import base64
import datetime
import pdfplumber
import os, shutil
from base64 import b64decode
from datetime import datetime
import openpyxl
import psycopg2
import dateutil.relativedelta
from dateutil.relativedelta import relativedelta
#Developed Functions
from pdf_parser import data_extractor_numbers,data_extractor_alphanumeric,data_extractor_string


conn= psycopg2.connect(database="Tata_Power", user='postgres',password='shubham1',host='localhost',port='5432')
cursor=conn.cursor()


path = r"D:\TATA POWER PDF\MPL"
os.chdir(path)

def add_sub_value(list_value):
    subtract_value = float(0)
    add_value = float(0)
    for x in list_value:
    	if x != '':
	        if float(x) < 0:
	            subtract_value+=float(x)
	        elif float(x) >= 0:
	            add_value+=float(x)
    return float(subtract_value), float(add_value)

def trigger_wbsedcl(path):
##    start_t = datetime.datetime.now()
##    start_t = datetime.datetime.now()
##    start_t = datetime.datetime.now()
    pdf = pdfplumber.open(path)
    text = ''
    for i in pdf.pages:
        text1=i.extract_text()
        text = text+text1
##    pdf = pdfplumber.open(path)
##    
##    page = pdf.pages[0]
##    
##    page1 = pdf.pages[1]
##    text1 = page.extract_text()
##    text2 = page1.extract_text()
##    text = text1 + text2
    data_dict = {}
    l = ['(', ')', '.', '/', '-', ',', '%']

##    print(text)

    mapping_sheet = 'MPL'
    start_time= datetime.now().time()
##
##    data_dict['Power at Delhi']=''
##    data_dict['Spare 1']=''
##    data_dict['Spare 2']=''
##    data_dict['Spare 3']=''
    data_dict['PPAC']='ISGS'
##    data_dict['FY']=''
##    data_dict['Quarter']=''
##    data_dict['Exclusions']=''
##    data_dict['Exclusion Classification']=''
##    data_dict['Classification']=''
    data_dict['TreasuryReco']='TPTCL'
    data_dict['MISHEAD']='MPL'
    data_dict['Particulars_PartyName']='Tata Power Trading Company Limited'
##    
##    data_dict['Board']=''
##    data_dict['For Energy Balance']=''
##    data_dict['For Energy Balance']=''

    data_dict['NameofStation']='Maithon Power'
    
    data_extractor_numbers(text,'Bill No.',1,data_dict,'\n','Bill No',l,'[0-9]{11}', 0)
##    data_dict['Nature of Invoice']=''
    
    data_extractor_numbers(text,'Bill Date',1,data_dict,'\n','Bill Date',l,'\d+\.\d+\.\d+', 0)

    RM = datetime.strptime(data_dict['Bill Date'], "%d.%m.%Y").month
    now = (data_dict['Bill Date'])
    now=datetime.strptime(now,'%d.%m.%Y').date()
    last_month = now - relativedelta(months=1)
    last_month  = last_month.strftime('%B-%y')
    data_dict['PertainingBillMonth'] = last_month
    data_dict['Month'] = last_month.split('-')[0]

    data_extractor_numbers(text,'Period',1,data_dict,'\n','Bill period from',l,'\d+\.\d+\.\d+', 0)
    data_extractor_numbers(text,'Period',1,data_dict,'\n','Bill period to',l,'\d+\.\d+\.\d+', 1)
############################------------------------Natureofinvoice-------------------------------------------#############################################
##    March21 = datetime.datetime(2021,3,1)
##    March20 = datetime.datetime(2020,3,1)
##
##    if datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y') >= March21:
##    	data_dict['NatureofInvoice'] = 'Current Month Bill'
##    elif datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y') < March21 or datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y') > March20:
##    	data_dict['NatureofInvoice'] ='CY Arrear'
##    elif datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y') < March20:
##    	data_dict['NatureofInvoice'] =  "Previous Arrear Bill"
##    # print(data_dict['NatureofInvoice'])

    current_year=datetime.now().year
    apr21=datetime(current_year,4,1)
    # print('Apr21: ',apr21)

    next_year=int(datetime.now().year)+1
    mar22=datetime(next_year,3,31)
    # print('Mar22: ',mar22)

    current_year2=int(datetime.now().year)-1
    apr20=datetime(current_year2,4,1)
    # print('Apr20: ',apr20)

    next_year2=datetime.now().year
    mar21=datetime(next_year2,3,1)
    print('Mar21: ',mar21,'.....',type(mar21))


##    print(data_dict['to'],'..........',type(datetime.strptime(data_dict['to'],'%d.%m.%Y')))
    if mar21<= datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y') <= mar22:
        data_dict['NatureofInvoice'] = 'Current Month Bill'
    elif apr20 <= datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y') <mar21:
        data_dict['NatureofInvoice'] = 'CY Arrear ' 
    elif datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y') < apr20:
        data_dict['NatureofInvoice'] = 'PY Arrear'
####################################################################################################################
    if data_dict['NatureofInvoice'] == 'Current Month Bill':
    	Power = "Power at Delhi"
    else:
    	Power = 'Arrears'

    data_dict['PoweratDelhi'] = Power
###################_____________________FY______________________________________################################
##    year = datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y').year
##
##    if datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y').year >= datetime.datetime(2021,1,1).year:
##        FY = 'FY 21-22'
##    elif datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y').year == datetime.datetime(2020,1,1).year:
##        FY = 'FY 20-21'
##    else:
##        FY = 'FY 19-20'
##    data_dict['FY'] = FY
    current_year=datetime.now().year
    apr21=datetime(current_year,4,1)
    # print('Apr21: ',apr21)

    next_year=int(datetime.now().year)+1
    mar22=datetime(next_year,3,31)
    # print('Mar22: ',mar22)

    current_year2=int(datetime.now().year)-1
    apr20=datetime(current_year2,4,1)
    # print('Apr20: ',apr20)

    next_year2=datetime.now().year
    mar21=datetime(next_year2,3,31)
    print('Mar21: ',mar21,'.....',type(mar21))

    # t=datetime.strptime('2021-04-01','%Y-%m-%d')

##    print(data_dict['Invoice_Date'],'..........',type(datetime.strptime(data_dict['to'],'%d.%m.%Y')))
    if apr21<= datetime.strptime(data_dict['Bill Date'],'%d.%m.%Y') - dateutil.relativedelta.relativedelta(months=1) <= mar22:
        data_dict['FY'] = 'FY 21-22'
    elif apr20 <= datetime.strptime(data_dict['Bill Date'],'%d.%m.%Y') - dateutil.relativedelta.relativedelta(months=1) <=mar21:
        data_dict['FY'] = ' FY 20-21' 
    elif datetime.strptime(data_dict['Bill Date'],'%d.%m.%Y') - dateutil.relativedelta.relativedelta(months=1) < apr20:
        data_dict['FY'] = 'FY 19-20'
#####################################################################################################

    if 'Power at Delhi' in data_dict['PoweratDelhi']:
    	data_dict['Classification'] = 'TPDDL Periphery'
    elif 'Arrear' in data_dict['PoweratDelhi']:
    	data_dict['Classification'] = 'Others'

    if 'Current Month Bill' in data_dict['NatureofInvoice']:
    	data_dict['Board'] = 'ISGS'

    else:
    	data_dict['Board'] = 'Others'


    if 'Current Month Bill' in data_dict['NatureofInvoice']:
    	data_dict['ForEnergyBalance'] = 'MPL'
    else:
    	data_dict['ForEnergyBalance'] = 'Arrears'

    if "April"  in data_dict['Month'] or "May" in data_dict['Month'] or "June" in data_dict['Month']:
        data_dict['Quarter'] = 'Q1'
    elif "July" in data_dict['Month'] or "August" in data_dict['Month'] or "September" in data_dict['Month']:
        data_dict['Quarter']= 'Q2'
    elif "October" in data_dict['Month'] or "November" in data_dict['Month'] or "December" in data_dict['Month']:
        data_dict['Quarter']= 'Q3'
    elif "January" in data_dict['Month'] or "February" in data_dict['Month'] or "March" in data_dict['Month']:
        data_dict['Quarter']= 'Q4'


   
    data_extractor_numbers(text,'Schedule Generation',1,data_dict,'\n','Units Coal',l,'\d+.*', 0)
    data_dict['Units Hydro']=''

    data_extractor_numbers(text,'Energy Charge',1,data_dict,'\n','Variable Cost',l,'\d+.*', 0)
    data_dict['Variable Cost Hydro']=''
    data_extractor_numbers(text,'Fixed Charge',1,data_dict,'\n','Fixed Cost',l,'\d+.*', 0)
    data_dict['Fixed Cost (Peak)']=''
    try:
        rra=data_extractor_numbers(text,'Differential Amount Amount  Fixed charges refund Rs. ',1,data_dict,'\n','RRAS Charges',l,'[0-9\,\.\-]+', 0).replace(',',"")
    except:
        rra=data_extractor_numbers(text,'Differential Amount Amount  Fixed charges refund Rs. ',1,data_dict,'\n','RRAS Charges',l,'[0-9\,\.\-]+', 0)
    print(rra)
    data_dict['RRAS Charges']=rra
    data_extractor_numbers(text,'Incentive Charges',1,data_dict,'\n','Incentive',l,'\d+.*', 0)
    data_dict['Incentive (off peak)']=0

    #70
    #71
    try:
        other=data_extractor_numbers(text,'SOC & MOC Charges',1,data_dict,'\n','Other',l,'\d+.*', 0).replace(',',"")
    except:
        other=data_extractor_numbers(text,'SOC & MOC Charges',1,data_dict,'\n','Other',l,'\d+.*', 0)
    data_dict['Other']=other

##    try:
    Rtm=data_extractor_numbers(text,'Sharing of Gains on account of RTM power sales Rs.',1,data_dict,'\n','RTM Trade Gain Share',l,'[0-9\,\.\-]+', 0)
    

    data_dict['RTM Trade Gain Share']=Rtm
    
    try:
        trading=data_extractor_numbers(text,'Trading Margin',1,data_dict,'\n','Trading Margin',l,'\d+.*', 0).replace(',',"")
    except:
        trading=data_extractor_numbers(text,'Trading Margin',1,data_dict,'\n','Trading Margin',l,'\d+.*', 0)

    data_dict['Trading Margin']=trading
    
    try:
        Total=data_extractor_numbers(text,'Bill Amount Rs.',1,data_dict,'\n','Total',l,'[0-9\,\.\-]+', 0).replace(',',"")
    except:
        Total=data_extractor_numbers(text,'Bill Amount Rs.',1,data_dict,'\n','Total',l,'[0-9\,\.\-]+', 0)
        
    data_dict['Total']=Total
    #119
    #121
    #137
    #Sharing of Gains on account of RTM power sales

    #############################################    ADDITION    ###################################################
    try:
        TotalUnitsbilled = [str(data_dict['Units Coal'].replace(",", "")), str(data_dict['Units Hydro'])]
        value1 = add_sub_value(TotalUnitsbilled)
    except:
        TotalUnitsbilled = [str(data_dict['Units Coal']), str(data_dict['Units Hydro'])]
        value1 = add_sub_value(TotalUnitsbilled)
    data_dict['TotalUnitsbilled'] = value1[0]+value1[1]

    

    try:
        VC = [str(data_dict['Variable Cost'].replace(",", "")), str(data_dict['Variable Cost Hydro'].replace(",", ""))]#, data_dict['Energy_Charges_Liquid_1'], data_dict['Energy_Charges_NAPM_Gas_1'], data_dict['Energy_Charges_Committed_Gas_1'], data_dict['Energy_Charges_Committed_LNG_1'], data_dict['EC_Adjustment_on_Actual_1'], data_dict['Variable_Cost'], data_dict['Variable_Cost_Hydro']]

        value1 = add_sub_value(VC)
        data_dict['VC'] = value1[0]+value1[1]
    except:
        VC = [data_dict['Variable Cost'], data_dict['Variable Cost Hydro']]
        value1 = add_sub_value(VC)
        data_dict['VC'] = value1[0]+value1[1]

    try:
        FC = [str(data_dict['Fixed Cost'].replace(',', "")), str(data_dict['Fixed Cost (Peak)']), str(data_dict['RRAS Charges'])]#, data_dict['Energy_Charges_Liquid_1'], data_dict['Energy_Charges_NAPM_Gas_1'], data_dict['Energy_Charges_Committed_Gas_1'], data_dict['Energy_Charges_Committed_LNG_1'], data_dict['EC_Adjustment_on_Actual_1'], data_dict['Variable_Cost'], data_dict['Variable_Cost_Hydro']]

        value1 = add_sub_value(FC)
        data_dict['FC'] = value1[0]+value1[1]
    except:
        FC = [data_dict['Fixed Cost'], data_dict['Fixed Cost (Peak)'],data_dict['RRAS Charges']]
        value1 = add_sub_value(FC)
        data_dict['FC'] = value1[0]+value1[1]

    try:
        incentive = [str(data_dict['Incentive'].replace(',', "")), str(data_dict['Incentive (off peak)'])]#, data_dict['Energy_Charges_Liquid_1'], data_dict['Energy_Charges_NAPM_Gas_1'], data_dict['Energy_Charges_Committed_Gas_1'], data_dict['Energy_Charges_Committed_LNG_1'], data_dict['EC_Adjustment_on_Actual_1'], data_dict['Variable_Cost'], data_dict['Variable_Cost_Hydro']]

        value1 = add_sub_value(incentive)
        data_dict['Incentive_'] = value1[0]+value1[1]
    except:
        incentive = [data_dict['Incentive'], data_dict['Incentive (off peak)']]
        value1 = add_sub_value(incentive)
        data_dict['Incentive_'] = value1[0]+value1[1]

    try:
        Others = [data_dict['Other'], data_dict['Trading Margin'], data_dict['RTM Trade Gain Share']]#, data_dict['Energy_Charges_Liquid_1'], data_dict['Energy_Charges_NAPM_Gas_1'], data_dict['Energy_Charges_Committed_Gas_1'], data_dict['Energy_Charges_Committed_LNG_1'], data_dict['EC_Adjustment_on_Actual_1'], data_dict['Variable_Cost'], data_dict['Variable_Cost_Hydro']]

        value1 = add_sub_value(Others)
        data_dict['Others_'] = value1[0]+value1[1]
    except:
        Others = [data_dict['Other'], data_dict['Trading Margin'], data_dict['RTM Trade Gain Share'].replace(',','')]#, data_dict['Energy_Charges_Liquid_1'], data_dict['Energy_Charges_NAPM_Gas_1'], data_dict['Energy_Charges_Committed_Gas_1'], data_dict['Energy_Charges_Committed_LNG_1'], data_dict['EC_Adjustment_on_Actual_1'], data_dict['Variable_Cost'], data_dict['Variable_Cost_Hydro']]

        value1 = add_sub_value(Others)
        data_dict['Others_'] = value1[0]+value1[1]

    lst = [data_dict['VC'], data_dict['FC'],  data_dict['Others_'], data_dict['Incentive_']]
    check = sum(lst)
    print(check)

##    data_dict['Total']=(data_dict['Total']).replace(',',"")

    checkum=(float(check)-float(Total))
    print(checkum)
    data_dict['check']=checkum

    ###########################################################################################################

    if re.search(r'Current.*?Bill',data_dict['NatureofInvoice']):
        pass
    else:
        data_dict['TotalUnitsbilled'] = ''

##############################################################################################
    if 'Current Month Bill' in data_dict['NatureofInvoice']:
    	data_dict['Remarks2'] = "Monthly energy bill for "+data_dict['PertainingBillMonth']
    else:
    	data_dict['Remarks2'] = 'Revision Bill'

########################################################################################
    for key1,value1 in data_dict.items():
        if value1=='0' or value1==0 or value1==0.0 or value1== ' 0' or value1 == '0 ' or value1 == '':
            data_dict[key1]=''
    
    end_time= datetime.now().time()

##    query = "insert into  pgcil values('{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}')".format(mapping_sheet, start_time,end_time, '','','','','','','','','','','','','','','','', data_dict['Bill No'], '', data_dict['Bill Date'],data_dict['Bill period from'], data_dict['Bill period to'], '','',data_dict['Units Coal'], '','','','','','','','','','','','','','', data_dict['Variable Cost'], '','','','','','','','','','','','','','','','','','','','','','','','','',data_dict['Fixed Cost'],'','','','', data_dict['RRAS Charges'],'','','','','',data_dict['Incentive'],'','','','','','','',data_dict['Other'],'','',data_dict['RTM Trade Gain Share'],'','',data_dict['Trading Margin'],'','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','')
    query = "insert into Tata_Power values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
    values=(data_dict['PoweratDelhi'],'','','',data_dict['PPAC'],data_dict['FY'],data_dict['Quarter'],'','',data_dict['Classification'],data_dict['TreasuryReco'],data_dict['MISHEAD'],data_dict['Particulars_PartyName'],data_dict['Board'],data_dict['ForEnergyBalance'],data_dict['NameofStation'],data_dict['Bill No'],data_dict['NatureofInvoice'],data_dict['Bill Date'],data_dict['Bill period from'],data_dict['Bill period to'],data_dict['PertainingBillMonth'],data_dict['Month'],data_dict['Units Coal'],'','','','','','','','','','','','','',data_dict['TotalUnitsbilled'],data_dict['Variable Cost'],'', '','','','','','','','','','','','','','','','','','','','','','','',data_dict['VC'], data_dict['Fixed Cost'],'','','','',data_dict['RRAS Charges'],data_dict['FC'],'','','','','',data_dict['Incentive'],'','',data_dict['Incentive_'],'','','','',data_dict['Other'],'','',data_dict['RTM Trade Gain Share'],'','',data_dict['Trading Margin'],'','','','', '','', '','','','','','','','','','','','','','','','','','','','','','','','',data_dict['Others_'],'','','','','','','','','','','','','',data_dict['Remarks2'],data_dict['check'],data_dict['Total'],'','','','','','','','','','','','','','','','', '', '','','','','','', '','', '','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','', '','','','','','','','','','','','','','','','','','','','','','','','', '','','','','','', '','','','','','','','','','','','', '','','','','','','','','','','','','','', '','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','',mapping_sheet,start_time,end_time)  
    cursor.execute(query,values)   
    conn.commit()
    print("record inserted")
  
                                            
    print("\n")
    print(data_dict)


#print(trigger_wbsedcl(r"D:\TATA POWER PDF\MPL\MPL Bill Jun_21 (TPTCL Format).pdf"))

##for file in os.listdir(path):
##    try:
##        print(file)
##        trigger_wbsedcl(file)
##    except:
##        pass
##
##path = r"D:\TATA POWER PDF\IPGCL"
##os.chdir(path)
for i in os.listdir(path):
    print(i)
    if '.txt' in i or '.pdf' in i or '.PDF' in i:
        i = os.path.join(path, i)
        trigger_wbsedcl(i)

        


            
        



