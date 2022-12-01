# Python Inbuilt Functions
import re
import base64
import datetime
import pdfplumber
import os, shutil
from base64 import b64decode
import openpyxl
from datetime import datetime
import psycopg2
import dateutil.relativedelta
from dateutil.relativedelta import relativedelta
#Developed Functions
from pdf_parser import data_extractor_numbers,data_extractor_alphanumeric,data_extractor_string


conn= psycopg2.connect(database="Tata_Power", user='postgres',password='shubham1',host='localhost',port='5432')
cursor=conn.cursor()

path = r"D:\TATA POWER PDF\CLP"
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
    start_t = datetime.now()
    
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

    print(text)

    mapping_sheet = 'CLP'
    start_time= datetime.now().time()

##    data_dict['Power at Delhi']=''
##    data_dict['Spare 1']=''
##    data_dict['Spare 2']=''
##    data_dict['Spare 3']=''
    
##    data_dict['FY']=''
##    data_dict['Quarter']=''
##    data_dict['Exclusions']=''
##    data_dict['Exclusion Classification']=''
##    data_dict['Classification']=''
    data_dict['TreasuryReco']='TPTCL'
    
    data_dict['Particulars_PartyName']='Tata Power Trading Company Limited'
    
##    data_dict['Board']=''
##    data_dict['For Energy Balance']=''

    

    	
##    data_extractor_string(text,'Source of Power',1,data_dict,'\n','Name of Station',l,'[A-Z].*', 0)
    data_dict['NameofStation']='CLP Jhajjar'
    
    data_extractor_numbers(text,'Bill No.',1,data_dict,'Bill','Bill No',l,'[0-9]{11}', 0)
##    data_dict['Nature of Invoice']=''
    data_extractor_numbers(text,'Bill Date',1,data_dict,'Due','BillDate',l,'\d+\.\d+\.\d+', 0)

##    March21 = datetime.datetime(2021,3,1)
##    March20 = datetime.datetime(2020,3,1)

    RM = datetime.strptime(data_dict['BillDate'], "%d.%m.%Y").month
    now = (data_dict['BillDate'])
    now=datetime.strptime(now,'%d.%m.%Y').date()
    last_month = now - relativedelta(months=1)
    last_month  = last_month.strftime('%B-%y')
    data_dict['PertainingBillMonth'] = last_month
    data_dict['Month'] = last_month.split('-')[0]
    
            
    data_extractor_numbers(text,'Bill for Period',1,data_dict,'\n','Bill period from',l,'\d+\.\d+\.\d+', 0)
    data_extractor_numbers(text,'Bill for Period',1,data_dict,'\n','Bill period to',l,'\d+\.\d+\.\d+', 1)
##########################-------------------Natureof invoice-------------------------########################
##    if datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y') >= March21:
##        data_dict['NatureofInvoice'] = 'Current Month Bill'
##    elif datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y') < March21 or datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y') > March20:
##        data_dict['NatureofInvoice'] ='CY Arrear'
##    elif datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y') < March20:
##        data_dict['NatureofInvoice'] =  "Previous Arrear Bill"

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

################################################################################################################    
    data_extractor_numbers(text,'Schedule Generation',1,data_dict,'\n','unit Coal',l,'\d.*', 0)
    data_dict['Units Hydro']=''
    data_dict['Units Hydro Beyond DE']=''
    
    data_dict['Units APM']=''
    data_dict['Units RLNG']=''
    data_dict['Units Liquid']=''
    data_dict['SG NAPM Gas']=''
    data_dict['SG Committed LNG']=''
    data_dict['SG Committed Gas']=''
    data_dict['Units Solar']=''

##    total_unit= [str(float(data_dict['unit Coal'])),str(float(data_dict['Units Hydro']))]#,(float(data_dict['Units Hydro Beyond DE'])),(float(data_dict['Units APM'])),(float(data_dict['Units RLNG'])),(float(data_dict['Units Liquid'])),(float(data_dict['SG NAPM Gas'])),(float(data_dict['SG Committed LNG'])),(float(data_dict['SG Committed Gas'])),(float(data_dict['Units Solar']))]
##    value1 = add_sub_value(total_unit)
##    value1 = float(value1[0]) + float(value1[1])
####    if value1 == 0.0:
####        data_dict['TotalUnitsbilled']= ''
####    else:
##    data_dict['TotalUnitsbilled'] = value1
   
##    data_dict['Variable Cost Hydro beyond DE']=''
##    
##    data_dict['Energy Charges Gas']=''
##    data_dict['Energy Charges LNG']=''
##    data_dict['Energy Charges Liquid']=''
##    data_dict['Energy Charges NAPM Gas']=''
##
##    data_dict['Energy Charges Committed Gas']=''
##    data_dict['Energy Charges Committed LNG']=''
##    data_dict['EC Adjustment on Actual']=''
##    data_dict['Bilateral']=''
##    data_dict['Banking']=''
##    data_dict['IEX Purchase']=''
##    data_dict['IEX Sale']=''
##    data_dict['UI DSM']=''
##    data_dict['UI ADSM']=''
##    data_dict['UI SDSM']=''
##    data_dict['Solar - Purchase']=''
##    data_dict['Solar Sale']=''
##    data_dict['IDT Sale']=''
##    data_dict['Banking Sale']=''
##    data_dict['VC']=''
    
    data_extractor_numbers(text,'Energy Charge',1,data_dict,'\n','Variable Cost',l,'\d.*', 0)
    data_dict['Variable Cost Hydro']=''
    

                                                                    
    data_extractor_numbers(text,'Fixed Charge',1,data_dict,'\n','Fixed Cost',l,'\d.*', 0)
    data_dict['Fixed Cost (Peak)']=''
##    data_dict['Fixed Cost (Off-Peak)']=''
##    data_dict['Fixed Cost (Off-Set)']=''
##    data_dict['RRAS Charges']='' 
##    data_dict['FC']=''
    

##    data_dict['Carrying Cost']=''
##    data_dict['Interest Payable']=''
##    data_dict['Interest Receivable']=''
##    data_dict['Interest']=''
##    
##    data_dict['Incentive']='' #69
##    data_dict['Incentive (off peak)']=''
##    data_dict['Incentive (off set)']=''
##    data_dict['Incentive_']=''
##    data_dict['Income Tax']=''
##    data_dict['IEX Other Expenses']=''
##    data_dict['Exchange Fee']=''
##    data_dict['Pension Trust']=''
##

    if 'Unitary Charge' in text:
        
        data_extractor_numbers(text,'Unitary Charges for April 2021',1,data_dict,'\n','Other',l,'\d.*', 0)
    elif 'SOC Charges Rs' in text:
        data_extractor_numbers(text,'SOC Charges Rs',1,data_dict,'\n','Other',l,'\d.*', 0)
    else:
        data_dict['Other']=0
##    data_dict['FERV']=''
##    data_dict['Deferred Tax']=''
##    data_dict['RTM Trade Gain Share']=''
##    data_dict['Filing Fee']=''
##    data_dict['Non Tariff Income Share']=''

    
    data_extractor_numbers(text,'Trading Margin',1,data_dict,'\n','Trading Margin',l,'\d.*', 0)
##    data_dict['Water Charges']=''
##    data_dict['Annual Adjustment -SFC Price']=''
##    data_dict['Annual Adjustment -SFC saving']=''
##    data_dict['Addl O&M Expenses']=''
##    data_dict['URS Trade Gain Share']=''
##    data_dict['Gain on Actual Aux. Consumption']=''

##    data_dict['Transit Fee']=''
##    
##    data_dict['DMF']=''
##    
##    data_dict['NMET']=''
##    
##    data_dict['EDC']=''
##    
##    data_dict['Energy Development Cess']=''
##
##    data_dict['Refund of Abolished Electricity Duty']=''
##    
##
##    data_dict['Sustenance Allowance']=''
##    
##
##    data_dict['Free Education']=''
##    
##
##    data_dict['Old Age Pension']=''
##    
##
##    data_dict['GST on Royalty']=''
##    
##
##    data_dict['Compensation Charges(Th)']=''
##    data_dict['Compensation Charges(Gas)']=''
##    data_dict['Compensation Charges(LNG)']=''
##    data_dict['Compensation Charges(Liq)']=''
##    data_dict['Compensation Charges(NGas)']=''
##    data_dict['Compensation Charges(CLNG)']=''
##    data_dict['Compensation Charges(CGas)']=''
##
##    data_dict['MOPA Gas']=''
##    data_dict['MOPA LNG']=''
##    data_dict['MOPA Liq']=''
##    data_dict['MOPA NGas']=''
##    data_dict['MOPA CLNG']=''
##    data_dict['MOPA CGas']=''
##    data_dict['Others']=''
    


##
##    data_dict['Reactive Energy']=''
##    data_dict['Congestion Charges']=''
##    data_dict['RLDC Charges']=''
##    data_dict['Bilateral STOA']=''
##    data_dict['Banking STOA']=''
    
    data_extractor_numbers(text,'Transmissions Charges',1,data_dict,'\n','Transmissions',l,'\d+.*', 0)
##    data_dict['NRLDC (TDS)']=''
    
    data_extractor_numbers(text,'SLDC Charges',1,data_dict,'\n','SLDC',l,'\d.*', 0)
    if 'Bill Amount Rs' in text:
        
        try:
            Total=data_extractor_numbers(text,'Bill Amount Rs',1,data_dict,'\n','TotalInvoiceAmount',l,'\d.*', 0).replace(',',"")
        except:
   
                Total=data_extractor_numbers(text,'Bill Amount Rs',1,data_dict,'\n','TotalInvoiceAmount',l,'\d.*', 0)
    elif 'Amount Payable Rs.' in text:
        try:
            Total=data_extractor_numbers(text,'Amount Payable Rs.',1,data_dict,'\n','TotalInvoiceAmount',l,'\d.*', 0).replace(',','')
        except:
            Total=data_extractor_numbers(text,'Amount Payable Rs.',1,data_dict,'\n','TotalInvoiceAmount',l,'\d.*', 0).replace(',','')
            
    data_dict['TotalInvoiceAmount']=Total
##    print(total)

    #data_dict['SLDC']=''
##    data_dict['STOA']=''
##    data_dict['Application Fees']=''
##    data_dict['Open Access (Reimb.)']=''
##    data_dict['Transmission Charges']=''
    
##    data_dict['Remarks']=''
##    data_dict['Due Date']=''
##    data_dict['Contracted Capacity of Plant']=''
##    data_dict['Allocated Contracted Capacity %']=''
##    data_dict['Allocated Contracted Capacity (MW)']=''
##    data_dict['Capacity Charge for the Settlement Period']='' #134
##    data_dict['Capacity Charges cumulative']=''#135
##    data_dict['Capacity Charges cumulative (Recoverable)']=''#136
##    data_dict['AFC']=''#137
##    data_dict['Capacity Charges Already Billed']=''#138
##    data_dict['Fixed charges recoverable ']=''#139
##    data_dict['Date of Commercial Operation']=''
##    data_dict['Maximum possible Energy Corresponding to Allocated Contracted Capacity for the month']='' #141
##    data_dict['Allocated Energy corresponding to Declared Capability by SPL till last day of the month']=''#142
##    data_dict['Actual availability for Settlement Period']=''
##    data_dict['Energy scheduled during the bill period']=''
##    
##    data_dict['Scheduled Energy in the Month']=''#145
##    data_dict['ECR Coal']=''
##    data_dict['ECR Hydro (Normative)']=''
##    data_dict['ECR Hydro (Actual)']=''
##    data_dict['ECR Gas APM']=''
##    data_dict['ECR RLNG']=''
##    data_dict['ECR CRLNG']=''
##    data_dict['ECR Liquid']=''
##    data_dict['ECR NAPM']=''
##    data_dict['ECR Committed Gas']=''
##    data_dict['ECR_Gas APM']=''
##    data_dict['ECR_RLNG']=''
##    data_dict['ECR_CRLNG']=''
##    data_dict['ECR_Liquid']=''
##    data_dict['ECR_NAPM']=''
##    data_dict['ECR_Committed_Gas']=''
##
##    data_dict['RLDC Charges for the plant']=''
##
##    
##    data_dict['PAFM']=''#162
##    data_dict['PAFN']=''#163
##
##    data_dict['NAPAF']=''
##    data_dict['Cummulative days in current FY']=''
##    data_dict['Demand Season']=''
##    data_dict['Demand Month']=''
##    data_dict['PAFN:(Peak- Low)']=''
##    data_dict['PAFN:(Off Peak- Low)']=''
##    data_dict['PAFN:(Peak- High)']=''
##    data_dict['PAFN:(Off Peak- High)']=''
##
##    data_dict['Monthly Entitlement']=''
##    data_dict['Cumulative Entitlement']=''
##    data_dict['Mthly Entitlement(Peak)']=''
##    data_dict['Mthly Entitlement(Off-Peak)']=''
##    data_dict['Inc Energy Peak Low (Cum)']=''
##    data_dict['Inc Energy-OffPeak Low (Cum)']=''
##    data_dict['Inc Energy-Peak High (Cum)']=''
##    data_dict['Inc Energy-OffPeak High (Cum)']=''
##    data_dict['GHR Gas']=''
##    data_dict['AUX Gas']=''
##    data_dict['GHR Gas O']=''
##    data_dict['Aux Gas O']=''
##    data_dict['LPPF Gas']=''
##    data_dict['CVPF Gas']=''
##    data_dict['POCM_G']=''
##    data_dict['GHR Liquid']=''
##    data_dict['Aux Liquid']=''
##    data_dict['GHR Liquid O']=''
##    data_dict['Aux Liquid O']=''
##    data_dict['LPPF Liquid']=''
##    data_dict['CVPF Liquid']=''
##    data_dict['POCM_LQ']=''
##
##    data_dict['GHR RLNG']=''
##    data_dict['Aux RLNG']=''
##
##    data_dict['GHR RLNG O']=''
##    data_dict['Aux RLNG O']=''
##    data_dict['LPPF RLNG']=''
##    data_dict['CVPF RLNG']=''
##    data_dict['POCM_LNG']=''
##    data_dict['GHR NAPM Gas']=''
##
##    data_dict['Aux NAPM Gas']=''
##    data_dict['GHR NAPM Gas O']=''
##    data_dict['Aux NAPM Gas O']=''
##    data_dict['LPPF NAPM Gas']=''
##    data_dict['CVPF NAPM Gas']=''
##
##    data_dict['POCM_NAPM GAS']=''
##    data_dict['GHR Committed LNG']=''
##    data_dict['Aux Committed LNG']=''
##    data_dict['GHR Committed LNG O']=''
##    data_dict['Aux Committed LNG O']=''
##    data_dict['LPPF Committed LNG']=''
##    data_dict['CVPF Committed LNG']=''
##    data_dict['POCM_Committed LNG']=''
##    data_dict['GHRCommittedGas']=''
##    data_dict['AuxCommittedGas']=''
##    data_dict['GHRCommittedGasO']=''
##    data_dict['AuxCommittedGasO']=''
##    data_dict['LPPFCommittedGas']=''
##
##    data_dict['CVPFCommittedGas']=''
##    data_dict['POCM_CommittedGas']=''
##    data_dict['AUL_DC']=''
##    data_dict['AUL_SG']=''
##    data_dict['Station Cumulative SG']=''
##
##    data_dict['Benef. En. Req(Below 85%)']=''
##    data_dict['Statio En. Req(Below 85%)']=''
##    data_dict['LPPF(Cumulative)']=''
##
##    
##    data_dict['CVPF(Cumulative)']=''#228
##    data_dict['CVPF (Domestic)']=''#229
##    data_dict['CVPF (Imported)']=''#230
##
##    
##    
##    data_dict['Cumulative ECR(Normative)']=''
##    data_dict['Cumulative ECR(Actual)']=''
##    data_dict['Cumulative ECR(DC)']=''
##    data_dict['Cumulative ECR(SG)']=''
##    data_dict['GHR Actual(Cumulative)']=''
##    data_dict['AUX Actual(Cumulative)']=''
##
##
##    data_dict['Station Cumulative_SG']=''
##    data_dict['Benef. En. Req(Below_85%)']=''
##    data_dict['Statio En. Req(Below_85%)']=''
##    data_dict['LPPF_(Cumulative)']=''
##    data_dict['CVPF_(Cumulative)']=''
##    data_dict['Cumulative_ECR(Normative)']=''
##
##    data_dict['Cumulative_ECR(Actual)']=''
##    data_dict['Cumulative_ECR(DC)']=''
##    data_dict['Cumulative_ECR(SG)']=''
##    data_dict['GHR Actual_(Cumulative)']=''
##    data_dict['AUX Actual_(Cumulative)']=''
##    data_dict['StationCumulativeSG__']=''
##    data_dict['Benef_En_Req_Below_85_%_']=''
##    data_dict['StatioEn_Req_Below_85%_']=''
##    data_dict['LPPF_Cumulative__']=''
##
##    data_dict['CVPF_Cumulative__']=''
##    data_dict['Cumulative_ECR_Normative__']=''
##
##    data_dict['Cumulative_ECR_Actual__']=''
##    data_dict['Cumulative_ECR_DC__']=''
##    data_dict['Cumulative_ECR_SG__']=''
##    data_dict['GHRActual_Cumulative__']=''
##
##    data_dict['AUXActual_Cumulative__']=''
##    data_dict['Station_CumulativeSG__']=''
##    data_dict['Benef_En_Req_Below_85%__']=''
##
##    data_dict['StatioEn_Req_Below_85%__']=''
##    data_dict['_LPPF_Cumulative']=''
##    data_dict['_CVPF_Cumulative']=''
##    data_dict['_Cumulative_ECR_Normative']=''
##    data_dict['_Cumulative_ECR_Actual']=''
##    data_dict['_Cumulative_ECR_DC']=''
##    data_dict['_Cumulative_ECR_SG']=''
##    data_dict['_GHRActual_Cumulative']=''
##
##    data_dict['_AUXActual_Cumulative']=''
##    data_dict['_StationCumulative_SG']=''
##    data_dict['_Benef_En_Req_Below_85%']=''
##    data_dict['_StatioEn_Req_Below_85%']=''
##    data_dict['_LPPF_Cumulative_']='' #273
##    data_dict['_CVPF_Cumulative_']='' #274
##    data_dict['_Cumulative_ECR_Normative_']=''
##    data_dict['_Cumulative_ECR_Actual_']=''
##
##    data_dict['_Cumulative_ECR_DC_']=''
##    data_dict['_Cumulative_ECR_SG_']=''
##    data_dict['_GHRActual_Cumulative_']='' #279
##    data_dict['_AUXActual_Cumulative_']=''
##    data_dict['_StationCumulative_SG_']=''
##    data_dict['_Benef_En_Req_Below_85%_']=''
##    data_dict['_StatioEn_Req_Below_85%_']=''
##    data_dict['LPPF__Cumulative']=''
##
##    data_dict['CVPF__Cumulative']=''
##    data_dict['CumulativeECR__Normative']=''
##    data_dict['CumulativeECR__Actual']=''
##    data_dict['CumulativeECR__DC']=''
##    data_dict['CumulativeECR__SG']=''
##    data_dict['GHRActual__Cumulative']=''
##
##    data_dict['AUXActual__Cumulative']=''
##    data_dict['TransmissionchargesRate']=''
##    data_dict['TAFM']=''
##    data_dict['Share']=''
##
##    data_dict['AUXCoal']=''
##    data_dict['GHRCoal']=''
##    data_dict['SFCCoal']=''
##
##    data_dict['IncentiveRate']=''
##
##    data_dict['IncentiveRate_Peak']=''
##    data_dict['IncentiveRate_Off_Peak']=''
##    data_dict['LPSF']=''
##    data_dict['CVSF']=''
##    data_dict['LPPF']=''
##
##    data_dict['CVPF_As_Received']=''
##    data_dict['CVPF_VAR']=''
##    data_dict['CVPF']=''
##    data_dict['AddlROERate_LDS']=''
##
##
##    data_dict['AddlROERate_HDS']=''
##
##    data_dict['EffectiveTaxRate']=''
##    data_dict['TotalSGforStation']=''
##
##    data_dict['ColonyConsumption']=''
##    data_dict['Const_CommissioningPower']=''
##    data_dict['Total']=''
##    data_dict['APC']=''
##
##    data_dict['CessRateonAPC']=''
##    data_dict['EDRateonAPC']=''
##    data_dict['SGforBeneficiary']=''
##
##    data_dict['CessonAPCforBeneficiary']=''
##    data_dict['EDonAPCforBeneficiary']=''
##    data_dict['StationCumulative_SG']=''
##    data_dict['Benef_En_Req_Below_85per']=''
##
##    data_dict['StatioEn_Req_Below_85per']=''
##    data_dict['LPPF__Cumulative__']=''
##
##    data_dict['CVPF__Cumulative__']=''
##    data_dict['LPSF__Cumulative__']=''
##    data_dict['Avg_LPL_Cumulative']=''
##    data_dict['CVSFCumulative']=''
##    data_dict['Cumulative__ECR_Normative']=''
##    data_dict['Cumulative__ECR_Actual']=''
##    data_dict['Cumulative__ECR_DC']=''
##
##    data_dict['Cumulative__ECR_SG']=''
##    data_dict['AUL_DC_']=''
##    data_dict['AUL_SG_']=''
##    data_dict['GHRActual__Cumulative_']=''
##    data_dict['AUXActual__Cumulative_']=''
##    data_dict['DateofCommercialOperation_COD']=''
##    data_dict['Projectage_P_Age']=''
##    data_dict['AnnualDE_ADE']=''
##
##    data_dict['AuxilliaryConsumption_Normative_AC_NOR']=''
##    data_dict['AuxilliaryConsumption_Actual_AC_ACT']=''
##    data_dict['AnnualFixedChargesBilled_AFC']=''
##    data_dict['NormativePlantAvailabilityFactor_NAPAF']=''
##    data_dict['SaleabeAnnualdesignenergy_SLDE']=''
##    data_dict['SaleabeAnnualdesignenergy_AC_Actu_SLDE_ACT']=''

##    data_dict['ProjectScheduledEnergyprevyear_PSCH_PY1']=''
##    data_dict['ProjectScheduledEnergyprevtoprevy_PSCH_PY2']=''
##    data_dict['EnergyChargeRate_AC_Normative']=''
##    data_dict['EnergyChargeRate_AC_Actual']=''
##
##    data_dict['SecondaryEnergyChargeRate']=''
##    data_dict['PlantAvailabilityFactorfortheMonth']=''
##    data_dict['SaleableDesignEnergyforthemonth']=''
##    data_dict['SaleableDesignEnergyforthemonth_AC']=''
##     
##    data_dict['Noofdaysforthemonth']=''#354
##    data_dict['ECR_NOR']=''
##    data_dict['ECR_ACT']=''
##    data_dict['SE_RATE1419']=''
##    data_dict['PAFM_']=''
##
##    data_dict['SLDEM']=''
##    data_dict['SLDEM_ACT']=''
##
##    data_dict['NDM']=''
##    data_dict['NDY']=''
##    data_dict['ScheduledEnergy']=''
##    data_dict['ProjectEnergyCharges_ECR']=''
##    
##    data_dict['FreeEnergy']=''
##    data_dict['Cumulative Capacity Charge till previous month']=''
##    data_dict['Cumulative Capacity charge till current month']=''
##    data_dict['Capacity Charges']=''
##    data_dict['SaleableEnergy']=''
##    data_dict['WaterUsageCharges']=''
##    data_dict['ProjectSaleableEnergyuptoDE']=''
##    data_dict['RLDCCharges_']=''
##    data_dict['ProjectSaleableEnergyuptoDE_AC']=''
##    data_dict['SaleableEnergyuptoDE_ECR']=''
##    data_dict['TotalCharges']=''
##    data_dict['BeneficiaryScheduledEnergy']=''
##    data_dict['SaleableEnergy_']=''
##    data_dict['BenifSaleableEnergy_ECR']=''
##    
##    
##    data_dict['Arrear']=''
##    data_dict['Non-Escalable Capacity Charges']=''
##    data_dict['Escalable Capacity Charges']=''
##    data_dict['Subtotal Capacity Charge']=''
##    data_dict['Quoted Non-Indexed Energy Charges']=''
##    data_dict['Quoted Indexed Energy Charges']=''
##    data_dict['Subtotal Energy Charge']='' #385
##
##    data_dict['Reduction in Capacity Charge (if any)']=''
##    data_dict['Final Applicable Capacity Charge']=''
##    data_dict['Final Applicable Energy Charge']=''
    
    data_extractor_numbers(text,'YTD Availability % ',1,data_dict,'\n','PAFN',l,'\d.*', 0)
    data_extractor_numbers(text,'Monthly Availability % ',1,data_dict,'\n','PAFM',l,'\d.*', 0)
    


    data_extractor_numbers(text,'Total Amount due for a ',1,data_dict,'\n','Total Cost',l,'\d.*', 0)
    data_extractor_numbers(text,'Provisional Bill Amount raised',1,data_dict,'\n','Provisional Bill',l,'\d.*', 0)
##    data_dict['Index Calculation (Capacity)']=''
##    data_dict['Index Calculation (Energy)']=''
##    data_dict['Energy as per declared capacity']=''#393
##    data_dict['No. of days']=''#394
##    data_dict['Domestice coal consumed']=''#395
##    data_dict['Imported coal consumed']=''#396
##    data_dict['ABR (Actual Blended Ratio)']=''#397
    ########################################################## Intial Column###############################################

##    March21 = datetime.datetime(2021,3,1)
##    March20 = datetime.datetime(2020,3,1)
##
##    if datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y') >= March21:
##    	data_dict['NatureofInvoice'] = 'Current Month Bill'
##    elif datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y') < March21 or datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y') > March20:
##    	data_dict['NatureofInvoice'] ='CY Arrear'
##    elif datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y') < March20:
##    	data_dict['NatureofInvoice'] =  "Previous Arrear Bill"


#############################################################################################
    if data_dict['NatureofInvoice'] == 'Current Month Bill' and not "Transmissions Charges" in text:
    	Power = "Power at Delhi"

    elif data_dict['NatureofInvoice'] == 'Current Month Bill' and "Transmissions Charges" in text:
        Power = "Gross PPC"
    else:
    	Power = 'Arrears'

    data_dict['PoweratDelhi'] = Power

######################--------------------FY-------------------------###########################
##    year = datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y').year
##    if datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y').year >= datetime.datetime(2021,1,1).year:
##    	FY = 'FY 21-22'
##    elif datetime.datetime.strptime(data_dict['Bill period to'],'%d.%m.%Y').year == datetime.datetime(2020,1,1).year:
##    	FY = 'FY 20-21'
##    else:
##    	FY = 'FY 19-20'
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
    if apr21<= datetime.strptime(data_dict['BillDate'],'%d.%m.%Y') - dateutil.relativedelta.relativedelta(months=1) <= mar22:
        data_dict['FY'] = 'FY 21-22'
    elif apr20 <= datetime.strptime(data_dict['BillDate'],'%d.%m.%Y') - dateutil.relativedelta.relativedelta(months=1) <=mar21:
        data_dict['FY'] = ' FY 20-21' 
    elif datetime.strptime(data_dict['BillDate'],'%d.%m.%Y') - dateutil.relativedelta.relativedelta(months=1) < apr20:
        data_dict['FY'] = 'FY 19-20'

#################-------------------------Board---------------------------------------####################################
    if 'Current Month Bill' in data_dict['NatureofInvoice'] and not "Transmissions Charges" in text:
    	data_dict['Board'] = 'ISGS'
    elif 'Current Month Bill' in data_dict['NatureofInvoice'] and "Transmissions Charges" in text:
        data_dict['Board'] = 'Other TX'
    else:
    	data_dict['Board'] = 'Others'


    if 'Current Month Bill' in data_dict['NatureofInvoice'] and not "Transmissions Charges" in text:
    	data_dict['ForEnergyBalance'] = 'CLP'
    elif 'Current Month Bill' in data_dict['NatureofInvoice'] and "Transmissions Charges" in text:
        data_dict['ForEnergyBalance'] = 'Other TX'

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


#######################-----------------Classification--------------------------######################
    if 'Power at Delhi' in data_dict['PoweratDelhi']:
    	data_dict['Classification'] = 'TPDDL Periphery'

    elif 'Gross PPC' in data_dict['PoweratDelhi']:
        data_dict['Classification'] = 'Transmission Total'
    elif 'Arrear' in data_dict['PoweratDelhi']:
    	data_dict['Classification'] = 'Others'
######################################################################################################

######################---------------PPAC-------------------------------------#######################
    if 'Power at Delhi' in data_dict['PoweratDelhi']: 
        data_dict['PPAC']='ISGS'
    elif 'Gross PPC' in data_dict['PoweratDelhi']:
        data_dict['PPAC']='Tx'
##########################################################################
    if 'Power at Delhi' in data_dict['PoweratDelhi']: 
        data_dict['MISHEAD']='CLP'
    elif 'Gross PPC' in data_dict['PoweratDelhi']:
        data_dict['MISHEAD']='Other TX'



    #############################################ADDITION#####################################
    try:
        TotalUnitsbilled = [str(data_dict['unit Coal'].replace(",", "")), str(data_dict['Units Hydro'].replace(",", "")), data_dict['Units Hydro Beyond DE'], data_dict['Units APM'], data_dict['Units RLNG'], data_dict['Units Liquid'], data_dict['SG NAPM Gas'],data_dict['SG Committed LNG'],data_dict['SG Committed Gas'],data_dict['Units Solar']]
        value1 = add_sub_value(TotalUnitsbilled)
        data_dict['TotalUnitsbilled'] = value1[0]+value1[1]
    except:
        TotalUnitsbilled = [data_dict['unit Coal'], data_dict['Units Hydro'], data_dict['Units Hydro Beyond DE'], data_dict['Units APM'], data_dict['Units RLNG'], data_dict['Units Liquid'], data_dict['SG NAPM Gas'],data_dict['SG Committed LNG'],data_dict['SG Committed Gas'],data_dict['Units Solar']]
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
        FC = [str(data_dict['Fixed Cost'].replace(",", "")), str(data_dict['Fixed Cost (Peak)'].replace(",", ""))]#, data_dict['Energy_Charges_Liquid_1'], data_dict['Energy_Charges_NAPM_Gas_1'], data_dict['Energy_Charges_Committed_Gas_1'], data_dict['Energy_Charges_Committed_LNG_1'], data_dict['EC_Adjustment_on_Actual_1'], data_dict['Variable_Cost'], data_dict['Variable_Cost_Hydro']]

        value1 = add_sub_value(FC)
        data_dict['FC'] = value1[0]+value1[1]
    except:
        FC = [data_dict['Fixed Cost'], data_dict['Fixed Cost (Peak)']]
        value1 = add_sub_value(FC)
        data_dict['FC'] = value1[0]+value1[1]

    try:
        Others = [int(str(data_dict['Other'])), str(data_dict['Trading Margin'].replace(",", ""))]#, data_dict['Energy_Charges_Liquid_1'], data_dict['Energy_Charges_NAPM_Gas_1'], data_dict['Energy_Charges_Committed_Gas_1'], data_dict['Energy_Charges_Committed_LNG_1'], data_dict['EC_Adjustment_on_Actual_1'], data_dict['Variable_Cost'], data_dict['Variable_Cost_Hydro']]

        value1 = add_sub_value(Others)
        data_dict['Others_'] = value1[0]+value1[1]
    except:
        try:
            Others = [str(data_dict['Other'].replace(",", "")), data_dict['Trading Margin']]
            value1 = add_sub_value(Others)
            data_dict['Others_'] = value1[0]+value1[1]
        except:
            Others = [data_dict['Other'], data_dict['Trading Margin']]
            value1 = add_sub_value(Others)
            data_dict['Others_'] = value1[0]+value1[1]

    
    try:
        Transmission_Charges = [data_dict['Transmissions'], data_dict['SLDC']]#, data_dict['Energy_Charges_Liquid_1'], data_dict['Energy_Charges_NAPM_Gas_1'], data_dict['Energy_Charges_Committed_Gas_1'], data_dict['Energy_Charges_Committed_LNG_1'], data_dict['EC_Adjustment_on_Actual_1'], data_dict['Variable_Cost'], data_dict['Variable_Cost_Hydro']]

        value1 = add_sub_value(Transmission_Charges)
        data_dict['TransmissionCharges'] = value1[0]+value1[1]
    except:
        try:
            Transmission_Charges = [str(data_dict['Transmissions'].replace(",", "")), str(data_dict['SLDC'].replace(",", ""))]
            value1 = add_sub_value(Transmission_Charges)
            data_dict['TransmissionCharges'] = value1[0]+value1[1]
        except:
            Transmission_Charges = [data_dict['Transmissions'], str(data_dict['SLDC'].replace(",", ""))]
            value1 = add_sub_value(Transmission_Charges)
            data_dict['TransmissionCharges'] = value1[0]+value1[1]


    lst = [data_dict['VC'], data_dict['FC'],  data_dict['Others_'], data_dict['TransmissionCharges']]
    check = sum(lst)
    print(check)

    try:
        data_dict['Total']=(data_dict['TotalInvoiceAmount'])
    except:
        data_dict['Total']=(data_dict['TotalInvoiceAmount']).replace(',',"")

##    print(data_dict['Total'],'____')
    checkum=(float(check)-float(Total))
    print(checkum)
    data_dict['check']=checkum

    #########################################################################################

    if re.search(r'Current.*?Bill',data_dict['NatureofInvoice']):
        pass
    else:
        data_dict['TotalUnitsbilled'] = ''
####################################################################
    if 'Current Month Bill' in data_dict['NatureofInvoice']:
    	data_dict['Remarks2'] = "Monthly energy bill for "+data_dict['PertainingBillMonth']
    else:
    	data_dict['Remarks2'] = 'Revision Bill'
##################################################################
    for key1,value1 in data_dict.items():
        if value1=='0' or value1==0 or value1==0.0 or value1== ' 0' or value1 == '0 ' or value1 == '':
            data_dict[key1]=''

    
    end_time= datetime.now().time()

##    query = "insert into  pgcil values('{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}')".format(mapping_sheet, start_time, end_time, '','','','','','','','','','','','','','','',data_dict['Name of Station'], data_dict['Bill No'],'', data_dict['Bill Date'], data_dict['Bill period from'], data_dict['Bill period to'],'','',data_dict['unit Coal'],'','','','','','','','','','','','','','', data_dict['Variable Cost'], '','','','','','','','','','','','','','','','','','','','','','','','','', data_dict['Fixed Cost'],'','','','','','','','','','','','','','','','','','',data_dict['other'],'','','','','',data_dict['Trading Margin'],'','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','',data_dict['Transmissions'],'',data_dict['SLDC'],'','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','',data_dict['Total Cost'],data_dict['Provisional Bill'],'','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','')


    query = "insert into Tata_Power values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
    values=(data_dict['PoweratDelhi'],'','','',data_dict['PPAC'],data_dict['FY'],data_dict['Quarter'],'','',data_dict['Classification'],data_dict['TreasuryReco'],data_dict['MISHEAD'],data_dict['Particulars_PartyName'],data_dict['Board'],data_dict['ForEnergyBalance'],data_dict['NameofStation'],data_dict['Bill No'],data_dict['NatureofInvoice'],data_dict['BillDate'],data_dict['Bill period from'],data_dict['Bill period to'],data_dict['PertainingBillMonth'],data_dict['Month'],data_dict['unit Coal'],'','','','','','','','','','','','','',data_dict['TotalUnitsbilled'],data_dict['Variable Cost'],'', '','','','','','','','','','','','','','','','','','','','','','','',data_dict['VC'], data_dict['Fixed Cost'],'','','','','',data_dict['FC'],'','','','','','','','','','','','','',data_dict['Other'],'','','','','',data_dict['Trading Margin'],'','','','', '','', '','','','','','','','','','','','','','','','','','','','','','','','',data_dict['Others_'],'','','','','',data_dict['Transmissions'],'',data_dict['SLDC'],'','','',data_dict['TransmissionCharges'],'',data_dict['Remarks2'],data_dict['check'],data_dict['TotalInvoiceAmount'],'','','','','','','','','','','','','','','','', '', '','','','','','', '','', '','','','','','','','','','','','','','','','','',data_dict['PAFM'],data_dict['PAFN'],'','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','', '','','','','','','','','','','','','','','','','','','','','','','','', '','','','','','', '','','','','','','','','','','','', '','','','','','','','','','','','','','', '','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','',data_dict['Total Cost'],data_dict['Provisional Bill'],'','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','',mapping_sheet,start_time,end_time)
    cursor.execute(query,values)   
    conn.commit()
    print("record inserted")

                                            
    print("\n")
    print(data_dict)


##print(trigger_wbsedcl(r"D:\TATA POWER PDF\CLP\CLP SOC April 2021.pdf"))

for file in os.listdir(path):
    #try:
    print(file)
    trigger_wbsedcl(file)
    #except:
     #   pass


##    data_extractor_numbers(text,'Gross Amount ',1,data_dict,'\n','EBAmountAfterDue',l,'\d{2,9}\.\d{2}', 0)

        


            
        



