#import local packages
from .Config import *
from .email_sender import send_email


def head():
    from os import system 
    #import global packages
    import pandas as pd
    import pickle
    import os
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from pandas.tseries.offsets import MonthEnd
    from datetime import datetime as dt
    import openpyxl
    from openpyxl.styles.borders import Border, Side
    from openpyxl.styles import Alignment

    pd.options.mode.chained_assignment = None
    system('cls')
    #Date variables
    #Findind the start and end date of the month to close and the two before

    inicitial_date = input('Strarting date of the month to close dd/mm/yyyy\n')
    inicitial_date = dt.strptime(inicitial_date, '%d/%m/%Y')
    periodo = str(inicitial_date.month+1)+'_'+str(inicitial_date.year)
    m_1 = inicitial_date.replace(month=inicitial_date.month-1)
    m_2 = inicitial_date.replace(month=inicitial_date.month-2)
    m_1 = m_1.strftime('%d/%m/%Y')
    m_2 = m_2.strftime('%d/%m/%Y')
    inicitial_date = pd.to_datetime(inicitial_date,dayfirst=True)
    m_1 = pd.to_datetime(m_1,dayfirst=True)
    m_2 = pd.to_datetime(m_2,dayfirst=True)
    closing_date = pd.to_datetime(inicitial_date,dayfirst=True)+ MonthEnd(1)
    closing_date_m_1 = pd.to_datetime(m_1,dayfirst=True)+ MonthEnd(1)
    closing_date_m_2 = pd.to_datetime(m_2,dayfirst=True)+ MonthEnd(1)
    

    #Concecting to the G-Sheet api and getting the values for the two DF, creating the pickle token
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)# If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials_BA.json', scopes)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    # Getting the data
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    result_data_month = sheet.values().get(spreadsheetId=sheet_id,
                                    range=periodo+range_data_month).execute()
    values_data_month = result_data_month.get('values', [])

    result_data_all = sheet.values().get(spreadsheetId=sheet_id,
                                    range=range_all_data).execute()
    values_data = result_data_all.get('values', [])


    #?Create the diferent DF

    #Create the DF with month information, the name of the G-sheet must be (m_yyyy), the month needs to be the next one
    # that you wanna close. Example :  July will be 8_2021

    month_data = pd.DataFrame.from_records(values_data_month)
    month_data.rename(columns=month_data.iloc[0],inplace=True)
    month_data.drop(month_data.index[0],inplace=True)
    month_data['booking_date'] = pd.to_datetime(month_data['booking_date'],dayfirst=True)
    month_data['contract_signed_on'] = pd.to_datetime(month_data['contract_signed_on'],dayfirst=True)
    month_data['car_handover_on'] = pd.to_datetime(month_data['car_handover_on'],dayfirst=True)
    month_data['Contribution Start Date'] = pd.to_datetime(month_data['Contribution Start Date'],dayfirst=True)
    month_data['nps_value'] = month_data['nps_value'].replace('', '-1')
    month_data['nps_value'] = month_data['nps_value'].fillna(value=-1)
    month_data['nps_value'] = month_data['nps_value'].astype('int')

    #Divide those DF into two and reshape it

    df = month_data[month_data.columns[0:month_data.columns.get_loc('Insurance')+1]]
    df.dropna(how='all',inplace=True)
    df['Inv Days'] = (df['contract_signed_on'] - df['booking_date']).dt.days
    df.replace(['ES - Autohero Oro - 24','ES - Autohero Diamante - 36','ES - Autohero Diamante - 24','ES - Autohero Diamante - 12',
                'ES - Autohero Oro - 12','ES - Autohero Oro - 36'],'Garantia I',inplace=True)
    df.replace(['ES - Autohero Plata - 36','ES - Autohero Plata - 24'],'Garantia II',inplace=True)
    df.replace(['ES - Autohero Plata - 12','ES - Autohero Plata - 24'],'Garantia AH (Básica)',inplace=True)
    df_objetive = month_data[month_data.columns[month_data.columns.get_loc('Email'):]]
    df_objetive.dropna(how='all',inplace=True)
    df_objetive['Days working'] =  (pd.Timestamp.today() - df_objetive['Contribution Start Date']).dt.days
    df_objetive.drop(columns=['Dashboard Link'],inplace=True)
    df_objetive['Extra'] = df_objetive['Extra'].fillna(value=0)
    df_objetive['Extra'] = df_objetive['Extra'].astype('int')
    df_objetive['Motivo'] = df_objetive['Motivo'].fillna('')
    

    #Create the DF with all data information
    
    all_data = pd.DataFrame.from_records(values_data)
    all_data.rename(columns=all_data.iloc[0],inplace=True)
    all_data.drop(all_data.index[0],inplace=True)
    all_data['booking_date'] = pd.to_datetime(all_data['booking_date'],dayfirst=True)
    all_data['contract_signed_on'] = pd.to_datetime(all_data['contract_signed_on'],dayfirst=True)
    all_data['car_handover_on'] = pd.to_datetime(all_data['car_handover_on'],dayfirst=True)
    all_data.replace(['ES - Autohero Oro - 24','ES - Autohero Diamante - 36','ES - Autohero Diamante - 24','ES - Autohero Diamante - 12',
                'ES - Autohero Oro - 12','ES - Autohero Oro - 36'],'Garantia I',inplace=True)
    all_data.replace(['ES - Autohero Plata - 36','ES - Autohero Plata - 24'],'Garantia II',inplace=True)
    all_data.replace(['ES - Autohero Plata - 12','ES - Autohero Plata - 24'],'Garantia AH (Básica)',inplace=True)

    all_data['nps_value'] = all_data['nps_value'].replace('', '-1')
    all_data['nps_value'] = all_data['nps_value'].fillna(value=-1)
    all_data['nps_value'] = all_data['nps_value'].astype('int')
    all_data['Inv Days'] = (all_data['contract_signed_on'] - all_data['booking_date']).dt.days


    #Create the final DF ant the inform
    df_head = pd.DataFrame(columns=['Email'])
    df_head['Email'] = df_objetive['Email'].unique()

    df_head_comp = pd.DataFrame(columns=['Email'])
    df_head_comp['Email'] = df_objetive['Email'].unique()

    #Calculate all the fields in the df to close the month

    for i in range(len(df_head)):
        #variables for the foor loop
        bonus = False
        comercial = df_head.loc[i,'Email']
        index_objetivo =df_objetive[df_objetive['Email']==comercial].index.values
        days_working = df_objetive.loc[index_objetivo,'Days working'].values
        if days_working <= config_reduce_objetive_days:
            df_head.loc[i,'Objetivo'] = config_reduce_objetive
        else:
            df_head.loc[i,'Objetivo'] = config_objective_number
        df_head.loc[i,'TL'] = df_objetive.loc[index_objetivo,'Team Leader'].values
        #sales
        sold_more_200 = sum((df['Inv Days']>=200) & (df['email_address']==comercial) )
        sold_more_100 = sum((df['Inv Days']>=100) & (df['Inv Days']<200) & (df['email_address']==comercial) )
        sold_less_100 = sum((df['Inv Days']<100) & (df['email_address']==comercial))
        ventas = sold_more_200+ sold_more_100 +sold_less_100
        df_head.loc[i,'Ventas (#)'] = ventas
        df_head.loc[i,'Ventas €'] = sold_more_200 * config_more_200 + sold_more_100 * config_more_100 + sold_less_100 * config_less_100
        df_head.loc[i,'Objetivo alcanzado'] = ventas / df_head.loc[i,'Objetivo']
        if df_head.loc[i,'Objetivo alcanzado'] >= config_objective:
            bonus = True
        if bonus == True:
            df_head.loc[i,'Bonus ventas €'] = int(df_head.loc[i,'Objetivo alcanzado'] * config_objective_sold)
        else:
            df_head.loc[i,'Bonus ventas €'] = 0
        
        #Warranty
        garantia_I = sum((df['warranty_title']=='Garantia I' ) & (df['email_address']==comercial))
        garantia_II = sum((df['warranty_title']=='Garantia II' ) & (df['email_address']==comercial))
        df_head.loc[i,'Garantia I (#)'] = garantia_I 
        df_head.loc[i,'Garantia II (#)'] = garantia_II
        df_head.loc[i,'Garantia (%)'] = (df_head.loc[i,'Garantia I (#)']+df_head.loc[i,'Garantia II (#)']) / df_head.loc[i,'Objetivo']
        if bonus == True and df_head.loc[i,'Garantia (%)'] >=config_bonus_warranty_objt:
            df_head.loc[i,'Bonus Garantia'] = config_bonus_warranty
        else:
            df_head.loc[i,'Bonus Garantia'] = 0

        #Finance
        df_head.loc[i,'Financiados (#)'] = sum((df['payment_type']=='CASH_AND_FINANCE' ) & (df['email_address']==comercial))
        df_head.loc[i,'Seguro (#)'] = sum((df['Insurance']=='Seguro de Vida' ) & (df['email_address']==comercial)) +sum((df['Insurance']=='Protección Total' ) & (df['email_address']==comercial))
        df_head.loc[i,'Financiados (%)'] = df_head.loc[i,'Financiados (#)']  / df_head.loc[i,'Objetivo']
        if bonus == True and df_head.loc[i,'Financiados (%)'] >= config_bonus_finance_objt:
                df_head.loc[i,'Bonus Financiación'] = config_bonus_finance
        else:
            df_head.loc[i,'Bonus Financiación'] =0

        #NPS
        promoters= sum((df['nps_value']>config_promoters ) & (df['email_address']==comercial))
        detractors= sum((df['nps_value']<config_detractors ) & (df['email_address']==comercial) & (df['nps_value']>=0 ))
        total= sum((df['nps_value']>=0 ) & (df['email_address']==comercial))
        try:
            df_head.loc[i,'Nota NPS']= int((promoters - detractors)/total*100)
        except:
            df_head.loc[i,'Nota NPS']=0

        if bonus == True and df_head.loc[i,'Nota NPS']>=65:
                df_head.loc[i,'Bonus NPS'] = config_bonus_nps_65
        elif bonus == True and df_head.loc[i,'Nota NPS']>=30:
            df_head.loc[i,'Bonus NPS'] = config_bonus_nps_30
        else:
            df_head.loc[i,'Bonus NPS'] = config_bonus_nps_less_30


        #Total
        df_head.loc[i,'Extra'] = df_objetive.loc[index_objetivo,'Extra'].values
        df_head.loc[i,'Motivo'] = df_objetive.loc[index_objetivo,'Motivo'].values
        df_head.loc[i,'Total Variable'] = [df_head.loc[i,'Ventas €'] + df_head.loc[i,'Bonus ventas €'] + df_head.loc[i,'Bonus NPS'] + df_head.loc[i,'Garantia I (#)'] * config_warranty_I 
        + df_head.loc[i,'Garantia II (#)'] * config_warranty_II + df_head.loc[i,'Financiados (#)']  * config_finance + df_head.loc[i,'Seguro (#)']  * config_insurance 
        + df_head.loc[i,'Bonus Garantia'] + df_head.loc[i,'Bonus Financiación']+df_head.loc[i,'Extra'] ]
        df_head.loc[i,'Devueltos']= sum((df['Devueltos']=='Devuelto' ) & (df['email_address']==comercial))
        df_head.loc[i,'Euro por coche']= df_head.loc[i,'Total Variable'] / df_head.loc[i,'Ventas (#)']
 


        #Calculate the values of DF


        #Sales and objetives
        df_head_comp.loc[i,'TL'] = df_objetive.loc[index_objetivo,'Team Leader'].values
        df_head_comp.loc[i,'Objetivo'] = df_head.loc[i,'Objetivo']
        df_head_comp.loc[i,'Ventas (#)'] = df_head.loc[i,'Ventas (#)']
        df_head_comp.loc[i,'Objetivo alcanzado'] = df_head.loc[i,'Objetivo alcanzado']
        if ((closing_date_m_1 - df_objetive.loc[index_objetivo,'Contribution Start Date']).dt.days).values <=45 :
            df_head_comp.loc[i,'Objetivo M-1'] = 16
        else:
            df_head_comp.loc[i,'Objetivo M-1'] = 20
        df_head_comp.loc[i,'Ventas (#) M-1'] = sum((all_data['Inv Days']>=0) & (all_data['email_address']==comercial) & (all_data['car_handover_on']>=m_1) & (all_data['car_handover_on']<= closing_date_m_1))
        df_head_comp.loc[i,'Objetivo alcanzado M-1'] = df_head_comp.loc[i,'Ventas (#) M-1'] /df_head_comp.loc[i,'Objetivo M-1'] 
        if ((closing_date_m_2 - df_objetive.loc[index_objetivo,'Contribution Start Date']).dt.days).values <=45 :
            df_head_comp.loc[i,'Objetivo M-2'] = 16
        else:
            df_head_comp.loc[i,'Objetivo M-2'] = 20
        df_head_comp.loc[i,'Ventas (#) M-2'] = sum((all_data['Inv Days']>=0) & (all_data['email_address']==comercial) & (all_data['car_handover_on']>=m_2) & (all_data['car_handover_on']<= closing_date_m_2))
        df_head_comp.loc[i,'Objetivo alcanzado M-2'] =df_head_comp.loc[i,'Ventas (#) M-2'] / df_head_comp.loc[i,'Objetivo M-2']

        #Warrantys

        df_head_comp.loc[i,'Garantia'] = df_head.loc[i,'Garantia I (#)'] + df_head.loc[i,'Garantia II (#)']
        df_head_comp.loc[i,'Garantia (%)'] = df_head.loc[i,'Garantia (%)']

        df_head_comp.loc[i,'Garantia  M-1'] = [sum((all_data['warranty_title']=='Garantia I') & (all_data['email_address']==comercial) & (all_data['car_handover_on']>=m_1) & (all_data['car_handover_on']<= closing_date_m_1)) +
                                                sum((all_data['warranty_title']=='Garantia II') & (all_data['email_address']==comercial) & (all_data['car_handover_on']>=m_1) & (all_data['car_handover_on']<= closing_date_m_1))]
        df_head_comp.loc[i,'Garantia (%) M-1'] = df_head_comp.loc[i,'Garantia  M-1'] / df_head_comp.loc[i,'Objetivo M-1']
        df_head_comp.loc[i,'Garantia  M-2'] = [sum((all_data['warranty_title']=='Garantia I') & (all_data['email_address']==comercial) & (all_data['car_handover_on']>=m_2) & (all_data['car_handover_on']<= closing_date_m_2)) +
                                                sum((all_data['warranty_title']=='Garantia II') & (all_data['email_address']==comercial) & (all_data['car_handover_on']>=m_2) & (all_data['car_handover_on']<= closing_date_m_2))]
        df_head_comp.loc[i,'Garantia (%) M-2'] = df_head_comp.loc[i,'Garantia  M-2'] / df_head_comp.loc[i,'Objetivo M-2']

        #Finance

        df_head_comp.loc[i,'Financiados (#)'] = df_head.loc[i,'Financiados (#)']
        df_head_comp.loc[i,'Financiados (%)'] = df_head.loc[i,'Financiados (%)']
        df_head_comp.loc[i,'Financiados (#) M-1'] = sum((all_data['payment_type']=='CASH_AND_FINANCE') & (all_data['email_address']==comercial) & (all_data['car_handover_on']>=m_1) & (all_data['car_handover_on']<= closing_date_m_1))
        df_head_comp.loc[i,'Financiados (%) M-1'] = df_head_comp.loc[i,'Financiados (#) M-1'] / df_head_comp.loc[i,'Objetivo M-1']
        df_head_comp.loc[i,'Financiados (#) M-2'] = sum((all_data['payment_type']=='CASH_AND_FINANCE') & (all_data['email_address']==comercial) & (all_data['car_handover_on']>=m_2) & (all_data['car_handover_on']<= closing_date_m_2))
        df_head_comp.loc[i,'Financiados (%) M-2'] = df_head_comp.loc[i,'Financiados (#) M-2'] / df_head_comp.loc[i,'Objetivo M-2']

        #NPS
        df_head_comp.loc[i,'NPS'] = df_head.loc[i,'Nota NPS']
        promoters_m_1= sum((all_data['nps_value']>config_promoters ) & (all_data['email_address']==comercial) & (all_data['car_handover_on']>=m_1) & (all_data['car_handover_on']<= closing_date_m_1))
        detractors_m_1= sum((all_data['nps_value']>=0) &(all_data['nps_value']<config_detractors ) & (all_data['email_address']==comercial) & (all_data['car_handover_on']>=m_1) & (all_data['car_handover_on']<= closing_date_m_1))
        total_m_1= sum((all_data['nps_value']>=0 ) & (all_data['email_address']==comercial) & (all_data['car_handover_on']>=m_1) & (all_data['car_handover_on']<= closing_date_m_1))
        try:
            df_head_comp.loc[i,'NPS M-1'] = int((promoters_m_1 - detractors_m_1)/total_m_1*100)
        except:
            df_head_comp.loc[i,'NPS M-1'] = 0
        promoters_m_2= sum((all_data['nps_value']>config_promoters ) & (all_data['email_address']==comercial) & (all_data['car_handover_on']>=m_2) & (all_data['car_handover_on']<= closing_date_m_2))
        detractors_m_2= sum((all_data['nps_value']>=0) &(all_data['nps_value']<config_detractors ) & (all_data['email_address']==comercial) & (all_data['car_handover_on']>=m_2) & (all_data['car_handover_on']<= closing_date_m_2))
        total_m_2= sum((all_data['nps_value']>=0 ) & (all_data['email_address']==comercial) & (all_data['car_handover_on']>=m_2) & (all_data['car_handover_on']<= closing_date_m_2))
        try:
            df_head_comp.loc[i,'NPS M-2'] = int((promoters_m_2 - detractors_m_2)/total_m_2*100)
        except:
            df_head_comp.loc[i,'NPS M-2'] = 0


    #Sort the DF in order

    df_head.sort_values(by=['Objetivo alcanzado'], ascending=False,inplace=True)
    df_head_comp.sort_values(by=['Objetivo alcanzado'], ascending=False,inplace=True)

    #Put all the information in the report (xlsx file)
    excel_name = 'Cierre_mensual_'+periodo+'.xlsx'
    range_excel_month = 'BCDEFGHIJKLMNOPQRSTUVW' #range to fill in the template to loop
    columns_with_euros = [4,6,10,14,16,19,21,23] #columns with €
    wb_template = openpyxl.load_workbook('Templates\Template_head.xlsx')
    thin_border = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thin'), 
                     bottom=Side(style='thin'))
    ws_template_month = wb_template.worksheets[0]
    ws_template_overview = wb_template.worksheets[1]

    contador_month = 3
    for i in range(len(df_head)):
        contador_df = 0
        for x in range_excel_month:
            ws_template_month[str(x)+str(contador_month)] = df_head.iloc[i,contador_df]
            if contador_df == 5 or contador_df == 9 or contador_df == 13:
                ws_template_month[str(x)+str(contador_month)].number_format = '0%'
            if contador_df in columns_with_euros :
                ws_template_month[str(x)+str(contador_month)].number_format = '0€'
            ws_template_month[str(x)+str(contador_month)].border = thin_border
            ws_template_month[str(x)+str(contador_month)].alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
            contador_df+=1
        contador_month+=1

    range_excel_overview = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'#range to fill in the template to loop
    columns_with_percentage = [4,7,10,12,14,16,18,20,22] #Columns with the format %
    contador_overview = 3
    for i in range(len(df_head_comp)):
        contador_df_overview = 0
        for x in range_excel_overview:
            ws_template_overview[str(x)+str(contador_overview)] = df_head_comp.iloc[i,contador_df_overview]
            if contador_df_overview in columns_with_percentage:
                ws_template_overview[str(x)+str(contador_overview)].number_format = '0%'
            ws_template_overview[str(x)+str(contador_overview)].border = thin_border
            ws_template_overview[str(x)+str(contador_overview)].alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
            contador_df_overview+=1
        contador_overview+=1

    wb_template.save(excel_name) 

    send_email(config_gmail_user,config_gmail_pass,config_head,[config_email_copia,config_vp],excel_name,'Cierre comisiones equipo comercial '+periodo,config_msg_head)
    os.remove(excel_name)





# if __name__ == '__main__':
#     head()
    
    