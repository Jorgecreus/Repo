from hashlib import algorithms_available
from .Config import *
from .email_sender import send_email


def comercial(only_one_sales = None):

    import pandas as pd
    import pickle
    import os
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from pandas.tseries.offsets import MonthEnd
    from datetime import datetime as dt
    from os import system 
    import openpyxl
    system('cls')

    pd.options.mode.chained_assignment = None

    #Date variables
    #Findind the start and end date of the month to close and the two before

    inicitial_date = input('Strarting date of the month to close dd/mm/yyyy\n')
    inicitial_date = dt.strptime(inicitial_date, '%d/%m/%Y')
    periodo = str(inicitial_date.month+1)+'_'+str(inicitial_date.year)
    closing_date = pd.to_datetime(inicitial_date,dayfirst=True)+ MonthEnd(1)



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
    df['Devueltos'] = month_data['Devueltos'].replace('', '')
    df['Devueltos'] = month_data['Devueltos'].fillna(value='')
    df['Insurance'] = month_data['Insurance'].replace('', '')
    df['Insurance'] = month_data['Insurance'].fillna(value='')

    df_objetive = month_data[month_data.columns[month_data.columns.get_loc('Email'):]]
    df_objetive.dropna(how='all',inplace=True)
    df_objetive['Days working'] =  (pd.Timestamp.today() - df_objetive['Contribution Start Date']).dt.days
    df_objetive.drop(columns=['Dashboard Link'],inplace=True)
    df_objetive['Extra'] = df_objetive['Extra'].fillna(value=0)
    df_objetive['Extra'] = df_objetive['Extra'].astype('int')
    df_objetive['Motivo'] = df_objetive['Motivo'].fillna('')

   
    sales_agents = df_objetive['Email'].unique()
    
    for agent in sales_agents:
        try:
            #If a email is pass to the function the program don´t loop over all the sales agents
            if only_one_sales != None:
                agent = only_one_sales
            bonus = False
            index_objetivo =df_objetive[df_objetive['Email']==agent].index.values
            days_working = df_objetive.loc[index_objetivo,'Days working'].values
            if days_working <= config_reduce_objetive_days:
                objective = config_reduce_objetive
            else:
                objective = config_objective_number
            df_comercial= df[df['email_address'] == agent]
            df_comercial.drop(columns=['booking_date','email_address'],inplace=True)
            df_comercial.reset_index(inplace=True)
            df_comercial.drop(columns=['index'],inplace=True)

            sold_more_200 = sum(df_comercial['Inv Days']>=200)
            sold_more_100 = sum((df_comercial['Inv Days']<200) & (df_comercial['Inv Days']>=100))
            sold_less_100 = sum(df_comercial['Inv Days']<100)
            units = sold_more_200 + sold_more_100 + sold_less_100
            units_euros = sold_more_200 * config_more_200 + sold_more_100 * config_more_100 + sold_less_100 * config_less_100
            objective_per = units / objective
            if objective_per >= config_objective:
                bonus = True
            if bonus == True:
                bonus_sales = int(objective_per * config_objective_sold)
            else:
                bonus_sales = 0

            warranty_I = sum(df_comercial['warranty_title']=='Garantia I' )
            warranty_II = sum(df_comercial['warranty_title']=='Garantia II' )
            warranty_per = (warranty_I+warranty_II) / objective
            if bonus == True and warranty_per >= config_bonus_warranty_objt:
                warranty_bonus = config_bonus_warranty
            else:
                warranty_bonus = 0

            finance = sum(df_comercial['payment_type']=='CASH_AND_FINANCE' ) 
            insurance = sum((df_comercial['Insurance']=='Seguro de Vida' )  + sum(df_comercial['Insurance']=='Protección Total' ))
            finance_per = finance / objective
            if bonus == True and  finance_per >= config_bonus_finance_objt:
                finance_bonus = config_bonus_finance
            else:
                finance_bonus = 0 

            promoters= sum(df_comercial['nps_value']>config_promoters )
            detractors= sum((df_comercial['nps_value']<config_detractors ) & (df_comercial['nps_value']>=0 ))
            total= sum(df_comercial['nps_value']>=0 ) 
            try:
                nps_score= int((promoters - detractors)/total*100)
            except:
                nps_score = 0

            if bonus == True and nps_score>=65:
                nps_bonus = config_bonus_nps_65

            elif bonus == True and nps_score>=30:
                nps_bonus = config_bonus_nps_30
            else:
                nps_bonus = config_bonus_nps_less_30
            tl = df_objetive.loc[index_objetivo,'Team Leader'].values
            extra = df_objetive.loc[index_objetivo,'Extra'].values
            motive = df_objetive.loc[index_objetivo,'Motivo'].values

            total_value = (units_euros + bonus_sales + warranty_I * config_warranty_I + warranty_II * config_warranty_II 
                        + warranty_bonus + finance * config_finance + insurance * config_insurance  + finance_bonus 
                        + nps_bonus + extra) 

            excel_name = agent+'_'+periodo+'.xlsx'
            wb_template = openpyxl.load_workbook('Templates\Template.xlsx')
            ws_template_resumen = wb_template.worksheets[0]
            ws_template_resumen['F6'] = periodo
            ws_template_resumen['F5'] = inicitial_date
            ws_template_resumen['G5'] = closing_date
            ws_template_resumen['F3'] = agent
            ws_template_resumen['F7'] = objective
            ws_template_resumen['F11'] = sold_more_200
            ws_template_resumen['G11'] = sold_more_200 * config_more_200
            ws_template_resumen['F12'] = sold_more_100
            ws_template_resumen['G12'] = sold_more_100  * config_more_100
            ws_template_resumen['F13'] = sold_less_100 
            ws_template_resumen['G13'] = sold_less_100 * config_less_100
            ws_template_resumen['F14'] = objective_per
            ws_template_resumen['G14'] = bonus_sales
            ws_template_resumen['F15'] = (bonus_sales + sold_more_200 * config_more_200 + sold_more_100  * config_more_100 + sold_less_100 * config_less_100)
            ws_template_resumen['F19'] = finance
            ws_template_resumen['F20'] = insurance
            ws_template_resumen['G19'] = finance * config_finance
            ws_template_resumen['G20'] = insurance * config_insurance
            ws_template_resumen['F21'] = finance_per
            ws_template_resumen['G21'] = finance_bonus
            ws_template_resumen['F22'] = finance * config_finance + insurance * config_insurance + finance_bonus
            ws_template_resumen['F26'] = warranty_I
            ws_template_resumen['G26'] = warranty_I * config_warranty_I
            ws_template_resumen['F27'] = warranty_II
            ws_template_resumen['G27'] = warranty_II * config_warranty_II
            ws_template_resumen['F28'] = warranty_per
            ws_template_resumen['G28'] = warranty_bonus
            ws_template_resumen['F29'] =  warranty_I * config_warranty_I +  warranty_II * config_warranty_II + warranty_bonus
            ws_template_resumen['F32'] = total
            ws_template_resumen['F33'] = nps_score
            ws_template_resumen['F34'] = nps_bonus
            ws_template_resumen['F37'] = int(extra)
            ws_template_resumen['F38'] = str(motive[0])
            ws_template_resumen['F41'] = int(total_value)
            contador = 51
            for i in range(len(df_comercial)):
                total = 0
                ws_template_resumen['E'+str(contador)] = df_comercial.loc[i,'stock_number']
                ws_template_resumen['F'+str(contador)] = df_comercial.loc[i,'order_number']
                if df_comercial.loc[i,'Inv Days'] >=200:
                    ws_template_resumen['I'+str(contador)] = config_more_200
                    total += config_more_200
                elif (df_comercial.loc[i,'Inv Days'] <200 and df_comercial.loc[i,'Inv Days'] >=100):
                    ws_template_resumen['K'+str(contador)] = config_more_100
                    total += config_more_100
                else:
                    ws_template_resumen['L'+str(contador)] = config_less_100
                    total += config_less_100
                ws_template_resumen['M'+str(contador)] = df_comercial.loc[i,'payment_type']
                if df_comercial.loc[i,'payment_type'] == 'CASH_AND_FINANCE':
                    total += config_finance
                ws_template_resumen['N'+str(contador)] = df_comercial.loc[i,'Insurance']
                if (df_comercial.loc[i,'Insurance'] == 'Seguro de Vida' or df_comercial.loc[i,'Insurance'] == 'Protección Total'):
                    total += config_insurance
                ws_template_resumen['P'+str(contador)] = df_comercial.loc[i,'warranty_title']
                if df_comercial.loc[i,'warranty_title'] == 'Garantia I':
                    total += config_warranty_I
                if df_comercial.loc[i,'warranty_title'] == 'Garantia II':
                    total += config_warranty_II
                ws_template_resumen['Q'+str(contador)] = total
                contador += 1
            wb_template.save(excel_name)
            send_email(config_gmail_user,config_gmail_pass,config_email_copia,config_email_tl[tl[0]],excel_name,'Cierre comisiones '+ agent + ' ' + periodo,config_msg_comercial)
            os.remove(excel_name)
            print('Comisiones de '+ agent + ' enviadas correctamente')
            if only_one_sales != None:
                break
        except:
            print('Error al enviar las comisiones de '+agent)






























