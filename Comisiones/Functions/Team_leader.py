from .Config import *
from .email_sender import send_email

def tl(only_one_tl=None):
     #import global packages
     import pandas as pd
     import pickle
     import os
     from google_auth_oauthlib.flow import InstalledAppFlow
     from google.auth.transport.requests import Request
     from googleapiclient.discovery import build
     from datetime import datetime as dt
     import openpyxl
     from openpyxl.styles.borders import Border, Side
     from openpyxl.styles import Alignment,Font
     from os import system 

     pd.options.mode.chained_assignment = None
     system('cls')
     #Date variables
     #Findind the start and end date of the month to close and the two before

     inicitial_date = input('Strarting date of the month to close dd/mm/yyyy\n')
     # inicitial_date = '01/08/2021'
     inicitial_date = dt.strptime(inicitial_date, '%d/%m/%Y')
     periodo = str(inicitial_date.month+1)+'_'+str(inicitial_date.year)
     # inicitial_date = pd.to_datetime(inicitial_date,dayfirst=True)
     # closing_date = pd.to_datetime(inicitial_date,dayfirst=True)+ MonthEnd(1)



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


   



     team_leader = list(df_objetive['Team Leader'].unique())
     for tl in team_leader:
          if only_one_tl != None:
                tl = only_one_tl
          df_tl = pd.DataFrame(columns=['Email'])
          df_tl['Email'] = df_objetive['Email'].where(df_objetive['Team Leader'] == tl )
          df_tl.dropna(inplace=True)
          df_tl.reset_index(inplace=True)
          df_tl.drop(columns=['index'],inplace=True)
          for i in range(len(df_tl)):
               bonus = False
               comercial = df_tl.loc[i,'Email']
               index_objetivo =df_objetive[df_objetive['Email']==comercial].index.values
               days_working = df_objetive.loc[index_objetivo,'Days working'].values
               if days_working <= config_reduce_objetive_days:
                    df_tl.loc[i,'Objetivo'] = config_reduce_objetive
               else:
                    df_tl.loc[i,'Objetivo'] = config_objective_number
               sold_more_200 = sum((df['Inv Days']>=200) & (df['email_address']==comercial) )
               sold_more_100 = sum((df['Inv Days']>=100) & (df['Inv Days']<200) & (df['email_address']==comercial) )
               sold_less_100 = sum((df['Inv Days']<100) & (df['email_address']==comercial))
               ventas = sold_more_200+ sold_more_100 +sold_less_100
               df_tl.loc[i,'Ventas (#)'] = ventas
               df_tl.loc[i,'Ventas €'] = sold_more_200 * config_more_200 + sold_more_100 * config_more_100 + sold_less_100 * config_less_100
               df_tl.loc[i,'Objetivo alcanzado'] = ventas / df_tl.loc[i,'Objetivo']
               if df_tl.loc[i,'Objetivo alcanzado'] >= config_objective:
                    bonus = True
               if bonus == True:
                    df_tl.loc[i,'Bonus ventas €'] = int(df_tl.loc[i,'Objetivo alcanzado'] * config_objective_sold)
               else:
                    df_tl.loc[i,'Bonus ventas €'] = 0

               #Warranty
               garantia_I = sum((df['warranty_title']=='Garantia I' ) & (df['email_address']==comercial))
               garantia_II = sum((df['warranty_title']=='Garantia II' ) & (df['email_address']==comercial))
               df_tl.loc[i,'Garantia I (#)'] = garantia_I 
               df_tl.loc[i,'Garantia II (#)'] = garantia_II
               df_tl.loc[i,'Garantia (%)'] = (df_tl.loc[i,'Garantia I (#)']+df_tl.loc[i,'Garantia II (#)']) / df_tl.loc[i,'Objetivo']
               if bonus == True and df_tl.loc[i,'Garantia (%)'] >=config_bonus_warranty_objt:
                    df_tl.loc[i,'Bonus Garantia'] = config_bonus_warranty
               else:
                    df_tl.loc[i,'Bonus Garantia'] = 0

               #Finance
               df_tl.loc[i,'Financiados (#)'] = sum((df['payment_type']=='CASH_AND_FINANCE' ) & (df['email_address']==comercial))
               df_tl.loc[i,'Seguro (#)'] = sum((df['Insurance']=='Seguro de Vida' ) & (df['email_address']==comercial)) +sum((df['Insurance']=='Protección Total' ) & (df['email_address']==comercial))
               df_tl.loc[i,'Financiados (%)'] = df_tl.loc[i,'Financiados (#)']  / df_tl.loc[i,'Objetivo']
               if bonus == True and df_tl.loc[i,'Financiados (%)'] >= config_bonus_finance_objt:
                         df_tl.loc[i,'Bonus Financiación'] = config_bonus_finance
               else:
                    df_tl.loc[i,'Bonus Financiación'] =0

               #NPS
               promoters= sum((df['nps_value']>config_promoters ) & (df['email_address']==comercial))
               detractors= sum((df['nps_value']<config_detractors ) & (df['email_address']==comercial) & (df['nps_value']>=0 ))
               total= sum((df['nps_value']>=0 ) & (df['email_address']==comercial))
               try:
                    df_tl.loc[i,'Nota NPS']= int((promoters - detractors)/total*100)
               except:
                    df_tl.loc[i,'Nota NPS']=0

               if bonus == True and df_tl.loc[i,'Nota NPS']>=65:
                         df_tl.loc[i,'Bonus NPS'] = config_bonus_nps_65
               elif bonus == True and df_tl.loc[i,'Nota NPS']>=30:
                    df_tl.loc[i,'Bonus NPS'] = config_bonus_nps_30
               else:
                    df_tl.loc[i,'Bonus NPS'] = config_bonus_nps_less_30


               #Total
               df_tl.loc[i,'Extra'] = df_objetive.loc[index_objetivo,'Extra'].values
               df_tl.loc[i,'Motivo'] = df_objetive.loc[index_objetivo,'Motivo'].values
               df_tl.loc[i,'Total Variable'] = [df_tl.loc[i,'Ventas €'] + df_tl.loc[i,'Bonus ventas €'] + df_tl.loc[i,'Bonus NPS'] + df_tl.loc[i,'Garantia I (#)'] * config_warranty_I 
               + df_tl.loc[i,'Garantia II (#)'] * config_warranty_II + df_tl.loc[i,'Financiados (#)']  * config_finance + df_tl.loc[i,'Seguro (#)']  * config_insurance 
               + df_tl.loc[i,'Bonus Garantia'] + df_tl.loc[i,'Bonus Financiación']+df_tl.loc[i,'Extra'] ]
               df_tl.loc[i,'Devueltos']= sum((df['Devueltos']=='Devuelto' ) & (df['email_address']==comercial))
               df_tl.sort_values(by=['Objetivo alcanzado'], ascending=False,inplace=True)

               #Put all the information in the report (xlsx file)
     for tl in team_leader:

          excel_name = 'Cierre_mensual_'+periodo+tl+'.xlsx'

          range_excel_month = 'BCDEFGHIJKLMNOPQRSTU' #range to fill in the template to loop
          wb_template = openpyxl.load_workbook('Templates\Template_TL.xlsx')
          thin_border = Border(left=Side(style='thin'), 
                    right=Side(style='thin'), 
                    top=Side(style='thin'), 
                    bottom=Side(style='thin'))
          ws_template_month = wb_template.worksheets[0]

          #Fill the month information to the excel
          contador_month = 3
          for i in range(len(df_tl)):
               contador_df = 0
               for x in range_excel_month:
                    ws_template_month[str(x)+str(contador_month)] = df_tl.iloc[i,contador_df]
                    if contador_df == 4 or contador_df == 8 or contador_df == 13:
                         ws_template_month[str(x)+str(contador_month)].number_format = '0%'
                    ws_template_month[str(x)+str(contador_month)].border = thin_border
                    ws_template_month[str(x)+str(contador_month)].alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
                    contador_df+=1
               contador_month+=1
          #create a sheet for every sales agent and fill it with their sales
          contador_page = 1
          for i in range(len(df_tl)):
               wb_template.create_sheet(df_tl.loc[i,'Email'])
               # wb_template.save(excel_name) 
               df_comercial= df[df['email_address'] == df_tl.loc[i,'Email']]
               df_comercial.drop(columns=['booking_date','contract_signed_on','car_handover_on','email_address'],inplace=True)
               df_comercial['nps_value'] = df_comercial['nps_value'].replace(-1, '')
               ws_template_month_comercial = wb_template.worksheets[contador_page]
               ws_template_month_comercial.column_dimensions['B'].width = 10
               ws_template_month_comercial['B2'] = 'ID'
               ws_template_month_comercial['B2'].border = thin_border
               ws_template_month_comercial['B2'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
               ws_template_month_comercial['B2'].font = Font(bold=True)
               ws_template_month_comercial.column_dimensions['C'].width = 15
               ws_template_month_comercial['C2'] = 'Order Number'
               ws_template_month_comercial['C2'].border = thin_border
               ws_template_month_comercial['C2'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
               ws_template_month_comercial['C2'].font = Font(bold=True)
               ws_template_month_comercial.column_dimensions['D'].width = 21
               ws_template_month_comercial['D2'] = 'Payment Type'
               ws_template_month_comercial['D2'].border = thin_border
               ws_template_month_comercial['D2'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
               ws_template_month_comercial['D2'].font = Font(bold=True)
               ws_template_month_comercial.column_dimensions['E'].width = 30
               ws_template_month_comercial['E2'] = 'Garantía'
               ws_template_month_comercial['E2'].border = thin_border
               ws_template_month_comercial['E2'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
               ws_template_month_comercial['E2'].font = Font(bold=True)
               ws_template_month_comercial.column_dimensions['F'].width = 6
               ws_template_month_comercial['F2'] = 'NPS'
               ws_template_month_comercial['F2'].border = thin_border
               ws_template_month_comercial['F2'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
               ws_template_month_comercial['F2'].font = Font(bold=True)
               ws_template_month_comercial.column_dimensions['G'].width = 10
               ws_template_month_comercial['G2'] = 'Devuelto'
               ws_template_month_comercial['G2'].border = thin_border
               ws_template_month_comercial['G2'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
               ws_template_month_comercial['G2'].font = Font(bold=True)
               ws_template_month_comercial.column_dimensions['H'].width = 25
               ws_template_month_comercial['H2'] = 'Seguro'
               ws_template_month_comercial['H2'].border = thin_border
               ws_template_month_comercial['H2'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
               ws_template_month_comercial['H2'].font = Font(bold=True)
               ws_template_month_comercial.column_dimensions['I'].width = 10
               ws_template_month_comercial['I2'] = 'INV Days'
               ws_template_month_comercial['I2'].border = thin_border
               ws_template_month_comercial['I2'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
               ws_template_month_comercial['I2'].font = Font(bold=True)
               contador_page += 1
               

               range_excel_overview_comercial = 'BCDEFGHI'#range to fill in the template to loop
               contador_overview_comercial = 3
               for i in range(len(df_comercial)):
                    contador_df_comercial = 0
                    for x in range_excel_overview_comercial:
                         if contador_df_comercial == 2:
                              ws_template_month_comercial[str(x)+str(contador_overview_comercial)] = str(df_comercial.iloc[i,contador_df_comercial]).lower()
                         else:
                              ws_template_month_comercial[str(x)+str(contador_overview_comercial)] = df_comercial.iloc[i,contador_df_comercial]
                         ws_template_month_comercial[str(x)+str(contador_overview_comercial)].border = thin_border
                         ws_template_month_comercial[str(x)+str(contador_overview_comercial)].alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
                         contador_df_comercial+=1
                    contador_overview_comercial+=1
          wb_template.save(excel_name) 
     
          send_email(config_gmail_user,config_gmail_pass,config_email_tl[tl],[config_email_copia,config_head],excel_name,'Cierre comisiones '+tl+' '+periodo,config_msg_tl)
          os.remove(excel_name)
          if only_one_tl != None:
                break





     