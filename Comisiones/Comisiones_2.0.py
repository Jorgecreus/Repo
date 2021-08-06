import pandas as pd
import config_2
import pickle
import os
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import openpyxl
from openpyxl.styles.borders import Border, Side
from openpyxl.styles.numbers import FORMAT_PERCENTAGE

pd.options.mode.chained_assignment = None


if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials_BA.json', config_2.scopes)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

# Conect to the G-Sheet and get the general data
service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()

result_data = sheet.values().get(spreadsheetId=config_2.sheet_id,
                                range=config_2.range_data).execute()
values_data = result_data.get('values', [])


#Reshape the DF and split in in two
data = pd.DataFrame.from_records(values_data)
data.rename(columns=data.iloc[0],inplace=True)
data.drop(data.index[0],inplace=True)
data['Booking date'] = pd.to_datetime(data['Booking date'],dayfirst=True)
data['Contract_signed'] = pd.to_datetime(data['Contract_signed'],dayfirst=True)
data['car_handover_date'] = pd.to_datetime(data['car_handover_date'],dayfirst=True)
data['Contribution Start Date'] = pd.to_datetime(data['Contribution Start Date'],dayfirst=True)
data['NPS_nota media'] = data['NPS_nota media'].replace('', '-1')
data['NPS_nota media'] = data['NPS_nota media'].fillna(value=-1)
data['NPS_nota media'] = data['NPS_nota media'].astype('int')


df = data[data.columns[0:data.columns.get_loc('Warranty_type')+1]]
df.dropna(how='all',inplace=True)
df['Inv Days'] = (df['Contract_signed'] - df['Booking date']).dt.days
df_objetive = data[data.columns[data.columns.get_loc('Email'):]]
df_objetive.dropna(how='all',inplace=True)
# df_objetive['today'] = pd.Timestamp.today()
df_objetive['Days working']  = (pd.Timestamp.today() - df_objetive['Contribution Start Date'] ).dt.days



def comercial_comis(df,df_objetive,comercial):

    #Calculate all the datapoints
    df_comercial= df[df['Comercial'] == comercial ]
    df_comercial.reset_index(inplace=True)
    df_comercial.drop(columns=['index'],inplace=True)
    # df_comercial['NPS_nota media'] = df_comercial['NPS_nota media'].astype('int64')
    # df_comercial['Inv Days'] = df_comercial['Inv Days'].astype('int64')
    more_200 = sum((df_comercial['Inv Days']>=200))
    more_100 = sum((df_comercial['Inv Days']<200) & (df_comercial['Inv Days']>=100))
    less_100 = sum((df_comercial['Inv Days']<100))
    more_200_euros = more_200 * config_2.more_200
    more_100_euros = more_100 * config_2.more_100
    less_100_euros = less_100 * config_2.less_100
    ventas = more_200 + more_100 + less_100
    index = df_objetive[df_objetive['Email']==comercial].index.values
    objetive = int(df_objetive.loc[index,['Objetivo']].values)
    percent_objective = ventas/objetive
    tl = df_objetive.loc[index,'Team Leader'].values
    try:
        extra = int(df_objetive.loc[index,['Extra']].values)
    except:
        extra = 0
    motive = str(df_objetive.loc[index,['Motivo']].values)
    warranty_I = sum((df_comercial['Warranty_type'] =='Garantia I'))
    warranty_II = sum((df_comercial['Warranty_type'] =='Garantia II'))
    finance = sum((df_comercial['Payment_type']=='CASH_AND_FINANCE'))
    insurance = sum((df_comercial['Insurance_type'] =='Seguro de Vida')) + sum((df_comercial['Insurance_type'] =='Protección Total'))
    promoters= sum((df_comercial['NPS_nota media']>8))
    detractors= sum((df_comercial['NPS_nota media']<7 ) &  (df_comercial['NPS_nota media']>=0 ))
    total= sum((df_comercial['NPS_nota media']>=0 ))
    percent_warranty = (warranty_I + warranty_II) / objetive
    percent_finance = finance / objetive
    try:
        nps_score = int((promoters - detractors)/total*100)
    except:
        nps_score = 0
    bonus_finan_euros = 0
    bonus_ventas_euros = 0
    bonus_warranty_euros = 0
    bonus_nps_euros = 0
    if percent_objective >= config_2.objective:
        bonus_ventas_euros = percent_objective * config_2.objective_sold
        if percent_finance >= config_2.bonus_finance_objt:
            bonus_finan_euros = config_2.bonus_finance
        if percent_warranty >= config_2.bonus_warranty_objt:
            bonus_warranty_euros = config_2.bonus_warranty
        if nps_score >= 65:
            bonus_nps_euros = config_2.bonus_nps_65
        elif nps_score >= 30:
            bonus_nps_euros = config_2.bonus_nps_30
        else:
            bonus_nps_euros = config_2.bonus_nps_less_30

    variable_ammount = int(more_200_euros+ more_100_euros + less_100_euros + warranty_I * config_2.warranty_I +  warranty_II * config_2.warranty_II + finance * config_2.finance + insurance * config_2.insurance + bonus_ventas_euros +bonus_finan_euros + bonus_warranty_euros + bonus_nps_euros + extra)
    #Reshape Excel

    wb_template = openpyxl.load_workbook('template.xlsx')
    ws_template_resumen = wb_template.worksheets[0]
    ws_template_resumen['F6'] = config_2.period
    ws_template_resumen['F5'] = config_2.inicial_date
    ws_template_resumen['G5'] = config_2.end_date
    ws_template_resumen['F3'] = comercial
    ws_template_resumen['F7'] = objetive
    ws_template_resumen['F11'] = more_200
    ws_template_resumen['G11'] = more_200_euros
    ws_template_resumen['F12'] = more_100
    ws_template_resumen['G12'] = more_100_euros 
    ws_template_resumen['F13'] = less_100
    ws_template_resumen['G13'] = less_100_euros
    ws_template_resumen['F14'] = percent_objective
    ws_template_resumen['G14'] = bonus_ventas_euros
    ws_template_resumen['F15'] = (bonus_ventas_euros + more_200_euros+ more_100_euros+ less_100_euros)
    ws_template_resumen['F19'] = finance
    ws_template_resumen['F20'] = insurance
    ws_template_resumen['G19'] = finance * config_2.finance
    ws_template_resumen['G20'] = insurance * config_2.insurance
    ws_template_resumen['F21'] = percent_finance
    ws_template_resumen['G21'] = bonus_finan_euros
    ws_template_resumen['F22'] = finance * config_2.finance + insurance * config_2.insurance + bonus_finan_euros
    ws_template_resumen['F26'] = warranty_I
    ws_template_resumen['G26'] = warranty_I * config_2.warranty_I
    ws_template_resumen['F27'] = warranty_II
    ws_template_resumen['G27'] = warranty_II * config_2.warranty_II
    ws_template_resumen['F28'] = percent_warranty
    ws_template_resumen['G28'] = bonus_warranty_euros
    ws_template_resumen['F29'] = bonus_warranty_euros +  warranty_I * config_2.warranty_I + warranty_II * config_2.warranty_II
    ws_template_resumen['F32'] = total
    ws_template_resumen['F33'] = nps_score
    ws_template_resumen['F34'] = bonus_nps_euros
    ws_template_resumen['F37'] = extra
    ws_template_resumen['F38'] = motive
    ws_template_resumen['F41'] = variable_ammount
    contador = 51
    for i in range(len(df_comercial)):
        total = 0
        ws_template_resumen['E'+str(contador)] = df_comercial.loc[i,'ID (stock_number)']
        ws_template_resumen['F'+str(contador)] = df_comercial.loc[i,'Nº de orden']
        if df_comercial.loc[i,'Inv Days'] >=200:
            ws_template_resumen['I'+str(contador)] = config_2.more_200
            total += config_2.more_200
        elif (df_comercial.loc[i,'Inv Days'] <200 and df_comercial.loc[i,'Inv Days'] >=100):
            ws_template_resumen['K'+str(contador)] = config_2.more_100
            total += config_2.more_100
        else:
            ws_template_resumen['L'+str(contador)] = config_2.less_100
            total += config_2.less_100
        ws_template_resumen['M'+str(contador)] = df_comercial.loc[i,'Payment_type']
        ws_template_resumen['N'+str(contador)] = df_comercial.loc[i,'Insurance_type']
        if (df_comercial.loc[i,'Insurance_type'] == 'Seguro de Vida' or df_comercial.loc[i,'Insurance_type'] == 'Protección Total'):
            total += config_2.insurance
        ws_template_resumen['P'+str(contador)] = df_comercial.loc[i,'Warranty_type']
        if df_comercial.loc[i,'Warranty_type'] == 'Garantia I':
            total += config_2.warranty_I
        if df_comercial.loc[i,'Warranty_type'] == 'Garantia II':
            total += config_2.warranty_II
        ws_template_resumen['Q'+str(contador)] = total
        contador += 1
    wb_template.save('name.xlsx')






def tl_comis(df,df_objetive,tl):
    df_tl = df_objetive[df_objetive['Team Leader'] == tl ]
    df_tl.reset_index(inplace=True)
    df_tl.drop(columns=['index'],inplace=True)
    for i in range(len(df_tl)):
        comercial = df_tl.loc[i,'Email']
        df_tl.loc[i,'More 200'] = sum((df['Inv Days']>=200) & (df['Comercial']==comercial) )
        df_tl.loc[i,'More 100'] = sum((df['Inv Days']>=100) & (df['Inv Days']<200) & (df['Comercial']==comercial) )
        df_tl.loc[i,'Less 100'] = sum((df['Inv Days']<100) & (df['Comercial']==comercial) )
        df_tl.loc[i,'Ventas €']  = df_tl.loc[i,'More 200'] * config_2.more_200 + df_tl.loc[i,'More 100'] * config_2.more_100 + df_tl.loc[i,'Less 100'] * config_2.less_100
        sales = df_tl.loc[i,'More 200'] + df_tl.loc[i,'More 100'] + df_tl.loc[i,'Less 100']
        df_tl.loc[i,'percent_objective'] = sales / int(df_tl.loc[i,'Objetivo'])
        promoters= sum((df['NPS_nota media']>8 ) & (df['Comercial']==comercial))
        detractors= sum((df['NPS_nota media']<7 ) & (df['Comercial']==comercial) & (df['NPS_nota media']>=0 ))
        total= sum((df['NPS_nota media']>=0 ) & (df['Comercial']==comercial))
        try:
            df_tl.loc[i,'Nota NPS']= int((promoters - detractors)/total*100)
        except:
            df_tl.loc[i,'Nota NPS']=0

    print(df_tl)

# comercial = 'antonio.magalhaes@autohero.com'
# comercial_comis(df,df_objetive,comercial)
tl= 'LGG'

if __name__ == '__main__':
    # tl_comis(df,df_objetive,tl)
    print(pd.Timestamp('2020-10')+ pd.offsets.MonthEnd(-1))
    print(pd.Timestamp('10-2020') + pd.offsets.MonthBegin(-1))
