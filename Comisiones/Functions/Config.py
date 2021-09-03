#Email


config_gmail_user= 'reports_es@auto1.com'
config_gmail_pass= 'RE8156ES'
config_head = 'jorge.creus@autohero.com'
config_email_copia = 'jorge.creus@autohero.com'
config_email_tl = {'RRU':'jorge.creus@autohero.com','LGG':'jorge.creus@autohero.com','JPV':'jorge.creus@autohero.com'}



#LOV

config_more_200 = 50
config_more_100 = 30
config_less_100 = 20
config_objective = 0.80
config_objective_sold = 300
config_bonus_nps_65 = 300
config_bonus_nps_30 = 200
config_bonus_nps_less_30 = 0
config_promoters = 8
config_detractors = 7
config_warranty_I = 15
config_warranty_II = 10
config_bonus_warranty_objt = 0.20
config_bonus_warranty = 100
config_finance = 15
config_insurance = 10
config_bonus_finance_objt = 0.50
config_bonus_finance = 200
config_reduce_objetive_days = 45
config_reduce_objetive = 16
config_objective_number = 20
#API
scopes = ['https://www.googleapis.com/auth/spreadsheets']
sheet_id='1TW9LpAHAutSeufcDE8780DbcsMdwR3vFEzLNIncHnrI'
range_data_month='!A1:X'
range_all_data = 'Data_comisiones!A2:L'



#Email Body

config_msg_tl = '''Buenas,

Este es un mensaje automático,

Adjuntar los extras de este mes para poder incluirlos (Festivos,horas extra)

IMPORTANTE, NO CONTESTAR AL CORREO DE {},

CONTESTAR SÓLO A {} ,

Este es el desglose final de vuestro equipo.

Un saludo
'''.format(config_gmail_user,config_email_copia)


config_msg_head = '''Buenas,

Adjunto excel con las comisiones del equipo comercial de Autohero 

Un saludo
'''


config_msg_comercial = '''Buenas,

Este es un mensaje automático,

IMPORTANTE, NO CONTESTAR AL CORREO DE {},

Para cualquier duda o consulta responder sálo a vuestro a TL.

Un saludo
'''.format(config_gmail_user)