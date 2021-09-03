from Functions.Head import head
from Functions.Team_leader import tl
from Functions.Comerciales import comercial
from Functions.Config import *

if __name__ == '__main__':
    print('Eliga una opcion:\n 1-Comercial\n 2-TL\n 3-Head')
    opcion = input('Eliga una opcion:') 
    if opcion == '1':
        print('\tEliga una opcion:\n \t1-Único comercial\n \t2-Todos los comerciales')
        sub_opcion = input('\tElija una opcion:')
        if sub_opcion == '1':
            sales = input('\t\tQue comercial quiere cerrar?:')
            comercial(sales)
        else:
            comercial()
    if opcion == '2':
        print('\tEliga una opcion:\n \t1-Único TL\n \t2-Todos los TL')
        sub_opcion = input('\tElija una opcion:')
        if sub_opcion == '1':
            for team_leader in config_email_tl.keys():
                print('\t\t'+team_leader)
            sales = input('\t\tQue tl quiere cerrar?:')
            tl(team_leader)
        else:
            tl()
    if opcion == '3':
        head()