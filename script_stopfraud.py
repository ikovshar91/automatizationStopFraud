import pandas as pd
import os

homedir = os.path.expanduser('~')

itog = pd.read_excel(homedir + '\\Desktop\\stopfraud\\stopfraud.xlsx', index_col=False)
zvonki = pd.read_excel(homedir + '\\Desktop\\stopfraud\\zvonki.xlsx', index_col=False)

itog2 = itog.copy()

itog2.insert(0, 'totalnumber', itog2['Номер А'].astype(str) + itog2['Номер В'].astype(str))

zvonki.insert(0, 'totalnumber',
              zvonki['Calling number normalized'].astype(str) + zvonki['Called number normalized'].astype(str))

itog2 = itog2.merge(zvonki, how='left', on='totalnumber')
itog2['Оператор А'] = itog2['FAS carrier A']
itog2['Дата звонка'] = itog2['Время начала вызова']

itog2['Транк'] = itog2['Carrier']

if itog2['Транк'][i] != 'TO_ROAM':
    itog2['Транк'][i] = itog2['Carrier'][i].str[2:].str[:3]
else:
    itog2['Транк'][i] = 'TO_ROAM'

itog2.drop(
    ['totalnumber', 'Calling number normalized', 'Called number normalized', 'Redirecting number normalized', 'Carrier',
     'FAS carrier A', 'FAS carrier B', 'FAS carrier C', 'Rule_ID', 'Rule name CDR',
     'День', 'Время начала вызова', 'VLR nmber', 'Action name', 'Events'], axis=1, inplace=True)


itog2.drop_duplicates(subset=['Дата обращения','От какого оператора','Номер А','Оператор А','Номер В','Оператор В',
                              'Дата звонка','Длительность','Дата запроса','Комментарий','Ответ оператора'],
                      keep='first',inplace=True)

itog2.to_excel(homedir + '\\Desktop\\stopfraud\\СтопФрод.xlsx', sheet_name='itog', index=False)

RTK = itog2['Транк'].isin(['RTK', 'VPM'])
GTK = itog2['Транк'].isin(['GTK', 'BEE'])
MGF = itog2['Транк'].isin(['MGF'])
MTS = itog2['Транк'].isin(['MTS'])
MTT = itog2['Транк'].isin(['MTT'])
TTK = itog2['Транк'].isin(['TTK'])
EQN = itog2['Транк'].isin(['EQN'])

itog2.drop(
    ['Дата обращения', 'От какого оператора', 'Транк', 'Дата запроса', 'Комментарий', 'Оператор В', 'Ответ оператора']
    , axis=1, inplace=True)

# ------------РТК--------------------------------------------------------------------------------------
if itog2[RTK].empty:
    print('у РТК пусто')
else:
    print('Сформирован файл на отправку РТК')
    itog2[RTK].to_excel(homedir + '\\Desktop\\stopfraud\\Письма на отправку\\Мошеннические вызовы.Соц.Инженерия_РТК.xlsx', sheet_name='itog', index=False)
# ------------------------------------------------------------------------------------------------------

# -------------Билайн-----------------------------------------------------------------------------------
if itog2[GTK].empty:
    print('у GTK пусто')
else:
    print('Сформирован файл на отправку GTK')
    itog2[GTK].to_excel(homedir + '\\Desktop\\stopfraud\\Письма на отправку\\Мошеннические вызовы.Соц.Инженерия_Билайн.xlsx', sheet_name='itog', index=False)
# ------------------------------------------------------------------------------------------------------

# -------------Мегафон----------------------------------------------------------------------------------
if itog2[MGF].empty:
    print('у MGF пусто')
else:
    print('Сформирован файл на отправку MGF')
    itog2[MGF].to_excel(homedir + '\\Desktop\\stopfraud\\Письма на отправку\\Мошеннические вызовы.Соц.Инженерия_Мегафон.xlsx', sheet_name='itog', index=False)
# -----------------------------------------------------------------------------------------------------

# -------------МТС----------------------------------------------------------------------------------
if itog2[MTS].empty:
    print('у MTS пусто')
else:
    print('Сформирован файл на отправку MTS')
    itog2[MTS].to_excel(homedir + '\\Desktop\\stopfraud\\Письма на отправку\\Мошеннические вызовы.Соц.Инженерия_МТС.xlsx', sheet_name='itog', index=False)
# -----------------------------------------------------------------------------------------------------

# -------------MTT----------------------------------------------------------------------------------
if itog2[MTT].empty:
    print('у MTT пусто')
else:
    print('Сформирован файл на отправку MTT')
    itog2[MTT].to_excel(homedir + '\\Desktop\\stopfraud\\Письма на отправку\\Мошеннические вызовы.Соц.Инженерия_MTT.xlsx', sheet_name='itog', index=False)
# -----------------------------------------------------------------------------------------------------

# -------------TTK----------------------------------------------------------------------------------
if itog2[TTK].empty:
    print('у TTK пусто')
else:
    print('Сформирован файл на отправку TTK')
    itog2[TTK].to_excel(homedir + '\\Desktop\\stopfraud\\Письма на отправку\\Мошеннические вызовы.Соц.Инженерия_TTK.xlsx', sheet_name='itog', index=False)
# -----------------------------------------------------------------------------------------------------

# -------------EQN----------------------------------------------------------------------------------
if itog2[EQN].empty:
    print('у EQN пусто')
else:
    print('Сформирован файл на отправку EQN')
    itog2[EQN].to_excel(homedir + '\\Desktop\\stopfraud\\Письма на отправку\\Мошеннические вызовы.Соц.Инженерия_EQN.xlsx', sheet_name='itog', index=False)
# -----------------------------------------------------------------------------------------------------

input('Нажмите Enter для завершения программы...\n')
