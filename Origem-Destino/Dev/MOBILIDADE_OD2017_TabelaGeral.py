# -*- coding: utf-8 -*-
'''
MOBILIDADE_OD2017_TabelaGeral.py
Processo de importação da Tabela 30 para o Python, mantendo a ordenação de 
origem e destino para fácil acesso.
'''

# Import Statments
import pandas as pd

# Paths configurations
folder_path = r'C:\_RAW\PROJETOS\201910_Mapeamento Gestao Urbana\XLS\\'
file_name = 'MOBILIDADE_OD2017_Tab30.xlsx'


# Cria o DataFrame com os dados da tabela 30 (Viagens totais)
df = pd.read_excel(folder_path+file_name, header=7, nrows=518)
## Ajusta o cabeçalho (primeira coluna)
columns_name = list(range(519))
columns_name[0] = 'z_o'
columns_name[518] = 'Total_o'
df.columns = columns_name

## Ajusta o Total dos destino para um nome mais claro
df.loc[517, 'z_o'] = 'Total_d'

## Ajusta o index para a primeira coluna
df = df.set_index('z_o')


zonas = [317, 340]

##### G L O B A L #####
# DESTINO
global_destino = df.loc['Total_d'][:-1]
global_destino = pd.DataFrame( {
            'z_d': global_destino.index.to_list(),
            'obs': global_destino
            })
# ORIGEM
global_origem = df['Total_o'][:-1]
global_origem = pd.DataFrame( {
            'z_o': global_origem.index.to_list(),
            'obs': global_origem
            })

##### L O C A L #####
# DESTINO nas zonas
local_destino = df.loc[:,zonas][:-1].T.sum()
local_destino = pd.DataFrame( {
            'z_o': local_destino.index.to_list(),
            'obs': local_destino
            })

# ORIGEM nas zonas
local_origem = df.loc[zonas].sum()[:-1]
    # Transformar em uma tabela
local_origem = pd.DataFrame( {
            'z_d': local_origem.index.to_list(),
            'obs': local_origem
            })
    
    
    
files_name = {'OD2017_Global_Destino.csv' : global_destino,
              'OD2017_Global_Origem.csv' : global_origem,
              'OD2017_Local_Destino.csv' : local_origem,
              'OD2017_Local_Origem.csv' : local_origem}

'''
FORMA INDIVIDUAL DE EXPORTAR

local_destino.to_csv(r'C:\_RAW\PROJETOS\201910_Mapeamento Gestao Urbana\CSV\OD2017_Local_Destino.csv', 
                        float_format='%.0f',
                        index=False)
                        

local_origem.to_csv(r'C:\_RAW\PROJETOS\201910_Mapeamento Gestao Urbana\CSV\OD2017_Local_Origem.csv', 
                        float_format='%.0f',
                        index=False)
'''

for i in files_name:
    files_name[i].to_csv(r'C:\_RAW\PROJETOS\201910_Mapeamento Gestao Urbana\CSV\{}'.format(i), 
                        float_format='%.0f',
                        index=False)
