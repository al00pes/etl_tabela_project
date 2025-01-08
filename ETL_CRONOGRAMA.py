#Ignorar os erros
import warnings
warnings.filterwarnings('ignore')
import pandas as pd
import aspose.tasks as tsk

#Lendo o arquivo project
path_file = 'CRONOGRAMA-PROJETO_BRRJ01.mpp'

#Abrindo projeto existente
project_subsea = tsk.Project(path_file)

#Convertendo o arquivo para CSV
dados = []

for resource in project_subsea.resources:
    if resource is not None and (resource.name):
        nome_recurso = resource.name
        inicio = resource.start
        termino = resource.finish
        trabalho_restante = resource.remaining_work
        horas_trabalhadas = resource.actual_work
        #status_trabalho = resource.peak_units
        
        
        dados.append({
            "Nome": nome_recurso,
            "Inicio": inicio,
            "Termino": termino,
            "trabalho_restante": trabalho_restante,
            "horas_trabalhadas": horas_trabalhadas
                        
        })

#Transformando em um dataframe
df = pd.DataFrame(dados)
#df.head(10)

#Tratando os dados e transformando em string
df['Inicio'] = df['Inicio'].astype(str)
df['Termino'] = df['Termino'].astype(str)

#Fatiando os dados da coluna e exibindo somente os 10 caracteres
df['Inicio'] = df['Inicio'].str.slice(0,10)
df['Termino'] = df['Termino'].str.slice(0,10)

#Invertendo os valores 
#str.split('-') -> Separa o compomente sempre que houver a ocorrência do - e cria um lista
#.str[::-1] - Inverte os compemente das lista
#str.join('-') -> Junta os elementos da lista e coloca  um delimitador -
df['Inicio'] =  df['Inicio'].str.split('-').str[::-1].str.join('-')
df['Termino'] =  df['Termino'].str.split('-').str[::-1].str.join('-')

#Exibindo o dataframe
print(df)

#Salvando o dataframe em excel

#Salvando o arquivo em excel
df.to_excel('CRONOGRAMA-PROJETO_BRRJ01.xlsx',index=False)
#print('Planilha exportada')
