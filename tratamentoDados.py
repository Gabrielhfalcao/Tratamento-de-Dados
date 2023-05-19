import csv
import os
import random
import pandas as pd
import locale
from datetime import date

pasta = r"C:\Users\B900780\Desktop\Lista Aleatória Python" 
pastaDef = pasta.replace("\\", "/")
extensoes = ['csv'] 
arq = []
arqManuais = []
arqAutomaticos = []
listaParaSelecionar = []
escolhas = []
escolhasAutomaticas = []

dados1 = []
dados2 = []
dados3 = []
dados4 = []
dados5 = []
dados6 = []
dados7 = []
dados8 = []
dados9 = []
dados10 = []

datasSelecionadas = []
codTarifaSelecionada = []
listafinal = []

arquivos = os.listdir(pastaDef)
for i in arquivos:
	extensao = i.split('.')[-1]
	if extensao in extensoes:
		x = i.split('.')[0]
		listaParaSelecionar.append(x)
		arq.append(i)

i = 0
while (i < 10):
    arqManuais.append(listaParaSelecionar[i])
    i += 1   

i2 = 10
while (i2 < 20):
    arqAutomaticos.append(listaParaSelecionar[i2])
    i2 += 1

while len(escolhas) < 5:
    escolha = random.choice(arqManuais)
    if not any([s.endswith(escolha[-3:]) for s in escolhas]):
        escolhas.append(escolha)

while len(escolhasAutomaticas) < 5:
    escolhaAuto = random.choice(arqAutomaticos)
    if not any([s.endswith(escolhaAuto[-3:]) for s in escolhas]) and not any([s.endswith(escolhaAuto[-3:]) for s in escolhasAutomaticas]):
        escolhasAutomaticas.append(escolhaAuto)    

for i in escolhasAutomaticas:
    escolhas.append(i)

def povoarListaDados(dadosListaAleatorio, listaDeDados):
    with open(str(dadosListaAleatorio + ".csv")) as arquivocsv:
        ler = csv.DictReader(arquivocsv, delimiter="\t")
        for linha in ler:
            dados1.append(linha)
            listaDeDados.append(linha)

def povoarVariaveis(listaDeDados):
    i = 0
    while i < len(listaDeDados):
        if (str(listaDeDados[i]['DT_LNC'])[0:10] in datasSelecionadas) or (listaDeDados[i]['CD_TAR'] in codTarifaSelecionada):
            i += 1
            continue
        else:
            listafinal.append(listaDeDados[i])
            datasSelecionadas.append(listaDeDados[i]['DT_LNC'][0:10]) 
            codTarifaSelecionada.append(listaDeDados[i]['CD_TAR'])
            break

povoarListaDados(escolhas[0], dados1)
povoarListaDados(escolhas[1], dados2)
povoarListaDados(escolhas[2], dados3)
povoarListaDados(escolhas[3], dados4)
povoarListaDados(escolhas[4], dados5)
povoarListaDados(escolhas[5], dados6)
povoarListaDados(escolhas[6], dados7)
povoarListaDados(escolhas[7], dados8)
povoarListaDados(escolhas[8], dados9)
povoarListaDados(escolhas[9], dados10)

povoarVariaveis(dados1)
povoarVariaveis(dados2)
povoarVariaveis(dados3)
povoarVariaveis(dados4)
povoarVariaveis(dados5)
povoarVariaveis(dados6)
povoarVariaveis(dados7)
povoarVariaveis(dados8)
povoarVariaveis(dados9)
povoarVariaveis(dados10)

indice = 0

while indice < len(listafinal):
    valor = float(listafinal[indice]['VR_COB'])
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    valor = locale.currency(valor, grouping=True, symbol=True)
    listafinal[indice]['VR_COB'] = valor

    indice += 1

data_atual = str(date.today())
nomeArquivoFinal = "autoverificacao" + "_" + data_atual[5:7] + "_" + data_atual[0:4] + ".xlsx"
print(nomeArquivoFinal)

df = (pd.DataFrame(listafinal)).style.set_properties(**{'text-align': 'center'})

writer = pd.ExcelWriter('C:/Users/B900780/Desktop/Lista Aleatória Python/Relatorios Registros Aleatórios/' + nomeArquivoFinal, engine='openpyxl')
df.to_excel(writer, sheet_name='Dados', index=False)
worksheet = writer.sheets['Dados']

for col in worksheet.columns:
    max_length = 0
    column = col[0].column_letter
    
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
        
    adjusted_width = (max_length + 5)
    worksheet.column_dimensions[column].width = adjusted_width

writer.close()

print()
print('linhas do arquivo gerado:')
print(listafinal[0])
print(listafinal[1])
print(listafinal[2])
print(listafinal[3])
print(listafinal[4])
print(listafinal[5])
print(listafinal[6])
print(listafinal[7])
print(listafinal[8])
print(listafinal[9])


print()
print('Arquivos escolhidos aleatóriamente:')
print(escolhas[0])
print(escolhas[1])
print(escolhas[2])
print(escolhas[3])
print(escolhas[4])
print(escolhas[5])
print(escolhas[6])
print(escolhas[7])
print(escolhas[8])
print(escolhas[9])
