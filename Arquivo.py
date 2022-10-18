# Passo 1: Importar Bibliotecas
import pandas as pd
import win32com.client as win32

#Importar tabela
tabela = pd.read_excel('iventario.xlsx')


# Visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela)

material = tabela[['MATERIAL','QUANTIDADE']].groupby('MATERIAL').count()
print(material)

# Enviar um e-mail com os dados
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'jacksonjunior2109@gmail.com'
mail.Subject = 'Iventário de Informática - URCE'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o relatório de equipamentos iventariados na sede URCE:</p>

<p>Iventário:</p>
{material.to_html()}

<p>Qualquer dúvida estou a disposição.</p>

<p>Att.,</p>

<p>Jackson Almeida</p>
'''
mail.Send()

print('E-mail Enviado')