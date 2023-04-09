import pandas as pd
import openpyxl
import win32com.client as win32

### Objetivo: Pegar todos os gastos presentes na tabela de cada tipo e soma-los. Apresentar o intervalo de tempo e qual detalhamento é o que mais gasta. Remover colunas mês, ano e fonte

# Importar a base de dados
planilha = pd.read_excel('TestesPandas.xlsx')
#print (planilha)

pd.set_option('display.max_columns', None)

# Colunas usadas: Tipo de Gasto; Detalhamento; Valor R$; Data

tipo = planilha[['Tipo de Gasto','Valor R$']].groupby('Tipo de Gasto').sum()
print (tipo)

### Qual detalhamento é o que mais gasta: Filtrar por detalhamento, valor, somar os valores e mostrar somente o detalhamento que mais gasta
detalhamento = planilha[['Detalhamento','Valor R$']].groupby('Detalhamento').sum()
print (detalhamento)

    ## Pegar o detalhamento e filtrar pelo mais caro, e mostrar o mais caro

        # Agrupar os dados por tipo e somar os valores
gastos_por_tipo = tipo.groupby('Tipo de Gasto')['Valor R$'].sum().reset_index()

        # Filtrar em ordem decrescente e pegar o tipo que mais gastou
tipo_mais_gastou = gastos_por_tipo.sort_values(by='Valor R$', ascending=False).iloc[0]['Tipo de Gasto']

        # Filtrar em ordem decrescente e pegar o valor que mais gastou
tipo_mais_gastou_valor = gastos_por_tipo.groupby('Valor R$').sum().sort_values(by='Valor R$', ascending=False).iloc[0].name
print (f'O tipo de gasto mais caro foi "{tipo_mais_gastou}" cujo preço foi: {f"R${tipo_mais_gastou_valor:,.2f}"}')

        # Pegar a primeira e a ultima data da tabela
data = planilha[['Data']]
primeira_data = data.sort_values(by='Data', ascending=True).iloc[0]['Data']
primeira_data_formatada = primeira_data.strftime('%d/%m/%Y')
ultima_data = data.sort_values(by='Data', ascending=False).iloc[0]['Data']
ultima_data_formatada = ultima_data.strftime('%d/%m/%Y')

# Enviar email

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'emailforexemple123@email.com'
mail.Subject = 'Relatório de Gastos Pessoais'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue relatório dos gastos pessoais enviados:</p>

<p> Tabela de tipos de gastos agrupados e seus respectivos valores somados: </p>
<p>{tipo.to_html(formatters={'Valor R$':'R${:,.2f}'.format})}</p>

<p> Tabela de detalhamento de gastos agrupados e seus respectivos valores somados: </p>
<p>{detalhamento.to_html(formatters={'Valor R$':'R${:,.2f}'.format})}</p>

<p> Outras informações: </p>

<p>O período analisado foi de {primeira_data_formatada} até {ultima_data_formatada}.</p>
<p>O tipo de gasto que saiu mais caro foi {tipo_mais_gastou} e o valor total desse tipo de gasto foi de {f"R${tipo_mais_gastou_valor:,.2f}"}</p>


Qualquer dúvida estou à disposição
Att.,
Caio
'''

mail.Send()
