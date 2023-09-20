'''Desafio - Automação de indicadores e processos de lojas

1) Importar bibliotecas e criar variáveis iniciais
2) Analisar planilhas do banco de dados e importá-las para o código
3) Trabalhar os arquivos do banco de dados
4) Criar um arquivo para cada loja
5) Salvar o backup em pastas do computador
6) Calcular os indicadores
7) Enviar e-mail para os gerentes (one page dos gerentes)
8) Criar ranking para a diretoria
9) Enviar e-mail para a diretoria (one page da diretoria)'''

# 1) Importar bibliotecas e criar variáveis iniciais
import pandas as pd
from IPython.display import display
import win32com.client as win32
import pathlib

dicionario_lojas = {}                                               # Criar dicionário - armazenar tabela de cada uma das lojas

meta_faturamento_dia = 1000
meta_faturamento_ano = 1_650_000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500

# 2) Analisar planilhas do banco de dados e importá-las para o código
emails = pd.read_excel(r'Emails.xlsx')
lojas = pd.read_csv(r'Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Vendas.xlsx')

# 3) Trabalhar os arquivos do banco de dados
vendas = vendas.merge(lojas, on='ID Loja')

# 4) Criar um arquivo para cada loja
for loja in lojas['Loja']:                                          # Para cada loja dentro da coluna 'Loja' do DF lojas...
    dicionario_lojas[loja] = vendas.loc[vendas['Loja']==loja, :]    # Adiciona a tabela de cada loja ao dicionário

# 5) Salvar o backup em pastas do computador
caminho_backup = pathlib.Path(r'Backup arquivos lojas')                   # Definir o caminho do backup
arquivos_pasta_backup = caminho_backup.iterdir()                          # Pegar todos os arquivos

lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]  # Criar uma lista com o nome dos arquivos
dia_indicador = vendas['Data'].max()                                      # Dia indicador é o dia mais recente

for loja in dicionario_lojas:                                             # Criar uma pasta com o nome de cada loja
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()

    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)
    local_arquivo = caminho_backup / loja / nome_arquivo
    dicionario_lojas[loja].to_excel(local_arquivo)                        # Salvar arquivo excel dentro de cada pasta

# 6) Calcular os indicadores
for loja in dicionario_lojas:

    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]

    # Indicador 1 - Faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()

    # Indicador 2 - Diversidade de produtos
    qtde_produtos_ano = len(vendas_loja['Produto'].unique())
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())

    # Indicador 3 - Ticket médio
    valor_venda_ano = vendas_loja.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_ano = valor_venda_ano['Valor Final'].mean()

    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()

    # 7) Enviar e-mail para os gerentes (one page dos gerentes)
    outlook = win32.Dispatch('outlook.application')

    nome = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]          # Nome do(a) gerente
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]        # Endereço de e-mail do(a) gerente
    mail.Subject = f'OnePage {nome} dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'
        
    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'
    if qtde_produtos_dia >= meta_qtdeprodutos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'
    if qtde_produtos_ano >= meta_qtdeprodutos_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'
    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    mail.HTMLBody = f'''
    <p>Bom dia, {nome}</p>

    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>

    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_dia:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_dia}</td>
        <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
        <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
      </tr>
    </table>
    <br>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_ano:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_ano}</td>
        <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
        <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
      </tr>
    </table>

    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

    <p>Qualquer dúvida estou à disposição.</p>
    
    <p>Att., </p>
    </p>Eduardo Ferreira</p>
    '''

    # Anexos dos e-mails para os gerentes:
    attachment  = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    mail.Attachments.Add(str(attachment))

    mail.Send()
    
    print('E-mail da Loja {} enviado'.format(loja))

# 8) Criar ranking para a diretoria
faturamento_lojas = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum(numeric_only=True)
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)

nome_arquivo = '{}_{}_ranking_anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_ano.to_excel(r'Backup arquivos lojas\{}'.format(nome_arquivo))

vendas_dia = vendas.loc[vendas['Data']==dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum(numeric_only=True)
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)

nome_arquivo = '{}_{}_ranking_dia.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_dia.to_excel(r'Backup arquivos lojas\{}'.format(nome_arquivo))

# 9) Enviar e-mail para a diretoria (one page diretoria)
outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja']=='Diretoria', 'E-mail'].values[0]
mail.Subject = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month}'
mail.Body = f'''
Prezados,

Melhor loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0, 0]:.2f}
Pior loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1, 0]:.2f}

Melhor loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0, 0]:.2f}
Pior loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1, 0]:.2f}

Segue em anexo os rankings do dia e do ano de todas as lojas.

Qualquer dúvida estou à disposição.

Att.,
Eduardo Ferreira
'''

# Anexos do e-mail para a diretoria:
attachment  = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_ranking_anual.xlsx'
mail.Attachments.Add(str(attachment))
attachment  = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_ranking_dia.xlsx'
mail.Attachments.Add(str(attachment))

mail.Send()

print('E-mail da Diretoria enviado')

print('Fim do código')