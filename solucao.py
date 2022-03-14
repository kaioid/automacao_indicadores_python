import pandas as pd
import pathlib
import win32com.client as win32

# importando as bases
emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')

vendas = vendas.merge(lojas, on='ID Loja')  # mesclando as duas planilhas para pegar o nome das lojas

dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja'] == loja, :]

dia_indicador = vendas['Data'].max()

caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')
arquivos_pasta_backup = caminho_backup.iterdir()
lista_nome_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

for loja in dicionario_lojas:
    if loja not in lista_nome_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()

    # salvando na pasta
    nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    local_arquivo = caminho_backup / loja / nome_arquivo
    dicionario_lojas[loja].to_excel(local_arquivo)

# definição de metas
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500

for loja in dicionario_lojas:
    # indicadores
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data'] == dia_indicador, :]

    # faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()

    faturamento_dia = vendas_loja_dia['Valor Final'].sum()

    # diversidade de produtos
    qtde_produtos_ano = len(vendas_loja['Produto'].unique())
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())

    # ticket medio
    valor_venda = vendas_loja.groupby('Código Venda').sum()
    ticket_medio_ano = valor_venda['Valor Final'].mean()

    # ticket medio_dia
    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()

    # enviando emails
    outlook = win32.Dispatch('outlook.application')

    nome = emails.loc[emails['Loja'] == loja, 'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]
    mail.Subject = f'One page dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'

    # corpo do email
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
        or_qtde_ano = 'red'
    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    # corpo do email
    mail.HTMLBody = f'''
    <p>Bom dia, {nome}!</p>
    
    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da loja <strong>{loja}</strong> foi:</p>
    
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
    <p>Att., Kaio</p>
    '''
    # anexo
    attachment = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    mail.Attachments.Add(str(attachment))
    mail.Send()
    print(f'E-mail da loja {loja} enviado com sucesso!')

# faturamento
faturamento_lojas_ano = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas_ano.sort_values(by='Valor Final', ascending=False)
nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
faturamento_lojas_ano.to_excel(fr'Backup Arquivos Lojas\{nome_arquivo}')

vendas_dia = vendas.loc[vendas['Data'] == dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)
nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking Diario.xlsx'
faturamento_lojas_dia.to_excel(fr'Backup Arquivos Lojas\{nome_arquivo}')

# enviar email para a diretoria
outlook = win32.Dispatch('outlook.application')


mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja'] == 'Diretoria', 'E-mail'].values[0]
mail.Subject = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month}'
mail.Body = f'''Prezados,

Segue em anexo os rankings do dia e do ano de todas as lojas.

Melhor loja em faturamento do dia: {faturamento_lojas_dia.index[0]} com faturamento de: R$ {faturamento_lojas_dia.iloc[0,0]:.2f}.
Pior loja em faturamento do dia: {faturamento_lojas_dia.index[-1]} com faturamento de: R$ {faturamento_lojas_dia.iloc[-1,0]:.2f}.

Melhor loja em faturamento do ano: {faturamento_lojas_ano.index[0]} com faturamento de: R$ {faturamento_lojas_ano.iloc[0,0]:.2f}.
Pior loja em faturamento do ano: {faturamento_lojas_ano.index[-1]} com faturamento de: R$ {faturamento_lojas_ano.iloc[-1,0]:.2f}

Qualquer dúvida, estou à disposição

Att.,
Kaio
'''
# anexos
attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
mail.Attachments.Add(str(attachment))
attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Diario.xlsx'
mail.Attachments.Add(str(attachment))

mail.Send()
print('E-mail da Diretoria enviado')
