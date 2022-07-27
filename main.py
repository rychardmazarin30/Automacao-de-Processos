
# Libs utilizadas
import pandas as pd
import win32com.client as win32 
import pathlib

# -- Caso utilize o Jupyter essas libs serão o suficiente, mas caso esteja usando algum outro -- 
# -- é importante que você tenha o openpyxl instalado para que o código funcione corretamente -- 

# Importando as Bases de Dados

email = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')

# Passo 2 - Criar uma Tabela para cada Loja e Definir o dia do Indicador

# Incluir Nome da loja em vendas
vendas = vendas.merge(lojas, on='ID Loja')

# Adicionando todas as lojas a um dicionário

dict_lojas = {}
for loja in lojas['Loja']:
    dict_lojas[loja] = vendas.loc[vendas['Loja']==loja, :]
    
# Definindo o dia do indicador para sempre como o ultimo dia da tabela, ou seja o dia mais recente

dia_indicador = vendas['Data'].max()

# Identificar se a pasta já existe para sempre ter o backup do ultimo dia de envio para cada loja

caminho_backup = pathlib.Path('Backup Arquivos Lojas') # Essa pasta você deve criar na mesma que estiver codando.
arquivos_pasta_backup = caminho_backup.iterdir()
lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

for loja in dict_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()
        
    # Salvar dentro da pasta    
    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)
    local_arquivo = caminho_backup / loja / nome_arquivo
    dict_lojas[loja].to_excel(local_arquivo)

# definição das variáveis de metas

meta_faturamento_dia = 1000

meta_faturamento_ano = 1650000
# Fiz isso só pra na hora de formatar o e-mail ficar bonitinho '.'
meta_faturamento_ano = f'<strong>R$</strong>{meta_faturamento_ano:_.2f}'
meta_faturamento_ano = meta_faturamento_ano.replace('.', ",").replace('_', '.')

meta_produtos_dia = 4
meta_produtos_ano = 120

meta_TM_dia = 500
meta_TM_ano = 500

# Calculando os indicadores e enviando os e-mails pros gerentes de cada loja

for loja in dict_lojas:
  
  #Calculando os indicadores
  vendas_loja = dict_lojas[loja]
  vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]

  # faturamento
  faturamento_ano = vendas_loja['Valor Final'].sum()
  faturamento_ano = f'<strong>R$</strong>{faturamento_ano:_.2f}'
  faturamento_ano = faturamento_ano.replace('.', ",").replace('_', '.')
  faturamento_dia = vendas_loja_dia['Valor Final'].sum()

  # diversidade de produtos
  # Esse unique(), tira os valores duplicados de uma mesma coluna.
  qtde_produtos_ano = len(vendas_loja['Produto'].unique()) 
  qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())

  # ticket médio
  # groupby(), serve para agrupar valores iguais e fazer akguma coisa com os outros valores
  valor_venda = vendas_loja.groupby("Código Venda").sum() 
  ticket_medio_ano = valor_venda['Valor Final'].mean()
  # ticket_medio_dia
  valor_venda_dia = vendas_loja_dia.groupby("Código Venda").sum()
  ticket_medio_dia = valor_venda_dia['Valor Final'].mean()
  
  
  # Enviando email para os gerentes

  #  -- é importante que você tenha o outlook instalado para enviar dessa maneira --
  outlook = win32.Dispatch('outlook.application')

  nome = email.loc[email['Loja']==loja, "Gerente"].values[0]
  mail = outlook.CreateItem(0)
  mail.To = email.loc[email['Loja']==loja, "E-mail"].values[0]
  mail.Subject = 'OnePage {}/{} - Loja {}'.format(dia_indicador.day, dia_indicador.month, loja)


  if faturamento_dia < meta_faturamento_dia:
      cor_faturamento_dia = 'red'
  else:
      cor_faturamento_dia = 'green'

  if faturamento_ano < meta_faturamento_ano:
      cor_faturamento_ano = 'red'
  else:
      cor_faturamento_ano = 'green'

  if qtde_produtos_dia < meta_produtos_dia:
      cor_produtos_dia = 'red'
  else:
      cor_produtos_dia = 'green'

  if qtde_produtos_ano < meta_produtos_ano:
      cor_produtos_ano = 'red'
  else:
      cor_produtos_ano = 'green'

  if ticket_medio_dia < meta_TM_dia:
      cor_TM_dia = 'red'
  else:
      cor_TM_dia = 'green'

  if ticket_medio_ano < meta_TM_ano:
      cor_TM_ano = 'red'
  else:
      cor_TM_ano = 'green'


  # Corpo do E-mail utilizando HTML
  mail.HTMLBody = f'''
  <p>Bom dia, {nome}</p>

  <p>O Resultado de <b>ontem ({dia_indicador.day}/{dia_indicador.month})</b> da <b>loja {loja}</b> foi:</p>

  <table>
    <tr>
      <th>Indicador</th>
      <th>Valor Dia</th>
      <th>Meta Dia</th>
      <th>Cenário Dia</th>
    </tr>
    <tr>
      <td>Faturamento</td>
      <td style="text-align: center"><strong>R$</strong>{faturamento_dia:.2f}</td>
      <td style="text-align: center"><strong>R$</strong>{meta_faturamento_dia:.2f}</td>
      <td style="text-align: center"><font color={cor_faturamento_dia}>◙</font></td>
    </tr>
    <tr>
      <td>Diversidade de Produtos</td>
      <td style="text-align: center">{qtde_produtos_dia}</td>
      <td style="text-align: center">{meta_produtos_dia}</td>
      <td style="text-align: center"><font color={cor_produtos_dia}>◙</font></td>
    </tr>
    <tr>
      <td>Ticket Médio</td>
      <td style="text-align: center"><strong>R$</strong>{ticket_medio_dia:.2f}</td>
      <td style="text-align: center"><strong>R$</strong>{meta_TM_dia:.2f}</td>
      <td style="text-align: center"><font color={cor_TM_dia}>◙</font></td>
    </tr>

  </table>
  <br>
  <table>
    <tr>
      <th >Indicador</th>
      <th>Valor Ano</th>
      <th>Meta Ano</th>
      <th>Cenário Ano</th>
    </tr>
    <tr>
      <td>Faturamento</td>
      <td style="text-align: center">{faturamento_ano}</td>
      <td style="text-align: center">{meta_faturamento_ano}</td>
      <td style="text-align: center"><font color={cor_faturamento_ano}>◙</font></td>
    </tr>
    <tr>
      <td>Diversidade de Produtos</td>
      <td style="text-align: center">{qtde_produtos_ano}</td>
      <td style="text-align: center">{meta_produtos_ano}</td>
      <td style="text-align: center"><font color={cor_produtos_ano}>◙</font></td>
    </tr>
    <tr>
      <td>Ticket Médio</td>
      <td style="text-align: center"><strong>R$</strong>{ticket_medio_ano:.3f}</td>
      <td style="text-align: center"><strong>R$</strong>{meta_TM_ano:.3f}</td>
      <td style="text-align: center"><font color={cor_TM_ano}>◙</font></td>
    </tr>

  </table>

  <p>Segue em anexo a planilha como todos os dados para a analise.</p>
  <p>Qualque dúvida estou a disposição.</p>
  <p>Atenciosamente,</p>
  <p>Rychard Mazarin</p>


  '''

  # Anexos (pode colocar quantos quiser):
  attachment = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
  mail.Attachments.Add(str(attachment))

  mail.Send()


# Rankings, feitos do ano atual e do ultimo dia de vendas para envio pra diretoria

faturamento_lojas = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas = faturamento_lojas.sort_values(by='Valor Final', ascending=False)

# Formatando para ficar bonitinho  '.' -- ANO --
melhor_faturamento_lojas = f'{faturamento_lojas.iloc[0,0]:_.2f}'
melhor_faturamento_lojas = melhor_faturamento_lojas.replace('.', ",").replace('_', '.')

pior_faturamento_lojas = f'{faturamento_lojas.iloc[-1,0]:_.2f}'
pior_faturamento_lojas = pior_faturamento_lojas.replace('.', ",").replace('_', '.')

nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
# crie uma pasta dentro da pasta de backups para ficar esses aquivos de rankings anuais e diarios em formato .xlsx
faturamento_lojas.to_excel(r'Backup Arquivos Lojas\Ranking Para Diretoria\{}'.format(nome_arquivo))

vendas_dia = vendas.loc[vendas['Data']==dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)

# Formatando para ficar bonitinho  '.' -- DIA --
melhor_faturamento_lojas_dia = f'{faturamento_lojas_dia.iloc[0,0]:_.2f}'
melhor_faturamento_lojas_dia = melhor_faturamento_lojas_dia.replace('.', ",").replace('_', '.')

pior_faturamento_lojas_dia = f'{faturamento_lojas_dia.iloc[-1,0]:_.2f}'
pior_faturamento_lojas_dia = pior_faturamento_lojas_dia.replace('.', ",").replace('_', '.')

nome_arquivo = '{}_{}_Ranking Dia.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_dia.to_excel(r'Backup Arquivos Lojas\Ranking Para Diretoria\{}'.format(nome_arquivo))


# Enviando email para os gerentes (Antes de enviar não esqueça de salvar na pasta de rankings)

#  -- é importante que você tenha o outlook instalado para enviar dessa maneira --
outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)
mail.To = email.loc[email['Loja']=='Diretoria', "E-mail"].values[0]
mail.Subject = 'Ranking de Vendas {}/{}'.format(dia_indicador.day, dia_indicador.month)
mail.Body = f'''
Prezados, bom dia 

Melhor loja do dia em faturamento: Loja {faturamento_lojas_dia.index[0]} com faturamento R${melhor_faturamento_lojas_dia}
Pior loja do dia em faturamento: Loja {faturamento_lojas_dia.index[-1]} com faturamento R${pior_faturamento_lojas_dia}

Melhor loja do ano em faturamento: Loja {faturamento_lojas.index[0]} com faturamento R${melhor_faturamento_lojas}
Pior loja do ano em faturamento: Loja {faturamento_lojas.index[-1]} com faturamento R${pior_faturamento_lojas}

Segue anexo os rankings do ano e de dia de todas as lojas.

Qualquer dúvida estou a disposição.

Atenciosamente,
Rychard Mazarin

'''

# Anexo dos arquivo dos rankings :
attachment = pathlib.Path.cwd() / caminho_backup / 'Ranking Para Diretoria' / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
mail.Attachments.Add(str(attachment))
attachment = pathlib.Path.cwd() / caminho_backup / 'Ranking Para Diretoria' / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
mail.Attachments.Add(str(attachment))

# .Send() enviara o E-mail
mail.Send()





