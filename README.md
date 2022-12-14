<h1>Automação de Processos</h1>

<p>Esse foi um projeto que fiz no curso Python Impressionador, nele conseguimos automatizar varias tarefas repetitivas, como:</p>
<ul>
  <li>Manipulação de Planilhas</li>
  <li>Envio de E-mails</li>
  <li>Criação e manipulação de pastas do computador</li>
  <li>Backups de tudo que é feito</li>
</ul>

<p>Imagine que você trabalha em uma grande rede de lojas de roupa com 25 unidades espalhadas por todo o Brasil.</p>

<p>Todo dia, pela manhã, a equipe de análise de dados calcula os chamados One Pages e envia para o gerente de cada loja o OnePage da sua loja, bem como todas as informações usadas no cálculo dos indicadores.</p>

<p>Um One Page é um resumo muito simples e direto ao ponto, usado pela equipe de gerência de loja para saber os principais indicadores de cada loja e permitir em 1 página (daí o nome OnePage) tanto a comparação entre diferentes lojas, quanto quais indicadores aquela loja conseguiu cumprir naquele dia ou não.</p>

<p>Então basicamente o objetivo é que o nosso código sempre que quisermos, ele analise a planilha de Vendas e calcule os indicadores, para que seja feito o envio dos e-mails para cada gerente de cada uma das 25 lojas e também para a diretoria com o rankings de melhor e pior loja.</p>

<p>As planilhas estão dentro da pasta 'Bases de Dados', crie uma pasta chamada 'Automação de Processos' e insira ela dentro, junto com o arquivo do código</p>

<p>É muito Importante também que você crie uma pasta para Backup, você pode colocar o nome que desejar mas no código eu coloquei 'Backup Arquivos Lojas', então se caso mudar não se esqueça de mudar no código também </p>

<p>Serão enviados 26 E-mails ao total, 25 para os respectivos gerentes de cada loja e um para a diretoria, abaixo os modelos dos e-mails:</p>

<h3>Diretoria</h3>

![Corpo E-mail Diretoria](https://user-images.githubusercontent.com/98194579/181138901-1eba8989-f574-4c8c-81f1-5b19e5761649.png)


<h3>Gerentes</h3>

![Corpo E-mail Gerentes](https://user-images.githubusercontent.com/98194579/181138928-eceb5499-a4b7-46cc-9aa9-697fbaf84545.png)

<p>Os Indicadores em vermelho e verde indicam se as metas foram batidas, e junto desses e-mails sempre vão em anexo a planilha de vendas para os gerentes e as planilhas de rankings para a diretoria.</p>

<h2>AVISO!</h2>

<p>Se você utilizar o Jupyter Notebook ele funcionará normalmente apenas com as libs: </p>
<ul>
  <li>Pandas</li>
  <li>win32.client</li>
  <li>pathlib</li>
</ul>
<p>Mas, caso esteja utilizando qualquer outro editor de código você precisará ter instalado também a lib <b>openpyxl</b></p>

<p>É isso, boas automações.</p>

<h4>Desenvolvido por Rychard Mazarin</h4>



