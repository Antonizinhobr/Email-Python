<div align="center">
  <h1> :rocket: Bem vindo ao meu repositório:rocket: </h1>
</div>

<br>
<br>

<div>
  <h1> Minhas redes sociais</h1>
  <a href="https://www.youtube.com/channel/UC88QEmxaSyY_V2vXn1RMgQQ" target="_blank"><img src="https://img.shields.io/badge/YouTube-FF0000?style=for-the-badge&logo=youtube&logoColor=white" target="_blank"></a>
<a href="https://www.instagram.com/_anthonny_michael_dev/" target="_blank"><img src="https://img.shields.io/badge/-Instagram-%23E4405F?style=for-the-badge&logo=instagram&logoColor=white" target="_blank"></a>
<a href="https://www.linkedin.com/in/anthonny-michael-64450a206/" target="_blank"><img src="https://img.shields.io/badge/-LinkedIn-%230077B5?style=for-the-badge&logo=linkedin&logoColor=white" target="_blank"></a> 
</div>



# Como baixar esse repositório? :sassy_man:

1. Baixando diretamente pelo github.

    <img src="/Email-Python/Email-Python/readme/Github Download Repo.png" />

2.  Baixando pelo "gitclone" no seu prompt de comando. (Sintaxe: git clone "https://www.site.com")

    <img src="/Email-Python/Email-Python/readme/Git clone.png" />
    
<br>

# O que é esse repositório?

1. Disparador de e-mails através da leitura de linhas/colunas de uma base EXCEL.

# Email-Python
Projeto em Python.

Sua funcionalidade é ler um arquivo/base EXCEL, e formar um e-mail a cada leitura de linha.

O código lê cada linha do banco de dados EXCEL e cria um email para remetentes específicos, sendo um email estruturado com os dados dessa linha lidos no EXCEL.

Projeto voltado ao monitoramento do atendimento ao cliente, onde informa a um operador de telemarketing sobre seu monitoramento pendente de assinatura na ferramenta de monitoramento, alertando ele e seu supervisor (gerente) sobre o monitoramento pendente via e-mail.

Contar a quantidade de e-mails enviados por cada supervisor e coordenador e-mail, onde ao final criar um e-mail, trazendo a contagem de cada e-mail enviado por supervisor e coordenador, criando um e-mail com essas contagens e enviá-lo ao gerente de operações como um relatório.

# Como funciona o disparo de e-mails?

O programa em Python, disparador de e-mails, lê linha a linha da base/arquivo em EXCEl. Dentro de um loop em For, ele cria um e-mail para cada linha, utilizando o atributo iterrows, da biblioteca Pandas, e com isso, é montado um e-mail a partir da leitura de cada linha do EXCEL, e enviado dentro do loop do For, e somente acaba quando a leitura finaliza por completo. São 2 loops, um de enviar e-mail para o supervisor, outro para o coordenador, e ao final de tudo isso, contabiliza o envio de e-mails para cada e-mail de supervisor e coordenador, e armazena isso como variável dentro de um argumento de função, pois como cada loop em For está em funções diferentes, foi preciso criar uma THREAD com varíaveis globais de contagens ee utilizadas como argumentos de cada função de envio de e-mail e com isso, trazer esses argumentos como variáveis dentro do e-mail de envio ao gerente, para printar a contagem de cada e-mail e colocar no e-mail ao gerente como forma de relatório.

1. <strong>Segue imagem de exemplo de e-mail enviado para o operador, notificando-o da pendência de monitoria de assinatura</strong>
<br>
   <img src="/Email-Python/Email-Python/readme/operador.png" />
<br>

2. <strong>Segue imagem de exemplo de e-mail enviado para o supervisor, notificando-o da pendência de monitoria de aplicação de feedback:</strong>
<br>
   <img src="/Email-Python/Email-Python/readme/supervisor.png" />
<br>

3. <strong>Segue imagem de exemplo de e-mail enviado para o coordenadoir, notificando-o da pendência de monitoria de seu supervisor:</strong>
<br>
   <img src="/Email-Python/Email-Python/readme/coordenador.jpg" />
<br>

4. <strong>Segue imagem de exemplo de e-mail enviado para o gerente, enviando um e-mail com contagem de monitorias por supervisor e coordenador:</strong>
<br>
   <img src="/Email-Python/Email-Python/readme/gerente.jpg" />
