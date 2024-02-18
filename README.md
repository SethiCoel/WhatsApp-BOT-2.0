<h1>WhatsApp BOT 2.0</h1> 
  Esta é uma versão melhorada do bot anterior, que estava apresentando algumas falhas então resolvir refazê-lo usando o Selenium que é mais recomendavél para automação.

### Atualizações e Melhorias
<ul>
  <li>Opção de reenvio de mensagens:</li>
    <br>-Ele irá pegar a planilha salva na pasta "Não enviados" e tentará enviar novamente a mensagem caso o mesmo esteja correto. Os que derem certo serão excluídos da planilha.

  <Br><li>Opção de ativação aos domingos de forma automatica:</li>
    <br>-Ao usar essa opção irá programar o bot para que inicie a automação aos domingos. Basta ativar a opção e deixar o computador ligao que ele irá enviar as mensagens a partir das 07:00h. O mesmo irá desligar o computador quando terminar.

<br><li>Melhorias:</li>
<br>-Usando selenium melhora bastante a perfomance do app pois ele consegue manipular e entender certos comando do navegador faciliando assim a automação e economizando mais tempo, não necessitando de tempos fixos para o funcionamento.
</ul>



<h2>Sobre</h2>
<p>Fiz esse bot para enviar mensagens a clientes cujo a data esteja próxima do vencimento com um dia de antecedência.<br>
A mensagem que é enviada contém o nome do cliente e sua data de vencimento, cada cliente recebe a mensagem com seu nome e sua data de vencimento, sendo possível a troca da mensagem e data de antecedência da forma que preferir.<br>

<h2>Como Funciona:</h2>
<p>O script foi feito para ser usado com planilhas. Logo é necessario ter uma planilha com os dados dos clientes, contendo as informações nessa ordem: <br> <br> Nome, Telefone, Dia de vencimento.<br>

É necessario também que a planilha esteja com o nome "Planilha Atualizada" para o funcionamento.<br> 

Caso você tenha uma planilha pronta, fora dessa ordem e com o nome diferente, o script tem uma funcionalidade que cria uma cópia dos dados dos clientes na ordem correta.<br>

Para isso é necessário colocar a planilha que deseja fazer a cópia na pasta "Planilha". Vale lembra que provavelmente sua planilha esteja com a posição diferente requerida no código, sendo necessário o ajuste do mesmo.<br> 

A pasta "planilha" é criada automaticamente após usar a opção "Criar lista de clientes".

Feito isso, criará uma planilha com o nome "Planilha Atualizada".<br> 

Agora está tudo pronto para iniciar o BOT.</p>

<h2>Funcionalidades:</h2>
<ul>
  <li> O bot abrirá o site para caso o celular não esteja conectado no WhatsApp Web para fazer essa configuração, após 20 segundos a pagina será fechada e o bot irá inciar o envio das mensagens </li><br>

  <li> O Bot tem uma função que verifica se algum cliente está sem o número de telefone, alertando e registrando-o na "Planilha de Reenvio" com as infomções do cliente na pasta "Não Enviados" que será criada logo após o ocorrido. </li><br>

  <li> O Bot também tem uma simples verificação caso o número seja possivelmente inválido. Vale mencionar que o mesmo não garante que seja realmente um número inválido,
  caso ocorra algo que o impeça de enviar a mensagem será acionado o mesmo problema de número inválido. Com isso será também adicionado a planilha "Planilha de Reenvio" nas pasta "Não Enviados". </li><br>

  <li> Com a lista de clientes que não receberam a mensagem terá um controle para a verificação do cliente e assim poder corrigi-lo.</li>
</ul>


<h2>Tecnologias</h2>
<div>
  <img src="https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54">
        Selenium | Openpyxl
</div>
