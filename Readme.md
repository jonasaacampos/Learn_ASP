# ASP Classic

> Anotaçõees de estudo sobre Active Server Pages, com finalidade de documentação de aprendizagem e compartilhamento de conhecimento

## Como o ASP funciona?

- ASP é a tecnologia que executa scripts no lado do servidor (backend)
- O navegador faz a requisição ao arquivo .asp, e este executa o script no backend e retorna um texto em html como resposta a essa requisição
- A linguagem de programação padrão do ASP é o VBScript, que é uma versão enxuta do Visual Basic da Microsoft

### Como escrever arquivos ASP

- Arquivos com a extensão .asp nada mais são que arquivos de html. Dentro do html são inseridas tags delimitadores de <% %> que indicam que naquele trecho serão executados scripts no backend.
- O método Response.Write() retorna um html ap servidor. Este método pode ser substituído pelo sinal de **=**

**cídigo do arquivo asp**

```html
<body>
    <%
    Response.Write ("<h1>Hello World!</h1>")
    %>
    <%="<p style='color:#0000ff'>O sinal de igualdade (=) tem a mesma funíão do método Response.Write().</p>"%>
</body>


</body>
```

**cídigo lido pelo navegador**

```html
<body>
    <h1>Hello World!</h1>
    <p style='color:#0000ff'>O sinal de igualdade (=) tem a mesma funíão do método Response.Write().</p>
</body>
```

### Variáveis em ASP

> Variáveis são receptáculos para guardar algo

- Para declarar uma variável em VBScript usamos os declaradores `Dim`, `Public` ou `Private`.
- Para declarar um array (variável especial) basta informar a quantidade de valores entre "( )" após a declaração da variável `Dim novoArrayDeNomes(2)`. Neste caso teremos uma lista contento 3 nomes.
- Para um array de dimensões múltiplas (matriz) basta informar as dimensões separadas por vírgulas dentro dos "( )" após a declaração da mesma: `Dim novoArrayMultidimensional(2,3)`, neste caso, temos uma matriz de 3 linhas e 4 colunas.
- Tipos de variáveis:
  - Variáveis de sessão: informações pertinentes é um único usuário, que estão disponível em toda a aplicação
  - Variávels de aplicação: armazenam informações de todos os usuários e estão disponíveis para uma aplicação específica

### Funíões e Condicionais

```vbscript
Function myfunction()
  some statements
  myfunction=some value
End Function
```

[Exemplo de funíões e condicionais aqui.](03-Condicionais-e-funcoes.asp)

### Laãos de repetição

- `For... Next`: repete a instrução em determinado numero de vezes
- `For Each... Next`: repete a instrução para cada item do array
- `Do... Loop`: repete enquanto for verdadeira determinada condição
- `While... Wend`: _semelhante ao Do...Loop_, mas **evite usar** este laão.

#### For... Next

```vb
For i = 0 To 10
  response.write("The number is " & i & "<br />")
Next

' A instrução step determina o incremento ou decremento do passo do contador (para decremento utilizar um valor negativo)'

For i = 0 To 10 Step 2
  response.write("The number is " & i & "<br />")
Next

'Para sair do laço For... Next utilizar a instrução Exit For'

For i = 0 To -100 Step 10
  response.write("The number is " & i & "<br />")
  If i = -50 Then Exit For
Next
```

#### For Each... Next

```vb
<%
Dim cars(2)
cars(0)="Volvo"
cars(1)="Saab"
cars(2)="BMW"

For Each x In cars
  response.write(x & "<br />")
Next
%>
```

#### Do... Loop

```vb
'Repete enquanto for verdadeiro'
Do While i > 10
    algumCodigo...
    i++
Loop

'Repete enquanto for verdadeiro, mas verifica a condição após a execução do primerio laão'
Do
  some code
Loop While i > 10

'Execute se i for DIFERENTE de 10'
Do Until i=10
  some code
Loop

'Executa pelo menos uma vez, se i=10 sai do laão'
Do
  some code
Loop Until i=10

'Sair do laão Do'
Do Until i=10
  i=i-1
  If i<10 Then Exit Do
Loop

```

### Entrada de dados pelo usuário

> Os métodos Request.QueryString e Request.Form são usados para receber dados que o usuário inseriu um um formulírio na p�gina

- `Request.QueryString` : Coleta valores utilizando a requisição **GET**
- `Request.Form` : Coleta valores utilizando a requisição **POST**

```vb
<form method="get" action="simpleform.asp">
    First Name: <input type="text" name="fname"><br>
    Last Name: <input type="text" name="lname"><br><br>
    <input type="submit" value="Submit">
</form>

'Ao clicar em enviar, é retornado do servidor a url com o valor das variáveis preenchidas, após o ponto de "?", com o nome dos parâmetros e o valor inserido pelo usuário'

'https://www.w3schools.com/simpleform.asp?fname=Jonas&lname=Campos'

```

### Cookies

> Cookies são usados para identificar as ações de um usuário
 
[Veja um exemplo de Cookie em ação clicando aqui.](06-Cookies.asp)

## Objeto de Sessão (ASP Session Object)

> O ASP cria um cookie único para cada usuário, após o fechamento do navegador, este cookie é descartado.

```vb
'O tempo padrão de fim da sessão é de 20 minutos. Para alterar usamos o parâmetro Timeout.'
<%
    Session.Timeout=5
%>

'Para finalizar a sessão imediatamente, usamos:'
<%
Session.Abandon
%>
```















## Para saber mais

https://www.w3schools.com/asp/asp_introduction.asp