<p align="center">
   <a href='https://jonasaacampos.github.io/portfolio/'>
      <img alt="ASP Classic - Badge" src="https://img.shields.io/static/v1?color=blue&label=ASP%20NET&message=VB-Script&style=for-the-badge&logo=classic-asp"/>
      </a>
</p>

<h1>ASP Classic</h1>

<img alt="brain" src="img/asp-logo.png" width=150 align=right>

<h2>Anotações e exemplos de sintáxe e lógica ASP em VBScript</h2>

![](https://img.shields.io/badge/VbScript-informational?style=flat&logo=ASP&logoColor=white&color=blue)
![](https://img.shields.io/badge/ASP-informational?style=flat&logo=ASP&logoColor=white&color=blue)

> Anotaçõees de estudo sobre Active Server Pages, com finalidade de documentação de aprendizagem e compartilhamento de conhecimento

[![](https://img.shields.io/badge/feito%20com%20%E2%9D%A4%20por-jaac-cyan)](https://jonasaacampos.github.io/portfolio/)
[![LinkedIn Badge](https://img.shields.io/badge/LinkedIn-Profile-informational?style=flat&logo=linkedin&logoColor=white&color=0D76A8)](https://www.linkedin.com/in/jonasaacampos)

## Índice do conteúdo

- [Índice do conteúdo](#índice-do-conteúdo)
- [Como o ASP funciona?](#como-o-asp-funciona)
  - [Como escrever arquivos ASP](#como-escrever-arquivos-asp)
  - [Variáveis em ASP](#variáveis-em-asp)
  - [Condicionais](#condicionais)
    - [If Then / Else | ElseIf Then](#if-then--else--elseif-then)
    - [Select Case](#select-case)
  - [Funções](#funções)
  - [Laços de repetição](#laços-de-repetição)
    - [For... Next](#for-next)
    - [For Each... Next](#for-each-next)
    - [Do... Loop](#do-loop)
  - [Entrada de dados pelo usuário](#entrada-de-dados-pelo-usuário)
  - [Cookies](#cookies)
- [Objeto de Sessão (ASP Session Object)](#objeto-de-sessão-asp-session-object)
- [Orientação a Objetos](#orientação-a-objetos)
  - [POO no VbScript](#poo-no-vbscript)
    - [Exemplo de classe em VBScript](#exemplo-de-classe-em-vbscript)
- [Funções](#funções-1)
  - [Tipos de dados (verificação)](#tipos-de-dados-verificação)
    - [VarType()](#vartype)
  - [Tipos de dados (conversão de tipos)](#tipos-de-dados-conversão-de-tipos)
    - [Int() e Fix()](#int-e-fix)
  - [Funções de Tratamento (Strings)](#funções-de-tratamento-strings)
  - [Funções de Tratamento de tempo (Data e Hora)](#funções-de-tratamento-de-tempo-data-e-hora)
    - [DateDiff()](#datediff)
    - [DateAdd()](#dateadd)
    - [DatePart()](#datepart)
    - [Parâmetros para data](#parâmetros-para-data)
      - [*interval*](#interval)
      - [*firstdayofweek*](#firstdayofweek)
    - [*firstweekofyear*](#firstweekofyear)
  - [Funções de Cálculo](#funções-de-cálculo)
- [Eventos no VBScript](#eventos-no-vbscript)
  - [Classes](#classes)
- [Para saber mais](#para-saber-mais)
  
---


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

[📖 voltar para o índice 📖](#índice-do-conteúdo)

---

### Condicionais

#### If Then / Else | ElseIf Then

```vbscript
 <%
    Function horaAtual(hora)
    
        If hora <  11 Then
            response.write("Bom dia")
        ElseIf hora <= 18 Then
            response.write("Boa tarde")
        Else
            Response.Write("Boa noite")
        End If

        Response.Write ("<br/>")
    
    End Function
  %>
```

#### Select Case

```vbscript
<!-- Condicional Select Case -->

<%
  Function diaDaSemana(numeroDiaDaSemana)
    Select Case numeroDiaDaSemana
      Case 1
        Response.Write ("Segundou galera!<br/>")
      Case 2
        Response.Write ("Terça de fazer feira<br/>")
      Case 3
        Response.Write ("Quarta da <strong>feijoada</strong><br/>")
      Case 4
        Response.Write ("Quinta do bom humor<br/>")
      Case 5
        Response.Write ("Sextou!<br/>")
      Case 6
        Response.Write ("Sábado do louvor<br/>")
      Case 7
        Response.Write ("Domingo da família<br/>")
      Case else
        Response.Write ("<br/>Dia da semana inválido<br/>")

    End Select
  End Function
%>
```

### Funções
> Funções e subrotinas ajudam para melhor manutenção do código

A diferença básica entre uma `Funcion` e `Sub` é que a função sempre retorna algo. Sub-rotinas seriam como funções `void` no Java.

```vbscript
Function myfunction()
  some statements
  myfunction=some value
End Function
```

[Exemplo de funções e condicionais aqui.](03-Condicionais-e-funcoes.asp)

[📖 voltar para o índice 📖](#índice-do-conteúdo)

---

### Laços de repetição

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

'Sair do laço Do'
Do Until i=10
  i=i-1
  If i<10 Then Exit Do
Loop

```

[📖 voltar para o índice 📖](#índice-do-conteúdo)

---

### Entrada de dados pelo usuário

> Os métodos Request.QueryString e Request.Form são usados para receber dados que o usuário inseriu um um formulírio na página

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

[📖 voltar para o índice 📖](#índice-do-conteúdo)

---

## Orientação a Objetos

> O Visual basic **não é** uma linguagem pensada para Programação Orientada a Objetos (POO), mas podemos utilizar _alguns dos princípios_ da POO em códigos VBScript dentro do ASP.

- Objeto é uma idéia, uma abstração escrita em código
- Instância é a representação lógica de um objeto, é 'o nascimento' do objeto
- Atributos: personalidade do objeto
- Métodos: ações que um objeto pode executar
- Construtor: é o método especial executado automaticamente quando o objeto é criado (instanciado)
- Herança: agrupamento lógico hierárquico de classes e objetos

### POO no VbScript

- Apenas um único construtor é aceito por classe
- O construtor não aceita parâmetros
- não aceita herança

#### Exemplo de classe em VBScript

```vb
'classe privada
Private m_Name

' métodos acessores
' GET
Public Property Get Name()  
    Name = m_Name
End Property

' LET (também podemos usar o SET como método)
Public Property Let Name(sName)
    m_Name = sName
End Property
```

[📖 voltar para o índice 📖](#índice-do-conteúdo)

---

## Funções

### Tipos de dados (verificação)

#### VarType()

> retorna o subtipo uma variável

| **Constant**      | **Value** | **Description**                               |
|-------------------|-----------|-----------------------------------------------|
| **vbEmpty**       | 0         | Empty (uninitialized)                         |
| **vbNull**        | 1         | Null (no valid data)                          |
| **vbInteger**     | 2         | Integer                                       |
| **vbLong**        | 3         | Long integer                                  |
| **vbSingle**      | 4         | Single-precision floating-point number        |
| **vbDouble**      | 5         | Double-precision floating-point number        |
| **vbCurrency**    | 6         | Currency                                      |
| **vbDate**        | 7         | Date                                          |
| **vbString**      | 8         | String                                        |
| **vbObject**      | 9         | Automation object                             |
| **vbError**       | 10        | Error                                         |
| **vbBoolean**     | 11        | Boolean                                       |
| **vbVariant**     | 12        | Variant (used only with arrays of Variants)   |
| **vbDataObject**  | 13        | A data-access object                          |
| **vbByte**        | 17        | Byte                                          |
| **vbArray**       | 8192      | Array                                         |

Função que verificam o tipo de dado contido na variável e retorna `True` ou `False`

```vb
IsArray()
IsNumeric()
IsDate()
IsEmpty()
IsNull()
IsObject()
```

### Tipos de dados (conversão de tipos)

```vb
Cboll() 'Converte uma expressão ou valor para Boolean
cByte()
cCur()  'Converte uma expressão oou valor para Currency
cDate()
cDbl()
CInt()
CLng()
CSng()  'Converte uma expressão ou variável para Single
CStr()
```

#### Int() e Fix()

> retornam somente a parte inteira de um número

- `Int()`: retorna o primeiro número **menor** ou igual
- `Fix()`: retorna o primeiro número **maior** ou igual

### Funções de Tratamento (Strings)

```vb
Asc()   'Retorna o código ANSI correspondente a primeira letra da string
Chr()   'Retorna um caracter ao receber um código ANSI
Len()   'Retorna o tamanho da string
LCase() 'Retorna uma string convertida para caixa alta
UCase() 'Retorna uma string convertida para caixa alta
Left(string, length)  'Retorna x caractes a partir da esquerda de uma string
Right(string, length) 'Retorna x caractes a partir da direita de uma string
Mid(string, start[, length])  'Retorna uma string de n até n'
String(number, character) 'Retorna uma string de n tamanho com o caracterer x
StrComp(string1, string2[, compare]) 'Verifica se a string x está contida na string y
                                      'o parâmetro opcional é 0 (texto exato) ou 1.
```

### Funções de Tratamento de tempo (Data e Hora)

```vb
Date()    'Retorna a data do sistema
Time()    'Retorna a hora do sistema
Day(date) 'Se o valor recebido por uma data, retorna o dia
Month(Now)  'Se o valor recebido por uma data, retorna o o número do mês
Now()    'Retorna a data e hora do sistema
MonthName(month[, abbreviate]) 'Retorna o nome do mês, padrão para abreviação é False. MonthName(10, True) = Oct
Hour()    'Se receber um valor time, retorna a hora
Year()    'Se receber um valor date, retorna o ano
WeekDay() 'Se receber um valor data, retorna o número do dia da semana
WeekdayName(weekday, abbreviate, firstdayofweek)  'Retorna o nome do dia da semana. Por padrão os param. *abbreviate* é False e *firstDayOfWeek* é 1 (Sunday). WeekDayName(6, True) = Fry
```

#### DateDiff()

>DateDiff(interval, date1, date2 [,firstdayofweek[, firstweekofyear]]) => recebe o tipo do intervalo, e calcula a difereça de valor entre duas datas

A função a seguir retorna quantos dias determinada data possui de diferença em relação a data atual

```vb
Function DiffADate(theDate)
   DiffADate = "Days from today: " & DateDiff("d", Now, theDate)
End Function
```

#### DateAdd()

> DateAdd(interval, number, date) => Adiciona ou remove um determinado intervalo de uma data

Abaixo, a nova variável de data recebe mais um mês. Para mais parâmetros de data consulte a sessão [Parâmetros para data](#parâmetros-para-data).

```vb
NewDate = DateAdd("m", 1, "31-Jan-95")
```

#### DatePart()

> DatePart(interval, date[, firstdayofweek[, firstweekofyear]]) => retorna um intervalo de tempo em uma medida específica.

É uma função única que agrega as funções

 - `Year `
 - `Month `
 - `Day `
 - `Hour`
 - `Minute`
 - `Second`

Basta inserir o parâmetro desejado que deseja que a função DatePart() retornará o trecho desejado da data (consulte a sessão [Parâmetros para data](#parâmetros-para-data).)

[Exemplo da Função DatePart()](functions/DatePart.asp)

#### Parâmetros para data

<details>
  <summary>
    <strong>Clique para expandir</strong>
  </summary>

##### *interval*

| **Setting** | **Description**            |
|-------------|----------------------------|
| **yyyy**    | Year                       |
| **q**       | Quarter                    |
| **m**       | Month                      |
| **y**       | Day of year (same as Day)  |
| **d**       | Day                        |
| **w**       | Weekday                    |
| **ww**      | Week of year               |
| **h**       | Hour                       |
| **n**       | Minute                     |
| **s**       | Second                     |

##### *firstdayofweek*

| **Constant**             | **Value** | **Description**                                   |
|--------------------------|-----------|---------------------------------------------------|
| **vbUseSystemDayOfWeek** | 0         | Use National Language Support (NLS) API setting.  |
| **vbSunday**             | 1         | Sunday (default)                                  |
| **vbMonday**             | 2         | Monday                                            |
| **vbTuesday**            | 3         | Tuesday                                           |
| **vbWednesday**          | 4         | Wednesday                                         |
| **vbThursday**           | 5         | Thursday                                          |
| **vbFriday**             | 6         | Friday                                            |
| **vbSaturday**           | 7         | Saturday                                          |

#### *firstweekofyear*

| **Constant**        | **Value** | **Description**                                                   |
|---------------------|-----------|-------------------------------------------------------------------|
| **vbUseSystem**     | 0         | Use National Language Support (NLS) API setting.                  |
| **vbFirstJan1**     | 1         | Start with the week in which January 1 occurs (default).          |
| **vbFirstFourDays** | 2         | Start with the week that has at least four days in the new year.  |
| **vbFirstFullWeek** | 3         | Start with the first full week of the new year.                   |

</details>



[📖 voltar para o índice 📖](#índice-do-conteúdo)

---

### Funções de Cálculo


[📖 voltar para o índice 📖](#índice-do-conteúdo)

---

## Eventos no VBScript
> evento é qualquer ação que o usuário realize em uma página.

Dentro do ASP, temos Quatro tipos de evento:

- Window
- Document
- Form
- Element

Para criar procedimentos (funções ou sub-rotinas) que responda a eventos, usamos a sintaxe:

```vb

<SCRIPT ID=clientEventHandlerVBS LANGUAGE= vbscript>
<!--
  Sub NomeDoObjeto_NomeDoEvento()
    ...
    ...
    ...
  End Sub
-->
</SCRIPT>
```

### Classes








[📖 voltar para o índice 📖](#índice-do-conteúdo)

---

## Para saber mais

- [ASP Tutorial - W3C Schools](https://www.w3schools.com/asp/asp_introduction.asp)
- [Using Object-Oriented Programming with VBScript](https://www.oreilly.com/library/view/designing-active-server/0596000448/ch04s02.html)
- [VbScript/ASP Classic good OOP Pattern](https://stackoverflow.com/questions/12246278/vbscript-asp-classic-good-oop-pattern)
- [Object Oriented ASP: Using Classes in Classic ASP](https://www.codeguru.com/dotnet/object-oriented-asp-using-classes-in-classic-asp/)
