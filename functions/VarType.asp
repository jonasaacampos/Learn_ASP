<html>
<head>
<title>Usos da funcao VarType</title>
</head>

<%
    Function TipoDescritivoDaVariavel(variavel)
        Select Case VarType(variavel):
            Case 0
                Response.Write("Código: 0 - Não iniciada (vazia)        : Valor da Variavel =>     " & variavel)
            Case 5
                Response.Write("Código: 5 - Numero de precisao dupla    : Valor da Variavel =>     " & variavel)
            Case 3
                Response.Write("Código: 3 - Inteiro longo               : Valor da Variavel =>     " & variavel)
            Case 8
                Response.Write("Código: 8 - Texto                       : Valor da Variavel =>     " & variavel)
            Case 7
                Response.Write("Código: 7 - Data                        : Valor da Variavel =>     " & variavel)
            Case 8192
                Response.Write("Código: 8204 - Array                    : Valor da Variavel =>     " & variavel)

            
        End Select
    End Function
%>


<body>
<h1>Exemplo da funcoes VarType</h1>

<%
Dim x, y, z
Dim a, b
Dim c(20)
Dim j

x = 12
y = 23.456
z = 123456789

a = "isso e uma variavel de texto"
b = Date()
%>

<ul>
    <li>Tipo da variavel j: <%Response.Write(TipoDescritivoDaVariavel( j ))%></li>
    <li>Tipo da variavel a: <%Response.Write(TipoDescritivoDaVariavel( a ))%></li>
    <li>Tipo da variavel b: <%Response.Write(TipoDescritivoDaVariavel( b ))%></li>
    <li>Tipo da variavel c: <%Response.Write(TipoDescritivoDaVariavel( c ))%></li>
    <li>Tipo da variavel x: <%Response.Write(TipoDescritivoDaVariavel( x ))%></li>
    <li>Tipo da variavel y: <%Response.Write(TipoDescritivoDaVariavel( y ))%></li>
    <li>Tipo da variavel z: <%Response.Write(TipoDescritivoDaVariavel( z ))%></li>
</ul>

</body>
</html>