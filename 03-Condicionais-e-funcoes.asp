<!DOCTYPE html>
<head>
    <title>Document</title>
</head>
<body>

<!-- Condicional IF -->
<h3>Funções e Condicional com If Then, ElseIf Then e Else</h3>

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

<%
Response.Write(horaAtual(10))
Response.Write(horaAtual(12))
Response.Write(horaAtual(18))
%>

<br/>
<hr>

<h3>Função e Condicional Select Case</h3>

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

<%=diaDaSemana(1)%>
<%=diaDaSemana(2)%>
<%=diaDaSemana(3)%>
<%=diaDaSemana(4)%>
<%=diaDaSemana(5)%>
<%=diaDaSemana(6)%>
<%=diaDaSemana(7)%>
<%=diaDaSemana(0)%>

</body>
</html>
