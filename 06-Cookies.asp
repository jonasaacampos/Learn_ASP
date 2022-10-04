<%
''———————————————————————————————————————————————————————————————————
''       	   Projeto: Contador de Visitas
''———————————————————————————————————————————————————————————————————
''	      Created At: 2022/10/04
''	       Create by: jaac
''———————————————————————————————————————————————————————————————————
''	 Funcionalidade: counter visit user based in brower session
''———————————————————————————————————————————————————————————————————
%>

<%
Function ContadorDeVisitas()
  dim numeroDeVisitasNaPagina
  response.cookies("numeroDeVisitasNaPagina").Expires=date+365
  numeroDeVisitasNaPagina=request.cookies("numeroDeVisitasNaPagina")

  if numeroDeVisitasNaPagina="" then
    response.cookies("numeroDeVisitasNaPagina")=1
    response.write("Bem Vindo! &eacute; a primeira vez que nos vemos nesta p&aacute;gina. xD")
  else
    response.cookies("numeroDeVisitasNaPagina")=numeroDeVisitasNaPagina+1
    response.write("Esta &eacute sua " & numeroDeVisitasNaPagina & "&ordf; visita nesta página.")
  end if

End Function
%>

<!DOCTYPE html>

<html>
</body>

<p>
  <%ContadorDeVisitas()%>
</p>

</body>
</html>