<!DOCTYPE html>
<html>
<body>

  <h1>Laços de repetição</h1>

  <h2>For</h2>
  <hr>

  <%
  For i = 1 To 6
    response.write("<h" & i & ">Este é o Estilo " & i & "</h" & i & ">")
  Next
  %>

  <hr>

  <h2>Do While</h2>
  <%
  i=0
  Do While i <  10
    response.write(i & "<br>")
    i=i+1
  Loop
  %>

  <hr>

  <h2>Do Until</h2>

  <p>
    - sendo i igual a 20, Faça enquanto i não valer 10
  </p>
  <%
  i = 20
  Do Until i=10
    response.write(i & "<br>")
    i=i-1
    If i < 10 Then Exit Do
  Loop
  %>

  <hr>

  <p>
    - sendo i igaul a 10, Execute, mas pare quando i valer 10
  </p>
  <%
  i = 10
  Do
    response.write(i & "<br>")
    i=i-1
    If i < 10 Then Exit Do
  Loop Until i=10
  %>

</body>
</html>