<!DOCTYPE html>
<head>
    <title>Document</title>
</head>
<body>

<form method="get" action="simpleform.asp">

    First Name: <input type="text" name="fname"><br>
    Last Name: <input type="text" name="lname"><br>
    <input type="submit" value="Submit">

</form>
<p>Seja bem vindo </p>


Welcome
<%
response.write(request.form("fname"))
response.write(" " & request.form("lname"))
%>


<%
response.write(request.querystring("fname"))
response.write(" " & request.querystring("lname"))
%>

</body>
</html>