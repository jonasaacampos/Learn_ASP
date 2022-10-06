<%
YearDate    = DatePart("yyyy"   , Now)
Quarter     = DatePart("q"      , Now)
MonthDate   = DatePart("m"      , Now)
DayOfYear   = DatePart("y"      , Now)
DayDate     = DatePart("d"      , Now)
WeekDayDate = DatePart("w"      , Now)
WeekOfYear  = DatePart("ww"     , Now)
HourDate    = DatePart("h"      , Now)
MinuteDate  = DatePart("n"      , Now)
SecondDate  = DatePart("s"      , Now)
%>

<h1>DatePart() function</h1>

<h2>
    Hora de execução da página: <% Response.Write Now %> 
</h2>

<ul>
    <li>          Ano: <%Response.Write YearDate%></li>
    <li>    Trimestre: <%Response.Write Quarter%></li>
    <li>          Mês: <%Response.Write MonthDate%></li>
    <li>   Dia do ano: <%Response.Write DayOfYear%></li>
    <li>   Dia do mês: <%Response.Write DayDate%></li>
    <li>Dia da semana: <%Response.Write WeekDayDate%></li>
    <li>         Hora: <%Response.Write HourDate%></li>
    <li>       Minuto: <%Response.Write MinuteDate%></li>
    <li>      Segundo: <%Response.Write SecondDate%></li>
    
</ul>

