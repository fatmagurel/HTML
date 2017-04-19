<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<%
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("Veritabanim.mdb"))
ssql="select * from BenimTablom1; "
Set oRS = oConn.Execute(sSQL)

Do While NOT oRS.EOF 
%>
<%=oRS("IsimSoyad")%>    <%=oRS("Yasi")%>    <%=oRS("KayitTarihi")%>
<br>
<%	
	oRS.MoveNext
Loop

oConn.Close
Set oRS = Nothing
Set oConn = Nothing
%>