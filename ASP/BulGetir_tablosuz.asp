<%
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("Veritabanim.mdb"))
ssql="select * from Benimtablom1 ORDER BY IsimSoyad;"
Set oRS = oConn.Execute(sSQL)

Do While NOT oRS.EOF 

if oRS("IsimSoyad") = Request.form("AdSoyad") then
%>
      <%=oRS("IsimSoyad")%> <br>
  <%=oRS("Yasi")%><br>

  <%=oRS("KayitTarihi")%><br>
   


<%	
end if
	oRS.MoveNext
Loop

oConn.Close
Set oRS = Nothing
Set oConn = Nothing
%>
<p><a href="menu.htm">Menü</a></p>