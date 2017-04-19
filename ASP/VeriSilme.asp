<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">

<%
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("Veritabanim.mdb"))
ssql="select * from Benimtablom1 ORDER BY IsimSoyad;"
Set oRS = oConn.Execute(sSQL)

Do While NOT oRS.EOF 

if oRS("IsimSoyad") = Request.form("AdSoyad") then
%>


<form action="VeriSilmeOK.asp" method="post">

<center>
<b>Veri Silme</b>

<br>
Bu kayıt silinsin mi?
<br>

<input type="hidden" value="<%=ors("IsimSoyad")%>" name="AdiSoyadi"> <br>

      <%=oRS("IsimSoyad")%> <br>
  <%=oRS("Yasi")%><br>

  <%=oRS("KayitTarihi")%><br><br>



<input type="submit" value="Sil" >
      
</form>



<%	
end if
	oRS.MoveNext
Loop

oConn.Close
Set oRS = Nothing
Set oConn = Nothing
%>
<p><a href="menu.htm">Menü</a></p>