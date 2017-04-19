<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<%
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("Veritabanim.mdb"))
ssql="select * from Benimtablom1 ORDER BY IsimSoyad;"
Set oRS = oConn.Execute(sSQL)

Do While NOT oRS.EOF 

if oRS("IsimSoyad") = Request.form("AdSoyad") then
%>


<form action="VeriDuzeltmeOK.asp" method="post">


<b>Veri Düzeltme</b>

<br><br>

Adı Soyadı: <%=ors("IsimSoyad")%> <br>

<input type=hidden name="AdiSoyadi" value=<%=ors("IsimSoyad")%> >



Yaşı <input type="text" value="<%=ors("Yasi")%>" name="Yas" ><br>

Kayıt Tarihi <input type="text" value="<%=ors("KayitTarihi")%>" name="KayitTr"><br> 


<input type="submit" value="Düzelt" >
      
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