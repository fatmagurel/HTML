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


<b>Veri D�zeltme</b>

<br><br>

Ad� Soyad�: <%=ors("IsimSoyad")%> <br>

<input type=hidden name="AdiSoyadi" value=<%=ors("IsimSoyad")%> >



Ya�� <input type="text" value="<%=ors("Yasi")%>" name="Yas" ><br>

Kay�t Tarihi <input type="text" value="<%=ors("KayitTarihi")%>" name="KayitTr"><br> 


<input type="submit" value="D�zelt" >
      
</form>



<%	
end if
	oRS.MoveNext
Loop

oConn.Close
Set oRS = Nothing
Set oConn = Nothing
%>
<p><a href="menu.htm">Men�</a></p>