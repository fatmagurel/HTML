<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<html>

<head>
<body bgcolor="#F7F7F7">

<%
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("Veritabanim.mdb"))
ssql="select * from Benimtablom1 ORDER BY IsimSoyad;"
Set oRS = oConn.Execute(sSQL)

%>
<p align="center"><b><font face="Tahoma" size="5">Arama Sonucu</font></b></p>

<%
Do While NOT oRS.EOF 

if oRS("IsimSoyad") = Request.form("AdSoyad") then
%>

  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse; text-align: center" bordercolor="#111111" width="457">
    <tr>
      <td width="171" style="text-align: right; border-right-style: none; border-right-width: medium" bgcolor="#C0C0C0">
      <b><font face="Tahoma">Adı Soyadı</font></b></td>
      <td width="14" style="border-left-style: none; border-left-width: medium" bgcolor="#C0C0C0">&nbsp;</td>
      <td width="14" style="border-right-style: none; border-right-width: medium">&nbsp;</td>
      <td width="255" style="text-align: left; border-left-style: none; border-left-width: medium"><font face="Tahoma"><%=oRS("IsimSoyad")%>&nbsp;</td>
    </tr>
    <tr>
      <td width="171" style="text-align: right; border-right-style: none; border-right-width: medium" bgcolor="#C0C0C0">
      <b><font face="Tahoma">Yaşı</font></b></td>
      <td width="14" style="border-left-style: none; border-left-width: medium" bgcolor="#C0C0C0">&nbsp;</td>
      <td width="14" style="border-right-style: none; border-right-width: medium">&nbsp;</td>
      <td width="255" style="text-align: left; border-left-style: none; border-left-width: medium"><font face="Tahoma"><%=oRS("Yasi")%>&nbsp;</td>
    </tr>
    <tr>
      <td width="171" style="text-align: right; border-right-style: none; border-right-width: medium" bgcolor="#C0C0C0">
      <b><font face="Tahoma">Kayıt Tarihi</font></b></td>
      <td width="14" style="border-left-style: none; border-left-width: medium" bgcolor="#C0C0C0">&nbsp;</td>
      <td width="14" style="border-right-style: none; border-right-width: medium">&nbsp;</td>
      <td width="255" style="text-align: left; border-left-style: none; border-left-width: medium"><font face="Tahoma"><%=oRS("KayitTarihi")%>&nbsp;</td>
    </tr>
  </table>
  </center>

<p align="center"><br>
<b><font face="Tahoma" size="4">
 </tr>


<%	
end if

	oRS.MoveNext
Loop



oConn.Close
Set oRS = Nothing
Set oConn = Nothing

%>

</table>


</font>

</b>

</body>
<p align="center"><a href="menu.htm">Menü</a></p>
</html>