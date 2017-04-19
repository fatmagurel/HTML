<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<%
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("Veritabanim.mdb"))
ssql="select * from BenimTablom1; "
Set oRS = oConn.Execute(sSQL)
%>

<table border="1" width="85%" cellspacing="0" cellpadding="0" bordercolor="#000000" style="border-collapse: collapse; text-align:center">
  <tr>
    <td width="36%" style="border-style:solid; border-width:1; " bgcolor="#C0C0C0" >
    <b><font face="Verdana">Adı Soyadı</font></b></td>
    <td width="36%" style="border-style:solid; border-width:1; " bgcolor="#C0C0C0" >
    <b><font face="Verdana">Yaşı</font></b></td>
    <td width="36%" style="border-style:solid; border-width:1; " bgcolor="#C0C0C0" >
    <b><font face="Verdana">Kayıt Tarihi</font></b></td>
  </tr>
<%
Do While NOT oRS.EOF 
%>
  <tr>
    <td width="36%" style="border-style:solid; border-width:1; " ><font face="Tahoma" size="2" ><%=oRS("IsimSoyad")%></font>&nbsp;</td>
    <td width="36%" style="border-style:solid; border-width:1; " ><font face="Tahoma" size="2" ><%=oRS("Yasi")%></font>&nbsp;</td>
    <td width="36%" style="border-style:solid; border-width:1; " ><font face="Tahoma" size="2" ><%=oRS("KayitTarihi")%></font>&nbsp;</td>
  </tr>
<%	
	oRS.MoveNext
Loop
%>
</table>
          </center>
<%
oConn.Close
Set oRS = Nothing
Set oConn = Nothing
%>