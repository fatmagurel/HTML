<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<%

dsn = "DBQ=" & Server.Mappath("veritabanim.mdb") & ";Driver={Microsoft Access Driver (*.mdb)};" 
Set conn=Server.CreateObject("ADODB.Connection")
conn.Open dsn

SQL = "Update BenimTablom1 Set Yasi = '"& Request.Form("Yas") &"' Where IsimSoyad = '" & trim(request.form("AdiSoyadi")) & "'"

pSQL = "Update BenimTablom1 Set KayitTarihi = '"& Request.Form("KayitTr") &"' Where IsimSoyad = '" & trim(request.form("AdiSoyadi")) & "'"


  Set RS = conn.Execute(SQL)
  Set RS = conn.Execute(pSQL)
  conn.Close
  Set conn = Nothing

%>

<p align="center"><b>
<%
response.write "Kayıt Düzeltilmiştir"


%>
<br>
<br>
<a href="menu.htm">Menu</a></font></p>