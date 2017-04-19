<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">

<%

'______________________  SİL:

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("veritabanim.mdb"))

		set kayit_sil = Server.CreateObject("ADODB.RecordSet")

		SQL = "delete * from BenimTablom1 Where IsimSoyad = '" & request.form("AdiSoyadi") & "'"

		kayit_sil.Open sql, oConn, 1, 2
		
		set kayit_sil=nothing
		oConn.CLOSE
		SET oConn = NOTHING

response.write "<center>Kayıt Silinmiştir"
%>

<br>
<br>
<a href="menu.htm">Menu</a>   