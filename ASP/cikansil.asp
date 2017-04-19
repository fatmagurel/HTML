<%
if Session("uname")="sturen" or Session("uname")="ikaras" then


Response.Expires = -1000

Dim oConn
Dim oRS
Dim rs
Dim sSQL
Dim sql
Dim nosu
Dim yariyil
Dim ogretimyili


dim kod(10)
dim sy
dim StrDizi
dim dkod
dim Teori
dim Uygulama
dim Toplam

dkod = Request.form("cekbaks")

'response.write dkod

StrDizi=split(dkod,",",-1,1)

'response.write StrDizi(0)
'response.write "-"
'response.write StrDizi(1)

'StrDizi(0)=ccur(StrDizi(0))
'StrDizi(1)=ccur(StrDizi(1))


dizi= UBound(StrDizi)



'response.write dizi
i=0

Do While not i=dizi+1

'______________________ SÝLÝNENLER TABLOSUNA AKTAR:

'TIRNaKLANMIÞ OLAN HER BÝR KAYDI AL:

set conn=server.createobject("adodb.connection")
path = "DRIVER={MICROSOFT ACCESS DRIVER (*.mdb)}; "
path = path & "DBQ=" & Server.MapPath("STOK.mdb")
conn.open path
ssql = "select * from MalzemeCikisi where Kimlik =" & strDizi(i) & ""
Set rs=conn.execute(ssql)

Do While Not rs.Eof

'________________ ALýnan KAYDI SÝLÝNENLER TABLOSUNA AKTAR:

Set alan = CreateObject("ADODB.Connection") 
alan.Open ("DRIVER={Microsoft Access Driver (*.mdb)};DBQ="& Server.MapPath("STOK.mdb"))
Set Grup = server. CreateObject("ADODB.Recordset")
Grup.Open "MalzemeCikisindanSilinenler", alan, 1 , 3

Grup.AddNew 
Grup("Cinsi") = rs("Cinsi")
Grup("Serino_Ozs") = rs("Serino_Ozs")
Grup("KimeVerildi") = rs("KimeVerildi")
Grup("zimmetlimi") = rs("zimmetlimi")
Grup("birimi") = rs("birimi")
Grup("Adedi") = rs("Adedi")
Grup("Tarihi") = rs("Tarihi")
Grup("AdetKiloVb") = rs("AdetKiloVb")
Grup("SilinmeTarihi") = date()
Grup.Update

  Grup.close
  set Grup = Nothing
  alan.close
  set alan = Nothing

'________________


rs.MoveNext
loop

'______________________ ESAS TABLODAN SÝL:

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("STOK.mdb"))

		set mesaj_sil = Server.CreateObject("ADODB.RecordSet")
		SQL = "delete * from MalzemeCikisi where Kimlik =" & strDizi(i) & ""
		mesaj_sil.Open sql, oConn, 1, 3
		
		set mesaj_sil=nothing
		oConn.CLOSE
		SET oConn = NOTHING

i=i+1

Loop
%>
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
<p align="center">
<p align="center">&nbsp;</p>
<p align="center"><b>
<%
response.write "Kayýt Silinmiþtir"

end if


%>



<p align="center">&nbsp;</p>
<p align="center"><b><font face="Arial">
<a href="cikansec.asp">Listeye Dön</a>   &nbsp; &nbsp; &nbsp;       <a href="menu.asp">Menüye Dön</a></font></b></p>