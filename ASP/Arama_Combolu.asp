



<%

'Bu �rnekte Veritaban�ndaki kayitlar� okuyarak alfabetik s�rayla ComboBox'in doldurulmasi g�r�lmektedir:


dim deg(5000)

set conn=server.createobject("adodb.connection")
path = "DRIVER={MICROSOFT ACCESS DRIVER (*.mdb)}; "
path = path & "DBQ=" & Server.MapPath("Veritabanim.mdb")
conn.open path
sql = "SELECT * FROM BenimTablom1 ORDER BY IsimSoyad;"
Set rs=conn.execute(sql)
sayac=1
krt=0
%>

<form action="Bulgetir.asp" method="post">

<select name="AdSoyad">

<%Do While Not rs.Eof
krt=0
for d=1 to sayac 'ayni degeri kutuya sadece bir kez eklemesi icin 
	if rs("IsimSoyad")=deg(d) then
		krt=1 
	end if
next
if krt<>1 then
	%><option value="<%=rs("IsimSoyad")%>"><%=rs("IsimSoyad")%></option><%
	deg(sayac)=rs("IsimSoyad")
	sayac=sayac+1
end if

rs.movenext
loop

conn.Close
Set rs = Nothing
Set conn = Nothing

%>
</select> 
<input type="submit" value="Getir"> </p>
</form>