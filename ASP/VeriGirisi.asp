
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">

<form action="VeriGirisiOK.asp" method="post">


<b>Veri Girişi </b>

<br><br>

Adı Soyadı <input type="text" name="AdiSoyadi"> <br>

Yaşı <input type="text" name="Yas" ><br>

Kayıt Tarihi <input type="text" name="KayitTr" value="<%=date()%>"><br> 

<%'Dikkat: üstteki date fonksiyonu olmasaydi, bu dosyayi asp olarak kaydetmeye gerek yoktu. htm uzantili olarak da kaydedilebilirdi.%>

<input type="submit" value="Kaydet" >
      
</form>