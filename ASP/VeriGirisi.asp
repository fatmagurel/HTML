
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">

<form action="VeriGirisiOK.asp" method="post">


<b>Veri Giri�i </b>

<br><br>

Ad� Soyad� <input type="text" name="AdiSoyadi"> <br>

Ya�� <input type="text" name="Yas" ><br>

Kay�t Tarihi <input type="text" name="KayitTr" value="<%=date()%>"><br> 

<%'Dikkat: �stteki date fonksiyonu olmasaydi, bu dosyayi asp olarak kaydetmeye gerek yoktu. htm uzantili olarak da kaydedilebilirdi.%>

<input type="submit" value="Kaydet" >
      
</form>