<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<% 
'kutuyu bo� b�rakmay� engelleme
If trim(request("AdiSoyadi"))="" then  
response.write ("<b>Eksik Bilgi! </b> L�tfen bo� b�rakmay�n�z.   [ <a href=""javascript:history.back()"">Geri</a> ]<br><br> ")
response.end  
end if

'--------------
'VT baglantisinin yapimasi:
Set Baglantim = CreateObject("ADODB.Connection") 
'VT'nin acilmasi:
Baglantim.Open ("DRIVER={Microsoft Access Driver (*.mdb)};DBQ="& Server.MapPath("Veritabanim.mdb"))
'Tablo nesnesinin olusturulmasi:
Set Tablom = server. CreateObject("ADODB.Recordset")
'Tablonun acilmasi:
Tablom.Open "BenimTablom1", Baglantim, 1, 3

'Tabloya veri eklemeye baslangic:
Tablom.AddNew 
'Tablodaki alanlara veri aktarma
Tablom("IsimSoyad") =  request("AdiSoyadi")
Tablom("Yasi") =  request("Yas")
Tablom("KayitTarihi") =  request("KayitTr")
'aktarma islemi birince tablonun guncellenmesi:
Tablom.Update

'tablonun kapatilmasi:
  Tablom.close
  set Tablom= Nothing
'baglantinin kesilmesi:
  Baglantim.close
  set Baglantim= Nothing

response.write "Veri Giri�i Yap�lm��t�r"
%>
<p><a href="menu.htm">Men�</a></p>