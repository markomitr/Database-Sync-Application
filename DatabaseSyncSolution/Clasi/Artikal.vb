Option Strict Off
Imports System.Text
Public Class Artikal

    Enum ArtikalStatus
        Onstock = 1
        Notonstock = 2
        Soonavailable = 3
        Lowquantity = 4
        Canbeordered = 5
        Nochange = -1
    End Enum
    Enum PolinjaAtributiURL
        GUID
        Customerid
        InternalCode
        Quantity
        Price
        Status
    End Enum
    Public Shared Function UpdateArtikalEdenPoEden(ByRef poraka As String) As Boolean
        Try
            Dim ds As DataSet
            Dim nepusesniArtikli As New List(Of String)
            Dim brUspesni, brNeuspesni As Integer
            Dim sb As New StringBuilder
            brUspesni = 0
            brNeuspesni = 0
            If AlatkaBaza.ZemiArtikliZaPrakanje(ds, poraka) Then
                If ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                    Console.WriteLine(DateTime.Now.ToString & " Najdov promeneti artikli:" & ds.Tables(0).Rows.Count)
                    sb.AppendLine("")
                    sb.AppendLine("Najdov promeneti artikli:" & ds.Tables(0).Rows.Count)
                    sb.AppendLine("")
                    For Each red As DataRow In ds.Tables(0).Rows
                        Dim SifraArt As String = Konv.ObjVoStr(red("Sifra_Art"))
                        Dim InterenKod As String = Konv.ObjVoStr(red("Alt_Sifra"))
                        Dim Cena As Decimal = Konv.ObjVoDec(red("CenaGolemo"))
                        Dim Kolic As Decimal = Konv.ObjVoDec(red("Kolic"))
                        Dim Status As ArtikalStatus = VratiStatus(Konv.ObjVoStr(red("Sostojba")))
                        If UpdateArtikalEDEN(InterenKod, Kolic, Cena, Status, poraka) Then
                            'Treba da se zapisi deka e uspesno
                            If Not AlatkaBaza.ZapisiUspesnoIzmenetArtikal(SifraArt, InterenKod, Kolic, Cena, Status, True, poraka) Then
                                Console.WriteLine(DateTime.Now.ToString & " Artikal: " & SifraArt & " - NEUSPESNO: " & poraka)
                                sb.AppendLine("Artikal: " & SifraArt & " Kolic: " & Kolic & " Cena: " & Cena & " - NEUSPESNO(KajNASvoSQL): " & poraka)
                                'Return False
                            End If
                            Console.WriteLine(DateTime.Now.ToString & " Artikal: " & SifraArt & " - USPESNO: ")
                            sb.AppendLine("Artikal: " & SifraArt & " Kolic: " & Kolic & " Cena: " & Cena & " - USPESNO")
                            brUspesni += 1
                            AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.IZMENI, TekParam.TipRezultat.Uspesno, AlatkaFile.StoPraka.Artikli, "ARTIKAL - Uspesno izmenat artikal.", "Artikal: " & SifraArt & " Alt_Sif:" & InterenKod & " Kolic: " & Kolic & " Cena: " & Cena)
                        Else
                            'Treba da se zapisi deka e neuspesno
                            nepusesniArtikli.Add("Sifra_Art")
                            Console.WriteLine(DateTime.Now.ToString & " Artikal: " & SifraArt & " - NEUSPESNO, PRICINA: " & poraka)
                            sb.AppendLine("Artikal: " & SifraArt & " - NEUSPESNO, PRICINA: " & poraka)
                            AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.IZMENI, TekParam.TipRezultat.NEUspesno, AlatkaFile.StoPraka.Artikli, "ARTIKAL - NE-Uspesno izmenat artikal!" & poraka, "Sif_Art: " & SifraArt & " Alt_Sif:" & InterenKod)
                            If Not AlatkaBaza.ZapisiUspesnoIzmenetArtikal(SifraArt, InterenKod, Kolic, Cena, Status, False, poraka) Then
                                ''''Return False
                            End If
                            brNeuspesni += 1

                        End If
                    Next
                    Console.WriteLine(DateTime.Now.ToString & " Broj na Uspesni zapisi:{0}, Neuspesni:{1}", brUspesni, brNeuspesni)
                    sb.AppendLine("")
                    sb.AppendLine("Broj na Uspesni zapisi:{" & brUspesni & "}, Neuspesni:{" & brNeuspesni & "}")
                    Console.WriteLine()
                    AlatkaFile.ZapisiVoFile(sb.ToString, AlatkaFile.KojPraka.WebShop, AlatkaFile.StoPraka.Artikli, poraka, True)
                    AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.ZACUVAJ, TekParam.TipRezultat.Uspesno, AlatkaFile.StoPraka.Artikli, "ARTIKAL-" & sb.ToString, "UPDATE ARTIKLI")
                    Return True
                Else
                    poraka = "Nema artikli za prakanje vo tabelata!"
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            poraka = "Greska UpdateArtikalEdenPoEden(): " & ex.Message
            AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.CatchEXCEPTION, TekParam.TipRezultat.NEUspesno, AlatkaFile.StoPraka.Artikli, "UpdateArtikalEdenPoEden() " & ex.Message, "UPDATE ARTIKLI")
            Return False
        End Try
    End Function
    Private Shared Function UpdateArtikalEDEN(ByVal interenKod As String, _
                                       ByVal kolicina As Decimal, _
                                       ByVal cena As Decimal, _
                                       ByVal status As Integer, _
                                       ByRef poraka As String) As Boolean
        'Ako vo polinjata Cena,Zaliha i State(Status) se prati NEGATIVNA vrednost taa vrednost nema da se izmeni
        Dim atributi As New ListaAtributi
        atributi.Add(PolinjaAtributiURL.GUID.ToString, TekParam.GuID)
        atributi.Add(PolinjaAtributiURL.InternalCode.ToString, interenKod)
        atributi.Add(PolinjaAtributiURL.Price.ToString, cena.ToString("0.00").Replace(",", ".")) 'Pazi na Regional Settings
        atributi.Add(PolinjaAtributiURL.Quantity.ToString, kolicina.ToString("0.00"))
        atributi.Add(PolinjaAtributiURL.Status.ToString, status.ToString("0"))

        Dim urlPrati As String = AlatkaURL.PripremiURLzaPrakanje("Updatearticleprice.aspx", "", atributi)
        Return AlatkaURL.PratiURL(urlPrati, poraka)
    End Function
    Public Shared Function UpdateArtikliMNOGU(ByRef poraka As String) As Boolean
        Dim ftpObj As AlatkaFTP
        If UpdateArtikliMNOGURabotiFtp(ftpObj, poraka) Then
            Return UpdateArtikliMNOGUProcesirajFTP(ftpObj, poraka)
        Else
            Return False
        End If
    End Function
    Public Shared Function UpdateArtikliCeniKolicMNOGU(ByRef poraka As String) As Boolean
        Dim ftpObj As AlatkaFTP
        If UpdateArtikliCeniKolicMNOGURabotiFtp(ftpObj, poraka) Then
            Return UpdateArtikliCeniKolicMNOGUProcesirajFTP(ftpObj, poraka)
        Else
            Return False
        End If
    End Function
    Private Shared Function UpdateArtikliMNOGURabotiFtp(ByRef ftpObj As AlatkaFTP, ByRef poraka As String) As Boolean
        Dim atributi As New ListaAtributi
        atributi.Add(PolinjaAtributiURL.GUID.ToString, TekParam.GuID)
        Dim urlPrati As String = AlatkaURL.PripremiURLzaPrakanje("Getuploadlocation.aspx", "Article", atributi)
        Dim rezultat As String = ""
        If AlatkaURL.ZemiFTPurl(urlPrati, rezultat) Then 'Ako uspesno e zemen FTP adresata za prakanje
            Dim ftpAdresa As String = rezultat
            ftpObj = New AlatkaFTP(ftpAdresa)
            If ftpObj.DaliEVoRed Then ' Vo red e adresata
                If ftpObj.PratiFileFtp("testfileArtNov.xls", "", poraka) Then
                    Return True
                Else
                    Return False
                End If
            Else
                poraka = "Ne e vo red FTP adresata! " + rezultat
                Return False
            End If

        Else
            poraka = rezultat
            Return False
        End If 'Kraj od ZemiFTPurl
        Return False
    End Function
    Private Shared Function UpdateArtikliMNOGUProcesirajFTP(ByVal ftpObj As AlatkaFTP, ByRef poraka As String) As Boolean
        If ftpObj Is Nothing Then
            poraka = "Ne inic FTP objekt!"
            Return False
        End If
        Dim atributi As New ListaAtributi
        atributi.Add(PolinjaAtributiURL.GUID.ToString, TekParam.GuID)
        Dim dodadtenUrl As String = ftpObj.SesijaZaFTPadresa & "/Article"
        Dim urlPrati As String = AlatkaURL.PripremiURLzaPrakanje("Processitems.aspx", dodadtenUrl, atributi)
        Return AlatkaURL.PratiURL(urlPrati, poraka)
    End Function
    Private Shared Function UpdateArtikliCeniKolicMNOGURabotiFtp(ByRef ftpObj As AlatkaFTP, ByRef poraka As String) As Boolean
        Dim atributi As New ListaAtributi
        atributi.Add(PolinjaAtributiURL.Customerid.ToString, TekParam.CustomerID)
        Dim urlPrati As String = AlatkaURL.PripremiURLzaPrakanje("Getuploadlocation.aspx", "BulkPrice", atributi)
        Dim rezultat As String = ""
        If AlatkaURL.ZemiFTPurl(urlPrati, rezultat) Then 'Ako uspesno e zemen FTP adresata za prakanje
            Dim ftpAdresa As String = rezultat
            ftpObj = New AlatkaFTP(ftpAdresa)
            If ftpObj.DaliEVoRed Then ' Vo red e adresata
                If ftpObj.PratiFileFtp("testfileArtNov.xls", "", poraka) Then
                    Return True
                Else
                    Return False
                End If
            Else
                poraka = "Ne e vo red FTP adresata! " + rezultat
                Return False
            End If

        Else
            poraka = rezultat
            Return False
        End If 'Kraj od ZemiFTPurl
        Return False
    End Function
    Private Shared Function UpdateArtikliCeniKolicMNOGUProcesirajFTP(ByVal ftpObj As AlatkaFTP, ByRef poraka As String) As Boolean
        If ftpObj Is Nothing Then
            poraka = "Ne inic FTP objekt!"
            Return False
        End If
        Dim atributi As New ListaAtributi
        atributi.Add(PolinjaAtributiURL.Customerid.ToString, TekParam.CustomerID)
        Dim dodadtenUrl As String = ftpObj.SesijaZaFTPadresa & "/BulkPrice"
        Dim urlPrati As String = AlatkaURL.PripremiURLzaPrakanje("Processitems.aspx", dodadtenUrl, atributi)
        Return AlatkaURL.PratiURL(urlPrati, poraka)
    End Function
    Private Shared Function VratiStatus(ByVal txt As String) As ArtikalStatus
        If Konv.TxtVoInt(txt) = 1 Then
            Return ArtikalStatus.Onstock
        Else
            Return ArtikalStatus.Notonstock
        End If
    End Function
End Class
