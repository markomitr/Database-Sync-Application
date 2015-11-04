Imports System.Xml
Imports System.Text.RegularExpressions
Public Class Faktura
    Enum DokumentZaglavje
        Guid
        Version
        Proformainvoiceerpid
        Numrows
        Customererpid
        Comment
        Orderedby
        Deliverytype
        Deliveryplace
        Deadline
        Orderid
        Completed
        Option1
        Option2
        Option3
        Option4
    End Enum
    Enum StavkiDokument
        SessionId
        Internalcode
        Dquant
        Price
        Vat
        Disc
        currencyid
        marketprice
    End Enum
    Public Structure GlavaDokument
        Dim WebShopDokID As String
        Dim SifraKup As String
        Dim WebShopNarID As String
        Dim CFMADokID As String
        Dim MetodIsoraka As String
        Dim MestoIsporaka As String
        Dim KraenRokIsporaka As String
        Dim Datum As Date
        Dim NaracanoOd As String
        Dim Komenatar As String
        Dim Status As String
        Dim NaracanoOdKorisnik As String
        Dim ProverenoOdKorisnik As String
        Dim ImePrezimeKupuvac As String
        Dim ImePrezimeKomercijalist As String
        Dim Opcija1 As String
        Dim Opcija2 As String
        Dim Opcija3 As String
        Dim Opcija4 As String
        Dim BrojArtikli As Integer
        Dim KompletiranaFaktura As Boolean
    End Structure
    Public Structure StavkiFaktura
        Dim SifraArt As String
        Dim Cena As Decimal
        Dim Kolic As Decimal
        Dim Popust As Decimal
        Dim WebShopDokID As String
        Dim Danok As Decimal
    End Structure
    Public Structure Stavr
        Public STRBr As Integer
        Public STVlIzl As String
        Public STSifra_Art As String
        Public STKolic As Decimal
        Public STDokCena As Decimal
        Public STMagCena As Decimal
        Public STNabCena As Decimal
        Public STPOsn As Decimal
        Public STPTar As String
        Public STUces As Decimal
        Public STDanDokCena As String
        Public STDanMagCena As String
        Public STCenaIznos As String
        Public STKod_Danok As String
        Public STTros As String
        Public STImeMat As String
        Public STEdMera As String
        Public STPOsnPren As Decimal
        Public STDobav As String
        Public STKontrSer As String
        Public STKaloProc As Decimal
        Public STUcesOsn As Decimal
        Public STUcesKol As Decimal
        Public STUcesDod As Decimal
        Public STAlt_Kolic As Decimal
        Public STProc_Rand As Decimal
        Public STTrsCena As Decimal
        Public STNarVrska As String
        Public STNarSifra_Nar As Integer
        Public STNabCenaZadad As String
        Public STBrAmbal As Integer
        Public STCrr As String
        Public STTokenKontr As String
        Public STSifra_KatStatus As String
        Public STRok As String
        Public STOE_WHM As String
        Public STSifBoja As String
        Public STimeBoja As String
        Public STSifVelic As String
        Public StImeVelic As String
        Public STSifra_Art_Paket As String
        Public STKolku_Akcii As Integer
        Public STLokacija As String
        Public STImeArt As String
        Public STImeArt2 As String
        Public STSEdmera As String
        Public STSMatUsl As String
        Public STTezina As Decimal
        Public STAlt_Sifra As String
    End Structure

    Dim stavkiFak() As StavkiFaktura
    Dim glavaDok As GlavaDokument
    Dim BrojDokCFMA As String = ""
    Public Function ObrabotiFaktura(ByVal podatokData() As Byte, ByRef poraka As String) As Boolean
        Dim pateka As String = ""
        Dim wUspeh As Boolean = False
        If Not AlatkaFile.ZacuvajFile(podatokData, AlatkaFile.TipFajlovi.XML, AlatkaFile.KojPraka.WebShop, AlatkaFile.StoPraka.Dokumenti, pateka, poraka) Then
            wUspeh = False
            Return wUspeh
        End If
        If ObrabotiXMLFaktura(pateka, poraka) Then 'Ovaa funkcija go Cita XMl dok(Fakutrata od WebShop) i gi kreira GlavaDok i  StavkiDok strukturi
            AlatkaFile.RenameNaFile(pateka, Me.glavaDok.WebShopDokID & "_" & AlatkaFile.ZemiImeFileOdPateka(pateka))
            Dim wBrojDok As String = ""
            Dim wIdentif_Br As String = Me.glavaDok.WebShopNarID ' = Me.glavaDok.WebShopDokID
            Dim wIspratnica As String = Me.glavaDok.WebShopDokID
            Dim wSifOe As String = "1" ' FIKSIRANO TREBA DA SE SREDI ' MARKO M
            Dim wSifDok As String = "1" ' FIKSIRANO TREBA DA SE SREDI ' MARKO M
            If AlatkaBaza.ProveriDaliPostoiDok(wIdentif_Br, Me.glavaDok.SifraKup, wBrojDok, poraka) Then 'Proverkata vraka TRUE - ne e vo funkcija(namerno)!
                wUspeh = True
                If wBrojDok <> "" Then
                    poraka = "Dokumentot " & wSifOe & "-" & wSifDok & "/" & wBrojDok & " postoi"
                    wUspeh = False
                Else
                    Dim wStavkiDok() As Stavr
                    Dim wSifra_Kup As String = glavaDok.SifraKup
                    Dim wDatum_Dok As Date = Date.Today
                    Dim teksKomentar As String = glavaDok.Komenatar & " Datum Veb[op: " & glavaDok.Datum.ToString() & " Operator:" & glavaDok.ImePrezimeKupuvac

                    Dim wBroj_Dok As String
                    Dim wRok_Dosp, wSifra_Nal As Integer
                    Dim Toc200, Marza, Danok1, Danok2, Magvr, MagvrDan, Pdanok1, Pdanok2, wUces As Decimal
                    Dim DanDokCena, DanMagCena, Kod_Danok, wKto As String
                    If wUspeh AndAlso AlatkaBaza.NajdiPodatociZaKomint(wSifra_Kup, wUces, wRok_Dosp, poraka) Then
                        wUspeh = True
                    Else
                        wUspeh = False
                    End If
                    If wUspeh AndAlso AlatkaBaza.NajdiPodatociZaTipDok(wSifDok, wSifra_Nal, wKto, poraka) Then
                        wUspeh = True
                    Else
                        wUspeh = False
                    End If
                    If wUspeh AndAlso AlatkaBaza.NajdiSledenBrojDok(wSifOe, wSifDok, wBroj_Dok, poraka) Then ' SledenBrojDOk
                        wUspeh = True
                    Else
                        wUspeh = False
                    End If
                    BrojDokCFMA = wSifOe & "-" & wSifDok & "/" & wBroj_Dok ' Opis na dokument od tip 1-1/1069
                    If wUspeh AndAlso DajStavkiDok(wStavkiDok, Toc200, Marza, wUces, Danok1, Danok2, Magvr, MagvrDan, Pdanok1, Pdanok2, DanDokCena, DanMagCena, Kod_Danok, poraka) Then 'Stavki
                        wUspeh = True
                    Else
                        wUspeh = False
                    End If
                    If wUspeh AndAlso AlatkaBaza.ZacuvajDokument(wSifOe, wSifDok, wBroj_Dok, wSifra_Kup, wIdentif_Br, wIspratnica, wDatum_Dok, wStavkiDok, _
                                           wKto, wSifra_Nal, wUces, wRok_Dosp, Toc200, Marza, Danok1, Danok2, Magvr, MagvrDan, Pdanok1, Pdanok2, DanDokCena, DanMagCena, Kod_Danok, teksKomentar, poraka) Then   'ZacuvajDokument
                        wUspeh = True
                    Else
                        wUspeh = False
                    End If
                End If
            Else
                wUspeh = False
            End If
        Else
            wUspeh = False
        End If
        If wUspeh Then
            AlatkaFile.RenameNaFile(pateka, "S_" & Me.glavaDok.WebShopDokID & AlatkaFile.ZemiImeFileOdPateka(pateka))
        Else
            AlatkaFile.DodadiPorakaVoXml(poraka, pateka)
            AlatkaFile.RenameNaFile(pateka, "F_" & AlatkaFile.ZemiImeFileOdPateka(pateka))
        End If
        Return wUspeh
    End Function
    Private Function DajStavkiDok(ByRef wStavkiDok() As Stavr, ByRef Toc200 As Decimal, ByRef Marza As Decimal, ByVal wUces As Decimal, ByRef Danok1 As Decimal, ByRef Danok2 As Decimal, ByRef Magvr As Decimal, ByRef MagvrDan As Decimal, _
        ByRef Pdanok1 As Decimal, ByRef Pdanok2 As Decimal, ByRef DanDokCena As String, ByRef DanMagCena As String, ByRef Kod_Danok As String, ByRef poraka As String) As Boolean
        Try
            Dim i, iW As Integer
            iW = -1
            For i = 0 To stavkiFak.Length - 1
                If stavkiFak(i).SifraArt <> "" Then
                    Dim wSifra_Art, wSifra_Tar, wEdMera As String
                    Dim wDogCena As Decimal
                    Dim wProcOsn As Decimal
                    iW += 1
                    ReDim Preserve wStavkiDok(iW)
                    With wStavkiDok(iW)
                        If Not AlatkaBaza.NajdiPodatociZaArtikl(stavkiFak(i).SifraArt, wSifra_Art, wSifra_Tar, wEdMera, poraka, wProcOsn, wDogCena) Then
                            Return False
                        End If
                        .STRBr = iW + 1
                        .STVlIzl = "I"
                        .STSifra_Art = wSifra_Art
                        .STKolic = stavkiFak(i).Kolic
                        .STDokCena = wDogCena
                        .STMagCena = wDogCena
                        .STDanDokCena = "D"
                        .STDanMagCena = "D"
                        .STPOsn = wProcOsn
                        .STPOsnPren = wProcOsn
                        .STUces = wUces
                        .STPTar = wSifra_Tar
                        .STCenaIznos = "N"
                        .STImeMat = ""
                        .STEdMera = ""
                        .STKontrSer = ""
                        .STUcesOsn = 0
                        .STUcesDod = 0
                        .STUcesKol = 0
                        .STAlt_Kolic = 0
                        .STCrr = ""
                        .STSifBoja = ""
                        .STSifVelic = ""
                        .STSifra_KatStatus = ""
                        .STDobav = ""
                        .STKaloProc = 0
                        .STTrsCena = 0
                        .STKolku_Akcii = 0
                        .STKod_Danok = ""
                        .STTros = ""
                    End With
                End If
            Next
            NajdiSmestiZbirovi(wStavkiDok, "D", "D", "", Toc200, Marza, Danok1, Danok2, Magvr, MagvrDan, Pdanok1, Pdanok2, DanDokCena, DanMagCena, Kod_Danok)
            Return True
        Catch ex As Exception
            poraka = "GRESKA vo DajStavkiDok(): " & ex.Message
            Return False
        End Try
    End Function
    Public Sub NajdiSmestiZbirovi(ByVal wStavkiDok() As Stavr, ByVal wDanDokCena As String, ByVal wDanMagCena As String, ByVal wKod_Danok As String, _
        ByRef Toc200 As Decimal, ByRef Marza As Decimal, ByRef Danok1 As Decimal, ByRef Danok2 As Decimal, ByRef Magvr As Decimal, ByRef MagvrDan As Decimal, _
        ByRef Pdanok1 As Decimal, ByRef Pdanok2 As Decimal, ByRef DanDokCena As String, ByRef DanMagCena As String, ByRef Kod_Danok As String)
        Dim Red As Integer
        Dim wNab, wMag, wMagDan, wDan1, wDan2, wpDan1, wpDan2 As Decimal
        Dim wSumNab As Decimal = 0
        Dim wSumMag As Decimal = 0
        Dim wSumMagDan As Decimal = 0
        Dim wSumDan1 As Decimal = 0
        Dim wSumDan2 As Decimal = 0
        Dim wSumpDan1 As Decimal = 0
        Dim wSumpDan2 As Decimal = 0
        Dim KolkuRedici As Integer '= MojList.BrojNaRedici
        If Not wStavkiDok Is Nothing Then
            For Red = 0 To UBound(wStavkiDok)
                IznosRed(wStavkiDok(Red), wNab, wMag, wMagDan, Red, wDan1, wDan2, wpDan1, wpDan2, wKod_Danok, 0)
                wSumNab += wNab
                wSumMag += wMag
                wSumMagDan += wMagDan
                wSumDan1 += wDan1
                wSumDan2 += wDan2
                wSumpDan1 += wpDan1
                wSumpDan2 += wpDan2
            Next
        End If
        'Dim ZaokrFra As Integer = TekParam.ZemiInt("zaokr_fra")
        'If ZaokrFra = 0 Then
        '    ZaokrFra = 2
        'End If
        Dim ZaokrFra As Integer = 0 'Fix po baranje od Diki

        wSumNab = Decimal.Round(wSumNab, ZaokrFra)
        wSumMag = Decimal.Round(wSumMag, 2)
        wSumMagDan = Decimal.Round(wSumMagDan, 2)
        wSumDan1 = Decimal.Round(wSumDan1, 2)
        wSumDan2 = Decimal.Round(wSumDan2, 2)
        wSumpDan1 = Decimal.Round(wSumpDan1, 2)
        wSumpDan2 = Decimal.Round(wSumpDan2, 2)
        Toc200 = wSumNab
        Marza = wSumMag - (wSumNab - wSumDan1 - wSumDan2)
        Danok1 = wSumDan1
        Danok2 = wSumDan2
        Magvr = wSumMag
        MagvrDan = wSumMagDan
        Pdanok1 = wSumpDan1
        Pdanok2 = wSumpDan2
        DanDokCena = wDanDokCena
        DanMagCena = wDanMagCena
        Kod_Danok = wKod_Danok
    End Sub
    Private Sub IznosRed(ByVal Stavka As Stavr, ByRef NabIznos As Decimal, ByRef MagIznos As Decimal, ByRef MagIznosDan As Decimal, ByVal KojRed As Integer, ByRef Danok1 As Decimal, ByRef Danok2 As Decimal, ByRef PDanok1 As Decimal, ByRef PDanok2 As Decimal, ByVal wKod_danok As String, ByVal ProcKasa As Decimal)
        Dim wPole As String
        Dim wKolic As Decimal
        Dim wCena As Decimal
        Dim wPCena As Decimal
        Dim wRabat As Decimal
        Dim wFaktorSoDDV As Decimal
        Dim wCFaktorNeto As Decimal
        Dim wMFaktorNeto As Decimal
        Dim wMFaktorBruto As Decimal
        Dim wPosn As Decimal
        Dim wPosnPren As Decimal
        With Stavka
            wRabat = .STUces
            wPosn = .STPOsn
            If wKod_danok = "C" Then
                wPosnPren = 0
            Else
                wPosnPren = .STPOsnPren
            End If
            If .STDanDokCena = "D" Then
                wFaktorSoDDV = 1
                wCFaktorNeto = 1 + wPosnPren / 100
            Else
                wFaktorSoDDV = 1 + wPosnPren / 100
                wCFaktorNeto = 1
            End If
            If .STDanMagCena = "D" Then
                wMFaktorNeto = 1 + wPosn / 100
                wMFaktorBruto = 1
            Else
                wMFaktorNeto = 1
                wMFaktorBruto = 1 + wPosn / 100
            End If
            wKolic = .STKolic
            wCena = .STDokCena
            wPCena = .STMagCena
        End With
        Danok1 = 0
        PDanok1 = 0
        Danok2 = wKolic * wCena * (1 - wRabat / 100) * (1 - ProcKasa / 100) / wCFaktorNeto * wPosnPren / 100
        PDanok2 = wKolic * wPCena / wMFaktorNeto * wPosn / 100
        NabIznos = Decimal.Round(wKolic * wCena * wFaktorSoDDV * (1 - wRabat / 100) * (1 - ProcKasa / 100), 2)
        MagIznos = Decimal.Round(wKolic * wPCena / wMFaktorNeto, 2)
        MagIznosDan = Decimal.Round(wKolic * wPCena * wMFaktorBruto, 2)
    End Sub
    Private Function ObrabotiXMLFaktura(ByVal pateka As String, ByRef poraka As String) As Boolean
        If ProcitajXML(pateka, poraka) Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Function ProcitajXML(ByVal pateka As String, ByRef poraka As String) As Boolean
        Dim brojacArtikli As Integer = 0
        Dim daliGlavaDok As Boolean = False
        Dim daliStavkiDok As Boolean = False
        Dim daliArtikal As Boolean = False
        Dim podatociGlavaDok As New Dictionary(Of String, String)
        Dim listaStavki As New List(Of Dictionary(Of String, String))
        Dim podatociStavka As New Dictionary(Of String, String)
        Dim tekovenElement As String = ""
        Dim reader As XmlTextReader = New XmlTextReader(pateka)
        Try
            Do While (reader.Read())
                Select Case reader.NodeType
                    Case XmlNodeType.Element  'Display beginning of element.
                        If reader.Name.Trim.ToUpper = "PROFORMA" Then
                            daliGlavaDok = True
                            daliStavkiDok = False
                        ElseIf reader.Name.Trim.ToUpper = "PROFORMACONTENT" Then
                            daliStavkiDok = True
                            daliGlavaDok = False
                            podatociStavka = New Dictionary(Of String, String)
                        End If
                        tekovenElement = reader.Name
                    Case XmlNodeType.Text 'Display the text in each element.
                        If daliStavkiDok Then
                            If Not podatociStavka Is Nothing Then
                                podatociStavka.Add(tekovenElement.ToUpper, reader.Value)
                            End If
                        End If
                        If daliGlavaDok Then
                            podatociGlavaDok.Add(tekovenElement.ToUpper, reader.Value)
                        End If
                    Case XmlNodeType.EndElement 'Display end of element.
                        If reader.Name.Trim.ToUpper = "PROFORMA" Then
                            daliGlavaDok = False
                        ElseIf reader.Name.Trim.ToUpper = "PROFORMACONTENT" Then
                            daliStavkiDok = False
                            listaStavki.Add(podatociStavka)
                            brojacArtikli = brojacArtikli + 1
                        End If
                End Select
            Loop
            reader.Close()
            If podatociGlavaDok.Count = 0 AndAlso brojacArtikli = 0 AndAlso listaStavki.Count = 0 Then
                poraka = "Prazen Faktura XML Dokument - nema Dok i Stavki!"
                Return False
            End If
            Return DodeliObjOdXml(podatociGlavaDok, brojacArtikli, listaStavki, poraka)
        Catch ex As Exception
            reader.Close()
            poraka = "GRESKA Parse-XML Invoice! " + ex.Message
            Return False
        End Try
    End Function
    Private Function DodeliObjOdXml(ByVal podatociGlavaDok As Dictionary(Of String, String), ByVal brArtikli As Integer, ByVal listaPodatociStavkiDok As List(Of Dictionary(Of String, String)), ByRef poraka As String) As Boolean
        Try
            Me.glavaDok.WebShopDokID = ProveriRecnik(podatociGlavaDok, "PROFORMAID")
            Me.glavaDok.SifraKup = Regex.Replace(ProveriRecnik(podatociGlavaDok, "CUSTOMERID"), "[^\d]", "").Trim 'Od WebShop se praka "ERP 000001" i zatoa pravime replace da go trgni ERP
            Me.glavaDok.WebShopNarID = ProveriRecnik(podatociGlavaDok, "ORDERID")
            Me.glavaDok.CFMADokID = ProveriRecnik(podatociGlavaDok, "PROFORMAERPID")
            Me.glavaDok.MetodIsoraka = ProveriRecnik(podatociGlavaDok, "DELIVERYMETHOD")
            Me.glavaDok.MestoIsporaka = ProveriRecnik(podatociGlavaDok, "DELIVERYPLACE")
            Me.glavaDok.KraenRokIsporaka = ProveriRecnik(podatociGlavaDok, "DEADLINE")
            Me.glavaDok.Datum = Konv.StrSamoBrojkiVoDate(ProveriRecnik(podatociGlavaDok, "DATE"))
            Me.glavaDok.NaracanoOd = ProveriRecnik(podatociGlavaDok, "ORDEREDBY")
            Me.glavaDok.Komenatar = ProveriRecnik(podatociGlavaDok, "COMMENT")
            Me.glavaDok.Status = ProveriRecnik(podatociGlavaDok, "STATUS")
            Me.glavaDok.NaracanoOdKorisnik = ProveriRecnik(podatociGlavaDok, "ORDEREDBYUSERNAME")
            Me.glavaDok.ProverenoOdKorisnik = ProveriRecnik(podatociGlavaDok, "SOLVEDBYUSERNAME")
            Me.glavaDok.Opcija1 = ProveriRecnik(podatociGlavaDok, "OPTION1")
            Me.glavaDok.Opcija2 = ProveriRecnik(podatociGlavaDok, "OPTION2")
            Me.glavaDok.Opcija3 = ProveriRecnik(podatociGlavaDok, "OPTION3")
            Me.glavaDok.Opcija4 = ProveriRecnik(podatociGlavaDok, "OPTION4")
            Me.glavaDok.BrojArtikli = brArtikli

            Dim imePrezimeKupKomerc() As String = Me.glavaDok.NaracanoOd.Split(";")
            If imePrezimeKupKomerc.Length = 2 Then
                Me.glavaDok.ImePrezimeKupuvac = imePrezimeKupKomerc(0)
                Me.glavaDok.ImePrezimeKomercijalist = imePrezimeKupKomerc(1)
            ElseIf imePrezimeKupKomerc.Length = 1 Then
                Me.glavaDok.ImePrezimeKupuvac = imePrezimeKupKomerc(0)
            End If
            Dim stavkiOdDok(listaPodatociStavkiDok.Count - 1) As StavkiFaktura

            For index As Integer = 0 To listaPodatociStavkiDok.Count - 1
                stavkiOdDok(index).SifraArt = ProveriRecnik(listaPodatociStavkiDok(index), "INTERNALCODE")
                stavkiOdDok(index).Cena = Konv.ObjVoDec(ProveriRecnik(listaPodatociStavkiDok(index), "PRICE"))
                stavkiOdDok(index).Kolic = Konv.ObjVoDec(ProveriRecnik(listaPodatociStavkiDok(index), "QUANTITY"))
                stavkiOdDok(index).Popust = Konv.ObjVoDec(ProveriRecnik(listaPodatociStavkiDok(index), "DISCOUNT"))
                stavkiOdDok(index).WebShopDokID = ProveriRecnik(listaPodatociStavkiDok(index), "PROFORMAID")
            Next
            Me.stavkiFak = stavkiOdDok
            Return True
        Catch ex As Exception
            poraka = "Greska vo DodeliObj Faktura! " & ex.Message
            Return False
        End Try

    End Function
    Private Function ProveriRecnik(ByVal recnik As Dictionary(Of String, String), ByVal kluc As String) As String
        If recnik.ContainsKey(kluc) Then
            Return recnik(kluc)
        Else
            Return ""
        End If
    End Function
    Public Function ZacuvajDokVoWEBSHOP(ByRef sesijaId As String, ByRef poraka As String) As Boolean ' TREBA DA SE TESTIRA !!!!
        Try
            Dim atributi As New ListaAtributi
            atributi.Add(DokumentZaglavje.Guid.ToString, TekParam.GuID)
            atributi.Add(DokumentZaglavje.Version.ToString, "2")
            atributi.Add(DokumentZaglavje.Proformainvoiceerpid.ToString, Me.glavaDok.CFMADokID)
            atributi.Add(DokumentZaglavje.Numrows.ToString, Me.glavaDok.BrojArtikli.ToString)
            atributi.Add(DokumentZaglavje.Customererpid.ToString, Me.glavaDok.SifraKup)
            atributi.Add(DokumentZaglavje.Comment.ToString, Me.glavaDok.Komenatar)
            atributi.Add(DokumentZaglavje.Orderedby.ToString, Me.glavaDok.NaracanoOdKorisnik)
            atributi.Add(DokumentZaglavje.Deliverytype.ToString, Me.glavaDok.MetodIsoraka)
            atributi.Add(DokumentZaglavje.Deliveryplace.ToString, Me.glavaDok.MestoIsporaka)
            atributi.Add(DokumentZaglavje.Deadline.ToString, Me.glavaDok.KraenRokIsporaka)
            atributi.Add(DokumentZaglavje.Orderid.ToString, Me.glavaDok.WebShopNarID)
            atributi.Add(DokumentZaglavje.Completed.ToString, Me.glavaDok.KompletiranaFaktura.ToString)
            atributi.Add(DokumentZaglavje.Option1.ToString, Me.glavaDok.Opcija1)
            atributi.Add(DokumentZaglavje.Option2.ToString, Me.glavaDok.Opcija2)
            atributi.Add(DokumentZaglavje.Option3.ToString, Me.glavaDok.Opcija3)
            atributi.Add(DokumentZaglavje.Option4.ToString, Me.glavaDok.Opcija4)

            Dim urlPrati As String = AlatkaURL.PripremiURLzaPrakanje("Uploadinvoice.aspx", "", atributi)
            If AlatkaURL.PratiURL(urlPrati, poraka, TekParam.TipObrakanje.Rezultat) Then
                If poraka.StartsWith("OK") AndAlso poraka.Contains("OK-") Then
                    Dim pom() As String = poraka.Split("-")
                    If pom.Length > 1 Then
                        sesijaId = pom(1)
                    End If
                End If
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            poraka = "GRESKA VO ZacuvajDokrVoWebShop: " & ex.Message
            Return False
        End Try
    End Function ' TREBA DA SE TESTIRA !!!!
    Public Function ZacuvajStavkiDokVoWEBSHOP(ByVal sesijaId As String, ByRef poraka As String) As Boolean ' TREBA DA SE TESTIRA !!!!
        Try
            If String.IsNullOrEmpty(sesijaId) Then
                poraka = "GRESKA ZacuvajStavkiDokWEBSHOP: Nema sesijaID! "
                Return False
            End If

            For index As Integer = 0 To Me.Stavki.Length - 1

                Dim atributi As New ListaAtributi
                atributi.Add(StavkiDokument.SessionId.ToString, sesijaId)
                atributi.Add(StavkiDokument.Internalcode.ToString, Me.Stavki(index).SifraArt)
                atributi.Add(StavkiDokument.Dquant.ToString, Me.Stavki(index).Kolic)
                atributi.Add(StavkiDokument.Price.ToString, Me.Stavki(index).Cena) 'TREBA DA SE PROVERI KOJA CENA DA E PRAKA
                atributi.Add(StavkiDokument.Vat.ToString, Me.Stavki(index).Danok)
                atributi.Add(StavkiDokument.Disc.ToString, Me.Stavki(index).Popust)
                atributi.Add(StavkiDokument.currencyid.ToString, Me.glavaDok.BrojArtikli.ToString)
                atributi.Add(StavkiDokument.marketprice.ToString, Me.Stavki(index).Cena) 'TREBA DA SE PROVERI KOJA CENA DA E PRAKA
                Dim dodUrl As String = Konv.VratiSesijaStringZaUrl(sesijaId)
                Dim urlPrati As String = AlatkaURL.PripremiURLzaPrakanje("uploadinvoicecontent.aspx", dodUrl, atributi)
                If AlatkaURL.PratiURL(urlPrati, poraka, TekParam.TipObrakanje.Rezultat) Then 'Vraka br na redeovi preostanati
                    If poraka.StartsWith("OK") Then 'TREBA DA SE TESTIRA
                        Return True
                    ElseIf Konv.TxtVoInt(poraka) = 0 Then ' TREBA DA SE TESTIRA 
                        Return False
                    End If
                Else
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception
            poraka = "GRESKA ZacuvajStavkiDokWEBSHOP: " & ex.Message
            Return False
        End Try
    End Function ' TREBA DA SE TESTIRA!!!!!
    Public ReadOnly Property GlavaNaDokument() As GlavaDokument
        Get
            Return Me.glavaDok
        End Get
    End Property
    Public ReadOnly Property Stavki() As StavkiFaktura()
        Get
            Return Me.stavkiFak
        End Get
    End Property
    Public ReadOnly Property BrojDokumentCFMA() As String
        Get
            Return BrojDokCFMA
        End Get
    End Property
End Class





