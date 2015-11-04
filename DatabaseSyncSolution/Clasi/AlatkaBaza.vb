Option Strict On
Imports System.Data.SqlClient
Public Class AlatkaBaza
    Public Shared Function ZemiArtikliZaPrakanje(ByRef ds As DataSet, ByRef poraka As String) As Boolean
        Dim sqlCn As New SqlConnection()
        Dim sqlCmd As New SqlCommand()
        Dim sqlAd As New SqlDataAdapter()
        Try
            sqlCn.ConnectionString = TekParam.KonekcijaString
            With sqlCmd
                .Connection = sqlCn
                .CommandText = "sp_ws_ZemiArtikliZaPrakanje"
                .CommandType = CommandType.StoredProcedure
            End With
            sqlAd.SelectCommand = sqlCmd
            If ds Is Nothing Then
                ds = New DataSet()
            End If
            sqlCn.Open()
            sqlAd.Fill(ds)
            sqlCn.Close()
            Return True
        Catch ex As Exception
            sqlCn.Close()
            poraka = "Greska ZemiArtikliZaPraknaje(): " & ex.Message
            Return False
        End Try
    End Function
    Public Shared Function ZapisiUspesnoIzmenetArtikal(ByVal sifraArt As String, ByVal InterenKod As String, ByVal Kolic As Decimal, ByVal Cena As Decimal, ByVal Status As Integer, ByVal uspesno As Boolean, ByRef poraka As String) As Boolean
        Dim sqlCn As New SqlConnection()
        Dim sqlCmd As New SqlCommand()
        Try
            sqlCn.ConnectionString = TekParam.KonekcijaString
            With sqlCmd
                .Connection = sqlCn
                .CommandText = "sp_ws_ZapisDaliPrenesenArtikal"
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@Sifra_Art", sifraArt)
                .Parameters.AddWithValue("@Alt_Sifra", InterenKod)
                .Parameters.AddWithValue("@Sostojba", Status)
                .Parameters.AddWithValue("@Kolic", Kolic)
                .Parameters.AddWithValue("@CenaGolemo", Cena)
                .Parameters.AddWithValue("@Uspeh", IIf(uspesno, "D", "N"))
                sqlCn.Open()
                .ExecuteNonQuery()
            End With
            sqlCn.Close()
            Return True
        Catch ex As Exception
            sqlCn.Close()
            poraka = "Greska ZapisiUspesnoIzmenetArtikal(): " & ex.Message
            Return False
        End Try
    End Function
    Public Shared Function ProveriDaliPostoiDok(ByVal wIdentif_Br As String, ByVal sifra_Kup As String, ByRef wBrojDok As String, ByRef poraka As String) As Boolean
        wBrojDok = ""
        'Morame 21.02.2014 imase Problem - Pred toa! ->Return True ' Proverkata Ne funkcionira!
        ' Bidejki brojot sto go dobivame od WebShopMatrix ne e ednoznacen!!! MarkoM

        Dim sqlCn As New SqlConnection()
        Dim sqlCmd As New SqlCommand()
        wBrojDok = ""
        Try
            sqlCn.ConnectionString = TekParam.KonekcijaString
            With sqlCmd
                .Connection = sqlCn
                .CommandText = "sp_ws_ProveriDaliPostoiDok"
                .CommandType = CommandType.StoredProcedure
                .Parameters.Clear()
                .Parameters.Add(New SqlParameter("@Identif_Br", wIdentif_Br))
                .Parameters.Add(New SqlParameter("@Sifra_Kup", sifra_Kup))
                sqlCn.Open()
                wBrojDok = Konv.ObjVoStr(.ExecuteScalar())
                If wBrojDok.Trim <> "" Then ' Ako VRATI nesto znaci deka postoi dokumentot
                    AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.ZACUVAJ, TekParam.TipRezultat.NEUspesno, AlatkaFile.StoPraka.Dokumenti, "GRESKA -  Zapis DUPLA Faktura - " & wBrojDok, wBrojDok)
                    poraka = "Dupli dokument: " & wBrojDok
                    Return False
                Else
                    Return True
                End If
            End With
            sqlCn.Close()
            Return True
        Catch ex As Exception
            sqlCn.Close()
            poraka = "Greska ProveriDaliPostoiDok(): " & ex.Message
            Return False
        End Try
    End Function
    Public Shared Function ZacuvajDokument(ByVal wSifra_Oe As String, ByVal wSifra_Dok As String, ByVal wBroj_Dok As String, ByVal wSifra_Kup As String, ByVal wIdentif_Br As String, ByVal wIspratnica As String, ByVal wDatum_Dok As Date, ByVal wStavkiDok() As Faktura.Stavr, _
        ByVal wKto As String, ByVal wSifra_Nal As Integer, ByVal wUces As Decimal, ByVal wRok As Integer, ByVal Toc200 As Decimal, ByVal Marza As Decimal, ByVal Danok1 As Decimal, ByVal Danok2 As Decimal, ByVal Magvr As Decimal, ByVal MagvrDan As Decimal, _
        ByVal Pdanok1 As Decimal, ByVal Pdanok2 As Decimal, ByVal DanDokCena As String, ByVal DanMagCena As String, ByVal Kod_Danok As String, ByVal komentar As String, ByRef poraka As String) As Boolean
        Dim wUspeh As Boolean
        Dim DokrId As Integer
        Dim wBrojDok As String
        Dim sqlTrans As SqlTransaction
        Dim sqlCn As New SqlConnection()

        sqlCn.ConnectionString = TekParam.KonekcijaString
        Try
            sqlCn.Open()
            sqlTrans = sqlCn.BeginTransaction

            If ZacuvajDokr(sqlTrans, sqlCn, wSifra_Oe, wSifra_Dok, wBroj_Dok, wSifra_Kup, wIdentif_Br, wIspratnica, wDatum_Dok, _
                        wKto, wSifra_Nal, wUces, wRok, Toc200, Marza, Danok1, Danok2, Magvr, MagvrDan, Pdanok1, Pdanok2, DanDokCena, DanMagCena, Kod_Danok, DokrId, wBrojDok, komentar, poraka) Then
                If ZacuvajStavr(sqlTrans, sqlCn, wSifra_Oe, wSifra_Dok, wBrojDok, wDatum_Dok, wStavkiDok, DokrId, poraka) Then
                    sqlTrans.Commit()
                    wUspeh = True
                Else
                    sqlTrans.Rollback()
                    sqlCn.Close()
                    Return False
                End If
            Else
                sqlTrans.Rollback()
                sqlCn.Close()
                Return False
            End If
            sqlCn.Close()
        Catch ex As Exception
            Try
                sqlTrans.Rollback()
            Catch ex1 As Exception
            End Try
            sqlCn.Close()
            poraka = "Greska vo ZacuvajDokument:" + ex.Message
            Return False
        End Try
        Return wUspeh
    End Function
    Public Shared Function ZacuvajDokr(ByVal SqlTransaction As SqlTransaction, ByVal sqlConn As SqlConnection, ByVal wSifra_Oe As String, ByVal wSifra_Dok As String, ByVal wBroj_Dok As String, ByVal wSifra_Kup As String, ByVal wIdentif_Br As String, ByVal wIspratnica As String, ByVal wDatum_Dok As Date, _
        ByVal wKto As String, ByVal wSifra_Nal As Integer, ByVal wUces As Decimal, ByVal wRok As Integer, ByVal Toc200 As Decimal, ByVal Marza As Decimal, ByVal Danok1 As Decimal, ByVal Danok2 As Decimal, ByVal Magvr As Decimal, ByVal MagvrDan As Decimal, _
        ByVal Pdanok1 As Decimal, ByVal Pdanok2 As Decimal, ByVal DanDokCena As String, ByVal DanMagCena As String, ByVal Kod_Danok As String, ByRef DokrId As Integer, ByRef wBrojDok As String, ByVal komentar As String, ByRef poraka As String) As Boolean

        Dim sqlCmd As New SqlCommand()
        Dim wNula As Decimal = 0

        Try
            With sqlCmd
                .Connection = sqlConn
                .Transaction = SqlTransaction
                .CommandText = "sp_ZacuvajDokr"
                .CommandType = CommandType.StoredProcedure
                .Parameters.Clear()
                .Parameters.Add(New SqlParameter("@Sifra_oe", Convert.ToInt16(wSifra_Oe)))
                .Parameters.Add(New SqlParameter("@Sifra_dok", wSifra_Dok))
                .Parameters.Add(New SqlParameter("@Broj_Dok", Convert.ToInt32(wBroj_Dok)))
                .Parameters.Add(New SqlParameter("@Sifra_Prim", DBNull.Value))
                .Parameters.Add(New SqlParameter("@Imadodatna", DBNull.Value))

                .Parameters.Add(New SqlParameter("@VlIzl", "I"))
                .Parameters.Add(New SqlParameter("@Sifra_Za", "1"))
                .Parameters.Add(New SqlParameter("@Sifra_nal", wSifra_Nal))
                .Parameters.Add(New SqlParameter("@Broj_nal", DBNull.Value))
                .Parameters.Add(New SqlParameter("@identif_br", wIdentif_Br))
                .Parameters.Add(New SqlParameter("@Ispratnica", wIspratnica))
                .Parameters.Add(New SqlParameter("@Opis", DBNull.Value))
                .Parameters.Add(New SqlParameter("@Sifra_Kup", wSifra_Kup))
                .Parameters.Add(New SqlParameter("@Sifra_obj", DBNull.Value))
                .Parameters.Add(New SqlParameter("@Sifra_mest", DBNull.Value))

                .Parameters.Add(New SqlParameter("@Datum_dok", wDatum_Dok))
                .Parameters.Add(New SqlParameter("@Uces", wUces))
                .Parameters.Add(New SqlParameter("@Kasa", wNula))
                .Parameters.Add(New SqlParameter("@KasaPoDDV", wNula))
                .Parameters.Add(New SqlParameter("@Rok", wRok))
                .Parameters.Add(New SqlParameter("@Sifra_Pat", 2)) 'FIKSIRANO KAJ NIV - Treba da se zacuva PATNIK - WEBSHOP
                .Parameters.Add(New SqlParameter("@SerBr", DBNull.Value))
                .Parameters.Add(New SqlParameter("@KojaSmetka", DBNull.Value))
                .Parameters.Add(New SqlParameter("@Kto", wKto))
                .Parameters.Add(New SqlParameter("@Kurs", wNula))
                .Parameters.Add(New SqlParameter("@KojaVal", DBNull.Value))

                .Parameters.Add(New SqlParameter("@Toc200", Toc200))
                .Parameters.Add(New SqlParameter("@Marza", Marza))
                .Parameters.Add(New SqlParameter("@Magvr", Magvr))
                .Parameters.Add(New SqlParameter("@Magvrdan", MagvrDan))
                .Parameters.Add(New SqlParameter("@Danok1", Danok1))
                .Parameters.Add(New SqlParameter("@Danok2", Danok2))
                .Parameters.Add(New SqlParameter("@Pdanok1", Pdanok1))
                .Parameters.Add(New SqlParameter("@Pdanok2", Pdanok1))
                .Parameters.Add(New SqlParameter("@PTrosok", wNula))
                .Parameters.Add(New SqlParameter("@Dandokcena", DanDokCena))
                .Parameters.Add(New SqlParameter("@DanMagCena", DanMagCena))
                .Parameters.Add(New SqlParameter("@CenaIznos", ""))
                .Parameters.Add(New SqlParameter("@ProcMarza", wNula))
                .Parameters.Add(New SqlParameter("@tekstposle", komentar))
                .Parameters.Add(New SqlParameter("@Kod_danok", ""))
                .Parameters.Add(New SqlParameter("@Sifra_Nivo", DBNull.Value))
                .Parameters.Add(New SqlParameter("@Blokiran", DBNull.Value))
                .Parameters.Add(New SqlParameter("@Sifra_OeNar", DBNull.Value))
                .Parameters.Add(New SqlParameter("@Sifra_Nar", DBNull.Value))
                .Parameters.Add(New SqlParameter("@Broj_Nar", DBNull.Value))
                .Parameters.Add(New SqlParameter("@Izrab_Nar", DBNull.Value))
                .Parameters.Add(New SqlParameter("@Spremil", DBNull.Value))
                .Parameters.Add(New SqlParameter("@Usluzna", "N"))
                .Parameters.Add(New SqlParameter("@TekstPred", DBNull.Value))
                .Parameters.Add(New SqlParameter("@Pec_Dok", DBNull.Value))
                .Parameters.Add(New SqlParameter("@Br_Kopii", DBNull.Value))
                .Parameters.Add(New SqlParameter("@Spec_Forma_Pec", DBNull.Value))
                .Parameters.Add(New SqlParameter("@Sifra_Prev", DBNull.Value))
                '.Parameters.Add(New SqlParameter("@DodadenNa", DatSoVreme))
                'Si ima default vo PROCEDURATA  .Parameters.Add(New SqlClient.SqlParameter("@MaxObidi", 50))

                .Parameters.Add(New SqlClient.SqlParameter("@dokrid", SqlDbType.Int))
                .Parameters("@dokrid").Direction = ParameterDirection.Output
                .Parameters("@Broj_Dok").Direction = ParameterDirection.InputOutput

                .Parameters.Add(New SqlClient.SqlParameter("@Greska", SqlDbType.Int))
                .Parameters("@Greska").Direction = ParameterDirection.ReturnValue

                .ExecuteNonQuery()
                'iGreska = Convert.ToInt32(.Parameters("@Greska").Value())

                DokrId = Convert.ToInt32(.Parameters("@dokrid").Value())
                wBrojDok = Convert.ToString(.Parameters("@Broj_dok").Value())
                'sqlConn.Open()
                '.ExecuteNonQuery()
            End With
            'sqlConn.Close()
            Return True
        Catch ex As Exception
            'sqlConn.Close()
            poraka = "Greska ZacuvajDokr(): " & ex.Message
            Return False
        End Try
    End Function
    Public Shared Function ZacuvajStavr(ByVal SqlTransaction As SqlTransaction, ByVal sqlConn As SqlConnection, ByVal wSifra_Oe As String, _
                    ByVal wSifra_Dok As String, ByVal wBroj_Dok As String, ByVal wDatum_Dok As Date, ByVal wStavkiDok() As Faktura.Stavr, ByVal DokrId As Integer, ByRef poraka As String) As Boolean
        Dim sqlCmd As New SqlCommand()
        Dim iW As Integer
        Dim wNula As Decimal = 0
        Dim wUspeh As Boolean = False
        Try
            For iW = 0 To UBound(wStavkiDok)
                With sqlCmd
                    .Connection = sqlConn
                    .Transaction = SqlTransaction
                    .CommandText = "sp_ZacuvajStavr"
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Clear()
                    .Parameters.Add(New SqlParameter("@Sifra_oe", Convert.ToInt16(wSifra_Oe)))
                    .Parameters.Add(New SqlParameter("@Sifra_dok", wSifra_Dok))
                    .Parameters.Add(New SqlParameter("@Broj_Dok", Convert.ToInt32(wBroj_Dok)))
                    .Parameters.Add(New SqlParameter("@Sifra_Prim", DBNull.Value))
                    .Parameters.Add(New SqlParameter("@Imadodatna", DBNull.Value))

                    .Parameters.Add(New SqlParameter("@RBr", wStavkiDok(iW).STRBr))
                    .Parameters.Add(New SqlParameter("@VlIzl", "I"))
                    .Parameters.Add(New SqlParameter("@Sifra_art", wStavkiDok(iW).STSifra_Art))
                    .Parameters.Add(New SqlParameter("@Kolic", wStavkiDok(iW).STKolic))
                    .Parameters.Add(New SqlParameter("@DokCena", wStavkiDok(iW).STDokCena))
                    .Parameters.Add(New SqlParameter("@MagCena", wStavkiDok(iW).STMagCena))
                    .Parameters.Add(New SqlParameter("@NabCena", wStavkiDok(iW).STNabCena))
                    .Parameters.Add(New SqlParameter("@POsn", wStavkiDok(iW).STPOsn))
                    .Parameters.Add(New SqlParameter("@POsnPren", wStavkiDok(iW).STPOsn))
                    .Parameters.Add(New SqlParameter("@PTar", wStavkiDok(iW).STPTar))

                    .Parameters.Add(New SqlParameter("@UcesDod", wNula))
                    .Parameters.Add(New SqlParameter("@UcesOsn", wNula))
                    .Parameters.Add(New SqlParameter("@UcesKol", wNula))
                    .Parameters.Add(New SqlParameter("@Uces", wStavkiDok(iW).STUces))

                    .Parameters.Add(New SqlParameter("@DanDokCena", wStavkiDok(iW).STDanDokCena))
                    .Parameters.Add(New SqlParameter("@DanMagCena", wStavkiDok(iW).STDanMagCena))
                    .Parameters.Add(New SqlParameter("@CenaIznos", "N"))
                    .Parameters.Add(New SqlParameter("@Kod_Danok", "D")) '?sto e ova
                    .Parameters.Add(New SqlParameter("@Datum_Dok", wDatum_Dok))

                    .Parameters.Add(New SqlParameter("@Tros", ""))
                    .Parameters.Add(New SqlParameter("@KontrSer", DBNull.Value))
                    .Parameters.Add(New SqlParameter("@DaliCenaSoa", "D"))
                    .Parameters.Add(New SqlParameter("@ImeMat", ""))
                    .Parameters.Add(New SqlParameter("@EdMera", ""))
                    .Parameters.Add(New SqlParameter("@StaviNabcKata", "N"))
                    .Parameters.Add(New SqlParameter("@Dobav", DBNull.Value))
                    .Parameters.Add(New SqlParameter("@KaloProc", DBNull.Value))
                    .Parameters.Add(New SqlParameter("@Alt_Kolic", DBNull.Value))

                    .Parameters.Add(New SqlClient.SqlParameter("@dokrid", DokrId))
                    'sqlConn.Open()
                    .ExecuteNonQuery()
                End With
            Next
            'sqlConn.Close()
            Return True
        Catch ex As Exception
            'sqlConn.Close()
            poraka = "Greska ZacuvajStavr(): " & ex.Message
            Return False
        End Try
        Return wUspeh
    End Function
    Public Shared Function NajdiSledenBrojDok(ByVal wSifra_Oe As String, ByVal wSifra_Dok As String, ByRef wBroj_Dok As String, ByRef poraka As String) As Boolean
        Dim wUspeh As Boolean
        Dim sqlCn As New SqlConnection()
        Dim sqlCmd As New SqlCommand()
        Dim sqlCitac As SqlDataReader

        Try
            sqlCn.ConnectionString = TekParam.KonekcijaString
            With sqlCmd
                .Connection = sqlCn
                .CommandText = "sp_PodigniBrDok"
                .CommandType = CommandType.StoredProcedure
                .Parameters.Add(New SqlParameter("@Sifra_Oe", wSifra_Oe))
                .Parameters.Add(New SqlParameter("@Sifra_Dok", wSifra_Dok))
                sqlCn.Open()
                sqlCitac = .ExecuteReader()
            End With

            With sqlCitac
                While .Read
                    wBroj_Dok = (Konv.ObjVoInt(.Item("Broj_Dok")) + 1).ToString
                End While
                .Close()
            End With

            sqlCn.Close()
            Return True
        Catch ex As Exception
            Try
                sqlCitac.Close()
            Catch ex1 As Exception
            End Try
            sqlCn.Close()
            poraka = "Greska NajdiSledenBrojDok(): " & ex.Message
            Return False
        End Try
    End Function
    Public Shared Function NajdiPodatociZaArtikl(ByVal wAlt_Sifra As String, ByRef Sifra_Art As String, ByRef Sifra_Tar As String, ByRef EdMera As String, ByRef poraka As String, ByRef ProcOsn As Decimal, ByRef DogCena As Decimal) As Boolean
        Dim wUspeh As Boolean
        Dim sqlCn As New SqlConnection()
        Dim sqlCmd As New SqlCommand()
        Dim sqlCitac As SqlDataReader
        Try
            sqlCn.ConnectionString = TekParam.KonekcijaString
            With sqlCmd
                .Connection = sqlCn
                .CommandText = "sp_ws_NajdiPodatociZaArtikl"
                .CommandType = CommandType.StoredProcedure
                .Parameters.Add(New SqlParameter("@Alt_Sifra", wAlt_Sifra))
                sqlCn.Open()
                sqlCitac = .ExecuteReader
            End With

            With sqlCitac
                While .Read
                    Sifra_Art = Konv.ObjVoStr(.Item("Sifra_Art"))
                    Sifra_Tar = Konv.ObjVoStr(.Item("Sifra_Tar"))
                    EdMera = Konv.ObjVoStr(.Item("EdMera"))
                    ProcOsn = Konv.ObjVoDec(.Item("ProcOsn"))
                    DogCena = Konv.ObjVoDec(.Item("DogCena"))
                End While
                .Close()
            End With

            sqlCn.Close()
            Return True
        Catch ex As Exception
            Try
                sqlCitac.Close()
            Catch ex1 As Exception
            End Try
            sqlCn.Close()
            poraka = "Greska NajdiPodatociZaArtikl(): " & ex.Message
            Return False
        End Try
    End Function
    Public Shared Function ZacuvajVoLOG(ByVal tOperacija As TekParam.TipAkcija, ByVal tRezultat As TekParam.TipRezultat, ByVal stoE As AlatkaFile.StoPraka, ByVal poraka As String, ByVal dodOpis As String) As Boolean
        Dim sqlCn As New SqlConnection()
        Dim cm As New SqlCommand()
        Dim param As SqlParameter

        Try
            sqlCn.ConnectionString = TekParam.KonekcijaString
            With cm
                .Connection = sqlCn
                .CommandText = "sp_ws_ZacuvajVoLOG"
                .CommandType = CommandType.StoredProcedure
                param = New SqlParameter("TipOperacija", tOperacija.ToString)
                .Parameters.Add(param)
                param = New SqlParameter("Uspeh", tRezultat.ToString)
                .Parameters.Add(param)
                param = New SqlParameter("Vreme", DateTime.Now)
                .Parameters.Add(param)
                If poraka.Length > 2000 Then 'Poleto vo baza e nvarchar(2000)
                    poraka = poraka.Substring(0, 1999)
                End If
                param = New SqlParameter("Poraka", poraka)
                .Parameters.Add(param)
                If dodOpis.Length > 100 Then 'Poleto vo baza e nvarchar(100)
                    dodOpis = dodOpis.Substring(0, 99)
                End If
                param = New SqlParameter("DodOpis", dodOpis)
                .Parameters.Add(param)
                param = New SqlParameter("Klasa", stoE.ToString)
                .Parameters.Add(param)
                sqlCn.Open()
                .ExecuteNonQuery()
            End With
            sqlCn.Close()
            Return True
        Catch ex As Exception
            sqlCn.Close()
            'Problem pri zapis vo log
            Return False
        End Try
    End Function
    Public Shared Function NajdiPodatociZaKomint(ByVal wSifra_Kup As String, ByRef Uces As Decimal, ByRef Rok_Dosp As Integer, ByRef poraka As String) As Boolean
        Dim wUspeh As Boolean
        Dim sqlCn As New SqlConnection()
        Dim sqlCmd As New SqlCommand()
        Dim sqlCitac As SqlDataReader
        Try
            sqlCn.ConnectionString = TekParam.KonekcijaString
            With sqlCmd
                .Connection = sqlCn
                .CommandText = "sp_PodigniMatic"
                .CommandType = CommandType.StoredProcedure
                .Parameters.Add(New SqlParameter("@Sifra", wSifra_Kup))
                .Parameters.Add(New SqlParameter("@Tabela", "KOMINT"))
                sqlCn.Open()
                sqlCitac = .ExecuteReader
            End With
            With sqlCitac
                While .Read
                    Uces = Konv.ObjVoDec(.Item("Uces"))
                    Rok_Dosp = Konv.ObjVoInt(.Item("Rok_Dosp"))
                End While
                .Close()
            End With

            sqlCn.Close()
            Return True
        Catch ex As Exception
            Try
                sqlCitac.Close()
            Catch ex1 As Exception
            End Try
            sqlCn.Close()
            poraka = "Greska NajdiPodatociZaKomint(): " & ex.Message
            Return False
        End Try
    End Function
    Public Shared Function NajdiPodatociZaTipDok(ByVal wSifra_Dok As String, ByRef Sifra_Nal As Integer, ByRef Kto As String, ByRef poraka As String) As Boolean
        Dim wUspeh As Boolean
        Dim sqlCn As New SqlConnection()
        Dim sqlCmd As New SqlCommand()
        Dim sqlCitac As SqlDataReader
        Dim KoeKto As String
        Try
            sqlCn.ConnectionString = TekParam.KonekcijaString
            With sqlCmd
                .Connection = sqlCn
                .CommandText = "sp_PodigniMatic"
                .CommandType = CommandType.StoredProcedure
                .Parameters.Add(New SqlParameter("@Sifra", wSifra_Dok))
                .Parameters.Add(New SqlParameter("@Tabela", "TIPDOK"))
                sqlCn.Open()
                sqlCitac = .ExecuteReader
            End With

            With sqlCitac
                While .Read
                    Sifra_Nal = Konv.ObjVoInt(.Item("Sifra_Nal"))
                    KoeKto = Konv.ObjVoStr(.Item("KoeKto"))

                    Dim wDoKade As Integer = KoeKto.IndexOf(" ")
                    If wDoKade > 0 Then
                        Kto = KoeKto.Substring(0, wDoKade)
                    ElseIf KoeKto.Length > 0 Then
                        Kto = KoeKto
                    End If
                End While
                .Close()
            End With

            sqlCn.Close()
            Return True
        Catch ex As Exception
            Try
                sqlCitac.Close()
            Catch ex1 As Exception
            End Try
            sqlCn.Close()
            poraka = "Greska NajdiPodatociZaKomint(): " & ex.Message
            Return False
        End Try
    End Function
End Class
