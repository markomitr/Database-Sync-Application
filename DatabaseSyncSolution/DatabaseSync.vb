Imports System.Net
Imports System.IO
Imports System.Collections.Generic
Imports System
Imports System.Timers
Imports System.Runtime.InteropServices

Module DatabaseSync
    Dim Verzija As String = "24.03.2014"
    Dim poraka As String = ""
    '23.12.2012 MarkoM
    '31.12.2012 MarkoM Krajni Izmeni - (Problem so Upload na FTP)
    '21.02.2014 Dupli dokumenti - pustena proverkata

    'Test za komunikacija so WebShop Matrix - http://matrixtest.autodelovi.rs
    'Testiranje na funkciite za Komunikacija pomegu CFMA(InfoBiro) - WebShop(MatrixAutoDeloviRS)
    'Site povici se odvivaat preku HTPP REQUEST/RESPONSE

    Dim WithEvents timerDokumenti, timerArtikli As Timers.Timer
    Dim kolkuTaktoviDokumenti As Integer = 0
    Dim kolkuTaktoviArtikli As Integer = 0
    Dim NAkolkuTaktoviZapisLogDokument As Integer = 0
    Dim NAkolkuTaktoviZapisLogArtikli As Integer = 0
    Sub Main()
        'RabotiArtikalEDEN() ' Prakanje na Artikal - azuranje na Cena i Zaliha
        'RabotiKomintent() ' Prakanje na Komintent
        'RabotiKorisnikZaKomint() ' Prakanje na KORISNIK za Komintent 
        'RabotiDokumenti()
        AddHandler AppDomain.CurrentDomain.ProcessExit, AddressOf CurrentDomain_ProcessExit
        AddHandler Console.CancelKeyPress, AddressOf Console_CancelKeyPress

        AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.KONTROLNO, TekParam.TipRezultat.Uspesno, AlatkaFile.StoPraka.Neodredeno, "STARTOVANO EXE za Prevzemanje", "EXE")
        PrikaziKonfiguracija()
        KreirajTimeri()
        Console.WriteLine("------------------------------------------------------------------------------")
        Do
            Dim keyinfo = Console.ReadKey()
            If keyinfo.Key = ConsoleKey.F2 Then 'Za IZLEZ se pritiska F2
                Exit Sub
            End If

        Loop
    End Sub
    Private Sub KreirajTimeri()
        timerDokumenti = New Timers.Timer()
        timerArtikli = New Timers.Timer()
        AddHandler timerDokumenti.Elapsed, AddressOf TimerRabotiDokumenti
        AddHandler timerArtikli.Elapsed, AddressOf TimerRabotiArtikli
        Try
            Dim timerDokumentiVoSec As Integer = Konv.TxtVoInt(TekParam.ZemiKonfig("TimerDokumentiVoSec")) ' Ako e 0 Timerot ne se uklucuva
            Dim timerArtikliVoSec As Integer = Konv.TxtVoInt(TekParam.ZemiKonfig("TimerArtikliVoSec")) ' Ako e 0 Timerot ne se uklucuva

            If timerDokumentiVoSec > 0 Then ' Ako e 0 Timerot ne se uklucuva
                AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.KONTROLNO, TekParam.TipRezultat.Uspesno, AlatkaFile.StoPraka.Neodredeno, "STARTOVAN TIMER za Prevzemanje DOKUMENTI! Time.sec:" & timerDokumentiVoSec.ToString, "TIMER DOKUMENTI")
                timerDokumenti.Interval = IIf(timerDokumentiVoSec = 0, 5, timerDokumentiVoSec) * 1000
                NAkolkuTaktoviZapisLogDokument = (86400 / timerDokumentiVoSec) / 24  ' Na sekoj 1h da zapisvua vo log deka raboti
                timerDokumenti.Enabled = True
                Console.WriteLine(DateTime.Now.ToString & " Timer DOKUMENTI: STARTOVAN!")
            Else
                Console.WriteLine(DateTime.Now.ToString & " Timer DOKUMENTI: NE E STARTOVAN! - Konfig-> TimerSec=0")
            End If
            If timerArtikliVoSec > 0 Then ' Ako e 0 Timerot ne se uklucuva
                AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.KONTROLNO, TekParam.TipRezultat.Uspesno, AlatkaFile.StoPraka.Neodredeno, "STARTOVAN TIMER za Prevzemanje ARTIKLI! Time.sec:" & timerArtikliVoSec.ToString, "TIMER ARTIKLI")
                timerArtikli.Interval = IIf(timerArtikliVoSec = 0, 60, timerArtikliVoSec) * 1000
                NAkolkuTaktoviZapisLogArtikli = (86400 / timerArtikliVoSec) / 24  ' Na sekoj 1h da zapisvua vo log deka raboti
                timerArtikli.Enabled = True
                Console.WriteLine(DateTime.Now.ToString & " Timer ARTIKLI: STARTOVAN!")
            Else
                Console.WriteLine(DateTime.Now.ToString & " Timer ARTIKLI: NE E STARTOVAN! - Konfig-> TimerSec=0")
            End If
        Catch ex As Exception
            Console.WriteLine(DateTime.Now.ToString & " Greska pri INICIJALIZACIJA NA TIMERI:" & ex.Message)
            AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.KONTROLNO, TekParam.TipRezultat.NEUspesno, AlatkaFile.StoPraka.Neodredeno, " Greska pri INICIJALIZACIJA NA TIMERI:" & ex.Message, "TIMER")
        End Try
    End Sub
    'TIMER DOKUMENT: Proverka dali ima Naracka/Dokument - treba da se pravi na sekoj pet(5) sec! Pravi obrakanje do WebShop!
    Private Sub TimerRabotiDokumenti(ByVal source As Object, ByVal e As ElapsedEventArgs)
        timerDokumenti.Enabled = False
        Console.WriteLine(DateTime.Now.ToString & " Timer: RabotamDokumenti!")
        If kolkuTaktoviDokumenti = NAkolkuTaktoviZapisLogDokument Then
            AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.KONTROLNO, TekParam.TipRezultat.Uspesno, AlatkaFile.StoPraka.Dokumenti, "FAKTURA - Povik na Funkcija za proverka na Fakturi()", "MatrixWebShopEXE")
            kolkuTaktoviDokumenti = 0
        Else
            kolkuTaktoviDokumenti += 1
        End If
        Try
            Dim podatok() As Byte
            If Dokumenti.ZemiFaktura(podatok, poraka) Then
                If podatok.Length > 57 Then ' Proverka dali ima nesto vo file-ot (56=Praznen XML)
                    Dim mojaFaktura As New Faktura()
                    If mojaFaktura.ObrabotiFaktura(podatok, poraka) Then
                        If Dokumenti.PratiPorakaZaFaktura("OK", mojaFaktura.GlavaNaDokument.WebShopDokID, poraka) Then
                            Console.WriteLine(DateTime.Now.ToString & " USPESNO OBRABOTENA FAKTURA")
                            AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.ZACUVAJ, TekParam.TipRezultat.Uspesno, AlatkaFile.StoPraka.Dokumenti, "Zapisana Faktura!", mojaFaktura.BrojDokumentCFMA)
                        End If
                    Else
                        If Dokumenti.PratiPorakaZaFaktura("GRESKA: " & poraka, mojaFaktura.GlavaNaDokument.WebShopDokID, "") Then
                            Console.WriteLine(DateTime.Now.ToString & " PratenoPoraka -> Greska vo Obrabotka Faktura" & poraka)
                        Else
                            Console.WriteLine(DateTime.Now.ToString & " NEuspeav_PratiPoraka do x-> Greska vo Obrabotka Faktura" & poraka)
                        End If
                        AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.ZACUVAJ, TekParam.TipRezultat.NEUspesno, AlatkaFile.StoPraka.Dokumenti, "GRESKA Zapis Faktura - " & poraka, mojaFaktura.BrojDokumentCFMA)
                    End If
                End If
            End If
        Catch ex As Exception
            AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.CatchEXCEPTION, TekParam.TipRezultat.NEUspesno, AlatkaFile.StoPraka.Neodredeno, "Greska vo TimerRabotiDokumenti(): " + ex.Message, "MatrixWebShopExe")
        End Try
        timerDokumenti.Enabled = True
    End Sub
    Private Sub TimerRabotiArtikli(ByVal source As Object, ByVal e As ElapsedEventArgs)
        timerArtikli.Enabled = False
        Console.WriteLine(DateTime.Now.ToString & " Timer: RabotamArtikli!")
        Console.WriteLine(DateTime.Now.ToString & " Update ARTIKAL EDEN po EDEN... ")
        Try
            If kolkuTaktoviArtikli = NAkolkuTaktoviZapisLogArtikli Then
                AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.KONTROLNO, TekParam.TipRezultat.Uspesno, AlatkaFile.StoPraka.Artikli, "ARTIKLI - Povik na Funkcija za Update na Artikli - EDEN po EDEN", "MatrixWebShopEXE")
                kolkuTaktoviArtikli = 0
            Else
                kolkuTaktoviArtikli += 1
            End If

            Artikal.UpdateArtikalEdenPoEden(poraka) ' NOVO dodadeno 02.01.2012 MarkoM

            If kolkuTaktoviArtikli = NAkolkuTaktoviZapisLogArtikli Then
                AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.IZMENI, TekParam.TipRezultat.Uspesno, AlatkaFile.StoPraka.Artikli, "UpdateArtikalEdenPoEden() " & poraka, "UPDATE ARTIKLI")
            End If
        Catch ex As Exception
            Console.WriteLine(DateTime.Now.ToString & " Zavrsiv Update za ARTIKAL EDEN! Status: " & poraka & " " & ex.Message)
            AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.CatchEXCEPTION, TekParam.TipRezultat.NEUspesno, AlatkaFile.StoPraka.Artikli, "UpdateArtikalEdenPoEden() " & ex.Message, "UPDATE ARTIKLI")
        End Try
        Console.WriteLine(DateTime.Now.ToString & " Zavrsiv Update za ARTIKAL EDEN! Status: " & poraka)
        Console.WriteLine()
        timerArtikli.Enabled = True
    End Sub
    Private Sub PrikaziKonfiguracija()
        Console.WriteLine("Ver: " & Verzija)
        Console.WriteLine()
        Console.WriteLine("KONFIG: " & "UserBaza=" & TekParam.ZemiKonfig("UserBaza"))
        Console.WriteLine("KONFIG: " & "BazaIme=" & TekParam.ZemiKonfig("BazaIme"))
        Console.WriteLine("KONFIG: " & "Server=" & TekParam.ZemiKonfig("Server"))
        Console.WriteLine()
        Console.WriteLine("KONFIG: " & "TimerDokumentiVoSec=" & TekParam.ZemiKonfig("TimerDokumentiVoSec"))
        Console.WriteLine("KONFIG: " & "TimerArtikliVoSec=" & TekParam.ZemiKonfig("TimerArtikliVoSec"))
        Console.WriteLine()
        Console.WriteLine("WebURL: " & TekParam.HttpURLWebShop)
        Console.WriteLine("GUID(Customer_ID): " & TekParam.GuID)
        Console.WriteLine()
        Console.WriteLine("FOLDER: " & AlatkaFile.PatekaFileWebShop)
        Console.WriteLine()
    End Sub
    Private Sub Console_CancelKeyPress(ByVal sender As Object, ByVal e As ConsoleCancelEventArgs)
        AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.KONTROLNO, TekParam.TipRezultat.Uspesno, AlatkaFile.StoPraka.Neodredeno, "ISKLUCENO EXE! Shutdown!", "EXE")
    End Sub
    Private Sub CurrentDomain_ProcessExit(ByVal sender As Object, ByVal e As EventArgs)
        AlatkaBaza.ZacuvajVoLOG(TekParam.TipAkcija.KONTROLNO, TekParam.TipRezultat.Uspesno, AlatkaFile.StoPraka.Neodredeno, "ISKLUCENO EXE! Shutdown!", "EXE")
    End Sub
End Module
