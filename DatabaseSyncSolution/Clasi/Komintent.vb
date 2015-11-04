Public Class Komintent
    Enum KomintentAtributi
        Customerid
        Clientid
        Clientname
        Address
        City
        Country
        Phone
        Fax
        Email
        Site
        Pin
        Active
    End Enum
    Enum RabatKomintentAtributi
        Guid
        Customererpid
        Discount
    End Enum
    Enum KorisnikKomintent
        Customerid
        Clientid
        Username
        Password
        Active
        Name
        Surname
        Country
        City
        Address
        Phone
        Pricetype
    End Enum
    Public Shared Function UpdateKomintentEDEN(ByVal SifraKup As String, _
                                   ByVal ImeKup As String, _
                                   ByVal Adresa As String, _
                                   ByVal Grad As String, _
                                   ByVal Drzava As String, _
                                   ByVal Tel As String, _
                                   ByVal Fax As String, _
                                   ByVal WebSite As String, _
                                   ByVal DDVBroj As String, _
                                   ByVal Aktiven As Boolean, _
                                   ByRef poraka As String) As Boolean
        If String.IsNullOrEmpty(SifraKup) Or String.IsNullOrEmpty(ImeKup) Then
            poraka = "Greska vo Klucevi: "
            If String.IsNullOrEmpty(SifraKup) Then
                poraka += "Prazen SifraKup!"
            End If
            If String.IsNullOrEmpty(ImeKup) Then
                poraka += "Prazen ImeKup!"
            End If
            Return False
        End If
        Dim atributi As New ListaAtributi

        atributi.Add(KomintentAtributi.Customerid.ToString, TekParam.CustomerID)
        atributi.Add(KomintentAtributi.Clientid.ToString, SifraKup)
        atributi.Add(KomintentAtributi.Clientname, ImeKup)
        atributi.Add(KomintentAtributi.Address.ToString, Adresa)
        atributi.Add(KomintentAtributi.City.ToString, Grad)
        atributi.Add(KomintentAtributi.Country.ToString, Drzava)
        atributi.Add(KomintentAtributi.Phone.ToString, Tel)
        atributi.Add(KomintentAtributi.Site.ToString, WebSite)
        atributi.Add(KomintentAtributi.Pin.ToString, DDVBroj)
        atributi.Add(KomintentAtributi.Active.ToString, Aktiven.ToString)

        Dim urlPrati As String = AlatkaURL.PripremiURLzaPrakanje("Uploadclient.aspx", "", atributi)
        Return AlatkaURL.PratiURL(urlPrati, poraka)
    End Function
    Public Shared Function UpdateRabatKomintentEDEN(ByVal SifraKup As String, ByVal popust As Integer, ByRef poraka As String) As Boolean
        If popust > 100 Then
            popust = -1
        End If
        If String.IsNullOrEmpty(SifraKup) Then
            poraka = "Greska vo Klucevi: "
            If String.IsNullOrEmpty(SifraKup) Then
                poraka += "Prazen SifraKup!"
            End If
            Return False
        End If

        Dim atributi As New ListaAtributi
        atributi.Add(RabatKomintentAtributi.Guid.ToString, TekParam.GuID)
        atributi.Add(RabatKomintentAtributi.Customererpid.ToString, SifraKup)
        atributi.Add(RabatKomintentAtributi.Discount.ToString, popust)

        Dim urlPrati As String = AlatkaURL.PripremiURLzaPrakanje("Updatedefaultdiscount.aspx", "Customer", atributi)
        Return AlatkaURL.PratiURL(urlPrati, poraka)
    End Function
    Public Shared Function UpdateKORISNIKzaKomint(ByVal SifraKup As String, _
                                                   ByVal EmailValiden_UserName As String, _
                                                   ByVal Lozinka As String, _
                                                   ByVal ImeKorisnik As String, _
                                                   ByVal PrezimeKorisnik As String, _
                                                   ByVal Adresa As String, _
                                                   ByVal Grad As String, _
                                                   ByVal Drzava As String, _
                                                   ByVal Tel As String, _
                                                   ByVal Aktiven As Boolean, _
                                                   ByRef poraka As String) As Boolean
        If String.IsNullOrEmpty(SifraKup) Or String.IsNullOrEmpty(EmailValiden_UserName) Or String.IsNullOrEmpty(Lozinka) Then
            poraka = "Greska vo Klucevi: "
            If String.IsNullOrEmpty(SifraKup) Then
                poraka += "Prazen SifraKup!"
            End If
            If String.IsNullOrEmpty(EmailValiden_UserName) Then
                poraka += "Prazen USERNAME/email adresa!"
            End If
            If String.IsNullOrEmpty(Lozinka) Then
                poraka += "Prazna Lozinka!"
            End If
            Return False
        ElseIf Not TekParam.emailExpression.IsMatch(EmailValiden_UserName) Then
            poraka = "Nevaliden UserName-Email:" & EmailValiden_UserName
            Return False
        End If
        Dim atributi As New ListaAtributi
        atributi.Add(KorisnikKomintent.Customerid.ToString, TekParam.CustomerID)
        atributi.Add(KorisnikKomintent.Clientid.ToString, SifraKup)
        atributi.Add(KorisnikKomintent.Username.ToString, EmailValiden_UserName)
        atributi.Add(KorisnikKomintent.Password.ToString, Lozinka)
        atributi.Add(KorisnikKomintent.Name.ToString, ImeKorisnik)
        atributi.Add(KorisnikKomintent.Surname.ToString, PrezimeKorisnik)
        atributi.Add(KorisnikKomintent.Address.ToString, Adresa)
        atributi.Add(KorisnikKomintent.City.ToString, Grad)
        atributi.Add(KorisnikKomintent.Country.ToString, Drzava)
        atributi.Add(KorisnikKomintent.Phone.ToString, Tel)
        atributi.Add(KorisnikKomintent.Active.ToString, Aktiven.ToString)
        Dim urlPrati As String = AlatkaURL.PripremiURLzaPrakanje("Uploadwebuser.aspx", "", atributi)
        Return AlatkaURL.PratiURL(urlPrati, poraka)
    End Function
End Class
