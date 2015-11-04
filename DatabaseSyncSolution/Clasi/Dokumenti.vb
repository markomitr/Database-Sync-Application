Public Class Dokumenti
    Enum FakturaAtributi
        Customerid
        Status
        Version
    End Enum
    Enum PorakaWebShop
        Customerid
        Messageid
        Message
    End Enum
    Public Shared Function ZemiFaktura(ByRef podatokData() As Byte, ByRef poraka As String) As Boolean
        Dim atributi As New ListaAtributi
        atributi.Add(FakturaAtributi.Customerid.ToString, TekParam.CustomerID)
        atributi.Add("date", "")
        atributi.Add(FakturaAtributi.Status.ToString, "8")
        atributi.Add(FakturaAtributi.Version.ToString, "2")

        Dim urlPrati As String = AlatkaURL.PripremiURLzaPrakanje("GetXmlInvoices.aspx", "", atributi)
        Return AlatkaURL.PratiURL(urlPrati, poraka, TekParam.TipObrakanje.ZemamFile, podatokData)
    End Function
    Public Shared Function PratiPorakaZaFaktura(ByVal porakaDoWebShop As String, ByVal dokId As String, ByRef poraka As String) As Boolean
        Dim atributi As New ListaAtributi
        atributi.Add(PorakaWebShop.Customerid.ToString, TekParam.CustomerID)
        atributi.Add(PorakaWebShop.Messageid.ToString, IIf(dokId Is Nothing, "", dokId))
        atributi.Add(PorakaWebShop.Message.ToString, porakaDoWebShop)
        Dim urlPrati As String = AlatkaURL.PripremiURLzaPrakanje("SendCustomMessage.aspx", "CustomMessage/", atributi)
        Return AlatkaURL.PratiURL(urlPrati, poraka)
    End Function
End Class
