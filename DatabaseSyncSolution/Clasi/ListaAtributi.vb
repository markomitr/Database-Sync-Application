Public Class ListaAtributi
    Inherits Dictionary(Of String, String)

    Public Sub New()
        MyBase.New()
    End Sub
    Public Sub DodadiParametri(ByVal imeAtribut As String, ByVal vrednostAtribut As String)
        Me.Add(imeAtribut, vrednostAtribut)
    End Sub

End Class
