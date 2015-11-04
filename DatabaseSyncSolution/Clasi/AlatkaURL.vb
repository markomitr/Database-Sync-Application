Option Strict On
Imports System.Text
Imports System.Web
Imports System.Net
Imports System.IO
Public Class AlatkaURL



    Public Shared Function PripremiURLzaPrakanje(ByVal funkcija As String, ByVal dodatenURLdel As String, ByVal podatoci As Dictionary(Of String, String)) As String
        Dim krajnoURLzaPrakanje As New StringBuilder

        krajnoURLzaPrakanje.Append(TekParam.HttpURLWebShop) ' So ova se dobiva ovoj String - http://{webshopURL}/ERPUpdate/

        If Not String.IsNullOrEmpty(dodatenURLdel) Then 'Proverka dali dodatniot String e OK - Primer: /Order/
            If dodatenURLdel.StartsWith("/") Then
                dodatenURLdel.Remove(0, 1)
            End If
            If Not dodatenURLdel.EndsWith("/") Then
                dodatenURLdel = dodatenURLdel + "/"
            End If
            krajnoURLzaPrakanje.Append(dodatenURLdel)
        End If

        If Not funkcija.Contains(".aspx") Then
            funkcija = funkcija & ".aspx"
        End If

        krajnoURLzaPrakanje.Append(funkcija) ' Se dodava funkcijata koja se povikuva UPDATEPrice, UpdateDiscount...
        krajnoURLzaPrakanje.Append("?") ' Se dodava vo url-to po konvencija http://{webshopURL}/ERPUpdate/UpdatePrice?

        Dim brojPodatoci As Integer = podatoci.Count
        Dim brojac As Integer = 0
        For Each kluc As KeyValuePair(Of String, String) In podatoci ' Se postavuvaat atributite - pr customerId=asdasdsadasd,price=sadsadasd
            Dim podatokCel As String = kluc.Key.Trim & "=" & kluc.Value.Trim
            krajnoURLzaPrakanje.Append(podatokCel)

            brojac = brojac + 1
            If brojac <> brojPodatoci Then
                krajnoURLzaPrakanje.Append("&")
            End If
        Next

        Return krajnoURLzaPrakanje.ToString()
    End Function
    Public Shared Function PratiURL(ByVal url As String, ByRef rezultat As String, Optional ByVal kakovRezultat As TekParam.TipObrakanje = TekParam.TipObrakanje.Rezultat, Optional ByRef podatokData As Byte() = Nothing) As Boolean
        'Instanca od klasta WebClient za komunikacija so HttpRequests WebRequests
        Dim webClient As New System.Net.WebClient
        Dim resultatPovik As String = ""
        Dim resultatDataPovik() As Byte
        Try
            If kakovRezultat = TekParam.TipObrakanje.ZemamFile Then
                resultatDataPovik = webClient.DownloadData(url)
            Else
            resultatPovik = webClient.DownloadString(url)
            End If
            '<---------------------END ------------------------------->
            If kakovRezultat = TekParam.TipObrakanje.ZemamFile Then
                podatokData = resultatDataPovik
                Return True
            End If
            If kakovRezultat = TekParam.TipObrakanje.ZemamPodatok Then
                rezultat = resultatPovik
                Return True
            End If
            If resultatPovik.ToUpper().Contains("OK") AndAlso resultatPovik.ToUpper.StartsWith("OK") Then
                rezultat = resultatPovik
                Return True
            Else
                rezultat = "GRESKA vo WebShopBaza: " & resultatPovik
                Return False
            End If

        Catch ex As Exception
            rezultat = "GRESKA vo URL: " & url & " PorakaAPP: " & ex.Message
            Return False
        End Try


    End Function
    Public Shared Function ZemiFTPurl(ByVal url As String, ByRef rezultat As String) As Boolean
        Return PratiURL(url, rezultat, TekParam.TipObrakanje.ZemamPodatok)
    End Function
End Class
