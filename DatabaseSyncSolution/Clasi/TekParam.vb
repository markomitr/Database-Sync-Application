Imports System.Text.RegularExpressions
Imports System.Xml
Public Class TekParam

    Private Const HTTP_PROTOKOL As String = "HTTP://"
    Private Const FTP_PROTOKOL As String = "FTP://"
    'Za TEST="matrixtest.autodelovi.rs"
    Private Const URL_WEB_SHOP As String = "erpservices.*********.net"
    Private Const HTTP_URL_WEBSHOP As String = HTTP_PROTOKOL & URL_WEB_SHOP & "/"
    Private Const CUSTOMER_ID As String = "c94e2bc6-****-4233-b180-****" '!!! ID koe e ednoznacno za MATRIX vo WebSHOP - Kluc za komunikacija so WebShop 

    Private Shared CnStringFIKSEN = "Server=(local);Database=******;User Id=****;Password=****;"

    Public Shared emailExpression As New Regex("^[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})$")

    Public Enum TipObrakanje
        ZemamPodatok
        ZemamFile
        Rezultat
    End Enum
    Public Enum TipAkcija
        PROVERKA
        ZACUVAJ
        IZMENI
        BRISI
        PODIGNI
        KONTROLNO
        CatchEXCEPTION
    End Enum
    Public Enum TipRezultat
        Uspesno
        NEUspesno
    End Enum
    Public Shared Function ZemiKonfig(ByVal sPole As String) As String
        Try


            Dim XMLCitac As New XmlTextReader("wsKonfig.xml")
            While XMLCitac.Read
                XMLCitac.MoveToContent()
                If XMLCitac.NodeType = XMLCitac.NodeType.Element And XMLCitac.Name = sPole Then
                    Dim rez As String = XMLCitac.ReadInnerXml
                    XMLCitac.Close()
                    Return rez
                End If
            End While
            XMLCitac.Close()
            Return ""
        Catch ex As Exception

            Return ""
        End Try
    End Function
    Private Shared Function SrediCnString()
        Dim serverIme As String = ZemiKonfig("Server")
        Dim bazaIme As String = ZemiKonfig("BazaIme")
        Dim user As String = ZemiKonfig("UserBaza")
        Dim pass As String = ZemiKonfig("Pass")
        If String.IsNullOrEmpty(serverIme) Or _
           String.IsNullOrEmpty(bazaIme) Or _
           String.IsNullOrEmpty(user) Or _
           String.IsNullOrEmpty(pass) Then
            Return CnStringFIKSEN

        Else
            Return "Server=" & serverIme & ";Database=" & bazaIme & ";User Id=" & user & ";Password=" & pass & ";"
        End If

    End Function

    Public Shared Function ZemiStr(ByVal ImePole As String) As String
        Dim cmdKonfig As New SqlClient.SqlCommand()
        Dim CitacKonfig As SqlClient.SqlDataReader
        Dim sqlCn As New SqlClient.SqlConnection()
        Dim wRez As String = ""
        Dim Otvoren As Boolean = False

        Try
            sqlCn.ConnectionString = KonekcijaString
            With cmdKonfig
                .Connection = sqlCn
                .CommandType = CommandType.Text
                .CommandText = "SELECT " + ImePole + " FROM Konfig"
                sqlCn.Open()
                CitacKonfig = .ExecuteReader
                sqlCn.Close()
            End With

            With CitacKonfig
                If .Read Then
                    wRez = Konv.ObjVoStr(.Item(ImePole))
                End If
                .Close()
            End With
        Catch ex As Exception
            sqlCn.Close()
        End Try
        Return wRez
    End Function

    Public Shared Function ZemiInt(ByVal ImePole As String) As Integer
        Dim cmdKonfig As New SqlClient.SqlCommand()
        Dim CitacKonfig As SqlClient.SqlDataReader
        Dim sqlCn As New SqlClient.SqlConnection()
        Dim wRez As Integer = 0
        Dim Otvoren As Boolean = False

        Try
            sqlCn.ConnectionString = KonekcijaString
            With cmdKonfig
                .Connection = sqlCn
                .CommandType = CommandType.Text
                ''''' Bese do Okt 22, 2005   .CommandText = "SELECT * FROM Konfig"
                .CommandText = "SELECT " + ImePole + " FROM Konfig"
                sqlCn.Open()
                CitacKonfig = .ExecuteReader()
            End With

            With CitacKonfig
                If .Read Then
                    wRez = Konv.ObjVoInt(.Item(ImePole))
                End If
                .Close()
            End With

            sqlCn.Close()
        Catch ex As Exception
            sqlCn.Close()
        End Try
        Return wRez
    End Function
    Public Shared ReadOnly Property HttpURLWebShop() As String
        Get
            Return HTTP_URL_WEBSHOP
        End Get
    End Property
    Public Shared ReadOnly Property CustomerID() As String
        Get
            Return CUSTOMER_ID
        End Get
    End Property

    Public Shared ReadOnly Property GuID() As String
        Get
            Return CUSTOMER_ID
        End Get
    End Property
    Public Shared ReadOnly Property FtpProtokol() As String
        Get
            Return FTP_PROTOKOL
        End Get
    End Property
    Public Shared ReadOnly Property KonekcijaString() As String
        Get
            Return SrediCnString()
        End Get
    End Property
End Class
