Imports System.Net
Imports System.IO

Public Class AlatkaFTP
    Private usrName As String
    Private pass As String
    Private sessionID As String
    Private ftpUrl As String
    Private daliSeVoRed = True
    Public Sub New(ByVal FtpAdresaSiteInfo As String)
        daliSeVoRed = Me.SrediFTPadresa(FtpAdresaSiteInfo)
    End Sub
    Private Function SrediFTPadresa(ByVal ftpString As String) As Boolean
        'Tuka go pravam parsiranjeto na FTPAdresata
        Dim ftpGrub_Cel() As String = ftpString.Split("@")

        If ftpGrub_Cel.Length < 2 Then
            Return False
        End If

        Dim celaNizaFtp() As String = ftpGrub_Cel(1).Split("\")
        sessionID = celaNizaFtp(celaNizaFtp.Length - 1) 'Ja zemame sesijata(string)

        If SesijaId.Length < 1 Then
            Return False
        End If

        ftpGrub_Cel(1) = ftpGrub_Cel(1).Replace("\", "/")
        ftpUrl = "ftp://" & ftpGrub_Cel(1) & "/" ' Kraen FtpString za Povik

        Dim FtpUserPass() As String = (ftpGrub_Cel(0).Replace("ftp://", "")).Split(":")
        If FtpUserPass.Length < 2 Then
            Return False
        End If

        usrName = FtpUserPass(0)
        pass = FtpUserPass(1)

        Return True
    End Function
    Public Function PratiFileFtp(ByVal imeFile As String, ByVal patekaFajl As String, ByRef poraka As String) As Boolean
        Try
            'Kreiram objekt za komunikacija so FTP
            Dim request As FtpWebRequest = DirectCast(FtpWebRequest.Create(Me.FtpAdresaZaPovik & imeFile), FtpWebRequest)

            'Dim request As FtpWebRequest = DirectCast(FtpWebRequest.Create("ftp://ftp.autodelovi.rs/" & imeFile), FtpWebRequest)

            request.Method = WebRequestMethods.Ftp.UploadFile
            request.Credentials = New NetworkCredential(Me.UserName, Me.Password)
            request.UsePassive = False ' Vazno
            request.UseBinary = True
            request.KeepAlive = False
            request.ReadWriteTimeout = 10000000
            request.Timeout = 10000000

            If Not String.IsNullOrEmpty(patekaFajl) AndAlso Not patekaFajl.EndsWith("\") Then
                patekaFajl = patekaFajl & "\"
            End If
            'Load file
            Dim stream As FileStream = File.OpenRead(patekaFajl & imeFile)
            Dim buffer As Byte() = New Byte(CInt(stream.Length - 1)) {}
            stream.Read(buffer, 0, buffer.Length)
            stream.Close()

            'Upload file
            Dim reqStream As Stream = request.GetRequestStream()
            reqStream.Write(buffer, 0, buffer.Length)
            reqStream.Close()

            Dim response As FtpWebResponse = CType(request.GetResponse(), FtpWebResponse)
            If response.StatusCode <> 226 Then
                poraka = "GRESKA UPLOAD FILE: " & response.StatusDescription
                response.Close()
                Return False
            End If
            response.Close()

            poraka = "USPESNO - UPLOAD FILE!" & vbCrLf & " FTP.url: " & Me.FtpAdresaZaPovik
            Return True
        Catch ex As Exception
            poraka = "GRESKA UPLOAD FILE!" & vbCrLf & " FTP.url: " & Me.FtpAdresaZaPovik & vbCrLf & " FilePateka: " & patekaFajl & vbCrLf & " ERORR:" & ex.Message
            Return False
        End Try
    End Function
    Public ReadOnly Property UserName() As String
        Get
            Return usrName
        End Get
    End Property
    Public ReadOnly Property Password() As String
        Get
            Return pass
        End Get
    End Property
    Public ReadOnly Property SesijaId() As String
        Get
            Return sessionID
        End Get
    End Property
    Public ReadOnly Property FtpAdresaZaPovik() As String
        Get
            Return ftpUrl
        End Get
    End Property
    Public ReadOnly Property DaliEVoRed() As Boolean
        Get
            Return daliSeVoRed
        End Get
    End Property
    Public ReadOnly Property SesijaZaFTPadresa()
        Get
            Return "(S(" & Me.SesijaId & "))"
        End Get
    End Property
End Class
