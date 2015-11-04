Imports System.Data.OleDb
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Reflection
Imports System.Threading.Thread
Imports System.Globalization
Imports System.Xml
Imports System.Text
Imports System.Security.AccessControl

Public Class AlatkaFile
    Private Const PatekaFolderGlaven = "C:\IB\WebShopUpDown\"
    Private Const ImeFolderWebShopGlaven = "WebShopUpDown"
    Private Const ImeFolderZaFileOdWebShop = "WebShopApp"
    Private Const ImeFolderZaFileOdCFMA = "CFMAApp"
    Private Const ArhivaImeFolderZaFileOdWebShop = "Arhiva_WebShopApp"
    Private Const ArhivaImeFolderZaFileOdCFMA = "Arhiva_CFMAApp"
    Enum TipFajlovi
        Excel_XLS
        Excel_XLSX
        XML
        Txt
    End Enum
    Enum KojPraka
        WebShop
        CFMA
    End Enum
    Public Enum StoPraka
        Artikli
        Komintenti
        Dokumenti
        Neodredeno
    End Enum
    Public Shared ReadOnly Property PatekaFileWebShop() As String
        Get
            Return PatekaFolderGlaven & ImeFolderZaFileOdWebShop
        End Get
    End Property
    Public Shared ReadOnly Property PatekaFileCFMA() As String
        Get
            Return PatekaFolderGlaven & ImeFolderZaFileOdCFMA
        End Get
    End Property
    Public Shared ReadOnly Property PatekaARHIVAFileWebShop() As String
        Get
            Return PatekaFolderGlaven & ArhivaImeFolderZaFileOdWebShop
        End Get
    End Property
    Public Shared ReadOnly Property PatekaARHIVAFileCFMA() As String
        Get
            Return PatekaFolderGlaven & ArhivaImeFolderZaFileOdCFMA
        End Get
    End Property
    Private Shared Sub ProveriFolderi()

        If (Not Directory.Exists(PatekaFolderGlaven)) Then
            System.IO.Directory.CreateDirectory(PatekaFolderGlaven)
        End If

        If (Not Directory.Exists(PatekaFileWebShop)) Then
            System.IO.Directory.CreateDirectory(PatekaFileWebShop)
        End If

        'If (Not Directory.Exists(PatekaFileCFMA)) Then
        '    System.IO.Directory.CreateDirectory(PatekaFileCFMA)
        'End If

        'If (Not Directory.Exists(PatekaARHIVAFileWebShop)) Then
        '    System.IO.Directory.CreateDirectory(PatekaARHIVAFileWebShop)
        'End If

        'If (Not Directory.Exists(PatekaARHIVAFileCFMA)) Then
        '    System.IO.Directory.CreateDirectory(PatekaARHIVAFileCFMA)
        'End If


    End Sub
    Public Shared Function ZacuvajFile(ByVal ds As DataSet, ByVal kakovFile As TipFajlovi, ByVal kojGoPraka As KojPraka, ByVal stoPusta As StoPraka, _
                                        ByRef pateka As String, ByRef poraka As String) As Boolean
        If ds Is Nothing Then
            poraka = "ZacuvajFile: Prazen DataSet!"
            Return False
        End If
        ProveriFolderi() ' Se pravi proverka dali postojat site folderi
        Try
            Dim patekaFile = VratiPatekaiImeZaFile(kakovFile, kojGoPraka, stoPusta)
            pateka = patekaFile
            If kakovFile = TipFajlovi.Excel_XLS Or kakovFile = TipFajlovi.Excel_XLSX Then
                If Not ExportToExcelNAJNOV(patekaFile, ds, "", poraka) Then
                    Return False
                End If
            ElseIf kakovFile = TipFajlovi.XML Then
                ds.WriteXml(patekaFile)
            End If
            Return True
        Catch ex As Exception
            poraka = "ZacuvajFile: " & pateka & " GRESKA:" & ex.Message
            Return False
        End Try
    End Function
    Public Shared Function ZacuvajFile(ByVal fileByteArray As Byte(), ByVal kakovFile As TipFajlovi, ByVal kojGoPraka As KojPraka, ByVal stoPusta As StoPraka, _
                                        ByRef pateka As String, ByRef poraka As String) As Boolean
        If fileByteArray Is Nothing Then
            poraka = "ZacuvajFile: Prazen FileByteArray!"
            Return False
        End If
        ProveriFolderi() ' Se pravi proverka dali postojat site folderi

        Try
            Dim patekaFile = VratiPatekaiImeZaFile(kakovFile, kojGoPraka, stoPusta, False, True)
            pateka = patekaFile
            ProveriDaliPostoiFolder(pateka, True)
            Dim oFileStream As New System.IO.FileStream(patekaFile, System.IO.FileMode.Create)
            oFileStream.Write(fileByteArray, 0, fileByteArray.Length)
            oFileStream.Close()
            Return True
        Catch ex As Exception
            poraka = "ZacuvajFile: " & pateka & " GRESKA:" & ex.Message
            Return False
        End Try

    End Function
    Private Shared Function VratiPatekaiImeZaFile(ByVal kakovFile As TipFajlovi, ByVal kojGoPraka As KojPraka, ByVal stoPusta As StoPraka, Optional ByVal kratkoIme As Boolean = False, Optional ByVal daliPoDatumFolderi As Boolean = False) As String
        Dim pateka = ""
        Dim imeFILE As String = ""
        Dim nastavkaFILE As String = ""
        Dim kojNiPraka As String = ""
        Dim stoNiPraka As String = ""

        If kakovFile = TipFajlovi.Excel_XLS Then
            nastavkaFILE = ".xls"
        ElseIf kakovFile = TipFajlovi.Excel_XLSX Then
            nastavkaFILE = ".xlsx"
        ElseIf kakovFile = TipFajlovi.XML Then
            nastavkaFILE = ".xml"
        ElseIf kakovFile = TipFajlovi.Txt Then
            nastavkaFILE = ".txt"
        End If

        If kojGoPraka = KojPraka.CFMA Then
            kojNiPraka = "CFMA"
            pateka = PatekaFileCFMA
        ElseIf kojGoPraka = KojPraka.WebShop Then
            kojNiPraka = "WebShop"
            pateka = PatekaFileWebShop
        End If

        If stoPusta = StoPraka.Artikli Then
            stoNiPraka = "Artikli"
        ElseIf stoPusta = StoPraka.Komintenti Then
            stoNiPraka = "Komintenti"
        ElseIf stoPusta = StoPraka.Dokumenti Then
            stoNiPraka = "Dokumenti"
        ElseIf stoPusta = StoPraka.Neodredeno Then
            stoNiPraka = "Neodredeno"
        End If
        If kratkoIme Then
            imeFILE = stoNiPraka & "_" & kojNiPraka & "_" & DateTime.Now.ToString("yyyy-MM-dd") & nastavkaFILE
        Else
            imeFILE = stoNiPraka & "_" & kojNiPraka & "_" & DateTime.Now.ToString("yyyy-MM-dd-hhmmss") & nastavkaFILE
        End If

        If daliPoDatumFolderi Then
            pateka = pateka & "\" & stoNiPraka & "\" & Konv.VratiDatumVoString & "\" & imeFILE
        Else
            pateka = pateka & "\" & stoNiPraka & "\" & imeFILE
        End If


        Return pateka
    End Function
    Private Shared Function ExportToExcelNAJNOV(ByVal a_sFilename As String, ByVal a_sData As DataSet, ByVal a_sFileTitle As String, ByRef a_sErrorMessage As String) As Boolean
        a_sErrorMessage = String.Empty
        Dim bRetVal As Boolean = False
        Dim dsDataSet As DataSet = Nothing
        Try
            dsDataSet = a_sData
            Dim Oldci As CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
            System.Threading.Thread.CurrentThread.CurrentCulture = New CultureInfo("en-us")

            Dim xlObject As Excel.Application = Nothing
            Dim xlWB As Excel.Workbook = Nothing
            Dim xlSh As Excel.Worksheet = Nothing
            Dim rg As Excel.Range = Nothing
            Try
                xlObject = New Excel.Application()
                xlObject.AlertBeforeOverwriting = False
                xlObject.DisplayAlerts = False

                ''This Adds a new woorkbook, you could open the workbook from file also
                xlWB = xlObject.Workbooks.Add()
                xlWB.SaveAs(a_sFilename, 56, Missing.Value, Missing.Value, Missing.Value, Missing.Value, _
                Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value)

                xlSh = DirectCast(xlObject.ActiveWorkbook.ActiveSheet, Excel.Worksheet)

                'Dim sUpperRange As String = "A1"
                'Dim sLastCol As String = "AQ"
                'Dim sLowerRange As String = sLastCol + (dsDataSet.Tables(0).Rows.Count + 1).ToString()

                For j As Integer = 0 To dsDataSet.Tables(0).Columns.Count - 1
                    xlSh.Cells(1, j + 1) = _
                        dsDataSet.Tables(0).Columns(j).ToString()
                    xlSh.Cells(1, j + 1).Font.Bold = True
                Next

                For i As Integer = 1 To dsDataSet.Tables(0).Rows.Count
                    For j As Integer = 0 To dsDataSet.Tables(0).Columns.Count - 1
                        xlSh.Cells(i + 1, j + 1) = _
                            dsDataSet.Tables(0).Rows(i - 1)(j).ToString()
                    Next
                Next
                xlSh.Columns.AutoFit()
                'rg = xlSh.Range(sUpperRange, sLowerRange)
                'rg.Value2 = GetData(dsDataSet.Tables(0))

                'xlSh.Range("A1", sLastCol & "1").Font.Bold = True
                'xlSh.Range("A1", sLastCol & "1").HorizontalAlignment = XlHAlign.xlHAlignCenter
                'xlSh.Range(sUpperRange, sLowerRange).EntireColumn.AutoFit()

                If String.IsNullOrEmpty(a_sFileTitle) Then
                    xlObject.Caption = "untitled"
                Else
                    xlObject.Caption = a_sFileTitle
                End If

                xlWB.Save()
                bRetVal = True
            Catch ex As System.Runtime.InteropServices.COMException
                If ex.ErrorCode = -2147221164 Then
                    a_sErrorMessage = "Error in export: Please install Microsoft Office (Excel) to use the Export to Excel feature."
                ElseIf ex.ErrorCode = -2146827284 Then
                    a_sErrorMessage = "Error in export: Excel allows only 65,536 maximum rows in a sheet."
                Else
                    a_sErrorMessage = (("Error in export: " & ex.Message) + Environment.NewLine & " Error: ") + ex.ErrorCode
                End If
            Catch ex As Exception
                a_sErrorMessage = "Error in export: " & ex.Message
            Finally
                Try
                    If xlWB IsNot Nothing Then
                        xlWB.Close(Nothing, Nothing, Nothing)
                    End If
                    xlObject.Workbooks.Close()
                    xlObject.Quit()
                Catch
                End Try
                xlSh = Nothing
                xlWB = Nothing
                xlObject = Nothing
                ' force final cleanup!
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        Catch ex As Exception
            a_sErrorMessage = "Error in export: " & ex.Message
        End Try

        Return bRetVal
    End Function
    Public Shared Function RenameNaFile(ByRef pateka As String, ByVal novoIme As String) As Boolean
        Try
            If File.Exists(pateka) Then
                My.Computer.FileSystem.RenameFile(pateka, novoIme)
                pateka = ZemiPatekaFolderODPateka(pateka) & novoIme
                Return True
            End If
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Shared Function ZemiImeFileOdPateka(ByVal pateka As String) As String
        If Not String.IsNullOrEmpty(pateka) Then
            Dim pom() As String = pateka.Split("\")
            Return pom(pom.Length - 1)
        End If
        Return ""
    End Function
    Public Shared Function ZemiPatekaFolderODPateka(ByVal pateka As String) As String
        If Not String.IsNullOrEmpty(pateka) Then
            Dim pom() As String = pateka.Split("\")
            Return pateka.Replace(pom(pom.Length - 1), "")
        End If
        Return ""
    End Function
    Public Shared Function DodadiPorakaVoXml(ByVal poraka As String, ByVal xmlPateka As String) As Boolean
        Try
            Dim writeStart As Boolean
            If Not IO.File.Exists(xmlPateka) Then Return False

            Dim xmlFile As IO.FileStream = New IO.FileStream(xmlPateka, IO.FileMode.Append)
            Dim myXmlTextWriter As New XmlTextWriter(xmlFile, System.Text.Encoding.Default)
            If writeStart Then
                With myXmlTextWriter
                    .Formatting = Formatting.Indented
                    .Indentation = 3
                    .IndentChar = CChar(" ")
                    .WriteStartDocument()
                    .WriteStartElement("Proformas3")
                    .WriteEndElement()
                End With
            End If
            myXmlTextWriter.Close()
            AddXmlData(xmlPateka, "Proforma", poraka, "Proformas3")
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Shared Sub AddXmlData(ByVal xmlfile As String, ByVal voKojElDaVnesam As String, ByVal poraka As String, ByVal rootElement As String)
        Try
            Dim myXmlDocument As New XmlDocument
            Dim myNodes, myChildren As XmlNodeList
            Dim node(3) As XmlNode

            myXmlDocument.Load(xmlfile)
            myNodes = myXmlDocument.GetElementsByTagName(rootElement)

            For Each n As XmlNode In myNodes
                If n.Name = rootElement Then
                    myChildren = n.ChildNodes
                    For Each n1 As XmlNode In myChildren
                        If n1.Name = voKojElDaVnesam Then
                            node(1) = myXmlDocument.CreateNode(System.Xml.XmlNodeType.Element, "PorakaCFMA", "")
                            node(1).InnerText = poraka
                            n1.AppendChild(node(1))
                            myXmlDocument.Save(xmlfile)
                            Exit Sub
                        End If
                    Next
                    node(0) = myXmlDocument.CreateNode(XmlNodeType.Element, voKojElDaVnesam, "")
                    node(1) = myXmlDocument.CreateNode(System.Xml.XmlNodeType.Element, "PorakaCFMA", "")
                    node(1).InnerText = poraka
                    node(0).AppendChild(node(1))
                    n.AppendChild(node(0))
                    myXmlDocument.Save(xmlfile)
                End If
            Next
        Catch ex As Exception
            'Nema potreba da se hendlira!
            'Ova e samo informativen podatok koj se zapisuva vo filot(XML)
        End Try
    End Sub
    Public Shared Function ZapisiVoFile(ByVal tekst As String, ByVal kojGoPraka As KojPraka, ByVal stoZapisuvam As StoPraka, ByRef poraka As String, Optional ByVal infoZaZapis As Boolean = False) As Boolean
        Dim outfile As StreamWriter
        Try
            Dim patekaFile As String = ""
            Dim sb As New StringBuilder
            patekaFile = VratiPatekaiImeZaFile(TipFajlovi.Txt, kojGoPraka, stoZapisuvam, True)
            ProveriDaliPostoiFolder(patekaFile, True)
            If infoZaZapis Then
                sb.AppendLine("")
                sb.AppendLine("<==================[ " & DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") & " ]==================>")
                sb.AppendLine(tekst)
                sb.AppendLine("<==========================[ KRAJ ]==========================>")
            Else
                sb.AppendLine(tekst)
            End If
            outfile = New StreamWriter(patekaFile, True)
            outfile.WriteLine(sb.ToString)
            outfile.Close()
            Return True
        Catch ex As Exception
            poraka = "Greska ZapisVoLog: " & ex.Message
            If Not outfile Is Nothing Then
                outfile.Close()
            End If
            Return False
        End Try
    End Function
    Private Shared Function ProveriDaliPostoiFolder(ByVal pateka As String, ByVal kreirajAkoNePostoi As Boolean) As Boolean
        If pateka.EndsWith(".xls") Or _
               pateka.EndsWith(".xlsx") Or _
               pateka.EndsWith(".txt") Or _
               pateka.EndsWith(".xml") Then
            Dim patekaArr() As String = pateka.Split("\")
            Dim patekaFolder As String = ""
            For ii As Integer = 0 To patekaArr.Length - 2 'Vazno -2
                If ii = 0 Then
                    patekaFolder = patekaArr(ii)
                Else
                    patekaFolder = patekaFolder & "\" & patekaArr(ii)
                End If
            Next
            pateka = patekaFolder
        End If

        If (Not Directory.Exists(pateka)) Then 
            If kreirajAkoNePostoi Then
                System.IO.Directory.CreateDirectory(pateka)
            End If
            Return False
        Else
            Return True
        End If

        'detect whether its a directory or file
        'Dim attr As FileAttributes = File.GetAttributes(pateka)
        'If ((attr & FileAttributes.Directory) <> FileAttributes.Directory) Then
        '    Dim patekaArr() As String = pateka.Split("\")
        '    Dim patekaFolder As String = ""
        '    For ii As Integer = 0 To patekaArr.Length - 2 'Vazno -2
        '        If ii = 0 Then
        '            patekaFolder = patekaArr(ii)
        '        Else
        '            patekaFolder = patekaFolder & "\" & patekaArr(ii)
        '        End If

        '    Next
        '    pateka = patekaFolder
        'End If

    End Function
End Class
