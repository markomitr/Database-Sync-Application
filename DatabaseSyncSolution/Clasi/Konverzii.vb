Option Strict On

Public Class Konv
    Public Shared DefDatum As Date

    Public Shared Function TxtVoDec(ByVal Txt As String) As Decimal
        Dim Rez As Decimal
        Try
            If Txt = "" Then        'Zaradi efikasnost (mnogu e spor exception handlingot)
                Return 0
            End If
            Rez = Convert.ToDecimal(Txt)
        Catch
            Return 0
        End Try
        Return Rez
    End Function

    Public Shared Function TxtVoSingle(ByVal Txt As String) As Single
        Dim Rez As Single
        Try
            If Txt = "" Then
                Return 0
            End If
            Rez = Convert.ToSingle(Txt)
        Catch
            Return 0
        End Try
        Return Rez
    End Function

    Public Shared Function TxtVoDouble(ByVal Txt As String) As Double
        Dim Rez As Double
        Try
            If Txt = "" Then
                Return 0
            End If
            Rez = Convert.ToDouble(Txt)
        Catch
            Return 0
        End Try
        Return Rez
    End Function

    Public Shared Function TxtVoInt(ByVal Txt As String) As Integer
        Dim RezDec As Decimal
        Dim Rez As Integer
        Try
            If Txt = "" Then
                Return 0
            End If
            Rez = Convert.ToInt32(Txt)
        Catch
            Try
                RezDec = Math.Round(Convert.ToDecimal(Txt), 0)
                Rez = Convert.ToInt32(RezDec)
            Catch
                Return 0
            End Try
        End Try
        Return Rez
    End Function

    Public Shared Function TxtVoInt16(ByVal Txt As String) As Int16
        Dim RezDec As Decimal
        Dim Rez As Int16
        Try
            If Txt = "" Then
                Return 0
            End If
            Rez = Convert.ToInt16(Txt)
        Catch
            Try
                RezDec = Math.Round(Convert.ToDecimal(Txt), 0)
                Rez = Convert.ToInt16(RezDec)
            Catch
                Return 0
            End Try
        End Try
        Return Rez
    End Function

    Public Shared Function DateTimeVoDat(ByVal txt As String) As String
        Dim i As Integer
        i = InStr(txt, " ")
        If i > 0 Then
            Return Left(txt, i - 1)
        Else
            Return txt
        End If
    End Function

    Public Shared Function IntVoStr(ByVal txt As Integer) As String
        Dim Rez As String = ""

        Try
            Rez = Convert.ToString(txt)
        Catch

        End Try
        If Rez.Trim = "0" Then Rez = " "
        Return Rez
    End Function

    Public Shared Function ObjVoInt(ByVal Sto As Object) As Integer
        Dim Rez As Integer
        If Sto Is DBNull.Value Then
            Return 0
        Else
            Try
                Rez = Convert.ToInt32(Sto)
                Return Rez
            Catch
                Return 0
            End Try
        End If
    End Function

    Public Shared Function ObjVoInt64(ByVal Sto As Object) As Int64
        Dim Rez As Int64
        If Sto Is DBNull.Value Then
            Return 0
        Else
            Try
                Rez = Convert.ToInt64(Sto)
                Return Rez
            Catch
                Return 0
            End Try
        End If
    End Function

    Public Shared Function ObjVoByte(ByVal Sto As Object) As Byte
        Dim Rez As Byte
        If Sto Is DBNull.Value Then
            Return 0
        Else
            Try
                Rez = Convert.ToByte(Sto)
                Return Rez
            Catch
                Return 0
            End Try
        End If
    End Function

    Public Shared Function ObjVoDec(ByVal Sto As Object) As Decimal
        Dim Rez As Decimal
        If Sto Is DBNull.Value Then
            Return 0
        Else
            Try
                Rez = Convert.ToDecimal(Sto)
                Return Rez
            Catch
                Return 0
            End Try
        End If
    End Function

    Public Shared Function ObjVoDouble(ByVal Sto As Object) As Double
        Dim Rez As Double
        If Sto Is DBNull.Value Then
            Return 0
        Else
            Try
                Rez = Convert.ToDouble(Sto)
                Return Rez
            Catch
                Return 0
            End Try
        End If
    End Function

    Public Shared Function ObjVoStr(ByVal Sto As Object) As String
        If Sto Is DBNull.Value Then
            Return ""
        Else
            Return Convert.ToString(Sto)
        End If
    End Function

    Public Shared Function ObjVoStrSoTocka(ByVal Sto As Object) As String
        If Sto Is DBNull.Value Then
            Return ""
        Else
            Return Replace(Convert.ToString(Sto), ",", ".")
        End If
    End Function

    Public Shared Function ObjVoDate(ByVal Sto As Object) As Date
        If Sto Is DBNull.Value Then
            Return DefDatum
        Else
            Try
                Return Convert.ToDateTime(Sto)
            Catch
                Return DefDatum
            End Try
        End If
    End Function

    Public Shared Function TxtVoDate(ByVal Sto As String) As Date
        If Sto Is Nothing OrElse Sto = "" Then
            Return DefDatum
        Else
            Try
                Return Convert.ToDateTime(Sto)
            Catch
                Return DefDatum
            End Try
        End If
    End Function

    Public Shared Function ObjVoSingle(ByVal Sto As Object) As Single
        Dim Rez As Single
        If Sto Is DBNull.Value Then
            Return 0
        Else
            Try
                Rez = Convert.ToSingle(Sto)
                Return Rez
            Catch
                Return 0
            End Try
        End If
    End Function

    Public Shared Function TxtVoLong(ByVal Txt As String) As Long
        Dim RezDec As Decimal
        Dim Rez As Long
        Try
            If Txt = "" Then
                Return 0
            End If
            Rez = Convert.ToInt64(Txt)
        Catch
            Try
                RezDec = Math.Round(Convert.ToDecimal(Txt), 0)
                Rez = Convert.ToInt64(RezDec)
            Catch
                Return 0
            End Try
        End Try
        Return Rez
    End Function

    Public Shared Function LongVoStr(ByVal txt As Long) As String
        Dim Rez As String = ""

        Try
            Rez = Convert.ToString(txt)
        Catch

        End Try
        If Rez.Trim = "0" Then Rez = " "
        Return Rez
    End Function
    Public Shared Function StrSamoBrojkiVoDate(ByVal txt As String) As DateTime
        Dim godina, mesec, den, cas, min, sek As Integer
        If txt.Length = 14 Then
            godina = Konv.TxtVoInt(txt.Substring(0, 4))
            mesec = Konv.TxtVoInt(txt.Substring(4, 2))
            den = Konv.TxtVoInt(txt.Substring(6, 2))
            cas = Konv.TxtVoInt(txt.Substring(8, 2))
            min = Konv.TxtVoInt(txt.Substring(10, 2))
            sek = Konv.TxtVoInt(txt.Substring(12, 2))
            Return New DateTime(godina, mesec, den, cas, min, sek)
        End If
        Return DateTime.Today.AddYears(-50)
    End Function
    Public Shared Function VratiSesijaStringZaUrl(ByVal sesijaID As String) As String
        Return "(S(" & sesijaID & "))"
    End Function
    Public Shared Function VratiDatumVoString() As String
        Dim den, mesec, godina As String

        den = Date.Today.Day.ToString
        mesec = Date.Today.Month.ToString
        godina = Date.Today.Year.ToString

        Return den & "-" & mesec & "-" & godina
    End Function
End Class