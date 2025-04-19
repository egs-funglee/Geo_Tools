Public Class CNV 'store and return cnv format
    Public day As String = "01/01/1900"
    Public fix As Integer = 0
    Public g_cosine As Double = 0.0#
    Public g_sine As Double = 0.0#
    Public gyro As Double = 0.0#
    Public hhmmsss As String = "00:00:00.000"
    Public rec As Integer = 0
    Public tow_bearing As Single = -999.9F
    Public wd As Double = 0.0#
    Public x As Double = 0.0#
    Public y As Double = 0.0#

    Public Function Sectime(ByVal digit As Byte) As String 'time in second sssss.ss
        Dim seconds As Single
        Dim min As Byte
        Dim hour As Byte
        If Byte.TryParse(Strings.Left(hhmmsss, 2), hour) And Byte.TryParse(Strings.Mid(hhmmsss, 4, 2), min) And Single.TryParse(Strings.Right(hhmmsss, 6), seconds) Then
            Return (hour * 3600 + min * 60 + seconds).ToString("F" & digit.ToString)
        Else
            Return "-1"
        End If
        Return "0"
    End Function

    Public Function To_CNV_String() As String 'output_option 0 new
        Return day & " " & RSet(hhmmsss, 13) & " " & RSet(rec.ToString, 7) & " " & RSet(x.ToString("F2"), 12) & " " & RSet(y.ToString("F2"), 12) & " " & RSet(gyro.ToString("F2"), 6) _
                & " " & RSet(tow_bearing.ToString("F2"), 7) & " " & RSet(fix.ToString("F2"), 8) & " " & wd.ToString("F2")
    End Function

    Public Function To_Mag_CNV_String() As String 'output_option 0 new
        Return day & RSet(hhmmsss, 14) & RSet(fix.ToString, 8) & RSet(x.ToString("F2"), 13) & RSet(y.ToString("F2"), 13) & RSet(gyro.ToString("F2"), 7)
    End Function

    Public Function To_PC_String() As String
        Return RSet(rec.ToString, 7) & RSet(Sectime(2), 9) & RSet(fix.ToString, 7) &
            "  |        0.00 E        0.00 N    0.00M    0.00M" & RSet(gyro.ToString("F2"), 7) &
            "     0.00M   0.00    0.00  |        0.00 E        0.00 N   0.00S   0.00 M   0.00    0.00 M/S |   0.00    0.00 M/S   0.00S   0.00S        0.00 E        0.00 N   0.00 M   0.00     0.00 M   0.00   |    0.00M     0.00M    0.00     0.00M   0.00    0.00  |        0.00 E        0.00 N    0.00M   0.00      0.00M    0.00M    0.00M  0.00     0.00M   0.00          0.00 E        0.00 N    0.00 M   0.00         0.00 E        0.00 N     0.00 M   0.00     0.00 M   0.00     0.00 M   0.00      0.00M   0.00         0.00 E        0.00 N    0.00 M   0.00      0.00 M   0.00 " &
            RSet(x.ToString("F2"), 12) & " E" & RSet(y.ToString("F2"), 12) & " N  "
    End Function

    Public Function Tptime() As Double 'time in hour of a ay
        Dim seconds As Single
        Dim min As Byte
        Dim hour As Byte
        If Byte.TryParse(Strings.Left(hhmmsss, 2), hour) And Byte.TryParse(Strings.Mid(hhmmsss, 4, 2), min) And Single.TryParse(Strings.Right(hhmmsss, 6), seconds) Then
            Return Math.Round((hour + min / 60 + seconds / 3600), 4)
        Else
            Return -1.0#
        End If
        Return 0.0#
    End Function

End Class

Public Class Fenwdkpro_hdr 'a simple class store fix and its x-positions in sbp chart
    Public already_updated As Boolean = False
    Public iheight As Long = 0
    Public iwidth As Long = 0
    Public jpgpath As String = ""
    Public leftfix As Single = -1.0F
    Public leftpx As Long = 0
    Public leftx As Double = -1.0#
    Public rightfix As Single = -1.0F
    Public rightpx As Long = 0
    Public rightx As Double = -1.0#
    Public ve As Single = 10.0F
    Public vres As Single = 0.04F

    Public Function Cad_height() As Double
        Cad_height = ve * vres * iheight
    End Function

    Public Function Cad_inspt() As Double()
        Dim inspt(2) As Double
        inspt(0) = rightx - (rightpx / iwidth * Cad_width())
        inspt(1) = -Cad_height()
        inspt(2) = 0
        Cad_inspt = inspt
    End Function

    Public Function Cad_width() As Double
        Cad_width = (leftx - rightx) / (leftpx - rightpx) * iwidth
    End Function

End Class

Public Class Fenwdkpro_line 'a simple class store fenwdkpro line info
    Public fix As Double = 0.0#
    Public kp As Double = 0.0F
    Public ro As Single = 0.0F
    Public wd As Single = 0.0F
    Public x As Double = 0.0#
    Public y As Double = 0.0#

    Public Function To_String() As String
        Return fix.ToString("F1") & vbTab & x.ToString("F3") & vbTab & y.ToString("F3") & vbTab & wd.ToString("F3") & vbTab & kp.ToString("F3") & vbTab & ro.ToString("F3")
    End Function

End Class

Public Class FixPosition 'a simple class store fix and its x-positions in sbp chart
    Public fix As Single = -1.0F
    Public x As Double = -1.0#
End Class

Public Class PC
    Public fix As Double = 0.0#
    Public gyro As Single = 0.0F
    Public rec As Integer = 0
    Public sectime As Double = 0.0#
    Public x As Double = 0.0#
    Public y As Double = 0.0#

    Public Function To_PC_String() As String
        Return RSet(rec.ToString, 7) & RSet(sectime.ToString("F2"), 9) & RSet(fix.ToString, 7) &
            "  |        0.00 E        0.00 N    0.00M    0.00M" & RSet(gyro.ToString("F2"), 7) &
            "     0.00M   0.00    0.00  |        0.00 E        0.00 N   0.00S   0.00 M   0.00    0.00 M/S |   0.00    0.00 M/S   0.00S   0.00S        0.00 E        0.00 N   0.00 M   0.00     0.00 M   0.00   |    0.00M     0.00M    0.00     0.00M   0.00    0.00  |        0.00 E        0.00 N    0.00M   0.00      0.00M    0.00M    0.00M  0.00     0.00M   0.00          0.00 E        0.00 N    0.00 M   0.00         0.00 E        0.00 N     0.00 M   0.00     0.00 M   0.00     0.00 M   0.00      0.00M   0.00         0.00 E        0.00 N    0.00 M   0.00      0.00 M   0.00 " &
            RSet(x.ToString("F2"), 12) & " E" & RSet(y.ToString("F2"), 12) & " N  "
    End Function

End Class

Public Class XYZ ' a simply array with x y z , 3 element
    Public x As Long = -99
    Public y As Long = -99
    Public z As Single = -99.0F
    Private z_updated As Boolean = False

    Public Sub Add(ByVal input As Double)
        Select Case True
            Case x = -99
                x = CLng(input * 100)
            Case y = -99
                y = CLng(input * 100)
            Case Else
                z = CSng(input * -1)
                z_updated = True
        End Select
    End Sub

    Public Function Count() As Byte
        If Not z_updated Then Return 0
        Return 3
    End Function

    Public Sub Invert()
        z *= -1
    End Sub

    Public Function To_String() As String
        To_String = (x / 100).ToString("F2") & vbTab & (y / 100).ToString("F2") & vbTab & (z * -1).ToString("F2")
    End Function

End Class

'Public Class XyzPoint
'    Implements IComparable(Of XyzPoint)

'    Public Sub New(ByVal x As Double, ByVal y As Double, ByVal z As Double)
'        Me._x = x
'        Me._y = y
'        Me._z = z
'    End Sub

'    Public Overrides Function ToString() As String
'        Return String.Format("{0} {1} {2}", Me.X, Me.Y, Me.Z)
'    End Function

'    Private _x As Double
'    Public Property X() As Double
'        Get
'            Return _x
'        End Get
'        Set(ByVal value As Double)
'            _x = value
'        End Set
'    End Property

'    Private _y As Double
'    Public Property Y() As Double
'        Get
'            Return _y
'        End Get
'        Set(ByVal value As Double)
'            _y = value
'        End Set
'    End Property

'    Private _z As Double
'    Public Property Z() As Double
'        Get
'            Return _z
'        End Get
'        Set(ByVal value As Double)
'            _z = value
'        End Set
'    End Property

'    Public Function CompareTo(ByVal other As XyzPoint) As Integer Implements System.IComparable(Of XyzPoint).CompareTo
'        Return New XyzPointComparer().Compare(Me, other)
'    End Function

'    Private Class XyzPointComparer
'        Implements IComparer(Of XyzPoint)

'        Public Function Compare(ByVal x As XyzPoint, ByVal y As XyzPoint) As Integer Implements System.Collections.Generic.IComparer(Of XyzPoint).Compare
'            If x.X.CompareTo(y.X) <> 0 Then Return x.X.CompareTo(y.X)
'            If x.Y.CompareTo(y.Y) <> 0 Then Return x.Y.CompareTo(y.Y)
'            If x.Z.CompareTo(y.Z) <> 0 Then Return x.Z.CompareTo(y.Z)
'            Return 0
'        End Function
'    End Class
'End Class
'----------method 1-----------
'lambda expression
'l = j
'MyDoubles_init.CopyTo(MyDoubles, 0)
'recordt.Clear()
'copy list with only the selected rows
'sort (ascending) the input list with its x value

'recordt = records.FindAll(Function(yy) yy(1) = miny + cellsize * (nrows - 1 - j))
'recordt.Sort(Function(xx, yy) xx(0).CompareTo(yy(0)))
'For m = 0 To recordt.Count - 1
'    MyDoubles((recordt(m)(0) - minx) / cellsize) = -1 * recordt(m)(2)
'Next

'For Each value As Double In MyDoubles
'    If value = 999 Then
'        MyStrings(j) = MyStrings(j) + "x "
'    Else
'        MyStrings(j) = MyStrings(j) + value.ToString() + " "
'    End If
'Next
'MyStrings(j).TrimEnd()
'Debug.Print(DateTime.Now.Second.ToString + "." + DateTime.Now.Millisecond.ToString)
'firsti = records.FindIndex(Function(yy) yy(1) = miny + cellsize * (nrows - 1 - l))
'lasti = records.FindLastIndex(Function(yy) yy(1) = miny + cellsize * (nrows - 1 - l))
'firsti = lasti + 1
'lasti = records.FindIndex(Function(yy) yy(1) = (miny + cellsize * (nrows - 2 - l))) - 1
'If lasti < 0 Then lasti = records.Count - 1 '<-1
'----------method 2-----------
'If j < nrows - 1 And row_skipper(j) = records(firsti)(1) Then
'    If y_changes_index(i) > 0 Then
'        lasti = y_changes_index(i) - 1
'    End If
'    i = i + 1
'    For m = firsti To lasti
'        MyDS((records(m)(0) - minx) / cellsize) = (-0.001 * records(m)(2)).ToString
'    Next
'    firsti = lasti + 1 'set the first index for next array
'    'records(firsti)(1) will be the y value matching with row_skipper(j)
'ElseIf row_skipper(j) = records(firsti)(1) Then
'    'when j = nrows -1 work on the last row
'    lasti = records.Count - 1
'    For m = firsti To lasti
'        MyDS((records(m)(0) - minx) / cellsize) = (-0.001 * records(m)(2)).ToString
'    Next
'End If
'sort (ascending) the input list with its x (0) value, lambda, found no need in the matching stage, find min max x by read loop is faster
'records.Sort(Function(xx, yy) xx(0).CompareTo(yy(0)))
'minx = records(0)(0)
'maxx = records(records.Count - 1)(0)