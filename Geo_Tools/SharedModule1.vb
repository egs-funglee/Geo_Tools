Module SharedModule1

    Public Function Calc_HDG(ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double) As Single
        Dim dx As Double, dy As Double
        dx = x1 - x0
        dy = y1 - y0
        Select Case True
            Case dx = 0 'on y-axis or both zero
                If dy < 0 Then Calc_HDG = 180 Else Calc_HDG = 0
            Case dy = 0 'on x-axis
                If dx > 0 Then Calc_HDG = 90 Else Calc_HDG = 270
            Case Else 'not on either axis
                If dx < 0 And dy > 0 Then 'Quadrant 2
                    Calc_HDG = CSng(450 - Math.Atan2(dy, dx) * (180 / Math.PI))
                Else 'Quadrant 1,3,4
                    Calc_HDG = CSng(90 - Math.Atan2(dy, dx) * (180 / Math.PI))
                End If
        End Select
        If Calc_HDG > 180.0 Then Calc_HDG -= 360 '+/- 180 deg, north = zero
    End Function

    Public Function Check_exist(ByVal file_path As String) As Byte
        Select Case True
            Case System.IO.File.Exists(file_path)
                Return 1
            Case System.IO.Directory.Exists(file_path)
                Return 2
            Case Else
                Return 0
        End Select
    End Function

    Public Sub InteropReleaseComObject(ByVal obj As Object)
        Try
            Dim intRel As Integer = 0
            Do
                intRel = System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            Loop While intRel > 0
        Catch ex As Exception
            MsgBox("Error releasing object" & ex.ToString)
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Public Sub Line_to_cnv(ByVal line As String, ByRef cnv_obj As CNV)
        Dim seperators(0) As Char
        seperators(0) = Chr(32)

        Dim cells As String() = line.Split(seperators, 9, StringSplitOptions.RemoveEmptyEntries)
        Dim catcher As Boolean

        cnv_obj.day = cells(0)
        cnv_obj.hhmmsss = cells(1)
        catcher = Integer.TryParse(cells(2), cnv_obj.rec)
        catcher = Double.TryParse(cells(3), cnv_obj.x)
        catcher = Double.TryParse(cells(4), cnv_obj.y)
        catcher = Double.TryParse(cells(5), cnv_obj.gyro)

        If UBound(cells) > 5 Then 'newer cnv format v1
            catcher = Single.TryParse(cells(6), cnv_obj.tow_bearing)
        End If

        If UBound(cells) > 6 Then 'newer cnv format v2
            Dim tdbl As Double
            catcher = Double.TryParse(cells(7), tdbl)
            cnv_obj.fix = CInt(tdbl) 'ignore odd even rounding
        End If

        If UBound(cells) > 7 Then 'newer cnv format v3
            catcher = Double.TryParse(cells(8), cnv_obj.wd)
        End If
    End Sub

    Public Sub Line_to_pc_object(ByVal line As String, ByRef pc_obj As PC)
        Dim seperators(9) As Char
        seperators(0) = Chr(9) 'ControlChars.Tab
        seperators(1) = Chr(32) 'space
        seperators(2) = Chr(44) 'comma
        seperators(3) = Chr(124) '|
        seperators(4) = Chr(45) '-
        seperators(5) = "E"c
        seperators(6) = "N"c
        seperators(7) = "M"c
        seperators(8) = "S"c
        seperators(9) = "/"c

        'Dim values(65) As String
        'values = line.Split(seperators, 66, StringSplitOptions.RemoveEmptyEntries)

        Dim values() As String
        values = Strings.Left(line, 80).Split(seperators, 9, StringSplitOptions.RemoveEmptyEntries)

        Dim catcher As Boolean
        catcher = Integer.TryParse(values(0), pc_obj.rec)
        catcher = Double.TryParse(values(1), pc_obj.sectime)
        catcher = Double.TryParse(values(2), pc_obj.fix)
        catcher = Single.TryParse(values(7), pc_obj.gyro)

        values = Strings.Right(line, 30).Split(seperators, 3, StringSplitOptions.RemoveEmptyEntries)
        catcher = Double.TryParse(values(0), pc_obj.x)
        catcher = Double.TryParse(values(1), pc_obj.y)
    End Sub

    Public Sub Remove_listbox_items_if_contain(input_listbox As System.Windows.Forms.ListBox, ByVal must_contain As String)
        For i = input_listbox.Items.Count - 1 To 0 Step -1
            If System.IO.Path.GetFileName(input_listbox.Items(i).ToString).ToUpper.Contains(must_contain.ToUpper) Then input_listbox.Items.RemoveAt(i)
        Next
        If input_listbox.Items.Count = 0 Then Baseform.Init_listbox(input_listbox)
    End Sub

    Public Sub Remove_listbox_items_if_ending_match(input_listbox As System.Windows.Forms.ListBox, ByVal filter As String)
        For i = input_listbox.Items.Count - 1 To 0 Step -1
            If Strings.Right((input_listbox.Items(i).ToString), Len(filter)).ToUpper = (filter.ToUpper) Then input_listbox.Items.RemoveAt(i)
        Next
        If input_listbox.Items.Count = 0 Then Baseform.Init_listbox(input_listbox)
    End Sub

    Public Sub Remove_listbox_items_if_not_contain(input_listbox As System.Windows.Forms.ListBox, ByVal must_contain As String)
        For i = input_listbox.Items.Count - 1 To 0 Step -1
            If Not System.IO.Path.GetFileName(input_listbox.Items(i).ToString).ToUpper.Contains(must_contain.ToUpper) Then input_listbox.Items.RemoveAt(i)
        Next
        If input_listbox.Items.Count = 0 Then Baseform.Init_listbox(input_listbox)
    End Sub

    Public Sub TryWriteAllLines(ByVal outputfilepath As String, ByVal strings As String())
        Try
            IO.File.WriteAllLines(outputfilepath, strings)
        Catch ex As Exception
            Dim mbr As MsgBoxResult = MsgBox(ex.Message & vbCrLf & vbCrLf _
                                        & "Please select output file path manually", MsgBoxStyle.OkCancel)
            If mbr = MsgBoxResult.Ok Then
                Using saveFileDialog1 As New SaveFileDialog With {
                    .Filter = "output files (*.output)|*.output|All files (*.*)|*.*",
                    .FilterIndex = 1,
                    .RestoreDirectory = True,
                    .FileName = "Output file name"
                }
                    If saveFileDialog1.ShowDialog() = DialogResult.OK Then
                        outputfilepath = saveFileDialog1.FileName
                        TryWriteAllLines(outputfilepath, strings)
                    End If
                End Using
            End If
        End Try
    End Sub

End Module