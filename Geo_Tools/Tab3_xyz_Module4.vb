Module Tab3_xyz_Module4

    Sub Convert_to_GDAL_XYZ(input_listbox As System.Windows.Forms.ListBox)
        Dim k As Integer 'file loop
        Dim seperators(2) As Char
        seperators(0) = Chr(9) 'ControlChars.Tab
        seperators(1) = Chr(32) 'space
        seperators(2) = Chr(44) 'comma

        For k = 0 To input_listbox.Items.Count - 1
            Dim dot As Integer = input_listbox.Items(k).ToString.LastIndexOf(".")
            Dim outputfile As String = ""

            Dim ncols As Integer
            Dim nrows As Integer
            Dim j As Integer 'record count, control the row number

            Dim MyStrings() As String, remaining_rows As Integer, rows_written As Integer
            Dim records As New List(Of XYZ)

            input_listbox.SetSelected(k, True)
            input_listbox.Update()

            'read the xyz/dtm to a list
            If dot > 5 Then
                nrows = 1 'nrows as line number in case of exception, ncols as msgbox.result
                For Each line As String In System.IO.File.ReadLines(input_listbox.Items(k).ToString)
                    Dim values As New XYZ
                    For Each field As String In line.Split(seperators, 3, StringSplitOptions.RemoveEmptyEntries)
                        Try
                            values.Add(Double.Parse(field))
                        Catch Ex As Exception
                            nrows += 1 'using nrows ncols for exception var
                            ncols = CInt(MsgBox(input_listbox.Items(k).ToString & vbCrLf & Ex.Message &
                                   vbCrLf & "At line no. " & records.Count + nrows & vbCrLf & line & vbCrLf &
                                   "This line was skipped. Continue ?", MsgBoxStyle.YesNo))
                            If ncols = 7 Then Exit Sub
                        End Try
                    Next
                    If values.Count = 3 Then
                        values.Invert()
                        records.Add(values)
                    End If
                Next
                outputfile = Strings.Left(input_listbox.Items(k).ToString, dot) + "_GDAL." + Strings.Right(input_listbox.Items(k).ToString, 3)
                Try
                    FileOpen(2, outputfile, OpenMode.Output, OpenAccess.Write)
                    FileClose(2)
                Catch Ex As Exception
                    MessageBox.Show(Ex.Message & vbCrLf & "This file was skipped.")
                    Continue For
                End Try
            Else
                MsgBox("Input file error")
                Continue For
            End If

            'sort (ascending) the input list with its y (1) value, lambda
            'records.Sort(Function(xx As XYZ, yy As XYZ) xx.x.CompareTo(yy.x))
            'records.Sort(Function(xx As XYZ, yy As XYZ) xx.y.CompareTo(yy.y))
            records = records.OrderBy(Function(xx) xx.y).ThenBy(Function(xx) xx.x).ToList()

            Baseform.BaseProgressBar.Maximum = records.Count - 1
            Baseform.BaseProgressBar.Value = 0
            Baseform.BaseProgressBar.Update()

            If records.Count > 10000 Then
                'for big files
                ReDim MyStrings(0 To 9999)
                remaining_rows = records.Count
                rows_written = 0

                For j = 0 To records.Count - 1 Step 1
                    MyStrings(rows_written) = records(j).To_String
                    rows_written += 1
                    If rows_written = 10000 Then
                        IO.File.AppendAllLines(outputfile, MyStrings)
                        remaining_rows -= rows_written
                        If remaining_rows > rows_written Then '10000 = rows_written
                            ReDim MyStrings(0 To rows_written - 1)
                        Else
                            ReDim MyStrings(0 To remaining_rows - 1)
                        End If
                        rows_written = 0
                    End If
                    Baseform.BaseProgressBar.Value = j
                    Baseform.BaseProgressBar.Update()
                Next j
                If rows_written > 0 Then IO.File.AppendAllLines(outputfile, MyStrings)
            Else
                'for small file
                ReDim MyStrings(0 To records.Count - 1)
                For j = 0 To records.Count - 1 Step 1
                    MyStrings(j) = records(j).To_String
                    Baseform.BaseProgressBar.Value = j
                    Baseform.BaseProgressBar.Update()
                Next j
                IO.File.WriteAllLines(outputfile, MyStrings)
            End If

            input_listbox.SetSelected(k, False)
            Baseform.BaseProgressBar.Value = records.Count - 1
            Baseform.BaseProgressBar.Update()
        Next k
    End Sub

End Module