Module Tab3_xyz_Module1

    Sub Conv_xyz_to_asc(input_listbox As System.Windows.Forms.ListBox)
        Dim k As Integer 'file loop
        Dim seperators(2) As Char
        seperators(0) = Chr(9) 'ControlChars.Tab
        seperators(1) = Chr(32) 'space
        seperators(2) = Chr(44) 'comma

        For k = 0 To input_listbox.Items.Count - 1
            Dim dot As Integer = input_listbox.Items(k).ToString.LastIndexOf(".")
            Dim outputfile As String = ""

            Dim cellsize As Short
            Dim ncols As Integer
            Dim nrows As Integer
            Dim mnrows As Integer
            Dim minx As Long 'xy can be negative in some job
            Dim miny As Long
            Dim maxx As Long
            Dim maxy As Long
            Dim j As Integer 'record count, control the row number
            Dim m As Integer 'record loop between row
            Dim yt As Long 'temp y for cellsize

            Dim firsti As Integer
            Dim rows_written As Byte
            Dim records As New List(Of XYZ)

            input_listbox.SetSelected(k, True)
            input_listbox.Update()

            'read the xyz/dtm to a list
            If dot > 5 Then
                nrows = 1 'nrows as line number in case of exception, ncols as msgbox.result
                maxx = -9223372036854775807
                minx = 9223372036854775807
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
                        records.Add(values)
                        If values.x > maxx Then maxx = values.x
                        If values.x < minx Then minx = values.x
                    End If
                Next
                outputfile = Strings.Left(input_listbox.Items(k).ToString, dot + 1) + "ASC"
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

            'sort (descending) the input list with its y (1) value, lambda
            records.Sort(Function(xx, yy) yy.y.CompareTo(xx.y))
            maxy = records(0).y
            miny = records(records.Count - 1).y

            'init the first xy and cellsize, shall always used sorted y to find the cell size
            yt = maxy
            cellsize = 0
            For j = 0 To records.Count - 1
                If yt <> records(j).y Then
                    cellsize = CShort(System.Math.Abs(yt - records(j).y)) 'assume everything is gridded no gap between first and second row
                    Exit For
                End If
                'todo: compare second and third row for fail safe, msgbox user input specific grid size
            Next

            'init each row's y values
            Dim row_skipper(0 To CInt((maxy - miny) / cellsize)) As Long 'every row y value for skipping no data row
            row_skipper(0) = maxy
            For j = 1 To row_skipper.Count - 1
                row_skipper(j) = row_skipper(j - 1) - cellsize
            Next

            ncols = CInt((maxx - minx) / cellsize) + 1
            nrows = CInt((maxy - miny) / cellsize) + 1

            Dim MyStrings(0 To 5) As String
            MyStrings(0) = "ncols " + ncols.ToString()
            MyStrings(1) = "nrows " + nrows.ToString()
            MyStrings(2) = "xllcorner " + ((minx - cellsize / 2) / 100).ToString()
            MyStrings(3) = "yllcorner " + ((miny - cellsize / 2) / 100).ToString()
            MyStrings(4) = "cellsize " + (cellsize / 100).ToString()
            MyStrings(5) = "NODATA_value 99"
            IO.File.AppendAllLines(outputfile, MyStrings)

            '===Code below for ascii file type===

            'Prepare string array to writeto file
            If nrows > 9 Then
                ReDim MyStrings(0 To 9)
                mnrows = nrows - 10
            Else
                ReDim MyStrings(0 To nrows - 1)
            End If
            Dim MyDS_init(0 To ncols - 1) As String
            Dim MyDS(0 To ncols - 1) As String
            For j = 0 To ncols - 1
                MyDS_init(j) = "99"
            Next

            Baseform.BaseProgressBar.Maximum = nrows
            Baseform.BaseProgressBar.Value = 0

            firsti = 0
            rows_written = 0
            For j = 0 To nrows - 1
                MyDS_init.CopyTo(MyDS, 0) 'init MyDS

                If row_skipper(j) = records(firsti).y Then 'If no data gap
                    For m = firsti To records.Count - 1
                        If row_skipper(j) <> records(m).y Then Exit For
                        MyDS(CInt((records(m).x - minx) / cellsize)) = records(m).z.ToString
                    Next
                    firsti = m 'set the first index for next array
                End If

                MyStrings(rows_written) = String.Join(" ", MyDS)

                rows_written = CByte(rows_written + 1)

                'write to file every 10 rows
                If rows_written = 10 Or j = nrows - 1 Then 'rows_written = 10 or working on last row
                    IO.File.AppendAllLines(outputfile, MyStrings)

                    rows_written = 0
                    If mnrows > 10 Then 'if more than 10 rows left
                        ReDim MyStrings(0 To 9)
                        mnrows -= 10
                    Else
                        ReDim MyStrings(0 To mnrows - 1)
                    End If
                End If
                Baseform.BaseProgressBar.Increment(1)
                Baseform.BaseProgressBar.Update()
            Next j
            MyStrings = Nothing
            MyDS = Nothing
            MyDS_init = Nothing
            row_skipper = Nothing
            records = Nothing
            input_listbox.SetSelected(k, False)
            input_listbox.Update()
            Baseform.BaseProgressBar.Value = nrows
            Baseform.BaseProgressBar.Update()
        Next k
    End Sub

End Module