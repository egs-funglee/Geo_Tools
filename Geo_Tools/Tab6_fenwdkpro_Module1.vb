Module Tab6_fenwdkpro_Module1

    Sub Pc_to_fenwdkpro(input_listbox As System.Windows.Forms.ListBox, ByVal sbp_img_left2right As Boolean, ByVal overwriting As Boolean)
        input_listbox.SelectionMode = SelectionMode.One
        input_listbox.Update()

        Dim outputstrings(4) As String
        outputstrings(0) = ("#0,0.0,0.0,0.0")
        outputstrings(1) = ("#0,0,0,0,0.0")
        outputstrings(2) = ("#0,0,0,0,0.0")
        outputstrings(3) = ("#0,0.0")
        outputstrings(4) = ("#0,0.0")

        For i = 0 To input_listbox.Items.Count - 1

            If Not Strings.Right(input_listbox.Items(i).ToString, 5).ToUpper.Contains(".PC") Then Continue For

            Dim dot As Integer = InStrRev(input_listbox.Items(i).ToString, ".")
            Dim outputfile As String = Strings.Left(input_listbox.Items(i).ToString, dot) & "fenwdkpro"

            If Check_exist(outputfile) = 2 Then
                MsgBox("File skipped" & vbCrLf & "This output file path is a folder" & vbCrLf & outputfile)
                Continue For
            End If
            If Check_exist(outputfile) = 1 And overwriting = False Then Continue For

            input_listbox.SetSelected(i, True)

            Dim e0, n0 As Double
            Dim kp As Double = 0.0R
            Dim lr As Short = 1000

            If Not sbp_img_left2right Then
                lr = -1000
                kp = 99.999R
            End If

            Dim isfirstline As Boolean = True
            Dim firstline_have_fix As Boolean = False
            Dim fenwdkpro_list As New List(Of Fenwdkpro_line)
            Dim pc_obj As New PC

            For Each line As String In System.IO.File.ReadLines(input_listbox.Items(i).ToString)
                If Not (Left(line, 1) = "-" Or Left(line, 1) = "|") Then

                    Dim fenwdkpro_obj As New Fenwdkpro_line

                    Line_to_pc_object(line, pc_obj)

                    If isfirstline Then
                        If pc_obj.fix > 0 Then firstline_have_fix = True
                        fenwdkpro_obj.fix = pc_obj.fix
                        fenwdkpro_obj.x = pc_obj.x
                        fenwdkpro_obj.y = pc_obj.y
                        fenwdkpro_obj.kp = kp
                        fenwdkpro_list.Add(fenwdkpro_obj)
                        isfirstline = False
                        e0 = pc_obj.x
                        n0 = pc_obj.y
                        Continue For
                    End If

                    'other lines which have fix
                    If pc_obj.fix > 0 Then
                        fenwdkpro_obj.fix = pc_obj.fix
                        fenwdkpro_obj.x = pc_obj.x
                        fenwdkpro_obj.y = pc_obj.y
                        kp += (((pc_obj.x - e0) ^ 2 + (pc_obj.y - n0) ^ 2) ^ 0.5) / lr
                        fenwdkpro_obj.kp = kp
                        fenwdkpro_list.Add(fenwdkpro_obj)
                        e0 = pc_obj.x
                        n0 = pc_obj.y
                        If Not firstline_have_fix Then
                            fenwdkpro_list(0).fix = fenwdkpro_obj.fix - 0.5
                            firstline_have_fix = True
                        End If
                    End If
                End If
            Next

            'lastline
            If pc_obj.fix = 0 Then
                kp += (((pc_obj.x - e0) ^ 2 + (pc_obj.y - n0) ^ 2) ^ 0.5) / lr

                Dim fenwdkpro_obj As New Fenwdkpro_line With {
                    .fix = fenwdkpro_list(fenwdkpro_list.Count - 1).fix + 0.5,
                    .x = pc_obj.x,
                    .y = pc_obj.y,
                    .kp = kp
                }
                'fenwdkpro_obj.To_String()
                fenwdkpro_list.Add(fenwdkpro_obj)
            End If
            pc_obj = Nothing

            If fenwdkpro_list.Count > 1 Then
                IO.File.WriteAllLines(outputfile, outputstrings)
                Dim outputstrings2(fenwdkpro_list.Count - 1) As String
                outputstrings2 = Array.ConvertAll(fenwdkpro_list.ToArray, New System.Converter(Of Fenwdkpro_line, String)(AddressOf Convert_fenwdkpro_obj_to_line))
                IO.File.AppendAllLines(outputfile, outputstrings2)
                fenwdkpro_list = Nothing
                outputstrings2 = Nothing
            End If
            input_listbox.SetSelected(i, False)
        Next
        outputstrings = Nothing
    End Sub

    Sub Update_fenwdkpro(ByVal inputfile As String, input_listbox As System.Windows.Forms.ListBox, ByVal overwriting As Boolean, ByVal ve As Integer)
        Dim fixlist As New List(Of FixPosition)
        Dim exception_c As Integer
        Dim nerr As Integer
        Dim seperators(3) As Char
        seperators(0) = Chr(9) 'ControlChars.Tab
        seperators(1) = Chr(32) 'space
        seperators(2) = Chr(44) 'comma
        seperators(3) = Chr(35) '#

        input_listbox.SelectionMode = SelectionMode.One
        input_listbox.Update()

        'check input file format if valid obtain fix profile positions into a list update each fenwdkprofile
        If Strings.Right(inputfile, 4).ToUpper.Contains(".DXF") Then

            Dim lines() As String = System.IO.File.ReadAllLines(inputfile)

            If Strings.Left(lines(7), 2) = "AC" Then

                Dim dxf_version, line_interval_fix, line_interval_textalx As Integer
                dxf_version = Integer.Parse(Strings.Right(lines(7), 4))

                If dxf_version > 1010 Then
                    line_interval_fix = 20
                    line_interval_textalx = 28
                Else
                    line_interval_fix = 14
                    line_interval_textalx = 22
                End If

                For i = 0 To lines.Count - 1
                    Dim values As New FixPosition
                    If lines(i) = "TEXT" Then
                        If Double.TryParse(lines(i + line_interval_textalx), values.x) And Single.TryParse(lines(i + line_interval_fix), values.fix) Then
                            values.x = Math.Round(values.x, 3)
                            fixlist.Add(values)
                        End If
                        'Try
                        '    values.x = Math.Round(Double.Parse(lines(i + line_interval_textalx)), 3)
                        '    values.fix = Single.Parse(lines(i + line_interval_fix))
                        '    fixlist.Add(values)
                        'Catch
                        'End Try
                    End If
                Next
            Else

                MsgBox("Check input DXF File. Please save it again with AutoCAD")
                Exit Sub

            End If
        Else

            For Each line As String In System.IO.File.ReadLines(inputfile)

                Dim values As New FixPosition
                Dim texts(3) As String

                Try
                    texts = line.Split(seperators, 4, StringSplitOptions.RemoveEmptyEntries)
                    If Double.TryParse(texts(0), values.x) And Single.TryParse(texts(2), values.fix) Then
                        values.x = Math.Round(values.x, 3)
                        fixlist.Add(values)
                    End If
                    'values.x = Math.Round(Double.Parse(texts(0)), 3)
                    'values.fix = Single.Parse(texts(2))
                    'fixlist.Add(values)
                Catch Ex As Exception
                    nerr += 1 'for exception
                    exception_c = CInt(MsgBox(inputfile & vbCrLf & Ex.Message &
                           vbCrLf & "At line no. " & fixlist.Count + nerr & vbCrLf & line & vbCrLf &
                           "This line was skipped. Continue ?", MsgBoxStyle.YesNo))
                    If exception_c = 7 Then Exit Sub
                End Try

            Next

        End If

        If fixlist.Count < 3 Then
            MsgBox("Not enough fix points, check input file")
            Exit Sub
        End If

        For i = 0 To input_listbox.Items.Count - 1

            If Not Strings.Right(input_listbox.Items(i).ToString, 5).ToUpper.Contains("DKPRO") Then Continue For
            input_listbox.SetSelected(i, True)

            Dim inputfile_fenwdkpro As String = input_listbox.Items(i).ToString
            Dim already_updated As Boolean = False
            Dim fenwdkpro_obj As New Fenwdkpro_hdr
            Dim lines_fenwdkpro() As String = System.IO.File.ReadAllLines(inputfile_fenwdkpro).ToArray()

            Try
                Dim texts(3) As String
                texts = lines_fenwdkpro(0).Split(seperators, 4, StringSplitOptions.RemoveEmptyEntries) 'line1
                If texts(0) = "1" Then already_updated = True
                texts = lines_fenwdkpro(1).Split(seperators, 4, StringSplitOptions.RemoveEmptyEntries) 'line2
                fenwdkpro_obj.leftfix = Single.Parse(texts(0))
                texts = lines_fenwdkpro(2).Split(seperators, 4, StringSplitOptions.RemoveEmptyEntries) 'line3
                fenwdkpro_obj.rightfix = Single.Parse(texts(0))
                fenwdkpro_obj.ve = ve
            Catch Ex As Exception
                exception_c = CInt(MsgBox(inputfile_fenwdkpro & vbCrLf & Ex.Message &
                       vbCrLf & vbCrLf & "Please check the file format, file skipped", MsgBoxStyle.OkOnly))
                Continue For
            End Try

            If fenwdkpro_obj.leftfix = 0 Or fenwdkpro_obj.rightfix = 0 Then
                MsgBox("File skipped. No Left and Right Fix in the header of this file : " & vbCrLf &
                        inputfile_fenwdkpro & vbCrLf &
                        "Please make sure the fenwdkpro file is the updated result after creating SBP Image")
                Continue For
            End If

            If already_updated = False Or overwriting = True Then
                If fixlist.Exists(Function(xx) xx.fix = fenwdkpro_obj.leftfix) And fixlist.Exists(Function(xx) xx.fix = fenwdkpro_obj.rightfix) Then
                    'sort with it's x, ascending (left to right)
                    fixlist.Sort(Function(xx, yy) xx.x.CompareTo(yy.x))
                    'left fix x (last)
                    fenwdkpro_obj.leftx = fixlist.FindLast(Function(xx) xx.fix = fenwdkpro_obj.leftfix).x
                    'right fix x (first)
                    fenwdkpro_obj.rightx = fixlist.Find(Function(xx) xx.fix = fenwdkpro_obj.rightfix).x

                    lines_fenwdkpro(0) = "#1," + fenwdkpro_obj.leftx.ToString("R3") + "," + fenwdkpro_obj.rightx.ToString("R3") + "," + fenwdkpro_obj.ve.ToString("N3")
                    System.IO.File.WriteAllLines(inputfile_fenwdkpro, lines_fenwdkpro)
                Else
                    MsgBox("The input file doesn't contain the Fix for updating the follwing fenwdkpro :" & vbCrLf &
                            inputfile_fenwdkpro & vbCrLf & vbCrLf &
                            "Please check the input file :" & vbCrLf & inputfile)
                End If
            End If
            input_listbox.SetSelected(i, False)
        Next
    End Sub

    Private Function Convert_fenwdkpro_obj_to_line(input As Fenwdkpro_line) As String
        Return input.To_String
    End Function

End Module