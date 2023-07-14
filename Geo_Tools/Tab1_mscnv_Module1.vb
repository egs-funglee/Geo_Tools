Module Tab1_mscnv_Module1

    'Merge and Split CNV Files for C-View Smoothing module
    Sub Merge_cnv(rawfilelistbox As System.Windows.Forms.ListBox)
        Dim i As Integer
        Dim rb As Integer
        Dim lb As Integer
        Dim outputfile As String = ""
        Dim file_path As String
        file_path = "DUMMY"
        Remove_listbox_items_if_contain(rawfilelistbox, "COMBINED")
        rawfilelistbox.SelectionMode = SelectionMode.One
        rawfilelistbox.Update()
        For i = 0 To rawfilelistbox.Items.Count - 1
            rawfilelistbox.SetSelected(i, True)
            rb = rawfilelistbox.Items(i).ToString.LastIndexOf(") ") 'InStrRev(file_path(nIndex), ") ")
            lb = rawfilelistbox.Items(i).ToString.LastIndexOf(" (") + 1 'InStrRev(file_path(nIndex), " (") + 1
            If lb > 5 And rb - lb = 2 Then 'if have (#) on the filename
                'If Not (Strings.Left(file_path, lb - 5) = Strings.Left(rawfilelistbox.SelectedItem.ToString, lb - 5)) then
                If IsDiffLine(file_path, rawfilelistbox.SelectedItem.ToString) Then
                    outputfile = Strings.Left(rawfilelistbox.SelectedItem.ToString, rb + 1) + " Combined" + Strings.Right(rawfilelistbox.SelectedItem.ToString, Strings.Len(rawfilelistbox.SelectedItem) - rb - 1)
                    FileOpen(2, outputfile, OpenMode.Output, OpenAccess.Write)
                    FileClose(2)
                End If
                Dim InputStrings() As String = IO.File.ReadAllLines(rawfilelistbox.SelectedItem.ToString)
                IO.File.AppendAllLines(outputfile, InputStrings)
                'ReDim InputStrings(0)
            Else
                outputfile = Strings.Left(rawfilelistbox.Items(i).ToString, Len(rawfilelistbox.Items(i).ToString) - 8) & " Combined.cnv"
                If System.IO.File.Exists(outputfile) Then
                    Dim InputStrings() As String = IO.File.ReadAllLines(rawfilelistbox.SelectedItem.ToString)
                    IO.File.AppendAllLines(outputfile, InputStrings)
                    'ReDim InputStrings(0)
                Else
                    Dim InputStrings() As String = IO.File.ReadAllLines(rawfilelistbox.SelectedItem.ToString)
                    IO.File.WriteAllLines(outputfile, InputStrings)
                    'ReDim InputStrings(0)
                End If
            End If
            file_path = rawfilelistbox.SelectedItem.ToString 'save this file path for appending in next loop
        Next
    End Sub

    Sub Split_cnv(ByVal inputfile As String, rawfilelistbox As System.Windows.Forms.ListBox)
        Dim outputfilelist() As String
        Dim startline() As String
        Dim endline() As String
        ReDim startline(0 To rawfilelistbox.Items.Count - 1)
        ReDim endline(0 To rawfilelistbox.Items.Count - 1)
        ReDim outputfilelist(0 To rawfilelistbox.Items.Count - 1)
        Dim i As Integer, j As Integer
        Dim sol As Integer, eol As Integer, arraylength As Integer
        Dim raw_cnv_match_pos, raw_cnv_match_len As Integer

        Remove_listbox_items_if_contain(rawfilelistbox, "COMBINED")

        'Loop to get SOL EOL of each raw cnv (the record numbers)
        For i = 0 To rawfilelistbox.Items.Count - 1
            FileOpen(1, rawfilelistbox.Items(i).ToString, OpenMode.Input, OpenAccess.Read)
            'Get SOL
            startline(i) = Strings.Left(LineInput(1), 32)
            'Get EOL
            Do While Not EOF(1)
                endline(i) = Strings.Left(LineInput(1), 32)
            Loop
            FileClose(1)
        Next i

        Dim InputStrings() As String = IO.File.ReadAllLines(inputfile)
        'Calc CMG And smooth with T1G2CheckBox1 Is clicked
        If Baseform.T1G2CheckBox1.Checked Then

            'Prepare CNV array and Calc inter-record bearing (Coarse CMG) in CNV array
            Dim cnv_array() As CNV
            ReDim cnv_array(0 To UBound(InputStrings))
            Dim cnv_obj0 As New CNV
            Line_to_cnv(InputStrings(0), cnv_obj0)
            cnv_array(0) = cnv_obj0
            cnv_obj0 = Nothing
            For i = 1 To UBound(InputStrings) Step 1
                Dim cnv_obj As New CNV
                Line_to_cnv(InputStrings(i), cnv_obj)
                cnv_array(i) = cnv_obj
                cnv_array(i).gyro = Calc_HDG(cnv_array(i - 1).x, cnv_array(i - 1).y, cnv_array(i).x, cnv_array(i).y)
                cnv_array(i).g_sine = Math.Sin(cnv_array(i).gyro * Math.PI / 180) 'to sin
                cnv_array(i).g_cosine = Math.Cos(cnv_array(i).gyro * Math.PI / 180) 'to cos
            Next i
            cnv_array(0).g_sine = cnv_array(1).g_sine
            cnv_array(0).g_cosine = cnv_array(1).g_cosine

            'Smooth headings rectangular (SMA by 3 points) from cnv_array.gyro to oheadings- preliminary
            Dim o_sine() As Double
            ReDim o_sine(0 To UBound(InputStrings))
            o_sine(0) = (cnv_array(0).g_sine + cnv_array(1).g_sine) / 2
            o_sine(UBound(o_sine)) = (cnv_array(UBound(InputStrings) - 1).g_sine + cnv_array(UBound(InputStrings)).g_sine) / 2
            Dim o_cosine() As Double
            ReDim o_cosine(0 To UBound(InputStrings))
            o_cosine(0) = (cnv_array(0).g_cosine + cnv_array(1).g_cosine) / 2
            o_cosine(UBound(o_cosine)) = (cnv_array(UBound(InputStrings) - 1).g_cosine + cnv_array(UBound(InputStrings)).g_cosine) / 2
            For i = 1 To UBound(InputStrings) - 1 Step 1
                o_sine(i) = (cnv_array(i - 1).g_sine + cnv_array(i).g_sine + cnv_array(i + 1).g_sine) / 3
                o_cosine(i) = (cnv_array(i - 1).g_cosine + cnv_array(i).g_cosine + cnv_array(i + 1).g_cosine) / 3
            Next i

            'Smoth headings triangular (Weighted TMA by selected points) from oheadings to headings
            Dim basen As Integer, sman_sine As Double, sman_cosine As Double
            Dim sweight As Integer, sweight1 As Integer, sweight2 As Integer
            Dim sbegin As Integer, send As Integer
            Dim swindow As Integer = CInt(Baseform.T1G2TextBox1.Text)

            For i = swindow \ 2 To UBound(o_sine) - swindow \ 2 Step 1 'mid section
                basen = 0
                sman_sine = 0
                sman_cosine = 0
                For j = 1 To swindow \ 2 Step 1
                    sweight = (swindow \ 2 - j + 1)
                    sman_sine = sman_sine + o_sine(i - j) * sweight + o_sine(i + j) * sweight
                    sman_cosine = sman_cosine + o_cosine(i - j) * sweight + o_cosine(i + j) * sweight
                    basen += j * 2
                Next j
                sman_sine += o_sine(i) * j
                sman_cosine += o_cosine(i) * j
                basen += j
                cnv_array(i).g_sine = CDbl(sman_sine / basen)
                cnv_array(i).g_cosine = CDbl(sman_cosine / basen)
            Next i

            For i = 0 To swindow \ 2 - 1 Step 1 'begin section
                basen = 0
                sman_sine = 0
                sman_cosine = 0
                For j = 1 To swindow \ 2 Step 1
                    sweight1 = (swindow \ 2 - j + 1)
                    sweight2 = sweight1
                    sbegin = i - j
                    If sbegin < 0 Then
                        sbegin = 0
                        sweight1 = 0
                    End If
                    sman_sine = sman_sine + o_sine(sbegin) * sweight1 + o_sine(i + j) * sweight2
                    sman_cosine = sman_cosine + o_cosine(sbegin) * sweight1 + o_cosine(i + j) * sweight2
                    basen = basen + sweight1 + sweight2
                Next j
                sman_sine += o_sine(i) * (swindow \ 2 + 1)
                sman_cosine += o_cosine(i) * (swindow \ 2 + 1)
                basen += j
                cnv_array(i).g_sine = CDbl(sman_sine / basen)
                cnv_array(i).g_cosine = CDbl(sman_cosine / basen)
            Next i

            For i = (UBound(o_sine) - swindow \ 2 + 1) To UBound(o_sine) Step 1 'end section
                basen = 0
                sman_sine = 0
                sman_cosine = 0
                For j = 1 To swindow \ 2 Step 1
                    sweight1 = (swindow \ 2 - j + 1)
                    sweight2 = sweight1
                    send = i + j
                    If send > UBound(o_sine) Then
                        send = UBound(o_sine)
                        sweight2 = 0
                    End If
                    sman_sine = sman_sine + o_sine(i - j) * sweight1 + o_sine(send) * sweight2
                    sman_cosine = sman_cosine + o_cosine(i - j) * sweight1 + o_cosine(send) * sweight2
                    basen = basen + sweight1 + sweight2
                Next j
                sman_sine += o_sine(i) * (swindow \ 2 + 1)
                sman_cosine += o_cosine(i) * (swindow \ 2 + 1)
                basen += j
                cnv_array(i).g_sine = CDbl(sman_sine / basen)
                cnv_array(i).g_cosine = CDbl(sman_cosine / basen)
            Next i

            'Replace InputStrings Vessel/Fish Gyro for split
            For i = 0 To UBound(InputStrings) Step 1
                cnv_array(i).gyro = Math.Atan2(cnv_array(i).g_sine, cnv_array(i).g_cosine) * (180 / Math.PI) 'back to deg
                If cnv_array(i).gyro < 0 Then cnv_array(i).gyro = 360 + cnv_array(i).gyro 'fix +/- 180 to compass
                InputStrings(i) = cnv_array(i).To_CNV_String
            Next i

        End If

        For i = 0 To rawfilelistbox.Items.Count - 1
            Dim OutputStrings() As String
            j = 0 'j is the index number of filteredcnv array
            sol = 1
            eol = 1
            If rawfilelistbox.Items(i).ToString.ToUpper.Contains("RAW.CNV") Then
                outputfilelist(i) = Strings.Replace(rawfilelistbox.Items(i).ToString.ToUpper, "RAW.CNV", "Filtered.cnv")
                raw_cnv_match_pos = 25 'match rec number, ignore time because of time interpolation
                raw_cnv_match_len = 8
            Else
                outputfilelist(i) = Strings.Replace(rawfilelistbox.Items(i).ToString.ToUpper, ".CNV", " Filtered.cnv")
                raw_cnv_match_pos = 1
                raw_cnv_match_len = 32
            End If

            Do While Not j = InputStrings.Count - 1 'while not eol of input filtered cnv, with index j
                If Strings.Mid(InputStrings(j), raw_cnv_match_pos, raw_cnv_match_len) = Strings.Mid(startline(i), raw_cnv_match_pos, raw_cnv_match_len) Then sol = j 'set sol index
                j += 1

                'Dim a As String = Strings.Mid(InputStrings(j), raw_cnv_match_pos, raw_cnv_match_len)
                'Dim b As String = Strings.Mid(endline(i), raw_cnv_match_pos, raw_cnv_match_len)

                If Strings.Mid(InputStrings(j), raw_cnv_match_pos, raw_cnv_match_len) = Strings.Mid(endline(i), raw_cnv_match_pos, raw_cnv_match_len) Then
                    eol = j 'set eol index
                    Exit Do 'exit do loop when eol is set
                End If
            Loop
            arraylength = eol - sol + 1 'set length of output string array
            rawfilelistbox.SelectionMode = SelectionMode.One
            rawfilelistbox.Update()
            If eol > sol Then
                rawfilelistbox.SetSelected(i, True)
                ReDim OutputStrings(0 To arraylength - 1)
                Array.Copy(InputStrings, sol, OutputStrings, 0, arraylength)
                IO.File.WriteAllLines(outputfilelist(i), OutputStrings)
            End If
            'ReDim OutputStrings(0)
        Next i
        'ReDim InputStrings(0)
    End Sub

    Private Function IsDiffLine(ByVal ofn As String, ByVal nfn As String) As Boolean
        Dim i As Integer
        IsDiffLine = True
        If ofn.Length <> nfn.Length Then Exit Function

        Dim o_fno As Integer, n_fno As Integer
        Dim slash As Integer, olb As Integer, nlb As Integer
        Dim ldiff As Integer, rdiff As Integer

        slash = ofn.LastIndexOf("\") + 2
        olb = ofn.LastIndexOf(") ") + 2
        nlb = nfn.LastIndexOf(") ") + 2
        If olb <> nlb Then Exit Function 'should be same anyway...
        'check right of )
        If (Right(ofn, ofn.Length - olb) <> Right(nfn, nfn.Length - nlb)) Then Exit Function

        'check channel
        If Mid(ofn, olb - 2, 1) <> Mid(nfn, nlb - 2, 1) Then Exit Function

        olb = ofn.LastIndexOf(" (")
        nlb = nfn.LastIndexOf(" (")
        If olb <> nlb Then Exit Function 'should be same anyway...

        'find diff on left side
        For i = slash To olb
            If Mid(ofn, i, 1) <> Mid(nfn, i, 1) Then
                Exit For
            End If
        Next i
        ldiff = i

        For i = olb To slash Step -1
            If Mid(ofn, i, 1) <> Mid(nfn, i, 1) Then
                Exit For
            End If
        Next i
        rdiff = i

        If (rdiff - ldiff) > 2 Then Exit Function 'allow 3 characters different

        Dim ofndiff As String = Mid(ofn, ldiff, rdiff - ldiff + 1)
        Dim nfndiff As String = Mid(nfn, ldiff, rdiff - ldiff + 1)

        If IsNumeric(ofndiff) And IsNumeric(nfndiff) Then
            o_fno = CInt(ofndiff)
            n_fno = CInt(nfndiff)
            If (n_fno - o_fno) = 1 Then IsDiffLine = False
        End If

    End Function

End Module