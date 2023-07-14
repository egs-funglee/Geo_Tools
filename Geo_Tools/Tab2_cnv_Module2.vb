Module Tab2_cnv_Module2

    Sub Recalc_smooth_cnv_cmg(input_listbox As System.Windows.Forms.ListBox)
        Remove_listbox_items_if_not_contain(input_listbox, ".CNV")
        Dim wkgfile As String
        Dim j As Integer

        input_listbox.SelectionMode = SelectionMode.One
        input_listbox.Update()

        For ii = 0 To input_listbox.Items.Count - 1

            'Set the working file
            input_listbox.SetSelected(ii, True)
            wkgfile = input_listbox.Items(ii).ToString
            If Not wkgfile.ToUpper.Contains(".CNV") Then
                MsgBox("Something wrong with input files")
                Exit Sub
            End If

            Dim InputStrings() As String = IO.File.ReadAllLines(wkgfile)
            Dim cnv_array(0 To UBound(InputStrings)) As CNV
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
            cnv_array(0).gyro = cnv_array(1).gyro
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
            Dim swindow As Integer = CInt(Baseform.T2TextBox1.Text)

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

            IO.File.WriteAllLines(wkgfile, InputStrings)
            input_listbox.SetSelected(ii, False)
        Next ii
    End Sub

End Module