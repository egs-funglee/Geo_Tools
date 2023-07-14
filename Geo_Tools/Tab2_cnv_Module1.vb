Module Tab2_cnv_Module1

    Sub Conv_cnv2magxtfcnv(input_listbox As System.Windows.Forms.ListBox)
        input_listbox.SelectionMode = SelectionMode.One
        input_listbox.Update()
        For i = 0 To input_listbox.Items.Count - 1
            If System.IO.Path.GetFileName(input_listbox.Items(i).ToString).ToUpper.Contains(") FILTERED.CNV") Then
                Dim lb, rb As Integer
                input_listbox.SetSelected(i, True)

                rb = input_listbox.Items(i).ToString.LastIndexOf(") ") 'InStrRev(file_path(nIndex), ") ")
                lb = input_listbox.Items(i).ToString.LastIndexOf(" (") + 1 'InStrRev(file_path(nIndex), " (") + 1

                If lb > 5 And rb - lb = 2 Then

                    Dim outputfile As String = Strings.Left(input_listbox.Items(i).ToString, lb) & "Filtered.cnv"
                    Dim outputlines As New List(Of String)

                    'Dim pre_fix As Integer = 0
                    For Each line As String In IO.File.ReadLines(input_listbox.Items(i).ToString)
                        Dim cnv_obj As New CNV
                        Line_to_cnv(line, cnv_obj)
                        outputlines.Add(cnv_obj.To_Mag_CNV_String)
                    Next

                    If outputlines.Count > 1 Then System.IO.File.WriteAllLines(outputfile, outputlines.ToArray)
                    'outputlines = Nothing
                End If
                input_listbox.SetSelected(i, False)
            End If
        Next
    End Sub

    Sub Conv_cnv2pc(input_listbox As System.Windows.Forms.ListBox)
        input_listbox.SelectionMode = SelectionMode.One
        input_listbox.Update()
        For i = 0 To input_listbox.Items.Count - 1
            Dim line_name As String
            input_listbox.SetSelected(i, True)
            Dim outputfile As String
            line_name = System.IO.Path.GetFileNameWithoutExtension(input_listbox.Items(i).ToString)
            outputfile = System.IO.Path.GetDirectoryName(input_listbox.Items(i).ToString) & "\" & line_name & ".PC9"

            Dim outputlines As New List(Of String)
            Dim pre_fix As Integer = 0

            outputlines.Add("| Version 1.36 Position Check File")

            For Each line As String In IO.File.ReadLines(input_listbox.Items(i).ToString)
                Dim cnv_obj As New CNV
                Line_to_cnv(line, cnv_obj)
                If pre_fix = cnv_obj.fix Then
                    cnv_obj.fix = 0
                Else
                    pre_fix = cnv_obj.fix
                End If
                outputlines.Add(cnv_obj.To_PC_String)
            Next

            If outputlines.Count > 1 Then System.IO.File.WriteAllLines(outputfile, outputlines.ToArray)
            'outputlines = Nothing
            input_listbox.SetSelected(i, False)
        Next
    End Sub

    Sub Conv_cnv2scr(input_listbox As System.Windows.Forms.ListBox)
        Dim outputfile As String
        Dim line_name As String
        Dim pre_x As Double = 0.0#
        Dim pre_y As Double = 0.0#

        For i = 0 To input_listbox.Items.Count - 1

            input_listbox.SetSelected(i, True)

            Dim outputlines As New List(Of String)
            line_name = System.IO.Path.GetFileNameWithoutExtension(input_listbox.Items(i).ToString)
            outputfile = System.IO.Path.GetDirectoryName(input_listbox.Items(i).ToString) & "\" & line_name & ".scr"
            outputlines.Add("PLINE")

            For Each line As String In IO.File.ReadLines(input_listbox.Items(i).ToString)
                Dim cnv_obj As New CNV
                Line_to_cnv(line, cnv_obj)
                If Not (pre_x = cnv_obj.x And pre_y = cnv_obj.y) Then
                    outputlines.Add(cnv_obj.x.ToString("F2") & "," & cnv_obj.y.ToString("F2"))
                    pre_x = cnv_obj.x
                    pre_y = cnv_obj.y
                End If
            Next
            outputlines.Add("")
            IO.File.WriteAllLines(outputfile, outputlines.ToArray)
            'outputlines = Nothing
            input_listbox.SetSelected(i, False)
        Next
    End Sub

    Sub Conv_cnv2tp(input_listbox As System.Windows.Forms.ListBox)
        Dim tpsol As Byte
        Dim outputfile As String
        Dim line_name As String
        Dim pre_x As Double = 0.0#
        Dim pre_y As Double = 0.0#

        For i = 0 To input_listbox.Items.Count - 1

            input_listbox.SetSelected(i, True)

            line_name = System.IO.Path.GetFileNameWithoutExtension(input_listbox.Items(i).ToString)
            outputfile = System.IO.Path.GetDirectoryName(input_listbox.Items(i).ToString) & "\" & line_name & ".tp"

            tpsol = 1
            Dim outputlines As New List(Of String)

            For Each line As String In IO.File.ReadLines(input_listbox.Items(i).ToString)
                Dim cnv_obj As New CNV
                Line_to_cnv(line, cnv_obj)
                If Not (pre_x = cnv_obj.x And pre_y = cnv_obj.y) Then
                    outputlines.Add(tpsol & RSet(cnv_obj.fix.ToString, 7) &
                         RSet(cnv_obj.x.ToString("F2"), 12) & RSet(cnv_obj.y.ToString("F2"), 12) &
                          RSet(cnv_obj.Tptime.ToString("F4"), 8) &
                         " " & line_name)
                    tpsol = 2
                    pre_x = cnv_obj.x
                    pre_y = cnv_obj.y
                End If

            Next

            IO.File.WriteAllLines(outputfile, outputlines.ToArray)
            'outputlines = Nothing
            input_listbox.SetSelected(i, False)
        Next
    End Sub

    Sub Conv_cnv2tpc(input_listbox As System.Windows.Forms.ListBox)
        input_listbox.SelectionMode = SelectionMode.One
        input_listbox.Update()
        For i = 0 To input_listbox.Items.Count - 1
            Dim line_name As String
            input_listbox.SetSelected(i, True)
            Dim outputfile As String
            line_name = System.IO.Path.GetFileNameWithoutExtension(input_listbox.Items(i).ToString)
            outputfile = System.IO.Path.GetDirectoryName(input_listbox.Items(i).ToString) & "\" & line_name & ".PC0"

            Dim outputlines As New List(Of String)
            Dim pre_fix As Integer = 0
            outputlines.Add("| Version 1.36 Position Check File")

            For Each line As String In IO.File.ReadLines(input_listbox.Items(i).ToString)
                Dim cnv_obj As New CNV
                Line_to_cnv(line, cnv_obj)
                If pre_fix = cnv_obj.fix Then
                    cnv_obj.fix = 0
                Else
                    pre_fix = cnv_obj.fix
                    outputlines.Add(cnv_obj.To_PC_String)
                End If
            Next
            If outputlines.Count > 1 Then System.IO.File.WriteAllLines(outputfile, outputlines.ToArray)
            'outputlines = Nothing
            input_listbox.SetSelected(i, False)
        Next
    End Sub

End Module