Module Tab5_pc_Module1

    Sub Pc_to_cnv(input_listbox As System.Windows.Forms.ListBox)
        Dim outputfile As String
        Dim line_name As String
        Dim seperators(9) As Char
        Dim ch_no As Char

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

        input_listbox.SelectionMode = SelectionMode.One
        input_listbox.Update()

        For i = 0 To input_listbox.Items.Count - 1
            input_listbox.SetSelected(i, True)

            ch_no = System.Convert.ToChar(Right(input_listbox.Items(i).ToString, 1))
            line_name = System.IO.Path.GetFileNameWithoutExtension(input_listbox.Items(i).ToString)
            outputfile = System.IO.Path.GetDirectoryName(input_listbox.Items(i).ToString) & "\" & line_name &
                " (" & ch_no & ") Raw.cnv"

            Dim Input_Strings As New List(Of String)
            Dim lastfix As Double = 0

            For Each line As String In IO.File.ReadLines(input_listbox.Items(i).ToString)
                If Not (Left(line, 1) = "-" Or Left(line, 1) = "|") Then
                    Dim values(65) As String
                    values = line.Split(seperators, 66, StringSplitOptions.RemoveEmptyEntries)
                    Dim rectime As DateTime = (New DateTime().AddHours(Double.Parse(values(1)) / 3600))

                    Dim thisfix As Double = Double.Parse(values(2))

                    If thisfix > 0 Then lastfix = thisfix

                    Input_Strings.Add(RSet(rectime.ToString("dd\/MM\/yyyy  HH:mm:ss.fff"), 24) &
                        RSet(values(0), 8) & RSet(values(63), 13) & RSet(values(64), 13) &
                        RSet(values(7), 7) & " -999.90" & RSet(lastfix.ToString("0.00"), 9) & "   0.00")
                End If
            Next

            IO.File.WriteAllLines(outputfile, Input_Strings)
            'Input_Strings = Nothing
            input_listbox.SetSelected(i, False)
        Next
    End Sub

    Sub Pc_to_cpc(input_listbox As System.Windows.Forms.ListBox)
        Dim outputfile As String = ""
        Dim line_name As String
        Dim pre_line_name As String = ""
        Dim seperators(4) As Char
        Dim pre_fix As String = ""

        seperators(0) = Chr(9) 'ControlChars.Tab
        seperators(1) = Chr(32) 'space
        seperators(2) = Chr(44) 'comma
        seperators(3) = Chr(124) '|
        seperators(4) = Chr(45) '-

        input_listbox.SelectionMode = SelectionMode.One
        input_listbox.Update()

        For i = 0 To input_listbox.Items.Count - 1
            If Strings.Right(input_listbox.Items(i).ToString, 5).ToUpper.Contains(".PC0") Then Continue For
            If Strings.Right(input_listbox.Items(i).ToString, 5).ToUpper.Contains(".PC9") Then Continue For
            input_listbox.SetSelected(i, True)

            line_name = System.IO.Path.GetFileNameWithoutExtension(input_listbox.Items(i).ToString)
            line_name = LSet(line_name, Len(line_name) - 4)

            If Not pre_line_name = line_name Then
                outputfile = System.IO.Path.GetDirectoryName(input_listbox.Items(i).ToString)
                outputfile = outputfile & "\" & line_name & ".PC9"
                FileOpen(1, outputfile, OpenMode.Output, OpenAccess.Write)
                PrintLine(1, "| Version 1.36 Position Check File")
                PrintLine(1, "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")
                PrintLine(1, "|  Rec    Time    Fix    |      Antenna Position       |   Antenna Offset                                 |                     Datum                                       |                             Smoothed Datum                                                         |                                               Fixed Offset                                          |                      Cable                       |              Towed Sensor  Using   CMG        |            USBL                                                                                                                                                     | Smooth Datum Position Used |")
                PrintLine(1, "|                        |                             |                                 Relative  True   |                                 Inter Rec                       |                  Time Update  Rec                                  Inter Rec     Offset from Raw   |      Fixed                         Relative  True                                    Inter Rec      | Counter   Manual   Input Vertical  Offset Offset |                               Inter Rec       |                                 Inter Rec       USBL Offset       Towed Offset    Smooth USBL Offset    Smooth USBL Position        Inter Rec       USBL to Towed   |       For Interp           |")
                PrintLine(1, "|  No.    Secs    No.    |       X           Y         |   X        Y    Gyro      Dist   Brg       Brg   |       X            Y          Time    DMG     CMG     Speed     |  CMG      Speed     Time     Time       X              Y         Dist    Brg     Dist      Brg     |   X        Y       Gyro      Dist    Brg     Brg     |     X             Y           Dist   Brg     |  Input    Input    Used   Angle     Dist   Brg   |        X          Y           Dist   Brg      |     X              Y           Dist   Brg       Dist   Brg         Dist   Brg         Dist  Brg          X             Y           Dist   Brg        Dist   Brg     |     X            Y         |")
                PrintLine(1, "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")
                FileClose(1)
            End If

            Dim Input_Strings As New List(Of String)

            For Each line As String In IO.File.ReadLines(input_listbox.Items(i).ToString)
                If Not (Left(line, 1) = "-" Or Left(line, 1) = "|") Then
                    Dim values(3) As String
                    values = line.Split(seperators, 4, StringSplitOptions.RemoveEmptyEntries)
                    If pre_fix = values(2) Then line = Strings.Left(line, 16) & "      0" & Strings.Right(line, 646)
                    If Not (values(2) = "0" Or values(2) = pre_fix) Then pre_fix = values(2)
                    Input_Strings.Add(line)
                End If
            Next

            IO.File.AppendAllLines(outputfile, Input_Strings)
            'Input_Strings = Nothing
            pre_line_name = line_name
            input_listbox.SetSelected(i, False)
        Next
    End Sub

    Sub Pc_to_ctp(input_listbox As System.Windows.Forms.ListBox)
        Dim tpsol As Byte
        Dim outputfile As String = ""
        Dim line_name As String
        Dim pre_line_name As String = ""
        Dim seperators(9) As Char
        Dim pre_fix As String = ""

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

        input_listbox.SelectionMode = SelectionMode.One
        input_listbox.Update()

        For i = 0 To input_listbox.Items.Count - 1
            If Strings.Right(input_listbox.Items(i).ToString, 5).ToUpper.Contains(".PC0") Then Continue For
            If Strings.Right(input_listbox.Items(i).ToString, 5).ToUpper.Contains(".PC9") Then Continue For
            input_listbox.SetSelected(i, True)

            line_name = System.IO.Path.GetFileNameWithoutExtension(input_listbox.Items(i).ToString)
            line_name = LSet(line_name, Len(line_name) - 4)

            If Not pre_line_name = line_name Then
                outputfile = System.IO.Path.GetDirectoryName(input_listbox.Items(i).ToString)
                outputfile = outputfile & "\" & line_name & ".tp"
                FileOpen(1, outputfile, OpenMode.Output, OpenAccess.Write)
                FileClose(1)
                tpsol = 1
            Else
                tpsol = 2
            End If

            Dim Input_Strings As New List(Of String)

            For Each line As String In IO.File.ReadLines(input_listbox.Items(i).ToString)
                If Not (Left(line, 1) = "-" Or Left(line, 1) = "|") Then
                    Dim values(65) As String
                    values = line.Split(seperators, 66, StringSplitOptions.RemoveEmptyEntries)
                    If Not (values(2) = "0" Or values(2) = pre_fix) Then
                        Input_Strings.Add(tpsol & RSet(values(2), 7) &
                                         RSet(values(63), 12) & RSet(values(64), 12) &
                                          RSet((Double.Parse(values(1)) / 3600).ToString("F4"), 8) &
                                         " " & line_name)
                        tpsol = 2
                        pre_fix = values(2)
                    End If
                End If
            Next

            IO.File.AppendAllLines(outputfile, Input_Strings)
            'Input_Strings = Nothing
            pre_line_name = line_name
            input_listbox.SetSelected(i, False)
        Next
    End Sub

    Sub Pc_to_scr(input_listbox As System.Windows.Forms.ListBox)
        Dim outputfile As String = ""
        Dim line_name As String
        Dim pre_line_name As String = ""
        Dim seperators(9) As Char
        Dim pre_fix As String = ""

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

        input_listbox.SelectionMode = SelectionMode.One
        input_listbox.Update()

        For i = 0 To input_listbox.Items.Count - 1
            If Strings.Right(input_listbox.Items(i).ToString, 5).ToUpper.Contains(".PC0") Then Continue For
            If Strings.Right(input_listbox.Items(i).ToString, 5).ToUpper.Contains(".PC9") Then Continue For
            input_listbox.SetSelected(i, True)

            line_name = System.IO.Path.GetFileNameWithoutExtension(input_listbox.Items(i).ToString)
            line_name = LSet(line_name, Len(line_name) - 4)

            If Not pre_line_name = line_name Then
                If pre_fix <> "" Then
                    FileOpen(1, outputfile, OpenMode.Append, OpenAccess.Write)
                    PrintLine(1, "")
                    FileClose(1)
                End If
                outputfile = System.IO.Path.GetDirectoryName(input_listbox.Items(i).ToString)
                outputfile = outputfile & "\" & line_name & ".scr"
                FileOpen(1, outputfile, OpenMode.Output, OpenAccess.Write)
                PrintLine(1, "PLINE")
                FileClose(1)
            End If

            Dim Input_Strings As New List(Of String)

            For Each line As String In IO.File.ReadLines(input_listbox.Items(i).ToString)
                If Not (Left(line, 1) = "-" Or Left(line, 1) = "|") Then
                    Dim values(65) As String
                    values = line.Split(seperators, 66, StringSplitOptions.RemoveEmptyEntries)
                    If Not (values(2) = "0" Or values(2) = pre_fix) Then
                        Input_Strings.Add(values(63) & "," & values(64))
                        pre_fix = values(2)
                    End If
                End If
            Next

            IO.File.AppendAllLines(outputfile, Input_Strings)
            'Input_Strings = Nothing
            pre_line_name = line_name
            input_listbox.SetSelected(i, False)
        Next

        If pre_fix <> "" Then
            FileOpen(1, outputfile, OpenMode.Append, OpenAccess.Write)
            PrintLine(1, "")
            FileClose(1)
        End If
    End Sub

    Sub Pc_to_tcpc(input_listbox As System.Windows.Forms.ListBox)
        Dim outputfile As String = ""
        Dim line_name As String
        Dim pre_line_name As String = ""
        Dim seperators(4) As Char
        Dim pre_fix As String = ""

        seperators(0) = Chr(9) 'ControlChars.Tab
        seperators(1) = Chr(32) 'space
        seperators(2) = Chr(44) 'comma
        seperators(3) = Chr(124) '|
        seperators(4) = Chr(45) '-

        input_listbox.SelectionMode = SelectionMode.One
        input_listbox.Update()

        For i = 0 To input_listbox.Items.Count - 1
            If Strings.Right(input_listbox.Items(i).ToString, 5).ToUpper.Contains(".PC0") Then Continue For
            If Strings.Right(input_listbox.Items(i).ToString, 5).ToUpper.Contains(".PC9") Then Continue For
            input_listbox.SetSelected(i, True)

            line_name = System.IO.Path.GetFileNameWithoutExtension(input_listbox.Items(i).ToString)
            line_name = LSet(line_name, Len(line_name) - 4)

            If Not pre_line_name = line_name Then
                outputfile = System.IO.Path.GetDirectoryName(input_listbox.Items(i).ToString)
                outputfile = outputfile & "\" & line_name & ".PC0"
                FileOpen(1, outputfile, OpenMode.Output, OpenAccess.Write)
                PrintLine(1, "| Version 1.36 Position Check File")
                PrintLine(1, "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")
                PrintLine(1, "|  Rec    Time    Fix    |      Antenna Position       |   Antenna Offset                                 |                     Datum                                       |                             Smoothed Datum                                                         |                                               Fixed Offset                                          |                      Cable                       |              Towed Sensor  Using   CMG        |            USBL                                                                                                                                                     | Smooth Datum Position Used |")
                PrintLine(1, "|                        |                             |                                 Relative  True   |                                 Inter Rec                       |                  Time Update  Rec                                  Inter Rec     Offset from Raw   |      Fixed                         Relative  True                                    Inter Rec      | Counter   Manual   Input Vertical  Offset Offset |                               Inter Rec       |                                 Inter Rec       USBL Offset       Towed Offset    Smooth USBL Offset    Smooth USBL Position        Inter Rec       USBL to Towed   |       For Interp           |")
                PrintLine(1, "|  No.    Secs    No.    |       X           Y         |   X        Y    Gyro      Dist   Brg       Brg   |       X            Y          Time    DMG     CMG     Speed     |  CMG      Speed     Time     Time       X              Y         Dist    Brg     Dist      Brg     |   X        Y       Gyro      Dist    Brg     Brg     |     X             Y           Dist   Brg     |  Input    Input    Used   Angle     Dist   Brg   |        X          Y           Dist   Brg      |     X              Y           Dist   Brg       Dist   Brg         Dist   Brg         Dist  Brg          X             Y           Dist   Brg        Dist   Brg     |     X            Y         |")
                PrintLine(1, "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")
                FileClose(1)
            End If

            Dim Input_Strings As New List(Of String)

            For Each line As String In IO.File.ReadLines(input_listbox.Items(i).ToString)
                If Not (Left(line, 1) = "-" Or Left(line, 1) = "|") Then
                    Dim values(3) As String
                    values = line.Split(seperators, 4, StringSplitOptions.RemoveEmptyEntries)
                    If Not (values(2) = "0" Or values(2) = pre_fix) Then
                        Input_Strings.Add(line)
                        pre_fix = values(2)
                    End If
                End If
            Next

            IO.File.AppendAllLines(outputfile, Input_Strings)
            'Input_Strings = Nothing
            pre_line_name = line_name
            input_listbox.SetSelected(i, False)
        Next
    End Sub

    Sub Pc_to_tp(input_listbox As System.Windows.Forms.ListBox)
        Dim tpsol As Byte
        Dim outputfile As String
        Dim line_name As String
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

        input_listbox.SelectionMode = SelectionMode.One
        input_listbox.Update()

        For i = 0 To input_listbox.Items.Count - 1
            input_listbox.SetSelected(i, True)

            line_name = System.IO.Path.GetFileNameWithoutExtension(input_listbox.Items(i).ToString)
            outputfile = System.IO.Path.GetDirectoryName(input_listbox.Items(i).ToString) & "\" & line_name & ".tp"

            tpsol = 1
            Dim Input_Strings As New List(Of String)

            For Each line As String In IO.File.ReadLines(input_listbox.Items(i).ToString)
                If Not (Left(line, 1) = "-" Or Left(line, 1) = "|") Then
                    Dim values(65) As String
                    values = line.Split(seperators, 66, StringSplitOptions.RemoveEmptyEntries)
                    If Not values(2) = "0" Then
                        Input_Strings.Add(tpsol & RSet(values(2), 7) &
                                         RSet(values(63), 12) & RSet(values(64), 12) &
                                          RSet((Double.Parse(values(1)) / 3600).ToString("F4"), 8) &
                                         " " & line_name)
                        tpsol = 2
                    End If
                End If
            Next

            IO.File.WriteAllLines(outputfile, Input_Strings)
            'Input_Strings = Nothing
            input_listbox.SetSelected(i, False)
        Next
    End Sub

End Module