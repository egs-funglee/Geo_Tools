Module Tab6_fenwdkpro_Module2

    Sub GetFixPos_fom_ActiveLayer(ByRef inputfiletb As System.Windows.Forms.TextBox)
        Dim NewDC As Object, A2Kdwg As Object
#Disable Warning BC42017 ' Late bound resolution
        Try
            NewDC = GetObject(, "AutoCAD.Application")
            A2Kdwg = NewDC.ActiveDocument
        Catch exx As Exception
            MsgBox("Error: AutoCAD isn't running?" & vbCrLf & vbCrLf & exx.Message)
            Exit Sub
        End Try

        Dim actlayername As String
        actlayername = CType(A2Kdwg.ActiveLayer.name, String)

        Dim pfSS As Object
        pfSS = A2Kdwg.PickfirstSelectionSet

        pfSS.Clear

        Dim gpCode(1) As Short
        Dim dataValue(1) As Object
        gpCode(0) = 8
        gpCode(1) = 0
        dataValue(0) = actlayername
        dataValue(1) = "Text"

        pfSS.Select(5, , , gpCode, dataValue) 'select all text in active layer to ss
        If CType(pfSS.Count, Integer) = 0 Then
            MsgBox("No text to be searched in Active Layer:" & vbLf & actlayername)
            GoTo Exit_Sub
        End If

        Dim inspt() As Double
        Dim fix_count As Integer = 0
        Dim outputstr(CType(pfSS.Count, Integer) - 1) As String
        Dim fixno As String

        For i = 0 To CType(pfSS.Count, Integer) - 1
            fixno = CType(pfSS(i).TextString, String)
            If IsNumeric(fixno) Then
                inspt = CType(pfSS(i).TextAlignmentPoint, Double())
                outputstr(fix_count) =
                Format(inspt(0), "0.000") & "," & Format(inspt(1), "0.000") & "," & fixno & ",0"
                fix_count += 1
            End If
        Next

        'pfSS = Nothing
        If fix_count = 0 Then
            MsgBox("Cannot find any Text which look like Fix Number")
            GoTo Exit_Sub
        End If

        ReDim Preserve outputstr(fix_count - 1)

        actlayername = actlayername & Format(Now(), " yyyyMMdd_hhmmss") & ".asc"
        inputfiletb.Text = WriteAllLineD(A2Kdwg.Path.ToString, actlayername, outputstr)
        If Len(inputfiletb.Text) > 0 Then
            MsgBox(inputfiletb.Text & vbCrLf & "has been set for input")
        End If

Exit_Sub:
        'NewDC = Nothing
        'A2Kdwg = Nothing
#Enable Warning BC42017 ' Late bound resolution
    End Sub

    Private Function WriteAllLineD(ByVal activedwgpath As String, ByVal ascfilename As String, ByRef writestrings As String()) As String
        Dim outputfilepath As String = activedwgpath & "\" & ascfilename
        Try
            IO.File.WriteAllLines(outputfilepath, writestrings)
            WriteAllLineD = outputfilepath
            Baseform.trigger_textbox = True
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & vbCrLf & "Cannot write the ASC file" & vbCrLf & outputfilepath & vbCrLf & "Writing to Desktop Instead")
            outputfilepath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & ascfilename
            Try
                IO.File.WriteAllLines(outputfilepath, writestrings)
                WriteAllLineD = outputfilepath
                Baseform.trigger_textbox = True
            Catch ex2 As Exception
                MsgBox(ex2.Message & vbCrLf & vbCrLf & "Cannot write the ASC file to Desktop" & vbCrLf & outputfilepath)
                WriteAllLineD = ""
            End Try
        End Try
    End Function

End Module