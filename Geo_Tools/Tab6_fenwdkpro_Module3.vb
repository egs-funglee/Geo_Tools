Module Tab6_fenwdkpro_Module3

    Sub SBP_Image_tocad(input_listbox As System.Windows.Forms.ListBox)

        Dim NewDC As Object, A2Kdwg As Object, araster As Object
        Try
            NewDC = GetObject(, "AutoCAD.Application")
#Disable Warning BC42017 ' Late bound resolution
            A2Kdwg = NewDC.ActiveDocument
        Catch exx As Exception
            MsgBox("Error: AutoCAD isn't running?" & vbCrLf & vbCrLf & exx.Message)
            Exit Sub
        End Try

        Dim exception_c As Integer
        'Dim nerr As Integer
        Dim seperators(3) As Char
        seperators(0) = Chr(9) 'ControlChars.Tab
        seperators(1) = Chr(32) 'space
        seperators(2) = Chr(44) 'comma
        seperators(3) = Chr(35) '#
        input_listbox.SelectionMode = SelectionMode.One
        input_listbox.Update()
        Dim fenwdkpro_obj(input_listbox.Items.Count - 1) As Fenwdkpro_hdr

        For i = 0 To input_listbox.Items.Count - 1 Step 1
            fenwdkpro_obj(i) = New Fenwdkpro_hdr
            If Not Strings.Right(input_listbox.Items(i).ToString, 5).ToUpper.Contains("DKPRO") Then Continue For
            input_listbox.SetSelected(i, True)

            Dim inputfile_fenwdkpro As String = input_listbox.Items(i).ToString
            fenwdkpro_obj(i).jpgpath = Strings.Left(inputfile_fenwdkpro, Len(inputfile_fenwdkpro) - 9) & "jpg"
            If Not System.IO.File.Exists(fenwdkpro_obj(i).jpgpath) Then
                MsgBox(inputfile_fenwdkpro & vbCrLf & "Corresponding JPG not found. File skipped", MsgBoxStyle.OkOnly)
                Continue For
            End If

            Dim lines_fenwdkpro() As String = System.IO.File.ReadAllLines(inputfile_fenwdkpro).ToArray()
            Try
                Dim texts(3) As String
                texts = lines_fenwdkpro(0).Split(seperators, 4, StringSplitOptions.RemoveEmptyEntries) 'line1
                If texts(0) = "1" Then
                    fenwdkpro_obj(i).already_updated = True
                    fenwdkpro_obj(i).leftx = Double.Parse(texts(1))
                    fenwdkpro_obj(i).rightx = Double.Parse(texts(2))
                    fenwdkpro_obj(i).ve = Single.Parse(texts(3))

                    texts = lines_fenwdkpro(1).Split(seperators, 4, StringSplitOptions.RemoveEmptyEntries) 'line2
                    fenwdkpro_obj(i).leftfix = Single.Parse(texts(0))
                    fenwdkpro_obj(i).leftpx = Long.Parse(texts(1))

                    texts = lines_fenwdkpro(2).Split(seperators, 4, StringSplitOptions.RemoveEmptyEntries) 'line3
                    fenwdkpro_obj(i).rightfix = Single.Parse(texts(0))
                    fenwdkpro_obj(i).rightpx = Long.Parse(texts(1))

                    texts = lines_fenwdkpro(3).Split(seperators, 4, StringSplitOptions.RemoveEmptyEntries) 'line4
                    fenwdkpro_obj(i).iwidth = Long.Parse(texts(0))

                    texts = lines_fenwdkpro(4).Split(seperators, 4, StringSplitOptions.RemoveEmptyEntries) 'line5
                    fenwdkpro_obj(i).iheight = Long.Parse(texts(0))
                    fenwdkpro_obj(i).vres = Single.Parse(texts(1))
                End If
            Catch Ex As Exception
                exception_c = CInt(MsgBox(inputfile_fenwdkpro & vbCrLf & Ex.Message &
                       vbCrLf & vbCrLf & "Please check the file format, file skipped", MsgBoxStyle.OkOnly))
                fenwdkpro_obj(i).already_updated = False
                Continue For
            End Try

            input_listbox.SetSelected(i, False)
        Next i

        Baseform.BaseProgressBar.Maximum = input_listbox.Items.Count
        Baseform.BaseProgressBar.Value = 0

        For i = 0 To input_listbox.Items.Count - 1 Step 1
            If fenwdkpro_obj(i).already_updated Then
                Baseform.BaseProgressBar.PerformStep()

                'import to autocad
                Try
                    'RetVal = object.AddRaster(ImageFileName, InsertionPoint, ScaleFactor, RotationAngle)
                    Dim inspt() As Double
                    inspt = fenwdkpro_obj(i).Cad_inspt()
                    araster = A2Kdwg.ModelSpace.AddRaster(fenwdkpro_obj(i).jpgpath, inspt, fenwdkpro_obj(i).Cad_width, 0)
                    araster.ImageHeight = fenwdkpro_obj(i).Cad_height
                Catch exx As Exception
                    MsgBox("Cannot import:" & vbCrLf & fenwdkpro_obj(i).jpgpath & vbCrLf & vbCrLf & exx.Message)
                End Try
                'Try
                '    araster.Name = System.IO.Path.GetFileNameWithoutExtension(fenwdkpro_obj(i).jpgpath)
                'Catch exx As Exception
                '    If Not exx.Message = "Invalid key" Then
                '        MsgBox("Cannot import:" & vbCrLf & fenwdkpro_obj(i).jpgpath & vbCrLf & vbCrLf & exx.Message)
                '    End If
                'End Try
                'araster = Nothing

            End If
        Next i
#Enable Warning BC42017 ' Late bound resolution

        Baseform.BaseProgressBar.Value = 1
        Baseform.BaseProgressBar.Maximum = 1
    End Sub

End Module