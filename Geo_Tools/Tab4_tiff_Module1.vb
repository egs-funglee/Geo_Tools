Imports BitMiracle.LibTiff.Classic

Module Tab4_tiff_Module1

    'Private Function MyCStr(ByVal d As Double) As String
    '    Return CStr(d)
    'End Function

    Sub Geotiff_boundary_maker(input_listbox As System.Windows.Forms.ListBox)

        Dim tifpath As String, tfwpath As String, tifname As String
        Dim iwidth As Integer, iheight As Integer
        Dim sx As Double, sy As Double 'NW
        Dim ex As Double, ey As Double 'ES
        Dim outputs() As String, tfws() As String, tfwd() As Double
        Dim i As Integer, slash As Integer
        Dim j As Integer = 3
        Dim oimage As System.Drawing.Image
        Dim tfw_exists As Boolean

        Baseform.BaseProgressBar.Maximum = input_listbox.Items.Count
        Baseform.BaseProgressBar.Value = 0

        'makeing script file
        slash = input_listbox.Items(0).ToString.LastIndexOf("\") + 1
        Dim scrpath As String = input_listbox.Items(0).ToString.Substring(0, slash) + "GeoTiff_Boundary.scr"
        i = (input_listbox.Items.Count + 1) * 3 - 1
        ReDim outputs(0 To i) ' resize to tfw i
        outputs(0) = "-TEXT STYLE Standard 0,0 "

        input_listbox.SelectionMode = SelectionMode.One
        input_listbox.Update()

        For i = 0 To input_listbox.Items.Count - 1
            If Not System.IO.File.Exists(input_listbox.Items(i).ToString) Then
                Baseform.BaseProgressBar.PerformStep()
                Continue For 'next iteration
            End If

            input_listbox.SetSelected(i, True)
            tifpath = input_listbox.Items(i).ToString
            Dim fext As String = "tfw"
            If tifpath.ToLower.EndsWith(".jpg") Then fext = "jgw"
            tfwpath = input_listbox.Items(i).ToString.Substring(0, input_listbox.Items(i).ToString.Length - 3) + fext
            tifname = tifpath.Substring(slash, tifpath.Length - slash - 4)

            'tfw, get scale and insertion xy
            Try
                If Not System.IO.File.Exists(tfwpath) And tfwpath.EndsWith(".tfw") Then Tfw_maker(tifpath)
            Catch exx As Exception
                MsgBox("Cannot find BitMiracle.LibTiff.NET.dll" & vbCrLf &
                       "Please put it with this EXE file." & vbCrLf & vbCrLf & exx.Message)
                Exit Sub
            End Try

            tfw_exists = False
            ReDim tfwd(5)
            If System.IO.File.Exists(tfwpath) Then 'the file may not be created by tfw_maker
                tfw_exists = True
                tfws = IO.File.ReadAllLines(tfwpath)
                ReDim Preserve tfws(5)
                Try
                    tfwd = Array.ConvertAll(tfws, New Converter(Of String, Double)(AddressOf Double.Parse))
                Catch err As Exception 'remake TFW
                    tfw_exists = False
                    IO.File.Move(tfwpath, tfwpath + "_Incorrect") 'Maybe error here due to rights
                    If tfwpath.EndsWith(".tfw") Then
                        Tfw_maker(tifpath)
                        If System.IO.File.Exists(tfwpath) Then
                            tfw_exists = True
                            tfws = IO.File.ReadAllLines(tfwpath)
                            ReDim Preserve tfws(5)
                            tfwd = Array.ConvertAll(tfws, New Converter(Of String, Double)(AddressOf Double.Parse))
                        End If
                    Else
                        IO.File.WriteAllText(tifpath + "_.txt", "Unable to Get info from JGW file" & vbCrLf)
                    End If
                End Try
                Erase tfws
            Else
                If tfwpath.EndsWith(".jgw") Then
                    IO.File.WriteAllText(tifpath + "_.txt", "No JGW file" & vbCrLf)
                End If
            End If

            'tif, get image size for boundary. tfw file must exist at this stage
            If tfw_exists Then
                'system.io.filestream from string, *may be error here due to file access
                Dim tifpaths As New System.IO.FileStream(tifpath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite)

                'open image from filestream without loading whole image into memory
                Try
                    oimage = System.Drawing.Image.FromStream(tifpaths, False, False)
                Catch err As Exception
                    oimage = Nothing
                    IO.File.AppendAllText(
                        tifpath + "_.txt", "The Image file appears to be corrupt, unable to get image dimension." & vbCrLf)
                End Try

                If oimage IsNot Nothing Then
                    'image dimension
                    iwidth = oimage.Width
                    iheight = oimage.Height
                    oimage.Dispose()

                    'start end points
                    sx = tfwd(4) - tfwd(0) / 2
                    sy = tfwd(5) - tfwd(3) / 2
                    ex = sx + iwidth * tfwd(0)
                    ey = sy + iheight * tfwd(3)

                    'set strings into output string array
                    outputs(j) = ";" & tifname & " , Pixel Size (m) :" & tfwd(0)
                    j += 1
                    outputs(j) = "-TEXT " & sx & "," & sy + 50 & " 100 0 " & tifname
                    j += 1
                    outputs(j) = "RECTANG " & sx & "," & sy & " " & ex & "," & ey
                    j += 1
                End If
                tifpaths.Dispose()
            End If
            Baseform.BaseProgressBar.PerformStep()
            input_listbox.SetSelected(i, False)
        Next i
        ReDim Preserve outputs(j - 1)
        TryWriteAllLines(scrpath, outputs)
    End Sub

    Sub Tfw_maker(tiffpath As String) 'make tfw file from geotiff tags
        Dim image As Tiff = Tiff.Open(tiffpath, "r")
        If image IsNot Nothing Then
            Dim isgeotiff As Boolean = True
            'Dim bytes() As Byte
            Dim doubles(5) As Double, tempdbls() As Double
            Dim value As FieldValue()
            value = image.GetField(DirectCast(33550, TiffTag))
            If value IsNot Nothing Then
                tempdbls = value(1).ToDoubleArray
                doubles(0) = tempdbls(0) 'x-component of the pixel width (x-scale)
                doubles(3) = 0 - tempdbls(1) 'y-component of the pixel width (y-skew)
                'bytes = value(1).GetBytes
                'doubles(0) = BitConverter.ToDouble(bytes, 0) 'x-component of the pixel width (x-scale)
                'doubles(3) = 0 - BitConverter.ToDouble(bytes, 8) 'y-component of the pixel width (y-skew)
            Else
                isgeotiff = False
            End If
            value = image.GetField(DirectCast(33922, TiffTag))
            If value IsNot Nothing Then
                tempdbls = value(1).ToDoubleArray
                doubles(4) = tempdbls(3) + (doubles(0) / 2) 'x of the center of the upper left pixel transformed to the map
                doubles(5) = tempdbls(4) + (doubles(3) / 2) 'y of the center of the upper left pixel transformed to the map
                'bytes = value(1).GetBytes
                'doubles(4) = BitConverter.ToDouble(bytes, 24) + doubles(0) / 2 'x of the center of the upper left pixel transformed to the map
                'doubles(5) = BitConverter.ToDouble(bytes, 32) + doubles(3) / 2 'y of the center of the upper left pixel transformed to the map
            Else
                isgeotiff = False
            End If
            image.Close()
            'Erase bytes
            Erase value
            Erase tempdbls
            If isgeotiff Then
                Dim outstrings(5) As String
                outstrings = Array.ConvertAll(doubles, Function(input As Double) CStr(input))
                'outstrings = Array.ConvertAll(doubles, New System.Converter(Of Double, String)(AddressOf MyCStr))
                'outstrings = Array.ConvertAll(doubles, New System.Converter(Of Double, String)(AddressOf Double.Parse))
                IO.File.WriteAllLines(tiffpath.Substring(0, tiffpath.Length - 3) + "tfw", outstrings)
                Erase outstrings
            Else
                IO.File.WriteAllText(
                    tiffpath + "_.txt", "The file doesn't appear to be Geotiff, unable to get geotiff tags." & vbCrLf)
            End If
            Erase doubles
        Else
            IO.File.WriteAllText(
                    tiffpath + "_.txt", "The Tiff file appears to be corrupt, unable to open." & vbCrLf)
        End If
    End Sub

End Module