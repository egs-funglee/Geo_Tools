Imports BitMiracle.LibTiff.Classic

Module Tab4_tiff_Module2

    Sub Geotiff_tocad(input_listbox As System.Windows.Forms.ListBox)
        Dim NewDC As Object, A2Kdwg As Object, araster As Object
        Dim tif_inspt() As Double
#Disable Warning BC42017 ' Late bound resolution
        Try
            NewDC = GetObject(, "AutoCAD.Application")
            A2Kdwg = NewDC.ActiveDocument
        Catch exx As Exception
            MsgBox("Error: AutoCAD isn't running?" & vbCrLf & vbCrLf & exx.Message)
            Exit Sub
        End Try

        Dim tifpath As String, tfwpath As String
        Dim iwidth As Integer, iheight As Integer
        Dim tfws() As String, tfwd() As Double
        Dim i As Integer
        Dim tfw_exists As Boolean

        Baseform.BaseProgressBar.Maximum = input_listbox.Items.Count
        Baseform.BaseProgressBar.Value = 0

        input_listbox.SelectionMode = SelectionMode.One
        input_listbox.Update()

        For i = 0 To input_listbox.Items.Count - 1
            If Not System.IO.File.Exists(input_listbox.Items(i).ToString) Then
                Baseform.BaseProgressBar.PerformStep()
                Continue For 'next iteration
            End If

            input_listbox.SetSelected(i, True)
            tifpath = input_listbox.Items(i).ToString
            tfwpath = input_listbox.Items(i).ToString.Substring(0, input_listbox.Items(i).ToString.Length - 3) + "tfw"

            '========================================================
            'try to load tiff image and see if there is geotiff tags
            '========================================================

            Dim timage As Tiff = Tiff.Open(tifpath, "r")
            If timage IsNot Nothing Then
                Dim value As FieldValue()
                Dim isgeotiff As Boolean = True
                Dim doubles(5) As Double, tempdbls() As Double

                value = timage.GetField(TiffTag.IMAGEWIDTH)
                iwidth = value(0).ToInt

                value = timage.GetField(TiffTag.IMAGELENGTH)
                iheight = value(0).ToInt

                value = timage.GetField(DirectCast(33550, TiffTag))
                If value IsNot Nothing Then
                    tempdbls = value(1).ToDoubleArray
                    doubles(0) = tempdbls(0) 'x-component of the pixel width (x-scale)
                    doubles(3) = 0 - tempdbls(1) 'y-component of the pixel width (y-skew)
                Else
                    isgeotiff = False
                End If

                value = timage.GetField(DirectCast(33922, TiffTag))
                If value IsNot Nothing Then
                    tempdbls = value(1).ToDoubleArray
                    doubles(4) = tempdbls(3) 'x of the upper left pixel transformed to the map
                    doubles(5) = tempdbls(4) 'y of the upper left pixel transformed to the map
                Else
                    isgeotiff = False
                End If
                timage.Close()
                Erase value
                Erase tempdbls

                If isgeotiff Then
                    '====================================================
                    'if tiff image with geotiff tags
                    '====================================================
                    ReDim tif_inspt(0 To 2) 'BL
                    tif_inspt(0) = doubles(4) 'L
                    tif_inspt(1) = doubles(5) + (iheight * doubles(3)) 'B = T + [ iheight * vscale(negative) ]
                    tif_inspt(2) = 0
                    Try
                        araster = A2Kdwg.ModelSpace.AddRaster(tifpath, tif_inspt, iwidth * doubles(0), 0)
                        araster.ImageHeight = -iheight * doubles(3)
                        araster = Nothing
                    Catch exx As Exception
                        MsgBox("Cannot import:" & vbCrLf & tifpath & vbCrLf & vbCrLf & exx.Message)
                    End Try
                    Erase doubles
                Else
                    '====================================================
                    'tiff image without geotiff tags, try tfw
                    '====================================================
                    ReDim tfwd(5)
                    If System.IO.File.Exists(tfwpath) Then
                        tfw_exists = True
                        tfws = IO.File.ReadAllLines(tfwpath)
                        ReDim Preserve tfws(5)
                        Try
                            tfwd = Array.ConvertAll(tfws, New Converter(Of String, Double)(AddressOf Double.Parse))
                        Catch err As Exception
                            tfw_exists = False
                            MsgBox("Cannot load tfw:" & vbCrLf & tifpath & vbCrLf & vbCrLf & err.Message)
                            IO.File.Move(tfwpath, tfwpath + "_Incorrect_TFW") 'Maybe error here due to rights
                        End Try
                        Erase tfws
                    Else
                        tfw_exists = False
                    End If

                    If tfw_exists Then
                        ReDim tif_inspt(0 To 2) 'BL
                        tif_inspt(0) = tfwd(4) - tfwd(0) / 2 'L
                        tif_inspt(1) = tfwd(5) - tfwd(3) / 2 + iheight * tfwd(3) 'B
                        tif_inspt(2) = 0 'tif_width = ex - sx 'Width = R - L = sx + iwidth * tfwd(0) - sx

                        Try 'import to autocad
                            araster = A2Kdwg.ModelSpace.AddRaster(tifpath, tif_inspt, iwidth * tfwd(0), 0)
                            araster.ImageHeight = -iheight * tfwd(3) 'Height = T - B || sy - ey || sy - sy - iheight * tfwd(3)
                            araster = Nothing
                        Catch exx As Exception
                            MsgBox("Cannot import:" & vbCrLf & tifpath & vbCrLf & vbCrLf & exx.Message)
                        End Try
                    Else
                        MsgBox("Cannot import:" & vbCrLf & tifpath & vbCrLf & vbCrLf & "Not Geotiff and no TFW")
                    End If
                End If
            Else
                If tifpath.ToLower.EndsWith(".jpg") Or tifpath.ToLower.EndsWith(".png") Then 'For GeoJPG

                    'open image from filestream without loading whole image into memory
                    Dim tifpaths As New System.IO.FileStream(tifpath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite)
                    Dim oimage As System.Drawing.Image
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

                        tfwpath = input_listbox.Items(i).ToString.Substring(0, input_listbox.Items(i).ToString.Length - 2) + "gw"

                        '====================================================
                        'try jgw same as tfw
                        '====================================================
                        ReDim tfwd(5)
                        If System.IO.File.Exists(tfwpath) Then
                            tfw_exists = True
                            tfws = IO.File.ReadAllLines(tfwpath)
                            ReDim Preserve tfws(5)
                            Try
                                tfwd = Array.ConvertAll(tfws, New Converter(Of String, Double)(AddressOf Double.Parse))
                            Catch err As Exception
                                tfw_exists = False
                                MsgBox("Cannot load world file:" & vbCrLf & tifpath & vbCrLf & vbCrLf & err.Message)
                                IO.File.Move(tfwpath, tfwpath + "_Incorrect_World_file") 'Maybe error here due to rights
                            End Try
                            Erase tfws
                        Else
                            tfw_exists = False
                        End If

                        If tfw_exists Then
                            ReDim tif_inspt(0 To 2) 'BL
                            tif_inspt(0) = tfwd(4) - tfwd(0) / 2 'L
                            tif_inspt(1) = tfwd(5) - tfwd(3) / 2 + iheight * tfwd(3) 'B
                            tif_inspt(2) = 0 'tif_width = ex - sx 'Width = R - L = sx + iwidth * tfwd(0) - sx

                            Try 'import to autocad
                                araster = A2Kdwg.ModelSpace.AddRaster(tifpath, tif_inspt, iwidth * tfwd(0), 0)
                                araster.ImageHeight = -iheight * tfwd(3) 'Height = T - B || sy - ey || sy - sy - iheight * tfwd(3)
                                araster = Nothing
                            Catch exx As Exception
                                MsgBox("Cannot import:" & vbCrLf & tifpath & vbCrLf & vbCrLf & exx.Message)
                            End Try
                        Else
                            MsgBox("Cannot import:" & vbCrLf & tifpath & vbCrLf & vbCrLf & "Not GeoJPG no JGW")
                        End If
                    Else
                        MsgBox("Cannot import:" & vbCrLf & tifpath & vbCrLf & vbCrLf & "Not JPG")
                    End If
                    tifpaths.Dispose()
                Else
                    MsgBox("Cannot open:" & vbCrLf & tifpath)
                End If
            End If

            Baseform.BaseProgressBar.PerformStep()
            input_listbox.SetSelected(i, False)
        Next i
#Enable Warning BC42017 ' Late bound resolution
        NewDC = Nothing
        A2Kdwg = Nothing
    End Sub

End Module