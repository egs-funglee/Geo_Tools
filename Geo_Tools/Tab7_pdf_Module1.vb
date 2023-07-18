'Imports Microsoft.Office.Interop
#Disable Warning BC42017 'Late bound resolution
#Disable Warning BC42016 'Implicit conversion
Module Tab7_pdf_Module1

    Sub Convert_to_pdf(input_listbox As System.Windows.Forms.ListBox, ByVal overwrite As Boolean, ByVal work_on_xls As Boolean, ByVal xls_output_type As Byte, ByVal extract_svp As Boolean)
        Dim oWord As Object = Nothing 'New Microsoft.Office.Interop.Word.Application
        Dim oXL As Object = Nothing 'As New Microsoft.Office.Interop.Excel.Application
        Dim oDoc As Object 'As Microsoft.Office.Interop.Word.Document
        Dim oWB As Object 'As Microsoft.Office.Interop.Excel.Workbook
        Dim hasDOC As Boolean = False
        Dim hasXLS As Boolean = False

        For Each inputfilename As String In input_listbox.Items
            If inputfilename.ToUpper.ToUpper.Contains(".DOC") Then hasDOC = True
            If inputfilename.ToUpper.ToUpper.Contains(".XLS") Then hasXLS = True
        Next

        Try
            If hasDOC Then oWord = CreateObject("Word.Application")
            If hasXLS Then oXL = CreateObject("Excel.Application")
        Catch ex As Exception
            MessageBox.Show("Error:" & vbCrLf & ex.Message)
            Return
        End Try

        Dim progressbar_value As Integer
        Dim results As New List(Of String)

        Baseform.BaseProgressBar.Maximum = input_listbox.Items.Count
        Baseform.BaseProgressBar.Value = 0
        progressbar_value = 0

        For Each inputfilename As String In input_listbox.Items
            If Strings.Right(inputfilename, 5).ToUpper.Contains(".DOC") Then
                Dim ofn_pdf As String
                Dim dot As Integer
                dot = InStrRev(inputfilename, ".")

                ofn_pdf = Strings.Left(inputfilename, dot) & "pdf"
                If Check_exist(ofn_pdf) = 2 Then
                    MsgBox("File skipped" & vbCrLf & "This output file path is a folder" & vbCrLf & ofn_pdf)
                    GoTo skip_this_file
                End If
                If Check_exist(ofn_pdf) = 1 And overwrite = False Then GoTo skip_this_file

                Try
                    oWord.Visible = False
                    oWord.DisplayAlerts = 0 'Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone
                    oDoc = oWord.Documents.Open(inputfilename.ToString, ReadOnly:=1)
                    oDoc.ExportAsFixedFormat(ofn_pdf, 17, False,'Word.WdExportFormat.wdExportFormatPDF
                             0,'wdExportOptimizeForPrint
                             0, , ,'Word.WdExportRange.wdExportAllDocument
                             0, True, True,'Word.WdExportItem.wdExportDocumentContent
                             0, True, True, False) 'Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks

                    oDoc.Close(SaveChanges:=False)
                    oWord.DisplayAlerts = -1 'Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsAll
                Catch ex As Exception
                    MessageBox.Show("Error:" & vbCrLf & inputfilename & vbCrLf & ex.Message)
                Finally

                End Try
            ElseIf Strings.Right(inputfilename, 5).ToUpper.Contains(".XLS") And work_on_xls Then
                Dim ofn_pdf As String
                Dim dot As Integer
                Dim lastdepth, lasttemp As Decimal

                dot = InStrRev(inputfilename, ".")

                Try
                    oXL.Visible = False
                    oXL.DisplayAlerts = False

                    Select Case xls_output_type
                        Case 1 'ASN/NEC SVP
                            If Not System.IO.Path.GetFileNameWithoutExtension(inputfilename).ToUpper.Contains("SV") Then GoTo skip_this_file
                            oWB = oXL.Workbooks.Open(inputfilename, ReadOnly:=1)

                            If extract_svp Then
                                For Each oSheeti As Object In oWB.Worksheets 'Microsoft.Office.Interop.Excel.Worksheet
                                    Dim oSheet As String = oSheeti.Name.ToString
                                    If oSheet.ToUpper = "PROCESSED" Then
                                        Dim lastrow As Integer
                                        lastrow = CInt(oSheeti.Range("A1").End(-4121).Row) 'Microsoft.Office.Interop.Excel.XlDirection.xlDown
                                        If Decimal.TryParse(oSheeti.Cells(lastrow, 1).Value.ToString, lastdepth) And
                                            Decimal.TryParse(oSheeti.Cells(lastrow, 3).Value.ToString, lasttemp) Then
                                            lastdepth = Decimal.Round(lastdepth, 0)
                                            lasttemp = Decimal.Round(lasttemp, 1)
                                            results.Add(System.IO.Path.GetFileNameWithoutExtension(inputfilename) & "," & lasttemp & "," & lastdepth)
                                        End If
                                    End If
                                Next
                            End If


                            ofn_pdf = Strings.Left(inputfilename, dot) & "pdf"
                            If Check_exist(ofn_pdf) = 2 Then
                                MsgBox("File skipped" & vbCrLf & "This output file path is a folder" & vbCrLf & ofn_pdf)
                                oWB.Close(SaveChanges:=False)
                                InteropReleaseComObject(oWB)
                                GoTo skip_this_file
                            End If
                            If Check_exist(ofn_pdf) = 1 And overwrite = False Then
                                oWB.Close(SaveChanges:=False)
                                InteropReleaseComObject(oWB)
                                GoTo skip_this_file
                            End If
                            Try
                                'MsgBox("trying to hide all sheets")
                                For Each ws As Object In oWB.Worksheets 'Microsoft.Office.Interop.Excel.Worksheet
                                    ws.Visible = 2 'Excel.XlSheetVisibility.xlSheetHidden
                                Next
                                'MsgBox("trying to hide some charts")
                                For Each charti As Object In oWB.Charts 'Excel.Chart
                                    Dim chart As String = charti.Name.ToString
                                    If Not (chart = "Velocity" Or chart = "Temperature") Then
                                        charti.Visible = 2 ' Excel.XlSheetVisibility.xlSheetHidden
                                    End If
                                Next
                                'MsgBox("trying to save the pdf")
                                oXL.ActiveWorkbook.ExportAsFixedFormat(0, 'Excel.XlFixedFormatType.xlTypePDF
                                       ofn_pdf, 0, 'Excel.XlFixedFormatQuality.xlQualityStandard
                                       IncludeDocProperties:=True, IgnorePrintAreas:=False,
                                       OpenAfterPublish:=False)
                            Catch ex As Exception
                                MsgBox("This file is not SVP Excel and is skipped" & vbCrLf & inputfilename) '& vbCrLf & vbCrLf & ex.Message)
                            End Try

                            oWB.Close(SaveChanges:=False)
                            InteropReleaseComObject(oWB)

                        Case 2 'TE SubCom SVP

                            If Not System.IO.Path.GetFileNameWithoutExtension(inputfilename).ToUpper.Contains("SV") Then GoTo skip_this_file
                            oWB = oXL.Workbooks.Open(inputfilename, ReadOnly:=1)
                            If extract_svp Then
                                For Each oSheet As Object In oWB.Worksheets 'Microsoft.Office.Interop.Excel.Worksheet
                                    Dim oSheets As String = oSheet.Name.ToString
                                    If oSheets.ToUpper = "PROCESSED" Then
                                        Dim lastrow As Integer
                                        lastrow = CInt(oSheet.Range("A1").End(-4121).Row) 'Microsoft.Office.Interop.Excel.XlDirection.xlDown
                                        If Decimal.TryParse(oSheet.Cells(lastrow, 1).Value.ToString, lastdepth) And
                                        Decimal.TryParse(oSheet.Cells(lastrow, 3).Value.ToString, lasttemp) Then
                                            lastdepth = Decimal.Round(lastdepth, 0)
                                            lasttemp = Decimal.Round(lasttemp, 1)
                                            results.Add(System.IO.Path.GetFileNameWithoutExtension(inputfilename) & "," & lasttemp & "," & lastdepth)
                                        End If
                                    End If
                                Next
                            End If

                            ofn_pdf = Strings.Left(inputfilename, dot - 1) & "_Velocity.pdf"
                            Dim ofn_pdf2 As String = Strings.Left(inputfilename, dot - 1) & "_Temperature.pdf"
                            Dim ofn_pdf3 As String = Strings.Left(inputfilename, dot - 1) & "_Salinity.pdf"

                            If Check_exist(ofn_pdf) = 2 Or Check_exist(ofn_pdf2) = 2 Or Check_exist(ofn_pdf3) = 2 Then
                                MsgBox("File skipped" & vbCrLf & "This output file path is a folder" & vbCrLf & inputfilename)
                                oWB.Close(SaveChanges:=False)
                                GoTo skip_this_file
                            End If
                            If Check_exist(ofn_pdf) = 1 Or Check_exist(ofn_pdf2) = 1 Or Check_exist(ofn_pdf3) = 1 Then
                                If overwrite = False Then
                                    oWB.Close(SaveChanges:=False)
                                    GoTo skip_this_file
                                End If
                            End If

                            Try
                                Dim rawsheet As Object = oWB.Sheets("Raw") 'CType(oWB.Sheets("Raw"), Excel.Worksheet)
                                Dim sal_range As Object = rawsheet.Range("D:D") 'Excel.Range
                                Dim func As Object = oXL.WorksheetFunction
                                Dim sal_avg As Double = func.Average(sal_range)
                                func = Nothing
                                sal_range = Nothing
                                rawsheet = Nothing

                                For Each charti As Object In oWB.Charts 'Excel.Chart
                                    Dim chart As String = charti.Name.ToString
                                    If chart = "Velocity" Then
                                        charti.ExportAsFixedFormat(0, 'Excel.XlFixedFormatType.xlTypePDF
                                                   ofn_pdf, 0, 'Excel.XlFixedFormatQuality.xlQualityStandard
                                                   IncludeDocProperties:=True, IgnorePrintAreas:=False,
                                                   OpenAfterPublish:=False)
                                    End If
                                    If chart = "Temperature" Then
                                        charti.ExportAsFixedFormat(0, 'Excel.XlFixedFormatType.xlTypePDF
                                                   ofn_pdf2, 0, 'Excel.XlFixedFormatQuality.xlQualityStandard
                                                   IncludeDocProperties:=True, IgnorePrintAreas:=False,
                                                   OpenAfterPublish:=False)
                                    End If
                                    If chart = "Salinity" And sal_avg > 0 Then
                                        charti.ExportAsFixedFormat(0, 'Excel.XlFixedFormatType.xlTypePDF
                                                   ofn_pdf3, 0, 'Excel.XlFixedFormatQuality.xlQualityStandard
                                                   IncludeDocProperties:=True, IgnorePrintAreas:=False,
                                                   OpenAfterPublish:=False)
                                    End If
                                Next
                            Catch ex As Exception
                                MsgBox("This file is not SVP Excel and is skipped" & vbCrLf & inputfilename)
                            End Try

                            oWB.Close(SaveChanges:=False)
                        Case 3 'Last Active Sheets

                            ofn_pdf = Strings.Left(inputfilename, dot) & "pdf"
                            If Check_exist(ofn_pdf) = 2 Then
                                MsgBox("File skipped" & vbCrLf & "This output file path is a folder" & vbCrLf & ofn_pdf)
                                GoTo skip_this_file
                            End If
                            If Check_exist(ofn_pdf) = 1 And overwrite = False Then
                                GoTo skip_this_file
                            End If
                            oWB = oXL.Workbooks.Open(inputfilename, ReadOnly:=1)
                            Dim ischart As Boolean = False
                            Try
                                Dim aobject As Object = oXL.ActiveSheet 'Excel.Worksheet
                                'aobject = DirectCast(oXL.ActiveSheet, Excel.Worksheet)
                                aobject.ExportAsFixedFormat(0, 'Excel.XlFixedFormatType.xlTypePDF
                                           ofn_pdf, 0, 'Excel.XlFixedFormatQuality.xlQualityStandard
                                           IncludeDocProperties:=True, IgnorePrintAreas:=False,
                                           OpenAfterPublish:=False)
                            Catch
                                ischart = True
                            End Try
                            If ischart Then
                                Try
                                    Dim aobject As Object = oXL.ActiveSheet 'Excel.Chart
                                    'aobject = DirectCast(oXL.ActiveSheet, Excel.Chart)
                                    aobject.ExportAsFixedFormat(0, 'Excel.XlFixedFormatType.xlTypePDF
                                               ofn_pdf, 0, 'Excel.XlFixedFormatQuality.xlQualityStandard
                                               IncludeDocProperties:=True, IgnorePrintAreas:=False,
                                               OpenAfterPublish:=False)
                                Catch
                                End Try
                            End If
                            oWB.Close(SaveChanges:=False)

                        Case 4 'All Sheets

                            ofn_pdf = Strings.Left(inputfilename, dot) & "pdf"
                            If Check_exist(ofn_pdf) = 2 Then
                                MsgBox("File skipped" & vbCrLf & "This output file path is a folder" & vbCrLf & ofn_pdf)
                                GoTo skip_this_file
                            End If
                            If Check_exist(ofn_pdf) = 1 And overwrite = False Then
                                GoTo skip_this_file
                            End If
                            oWB = oXL.Workbooks.Open(inputfilename, ReadOnly:=1)

                            ofn_pdf = Strings.Left(inputfilename, dot) & "pdf"
                            If Check_exist(ofn_pdf) = 2 Then
                                MsgBox("File skipped" & vbCrLf & "This output file path is a folder" & vbCrLf & ofn_pdf)
                                GoTo skip_this_file
                            End If
                            If Check_exist(ofn_pdf) = 1 And overwrite = False Then GoTo skip_this_file

                            oXL.ActiveWorkbook.ExportAsFixedFormat(0, 'Excel.XlFixedFormatType.xlTypePDF
                                       ofn_pdf, 0, 'Excel.XlFixedFormatQuality.xlQualityStandard
                                       IncludeDocProperties:=True, IgnorePrintAreas:=False,
                                       OpenAfterPublish:=False)
                            oWB.Close(SaveChanges:=False)
                            InteropReleaseComObject(oWB)
                    End Select
                    oXL.DisplayAlerts = True
                Catch ex As Exception
                    MessageBox.Show("Error:" & vbCrLf & inputfilename & vbCrLf & ex.Message)
                End Try
            End If
skip_this_file:
            progressbar_value += 1
            Baseform.BaseProgressBar.Value = progressbar_value
        Next

        If hasDOC Then oWord.Quit()
        If hasXLS Then oXL.Quit()

        If results.Count > 0 Then
            Dim svp_info_path As String = System.IO.Path.GetTempPath & "SVP_Info.csv"
            System.IO.File.WriteAllText(svp_info_path, "SVP Name,Bottom Temperature,Bottom Depth,Remark,Probe Type,Easting,Northing" & vbCrLf)
            System.IO.File.AppendAllLines(svp_info_path, results.ToArray)
            Try
                oXL = CreateObject("Excel.Application")
                oXL.Visible = True
                oXL.DisplayAlerts = True
                oWB = oXL.Workbooks.Open(svp_info_path, ReadOnly:=1)
            Catch ex As Exception
                MessageBox.Show("Error:" & vbCrLf & ex.Message)
                Return
            End Try
        End If

    End Sub

    Sub Get_svp_info(input_listbox As System.Windows.Forms.ListBox)
        Dim lastdepth, lasttemp As Decimal
        Dim results As New List(Of String)

        Baseform.BaseProgressBar.Maximum = input_listbox.Items.Count
        Baseform.BaseProgressBar.Value = 0

        Dim inputfilenames() As String
        ReDim inputfilenames(input_listbox.Items.Count - 1)
        input_listbox.Items.CopyTo(inputfilenames, 0)
        Dim filecount As Integer
        filecount = 0

        Parallel.ForEach(inputfilenames,
                         Sub(currentFile)
                             If System.IO.Path.GetFileNameWithoutExtension(currentFile).ToUpper.Contains("SV") And Strings.Right(currentFile, 5).ToUpper.Contains(".XLS") Then
                                 Dim oXL As Object 'New Microsoft.Office.Interop.Excel.Application
                                 Dim oWB As Object 'Microsoft.Office.Interop.Excel.Workbook
                                 Try
                                     oXL = CreateObject("Excel.Application")
                                 Catch ex As Exception
                                     MessageBox.Show("Error:" & vbCrLf & ex.Message)
                                     Return
                                 End Try

                                 oXL.Visible = False
                                 oXL.DisplayAlerts = False
                                 oWB = oXL.Workbooks.Open(currentFile, ReadOnly:=1)
                                 Dim aaa As Object = oWB.Worksheets

                                 For Each oSheeti As Object In oWB.Worksheets 'Microsoft.Office.Interop.Excel.Worksheet
                                     Dim oSheet As String = oSheeti.Name.ToString
                                     If oSheet.ToUpper = "PROCESSED" Then
                                         Dim lastrow As Integer
                                         lastrow = CInt(oSheeti.Range("A1").End(-4121).Row) 'Microsoft.Office.Interop.Excel.XlDirection.xlDown
                                         If Decimal.TryParse(oSheeti.Cells(lastrow, 1).Value.ToString, lastdepth) And
                                             Decimal.TryParse(oSheeti.Cells(lastrow, 3).Value.ToString, lasttemp) Then
                                             lastdepth = Decimal.Round(lastdepth, 0)
                                             lasttemp = Decimal.Round(lasttemp, 1)
                                             results.Add(System.IO.Path.GetFileNameWithoutExtension(currentFile) & "," & lasttemp & "," & lastdepth)
                                         End If
                                     End If
                                 Next
                                 oWB.Close(SaveChanges:=False)
                                 oXL.Quit()
                             End If
                             filecount += 1
                             Baseform.BaseProgressBar.Value = filecount
                         End Sub)
        Baseform.BaseProgressBar.Value = filecount
        results.Sort(Function(xx, yy) xx.CompareTo(yy))

        If results.Count > 0 Then
            Dim svp_info_path As String = System.IO.Path.GetTempPath & "SVP_Info.csv"
            System.IO.File.WriteAllText(svp_info_path, "SVP Name,Bottom Temperature,Bottom Depth,Remark,Probe Type,Easting,Northing" & vbCrLf)
            System.IO.File.AppendAllLines(svp_info_path, results.ToArray)
            Dim oXL As Object 'New Microsoft.Office.Interop.Excel.Application
            Dim oWB As Object 'Microsoft.Office.Interop.Excel.Workbook
            Try
                oXL = CreateObject("Excel.Application")
                oXL.Visible = True
                oXL.DisplayAlerts = True
                oWB = oXL.Workbooks.Open(svp_info_path, ReadOnly:=1)
            Catch ex As Exception
                MessageBox.Show("Error:" & vbCrLf & ex.Message)
                Return
            End Try
        End If

    End Sub

    Sub Remove_selected(input_listbox As System.Windows.Forms.ListBox)
        Do While (input_listbox.SelectedItems.Count > 0)
            input_listbox.Items.Remove(input_listbox.SelectedItem)
        Loop
        If input_listbox.Items.Count = 0 Then Baseform.Init_listbox(input_listbox)
    End Sub

End Module