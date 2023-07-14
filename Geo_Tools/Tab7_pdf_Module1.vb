Imports Microsoft.Office.Interop

Module Tab7_pdf_Module1

    Sub Convert_to_pdf(input_listbox As System.Windows.Forms.ListBox, ByVal overwrite As Boolean, ByVal work_on_xls As Boolean, ByVal xls_output_type As Byte, ByVal extract_svp As Boolean)
        Dim oWord As New Microsoft.Office.Interop.Word.Application
        Dim oXL As New Microsoft.Office.Interop.Excel.Application
        Dim oDoc As Microsoft.Office.Interop.Word.Document
        Dim oWB As Microsoft.Office.Interop.Excel.Workbook
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
                    oWord.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone

                    oDoc = oWord.Documents.Open(inputfilename.ToString, ReadOnly:=1)
                    oDoc.ExportAsFixedFormat(ofn_pdf, Word.WdExportFormat.wdExportFormatPDF, False,
                             Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                             Word.WdExportRange.wdExportAllDocument, , ,
                             Word.WdExportItem.wdExportDocumentContent, True, True,
                             Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks, True, True, False)

                    oDoc.Close(SaveChanges:=False)
                    InteropReleaseComObject(oDoc)
                Catch ex As Exception
                    MessageBox.Show("Error:" & vbCrLf & inputfilename & vbCrLf & ex.Message)
                Finally
                    oWord.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsAll

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
                                For Each oSheet As Microsoft.Office.Interop.Excel.Worksheet In oWB.Worksheets
                                    If oSheet.Name.ToUpper = "PROCESSED" Then
                                        Dim lastrow As Integer
                                        lastrow = oSheet.Range("A1").End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
                                        If Decimal.TryParse(CType(oSheet.Cells(lastrow, 1), Excel.Range).Value.ToString, lastdepth) And
                                        Decimal.TryParse(CType(oSheet.Cells(lastrow, 3), Excel.Range).Value.ToString, lasttemp) Then
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
                                For Each ws As Microsoft.Office.Interop.Excel.Worksheet In oWB.Worksheets
                                    ws.Visible = Excel.XlSheetVisibility.xlSheetHidden
                                Next
                                'MsgBox("trying to hide some charts")
                                For Each chart As Excel.Chart In oWB.Charts
                                    If Not (chart.Name = "Velocity" Or chart.Name = "Temperature") Then
                                        chart.Visible = Excel.XlSheetVisibility.xlSheetHidden
                                    End If
                                Next
                                'MsgBox("trying to save the pdf")
                                oXL.ActiveWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF,
                                       ofn_pdf, Excel.XlFixedFormatQuality.xlQualityStandard,
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
                                For Each oSheet As Microsoft.Office.Interop.Excel.Worksheet In oWB.Worksheets
                                    If oSheet.Name.ToUpper = "PROCESSED" Then
                                        Dim lastrow As Integer
                                        lastrow = oSheet.Range("A1").End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
                                        If Decimal.TryParse(CType(oSheet.Cells(lastrow, 1), Excel.Range).Value.ToString, lastdepth) And
                                        Decimal.TryParse(CType(oSheet.Cells(lastrow, 3), Excel.Range).Value.ToString, lasttemp) Then
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
                                InteropReleaseComObject(oWB)
                                GoTo skip_this_file
                            End If
                            If Check_exist(ofn_pdf) = 1 Or Check_exist(ofn_pdf2) = 1 Or Check_exist(ofn_pdf3) = 1 Then
                                If overwrite = False Then
                                    oWB.Close(SaveChanges:=False)
                                    InteropReleaseComObject(oWB)
                                    GoTo skip_this_file
                                End If
                            End If

                            Try

                                Dim rawsheet As Excel.Worksheet = CType(oWB.Sheets("Raw"), Excel.Worksheet)
                                Dim sal_range As Excel.Range = rawsheet.Range("D:D")
                                Dim func As Excel.WorksheetFunction
                                func = oXL.WorksheetFunction
                                Dim testval As Integer = Integer.Parse(func.Average(sal_range).ToString("F0"))
                                func = Nothing
                                sal_range = Nothing
                                rawsheet = Nothing

                                For Each chart As Excel.Chart In oWB.Charts
                                    If chart.Name = "Velocity" Then
                                        chart.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF,
                                                   ofn_pdf, Excel.XlFixedFormatQuality.xlQualityStandard,
                                                   IncludeDocProperties:=True, IgnorePrintAreas:=False,
                                                   OpenAfterPublish:=False)
                                    End If
                                    If chart.Name = "Temperature" Then
                                        chart.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF,
                                                   ofn_pdf2, Excel.XlFixedFormatQuality.xlQualityStandard,
                                                   IncludeDocProperties:=True, IgnorePrintAreas:=False,
                                                   OpenAfterPublish:=False)
                                    End If
                                    If chart.Name = "Salinity" And testval > 0 Then
                                        chart.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF,
                                                   ofn_pdf3, Excel.XlFixedFormatQuality.xlQualityStandard,
                                                   IncludeDocProperties:=True, IgnorePrintAreas:=False,
                                                   OpenAfterPublish:=False)
                                    End If
                                Next
                            Catch ex As Exception
                                MsgBox("This file is not SVP Excel and is skipped" & vbCrLf & inputfilename)
                                'MsgBox("This file is not SVP Excel" & vbCrLf & inputfilename & vbCrLf & vbCrLf & ex.Message)
                            End Try

                            oWB.Close(SaveChanges:=False)
                            InteropReleaseComObject(oWB)
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
                                Dim aobject As Excel.Worksheet
                                aobject = DirectCast(oXL.ActiveSheet, Excel.Worksheet)
                                aobject.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF,
                                           ofn_pdf, Excel.XlFixedFormatQuality.xlQualityStandard,
                                           IncludeDocProperties:=True, IgnorePrintAreas:=False,
                                           OpenAfterPublish:=False)
                            Catch
                                ischart = True
                            End Try
                            If ischart Then
                                Try
                                    Dim aobject As Excel.Chart
                                    aobject = DirectCast(oXL.ActiveSheet, Excel.Chart)
                                    aobject.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF,
                                               ofn_pdf, Excel.XlFixedFormatQuality.xlQualityStandard,
                                               IncludeDocProperties:=True, IgnorePrintAreas:=False,
                                               OpenAfterPublish:=False)
                                Catch
                                End Try
                            End If
                            oWB.Close(SaveChanges:=False)
                            InteropReleaseComObject(oWB)

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

                            oXL.ActiveWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF,
                                       ofn_pdf, Excel.XlFixedFormatQuality.xlQualityStandard,
                                       IncludeDocProperties:=True, IgnorePrintAreas:=False,
                                       OpenAfterPublish:=False)
                            oWB.Close(SaveChanges:=False)
                            InteropReleaseComObject(oWB)
                    End Select
                Catch ex As Exception
                    MessageBox.Show("Error:" & vbCrLf & inputfilename & vbCrLf & ex.Message)
                Finally
                    oXL.DisplayAlerts = True
                End Try

            End If

skip_this_file:
            progressbar_value += 1
            Baseform.BaseProgressBar.Value = progressbar_value
        Next

        oWord.Quit()
        InteropReleaseComObject(oWord)

        If results.Count > 0 Then
            Dim svp_info_path As String = System.IO.Path.GetTempPath & "SVP_Info.csv"
            System.IO.File.WriteAllText(svp_info_path, "SVP Name,Bottom Temperature,Bottom Depth,Remark,Probe Type,Easting,Northing" & vbCrLf)
            System.IO.File.AppendAllLines(svp_info_path, results.ToArray)
            'oXL = New Microsoft.Office.Interop.Excel.Application
            'oWB = New Microsoft.Office.Interop.Excel.Workbook
            oXL.Visible = True
            oXL.DisplayAlerts = True
            oWB = oXL.Workbooks.Open(svp_info_path, ReadOnly:=1)
        Else
            oXL.Quit()
            InteropReleaseComObject(oXL)
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
                                 Dim oXL As New Microsoft.Office.Interop.Excel.Application
                                 Dim oWB As Microsoft.Office.Interop.Excel.Workbook
                                 oXL.Visible = False
                                 oXL.DisplayAlerts = False
                                 oWB = oXL.Workbooks.Open(currentFile, ReadOnly:=1)
                                 'Try
                                 For Each oSheet As Microsoft.Office.Interop.Excel.Worksheet In oWB.Worksheets
                                     If oSheet.Name.ToUpper = "PROCESSED" Then
                                         Dim lastrow As Integer
                                         lastrow = oSheet.Range("A1").End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
                                         If Decimal.TryParse(CType(oSheet.Cells(lastrow, 1), Excel.Range).Value.ToString, lastdepth) And
                                         Decimal.TryParse(CType(oSheet.Cells(lastrow, 3), Excel.Range).Value.ToString, lasttemp) Then
                                             lastdepth = Decimal.Round(lastdepth, 0)
                                             lasttemp = Decimal.Round(lasttemp, 1)
                                             results.Add(System.IO.Path.GetFileNameWithoutExtension(currentFile) & "," & lasttemp & "," & lastdepth)
                                         End If
                                     End If
                                 Next
                                 oWB.Close(SaveChanges:=False)
                                 oXL.Quit()
                                 InteropReleaseComObject(oWB)
                                 InteropReleaseComObject(oXL)
                             End If
                             filecount += 1
                             Baseform.BaseProgressBar.Value = filecount
                         End Sub)
        GC.Collect() 'kill the last interop
        Baseform.BaseProgressBar.Value = filecount
        results.Sort(Function(xx, yy) xx.CompareTo(yy))

        If results.Count > 0 Then
            Dim svp_info_path As String = System.IO.Path.GetTempPath & "SVP_Info.csv"
            System.IO.File.WriteAllText(svp_info_path, "SVP Name,Bottom Temperature,Bottom Depth,Remark,Probe Type,Easting,Northing" & vbCrLf)
            System.IO.File.AppendAllLines(svp_info_path, results.ToArray)
            Dim oXL As New Microsoft.Office.Interop.Excel.Application
            Dim oWB As Microsoft.Office.Interop.Excel.Workbook
            oXL.Visible = True
            oXL.DisplayAlerts = True
            oWB = oXL.Workbooks.Open(svp_info_path, ReadOnly:=1)
        End If

    End Sub

    Sub Remove_selected(input_listbox As System.Windows.Forms.ListBox)
        Do While (input_listbox.SelectedItems.Count > 0)
            input_listbox.Items.Remove(input_listbox.SelectedItem)
        Loop
        If input_listbox.Items.Count = 0 Then Baseform.Init_listbox(input_listbox)
    End Sub

End Module