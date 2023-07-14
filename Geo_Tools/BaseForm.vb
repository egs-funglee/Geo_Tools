Imports System.IO
Imports IniParser
Imports IniParser.Model
Public Class Baseform
    Public textbox_filedialog_title As String
    Public textbox_filedialog_filter As String
    Public textbox_file_filter1 As String
    Public textbox_file_filter2 As String
    Public textboxstring As String

    Public listbox_filedialog_title As String
    Public listbox_filedialog_filter As String
    Public listbox_file_filter1 As String
    Public listbox_file_filter2 As String
    Public listbox_file_exclude1 As String
    Public listbox_file_exclude2 As String
    Public listboxstring() As String

    Public trigger_textbox As Boolean
    Public trigger_listbox As Boolean

    Dim ini_data As New IniData
    Dim ini_filepath As String
    Dim lastpath As String = ""


    Private Sub Exit_Click(sender As Object, e As EventArgs) Handles _
        T1BX.Click, T2BX.Click, T3BX.Click, T4BX.Click, T5BX.Click, T6BX.Click, T7BX.Click
        Me.Close()
    End Sub

    Private Sub TabPage1_Enter(sender As Object, e As EventArgs) Handles TabPage1.Enter
        textbox_filedialog_title = "Open FILTERED.CNV File To Split"
        textbox_filedialog_filter = "All Files (*.*)| *.*|FILTERED.CNV Files (*.CNV*)|*.CNV*"
        textbox_file_filter1 = "FILTERED.CNV"
        textbox_file_filter2 = "FILTERED.CNV"
        listbox_filedialog_title = "Open RAW.CNV Files To Merge"
        listbox_filedialog_filter = "All Files (*.*)| *.*|RAW.CNV Files (*.CNV*)|*.CNV*"
        listbox_file_filter1 = ") RAW.CNV"
        listbox_file_filter2 = ") RAW.CNV"
        listbox_file_exclude1 = "FILTERED.CNV"
        listbox_file_exclude2 = "DUMMY"
        textboxstring = "Double Click To browse Or Drag Combined Smoothed CNV Here"
        ReDim listboxstring(4)
        listboxstring(0) = "Double Click To browse Or Drag Raw CNV files Here"
        listboxstring(1) = ""
        listboxstring(2) = "This Function version Is v20200731"
        listboxstring(3) = "Support time interpolated CNV And MagLog CNV"
        listboxstring(4) = "Re-calculate And smooth CMG from Filtered.CNV before splitting"
        Init_textbox(T1TextBox1)
        Init_listbox(T1ListBox1)
    End Sub
    Private Sub TabPage2_Enter(sender As Object, e As EventArgs) Handles TabPage2.Enter
        listbox_filedialog_title = "Open CNV Files"
        listbox_filedialog_filter = "All Files (*.*)| *.*|CNV Files (*.CNV)|*.CNV"
        listbox_file_filter1 = ".CNV"
        listbox_file_filter2 = ".CNV"
        listbox_file_exclude1 = "DUMMY"
        listbox_file_exclude2 = "DUMMY"
        ReDim listboxstring(2)
        listboxstring(0) = "Double Click To browse Or Drag C-View CNV files Here"
        listboxstring(1) = ""
        listboxstring(2) = "This convert files To target file type And save In the same folder"
        Init_listbox(T2ListBox1)
    End Sub
    Private Sub TabPage3_Enter(sender As Object, e As EventArgs) Handles TabPage3.Enter
        listbox_filedialog_title = "Open DTM / XYZ files"
        listbox_filedialog_filter = "All Files (*.*)| *.*|DTM/XYZ files (*.DTM;*.XYZ)|*.DTM;*.XYZ"
        listbox_file_filter1 = ".DTM"
        listbox_file_filter2 = ".XYZ"
        listbox_file_exclude1 = "DUMMY"
        listbox_file_exclude2 = "DUMMY"
        ReDim listboxstring(12)
        listboxstring(0) = "Double Click To browse Or Drag Raw CNV files Here"
        listboxstring(1) = ""
        listboxstring(2) = "This Function version Is v20190114"
        listboxstring(3) = ""
        listboxstring(4) = "CARIS EasyView support both Binary (FLT+HDR) And"
        listboxstring(5) = "ASCII (ASC) file As ESRI Grid input."
        listboxstring(6) = "HDR Is the header file Of Binary Grid FLT file."
        listboxstring(7) = "Binary files generally loads faster."
        listboxstring(8) = "File size Of compressed ASC would be smaller."
        listboxstring(9) = ""
        listboxstring(10) = "XYZ(XY-Depth) To GDAL ASCII Gridded XYZ (XY-Elevation, Y-Sorted)"
        listboxstring(11) = "EasyView 5 performance On big GDAL XYZ Is dissatisfying."
        listboxstring(12) = "ESRI formats are prefered On big files."
        Init_listbox(T3ListBox1)
    End Sub
    Private Sub TabPage4_Enter(sender As Object, e As EventArgs) Handles TabPage4.Enter
        listbox_filedialog_title = "Open GeoTiff/JPG/PNG files"
        listbox_filedialog_filter = "All Files (*.*)| *.*|GeoTiff/JPG/PNG Files (*.TIFF;*.TIF;*.JPG;*.PNG)|*.TIFF;*.TIF;*.JPG;*.PNG"
        listbox_file_filter1 = ".TIF"
        listbox_file_filter2 = "G"
        listbox_file_exclude1 = ".JGW"
        listbox_file_exclude2 = ".TFW"
        ReDim listboxstring(5)
        listboxstring(0) = "Double Click To browse Or Drag GeoTiff/JPG/PNG files Here"
        listboxstring(1) = ""
        listboxstring(2) = "This will create a boundary script (SCR) file For AutoCAD"
        listboxstring(3) = "And create TFW file from Geotiff tags, When it Is missing"
        listboxstring(4) = ""
        listboxstring(5) = "This Function version Is v20220119"
        Init_listbox(T4ListBox1)
    End Sub
    Private Sub TabPage5_Enter(sender As Object, e As EventArgs) Handles TabPage5.Enter
        listbox_filedialog_title = "Open PC Files"
        listbox_filedialog_filter = "All Files (*.*)| *.*|PC Files (*.PC*)|*.PC*"
        listbox_file_filter1 = ".PC"
        listbox_file_filter2 = ".PC"
        listbox_file_exclude1 = "DUMMY"
        listbox_file_exclude2 = "DUMMY"
        ReDim listboxstring(4)
        listboxstring(0) = "Double Click To browse Or Drag PC files Here"
        listboxstring(1) = ""
        listboxstring(2) = "This convert files To target file type And save In the same folder"
        listboxstring(3) = ""
        listboxstring(4) = ".PC0, .PC9 files will be skipped When target file type Is merged"
        Init_listbox(T5ListBox1)
    End Sub
    Private Sub TabPage6_Enter(sender As Object, e As EventArgs) Handles TabPage6.Enter
        textbox_filedialog_title = "Open DXF / e-SOFT ASCII Out (.ASC) File"
        textbox_filedialog_filter = "ASCII Files (*.ASC)| *.ASC|DXF Files (*.DXF)|*.DXF"
        textbox_file_filter1 = ".DXF"
        textbox_file_filter2 = ".ASC"
        listbox_filedialog_title = "Open FENWDKPRO / PC File"
        listbox_filedialog_filter = "All Files (*.*)| *.*|FENWDKPRO Files (*.FENWDKPRO;*.PC*)|*.FENWDKPRO;*.PC*"
        listbox_file_filter1 = ".FENWDKPRO"
        listbox_file_filter2 = ".FENWDKPRO"
        listbox_file_exclude1 = "DUMMY"
        listbox_file_exclude2 = "DUMMY"
        textboxstring = "Double Click To browse Or Drag DXF" & vbCrLf &
                        "Or e-SOFT ASCII Out (.ASC) File Here" & vbCrLf & vbCrLf &
                        "File Data Format: Profile X, Profile Y, Fix Number, Rotation Angle"
        ReDim listboxstring(10)
        listboxstring(0) = "Double Click to browse or Drag FENWDKPRO / PC files Here"
        listboxstring(1) = ""
        listboxstring(2) = "The DXF file for Fix Positions can be prepared by plotting"
        listboxstring(3) = "track in Profile view and save in DXF format (R12-2010)"
        listboxstring(4) = ""
        listboxstring(5) = "The ASC file can be prepared by e-SOFT ASCII Out Function."
        listboxstring(6) = "Exporting track in Profile view by 'Texts (IP & AP)'"
        listboxstring(7) = ""
        listboxstring(8) = "For some cases like coast line or without RPL"
        listboxstring(9) = "You may also Create FENWDKPRO file"
        listboxstring(10) = "by along track distance from PC file"
        Init_textbox(T6TextBox1)
        Init_listbox(T6ListBox1)
    End Sub
    Private Sub TabPage7_Enter(sender As Object, e As EventArgs) Handles TabPage7.Enter
        listbox_filedialog_title = "Open documents"
        listbox_filedialog_filter = "All Files (*.*)| *.*|Documents files (*.DOC*;*.XLS*)|*.DOC*;*.XLS*|Word files (*.DOC*)|*.DOC*|Excel files (*.XLS*)|*.XLS*"
        listbox_file_filter1 = ".DOC"
        If Not listbox_file_filter2 = ".XLS" Then listbox_file_filter2 = ".DOC"
        listbox_file_exclude1 = "DUMMY"
        listbox_file_exclude2 = "DUMMY"
        ReDim listboxstring(5)
        listboxstring(0) = "Double Click to browse or Drag Documents / Folders Here"
        listboxstring(1) = ""
        listboxstring(2) = "For SVP Excel Files:"
        listboxstring(3) = "File name must contains 'SV'"
        listboxstring(4) = "Please make sure All active printers (including XPS Writer)"
        listboxstring(5) = "and All Charts in SVP Excel files paper size are A4"
        Init_listbox(T7ListBox1)
    End Sub

    Private Sub Init_textbox(input_textbox As System.Windows.Forms.TextBox)
        If Len(input_textbox.Text) = 0 Then
            input_textbox.Text = textboxstring
            trigger_textbox = False
        ElseIf Not input_textbox.Text = textboxstring Then
            trigger_textbox = True
        End If
    End Sub

    Sub Init_listbox(input_listbox As System.Windows.Forms.ListBox)
        If input_listbox.Items.Count = 0 Then
            input_listbox.Sorted = False
            input_listbox.SelectionMode = SelectionMode.None
            For i = 0 To listboxstring.Count - 1 Step 1
                input_listbox.Items.Add(listboxstring(i))
            Next
            trigger_listbox = False
        ElseIf Not input_listbox.Items(0).ToString = listboxstring(0) Then
            trigger_listbox = True
        End If
        BaseProgressBar.Maximum = 1
        BaseProgressBar.Value = 0
        GC.Collect()
    End Sub

    Private Sub Input_box_dragEnter(sender As Object, e As DragEventArgs) Handles T1TextBox1.DragEnter, T6TextBox1.DragEnter,
        T1ListBox1.DragEnter, T2ListBox1.DragEnter, T3ListBox1.DragEnter, T4ListBox1.DragEnter, T5ListBox1.DragEnter, T6ListBox1.DragEnter, T7ListBox1.DragEnter
        If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
            e.Effect = DragDropEffects.All
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub Input_textbox_dragdrop(sender As Object, e As DragEventArgs) Handles T1TextBox1.DragDrop, T6TextBox1.DragDrop
        Dim s() As String = DirectCast(e.Data.GetData("FileDrop", False), String())
        Dim sender_object As System.Windows.Forms.TextBox
        sender_object = CType(sender, System.Windows.Forms.TextBox)
        sender_object.Clear()
        trigger_textbox = False
        For i = 0 To s.Length - 1
            If s(i).ToUpper.Contains(textbox_file_filter1) Or s(i).ToUpper.Contains(textbox_file_filter2) Then
                sender_object.Text = s(i)
                trigger_textbox = True
            End If
        Next i
        If trigger_textbox = False Then Init_textbox(sender_object)
        'update lastpath
        lastpath = Path.GetDirectoryName(s(0))
    End Sub

    Private Sub Input_listbox_dragdrop(sender As Object, e As DragEventArgs) Handles _
        T1ListBox1.DragDrop, T2ListBox1.DragDrop, T3ListBox1.DragDrop, T4ListBox1.DragDrop, T5ListBox1.DragDrop, T6ListBox1.DragDrop, T7ListBox1.DragDrop
        Dim s() As String = DirectCast(e.Data.GetData("FileDrop", False), String())
        Dim sender_object As System.Windows.Forms.ListBox
        sender_object = CType(sender, System.Windows.Forms.ListBox)
        sender_object.Items.Clear()
        sender_object.SelectionMode = SelectionMode.MultiExtended
        trigger_listbox = False
        sender_object.Sorted = True
        For i = 0 To s.Length - 1
            'folder
            If File.GetAttributes(s(i)) = FileAttributes.Directory Then
                Dim stack As New Stack(Of String)
                Dim fs() As String
                Dim no_file As Integer
                stack.Push(s(i))
                Do While (stack.Count > 0)
                    Dim dir As String = stack.Pop
                    Try
                        no_file = Directory.GetFiles(dir, "*.*").Count - 1
                        ReDim fs(no_file)
                        Directory.GetFiles(dir, "*.*").CopyTo(fs, 0)
                        For j = 0 To no_file Step 1
                            If File.Exists(fs(j)) And Not fs(j).Contains("\~$") Then
                                If (fs(j).ToUpper.Contains(listbox_file_filter1) Or fs(j).ToUpper.Contains(listbox_file_filter2)) And
                                    Not (fs(j).ToUpper.Contains(listbox_file_exclude1) Or fs(j).ToUpper.Contains(listbox_file_exclude2)) Then
                                    sender_object.Items.Add(fs(j))
                                    trigger_listbox = True
                                End If
                            End If
                        Next
                        Dim directoryName As String
                        For Each directoryName In Directory.GetDirectories(dir)
                            stack.Push(directoryName)
                        Next
                    Catch ex As Exception
                    End Try
                Loop
            End If
            'file
            If File.Exists(s(i)) And Not s(i).Contains("\~$") Then
                If (s(i).ToUpper.Contains(listbox_file_filter1) Or s(i).ToUpper.Contains(listbox_file_filter2)) And
                                    Not (s(i).ToUpper.Contains(listbox_file_exclude1) Or s(i).ToUpper.Contains(listbox_file_exclude2)) Then
                    sender_object.Items.Add(s(i))
                    trigger_listbox = True
                End If
            End If
        Next i
        If trigger_listbox = False Then Init_listbox(sender_object)
        'update lastpath
        lastpath = Path.GetDirectoryName(s(0))
    End Sub

    Private Sub Input_textbox_doubleclick(sender As Object, e As EventArgs) Handles T1TextBox1.DoubleClick, T6TextBox1.DoubleClick
        Dim myStream As Stream = Nothing
        Dim openFileDialog As New OpenFileDialog()
        Dim datfile As System.Collections.IEnumerable
        Dim sender_object As System.Windows.Forms.TextBox
        sender_object = CType(sender, System.Windows.Forms.TextBox)
        sender_object.Clear()
        openFileDialog.Title = textbox_filedialog_title
        openFileDialog.Filter = textbox_filedialog_filter
        openFileDialog.Multiselect = False
        openFileDialog.FilterIndex = 2
        openFileDialog.InitialDirectory = lastpath 'init lastpath
        trigger_textbox = False
        If openFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                myStream = openFileDialog.OpenFile()
                datfile = openFileDialog.FileNames
                If (myStream IsNot Nothing) Then
                    For Each filename As String In datfile
                        If filename.ToUpper.Contains(textbox_file_filter1) Or filename.ToUpper.Contains(textbox_file_filter2) Then
                            sender_object.Text = filename
                            lastpath = Path.GetDirectoryName(openFileDialog.FileName) 'update lastpath
                            trigger_textbox = True
                        End If
                    Next
                End If
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
                If myStream IsNot Nothing Then
                    myStream.Close()
                End If
            End Try
        End If
        If trigger_textbox = False Then Init_textbox(sender_object)
    End Sub

    Private Sub Input_listbox_doubleclick(sender As Object, e As EventArgs) Handles _
        T1ListBox1.DoubleClick, T2ListBox1.DoubleClick, T3ListBox1.DoubleClick, T4ListBox1.DoubleClick, T5ListBox1.DoubleClick, T6ListBox1.DoubleClick, T7ListBox1.DoubleClick
        Dim myStream As Stream = Nothing
        Dim openFileDialog As New OpenFileDialog()
        Dim datfile As System.Collections.IEnumerable
        Dim sender_object As System.Windows.Forms.ListBox
        sender_object = CType(sender, System.Windows.Forms.ListBox)
        sender_object.Items.Clear()
        sender_object.SelectionMode = SelectionMode.MultiExtended
        openFileDialog.Title = listbox_filedialog_title
        openFileDialog.Filter = listbox_filedialog_filter
        openFileDialog.Multiselect = True
        openFileDialog.FilterIndex = 2
        openFileDialog.InitialDirectory = lastpath 'init lastpath
        trigger_listbox = False
        sender_object.Sorted = True
        If openFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                myStream = openFileDialog.OpenFile()
                datfile = openFileDialog.FileNames
                If (myStream IsNot Nothing) Then
                    For Each filename As String In datfile
                        If (filename.ToUpper.Contains(listbox_file_filter1) Or filename.ToUpper.Contains(listbox_file_filter2)) And
                            Not (filename.ToUpper.Contains(listbox_file_exclude1) Or filename.ToUpper.Contains(listbox_file_exclude2)) Then
                            sender_object.Items.Add(filename)
                            lastpath = Path.GetDirectoryName(openFileDialog.FileName) 'update lastpath
                            trigger_listbox = True
                        End If
                    Next
                End If
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
                If myStream IsNot Nothing Then
                    myStream.Close()
                End If
            End Try
        End If
        If trigger_listbox = False Then Init_listbox(sender_object)
    End Sub


    '============== Drag over ==============


    Public Function GetTabPageIndex(ByVal pt As System.Drawing.Point, ByVal tc As TabControl) As Integer
        For index As Integer = 0 To tc.TabCount - 1
            If tc.GetTabRect(index).Contains(pt.X, pt.Y) Then Return index
        Next
        Return 0
    End Function

    Private Sub TcFolders_DragOver(ByVal sender As Object, ByVal e As DragEventArgs) Handles TabControl1.DragOver
        If e.Data.GetDataPresent(DataFormats.FileDrop) And (TabControl1.TabCount > 0) Then
            e.Effect = If(e.KeyState = 9, DragDropEffects.Copy, DragDropEffects.Move)
            Me.Activate()
        Else
            e.Effect = DragDropEffects.None
        End If

        Dim tc As TabControl = CType(sender, System.Windows.Forms.TabControl)
        Dim pt As System.Drawing.Point = tc.PointToClient(New System.Drawing.Point(e.X, e.Y))
        Dim IndexOld As Integer = tc.SelectedIndex
        Dim IndexNew As Integer = tc.SelectedIndex

        For index As Integer = 0 To TabControl1.TabCount - 1
            If TabControl1.GetTabRect(index).Contains(pt) Then
                IndexNew = index
                Exit For
            End If
        Next

        If IndexNew <> IndexOld Then
            tc.SelectedIndex = IndexNew
            tc.Focus()
            'Dim tr As Rectangle = tc.GetTabRect(IndexNew)
            'Windows.Forms.Cursor.Position = tc.PointToScreen(New Point(tr.X + (tr.Width \ 2), tr.Y + (tr.Height \ 2)))
        End If
    End Sub

    Private Sub Whatever_DragOver(sender As Object, e As DragEventArgs) Handles MyBase.DragOver, T1TextBox1.DragOver, T6TextBox1.DragOver,
        T1ListBox1.DragOver, T2ListBox1.DragOver, T3ListBox1.DragOver, T4ListBox1.DragOver, T5ListBox1.DragOver, T6ListBox1.DragOver, T7ListBox1.DragOver
        Me.Activate()
    End Sub




    '============== TAB 1 ==============


    Private Sub T1B1_Click(sender As Object, e As EventArgs) Handles T1B1.Click
        If trigger_listbox Then
            BaseProgressBar.Value = 0
            Merge_cnv(T1ListBox1)
            BaseProgressBar.Value = 1
        End If
    End Sub

    Private Sub T1B2_Click(sender As Object, e As EventArgs) Handles T1B2.Click
        If trigger_textbox And trigger_listbox Then
            BaseProgressBar.Value = 0
            Split_cnv(T1TextBox1.Text, T1ListBox1)
            BaseProgressBar.Value = 1
        End If
    End Sub

    Private Sub T1G1RadioButtons_Click(sender As Object, e As EventArgs) Handles T1G1RadioButton1.Click, T1G1RadioButton2.Click
        If T1G1RadioButton1.Checked Then
            listbox_file_filter2 = ") RAW.CNV"
            listbox_file_exclude2 = "DUMMY"
            Remove_listbox_items_if_not_contain(T1ListBox1, "RAW.CNV")
        Else
            listbox_file_filter2 = ".CNV"
            listbox_file_exclude2 = "RAW.CNV"
            Remove_listbox_items_if_ending_match(T1ListBox1, "RAW.CNV")
        End If
    End Sub

    Private Sub T1G2TextBox1_Leave(sender As Object, e As EventArgs) Handles T1G2TextBox1.Leave, T2TextBox1.Leave
        Dim swindowsize As Short 'update both numbers on tab1 and tab2
        Try
            swindowsize = CShort(CType(sender, System.Windows.Forms.TextBox).Text) 'CShort(T1G2TextBox1.Text)
            If swindowsize < 7 Then
                T1G2TextBox1.Text = "7"
                T2TextBox1.Text = "7"
                Exit Sub
            End If
            T1G2TextBox1.Text = ((swindowsize \ 2) * 2 + 1).ToString
            T2TextBox1.Text = T1G2TextBox1.Text
        Catch ex As Exception
            MsgBox("Please input a valid number")
            T1G2TextBox1.Text = "81"
            T2TextBox1.Text = "81"
        End Try
    End Sub

    Private Sub T1G2CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles T1G2CheckBox1.CheckedChanged
        If T1G2CheckBox1.Checked = False Then
            T1G2TextBox1.Enabled = False
        Else
            T1G2TextBox1.Enabled = True
        End If
    End Sub



    '============== TAB 2 ==============



    Private Sub T2B1_Click(sender As Object, e As EventArgs) Handles T2B1.Click
        If trigger_listbox Then
            BaseProgressBar.Value = 0
            Conv_cnv2pc(T2ListBox1)
            BaseProgressBar.Value = 1
        End If
    End Sub
    Private Sub T2B2_Click(sender As Object, e As EventArgs) Handles T2B2.Click
        If trigger_listbox Then
            BaseProgressBar.Value = 0
            Conv_cnv2tpc(T2ListBox1)
        End If
    End Sub
    Private Sub T2B3_Click(sender As Object, e As EventArgs) Handles T2B3.Click
        If trigger_listbox Then
            BaseProgressBar.Value = 0
            Conv_cnv2magxtfcnv(T2ListBox1)
            BaseProgressBar.Value = 1
        End If
    End Sub
    Private Sub T2B4_Click(sender As Object, e As EventArgs) Handles T2B4.Click
        If trigger_listbox Then
            BaseProgressBar.Value = 0
            Conv_cnv2scr(T2ListBox1)
            BaseProgressBar.Value = 1
        End If
    End Sub
    Private Sub T2B5_Click(sender As Object, e As EventArgs) Handles T2B5.Click
        If trigger_listbox Then
            BaseProgressBar.Value = 0
            Conv_cnv2tp(T2ListBox1)
            BaseProgressBar.Value = 1
        End If
    End Sub

    Private Sub T2B6_Click(sender As Object, e As EventArgs) Handles T2B6.Click
        If trigger_listbox Then
            BaseProgressBar.Value = 0
            Recalc_smooth_cnv_cmg(T2ListBox1)
            BaseProgressBar.Value = 1
        End If
    End Sub



    '============== TAB 3 ==============


    Private Sub T3B1_Click(sender As Object, e As EventArgs) Handles T3B1.Click
        If trigger_listbox Then Conv_xyz_to_bin(T3ListBox1)
    End Sub

    Private Sub T3B2_Click(sender As Object, e As EventArgs) Handles T3B2.Click
        If trigger_listbox Then Conv_xyz_to_asc(T3ListBox1)
    End Sub

    Private Sub T3B3_Click(sender As Object, e As EventArgs) Handles T3B3.Click
        If trigger_listbox Then Invert_XYZ(T3ListBox1)
    End Sub

    Private Sub T3B4_Click(sender As Object, e As EventArgs) Handles T3B4.Click
        If trigger_listbox Then Convert_to_GDAL_XYZ(T3ListBox1)
    End Sub



    '============== TAB 4 ==============


    Private Sub T4B1_Click(sender As Object, e As EventArgs) Handles T4B1.Click
        If trigger_listbox Then Geotiff_boundary_maker(T4ListBox1)
    End Sub
    Private Sub T4B2_Click(sender As Object, e As EventArgs) Handles T4B2.Click
        If trigger_listbox Then Geotiff_tocad(T4ListBox1)
    End Sub



    '============== TAB 5 ==============


    Private Sub T5B1_Click(sender As Object, e As EventArgs) Handles T5B1.Click
        If trigger_listbox Then
            BaseProgressBar.Value = 0
            Pc_to_tp(T5ListBox1)
            BaseProgressBar.Value = 1
        End If
    End Sub

    Private Sub T5B2_Click(sender As Object, e As EventArgs) Handles T5B2.Click
        If trigger_listbox Then
            BaseProgressBar.Value = 0
            Pc_to_ctp(T5ListBox1)
            BaseProgressBar.Value = 1
        End If
    End Sub

    Private Sub T5B3_Click(sender As Object, e As EventArgs) Handles T5B3.Click
        If trigger_listbox Then
            BaseProgressBar.Value = 0
            Pc_to_scr(T5ListBox1)
            BaseProgressBar.Value = 1
        End If
    End Sub

    Private Sub T5B4_Click(sender As Object, e As EventArgs) Handles T5B4.Click
        If trigger_listbox Then
            BaseProgressBar.Value = 0
            Pc_to_tcpc(T5ListBox1)
            BaseProgressBar.Value = 1
        End If
    End Sub

    Private Sub T5B5_Click(sender As Object, e As EventArgs) Handles T5B5.Click
        If trigger_listbox Then
            BaseProgressBar.Value = 0
            Pc_to_cpc(T5ListBox1)
            BaseProgressBar.Value = 1
        End If
    End Sub

    Private Sub T5B6_Click(sender As Object, e As EventArgs) Handles T5B6.Click
        If trigger_listbox Then
            BaseProgressBar.Value = 0
            Pc_to_cnv(T5ListBox1)
            BaseProgressBar.Value = 1
        End If
    End Sub



    '============== TAB 6 ==============


    Private Sub T6B1_Click(sender As Object, e As EventArgs) Handles T6B1.Click
        Dim ve As Integer
        If Not Integer.TryParse(T6TextBox2.Text, ve) Then
            MsgBox("Invalid vertical exaggreation")
            Exit Sub
        End If
        If trigger_textbox And trigger_listbox Then
            BaseProgressBar.Value = 0
            Update_fenwdkpro(T6TextBox1.Text, T6ListBox1, T6CheckBox2.Checked, ve)
            BaseProgressBar.Value = 1
        End If
    End Sub

    Private Sub T6B2_Click(sender As Object, e As EventArgs) Handles T6B2.Click
        If trigger_listbox Then
            BaseProgressBar.Value = 0
            Pc_to_fenwdkpro(T6ListBox1, T6G1RadioButton1.Checked, T6CheckBox2.Checked)
            BaseProgressBar.Value = 1
        End If
    End Sub
    Private Sub T6CheckBox1_Click(sender As Object, e As EventArgs) Handles T6CheckBox1.Click
        If T6CheckBox1.Checked Then
            T6GroupBox1.Enabled = True
            Remove_listbox_items_if_not_contain(T6ListBox1, ".PC")
            listbox_file_filter1 = ".PC"
            listbox_file_filter2 = ".PC"
            T6B1.Enabled = False
        Else
            T6GroupBox1.Enabled = False
            Remove_listbox_items_if_not_contain(T6ListBox1, ".FENWDKPRO")
            listbox_file_filter1 = ".FENWDKPRO"
            listbox_file_filter2 = ".FENWDKPRO"
            T6B1.Enabled = True
        End If
    End Sub
    Private Sub T6B3_Click(sender As Object, e As EventArgs) Handles T6B3.Click
        BaseProgressBar.Value = 0
        GetFixPos_fom_ActiveLayer(T6TextBox1)
        BaseProgressBar.Value = 1
    End Sub
    Private Sub T6B4_Click(sender As Object, e As EventArgs) Handles T6B4.Click
        BaseProgressBar.Value = 0
        SBP_Image_tocad(T6ListBox1)
        BaseProgressBar.Value = 1
    End Sub

    '============== TAB 7 ==============


    Private Sub T7B1_Click(sender As Object, e As EventArgs) Handles T7B1.Click
        Dim xls_output_type As Byte
        Select Case True
            Case T7CheckBox3.Checked And T7G2RadioButton1.Checked 'ASN/NEC
                xls_output_type = 1
            Case T7CheckBox3.Checked And T7G2RadioButton2.Checked 'TE Subcom
                xls_output_type = 2
            Case T7G1RadioButton1.Checked 'Last Active Sheets
                xls_output_type = 3
            Case T7G1RadioButton2.Checked 'All Sheets
                xls_output_type = 4
        End Select
        If trigger_listbox Then
            Convert_to_pdf(T7ListBox1, T7CheckBox1.Checked, T7CheckBox2.Checked, xls_output_type, T7G2CheckBox1.Checked)
            GC.Collect()
        End If
    End Sub

    Private Sub T7B2_Click(sender As Object, e As EventArgs) Handles T7B2.Click
        If trigger_listbox And T7CheckBox2.Checked Then
            Get_svp_info(T7ListBox1)
            GC.Collect()
        End If
    End Sub

    Private Sub T7B3_Click(sender As Object, e As EventArgs) Handles T7B3.Click
        Remove_selected(T7ListBox1)
    End Sub

    Private Sub T7CheckBox2_Click(sender As Object, e As EventArgs) Handles T7CheckBox2.Click
        If T7CheckBox2.Checked Then
            T7CheckBox3.Enabled = True
            listbox_file_filter2 = ".XLS"
            T7CheckBox3_Click()
        Else
            If trigger_listbox Then
                MsgBox("Excel files will be removed from the list")
                For i = T7ListBox1.Items.Count - 1 To 0 Step -1
                    If Strings.Right((T7ListBox1.Items(i).ToString), 5).ToUpper.Contains(".XLS") Then T7ListBox1.Items.RemoveAt(i)
                Next
            End If
            listbox_file_filter2 = ".DOC"
            T7CheckBox3.Enabled = False
            T7GroupBox1.Enabled = False
            T7GroupBox2.Enabled = False
        End If
    End Sub
    Private Sub T7CheckBox3_Click() Handles T7CheckBox3.Click
        If T7CheckBox3.Checked Then
            T7GroupBox1.Enabled = False
            T7GroupBox2.Enabled = True
        Else
            T7GroupBox1.Enabled = True
            T7GroupBox2.Enabled = False
        End If
    End Sub



    Private Sub Baseform_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'MsgBox("This tool is under test, please inform author when there is any bug", , "Under test")
        ToolTip1.SetToolTip(T1G2CheckBox1, "Enable this will recalculate the CMG when spliting the Combined Filtered.CNV. Replacing Vessel/Fish gyro")
        ToolTip1.SetToolTip(T1G2TextBox1, "This is related with the ping rate and speed. 71-101 is suggested for 200m SSS range")

        ToolTip1.SetToolTip(T2B1, "Export CNV to PC File (Full data density and only PUFI, heading as gyro)")
        ToolTip1.SetToolTip(T2B2, "Export CNV to PC File (Thin data to each Fix and only PUFI, heading as gyro)")
        ToolTip1.SetToolTip(T2B3, "Only Filtered.CNV files will be processed to Maggy Filtered CNV for Exchange")
        ToolTip1.SetToolTip(T2B4, "Export CNV to AutoCAD SCR Script file. Duplicate coordinates will be removed")
        ToolTip1.SetToolTip(T2B5, "Export CNV to Trackplot file. Duplicate coordinates will be removed")
        ToolTip1.SetToolTip(T2B6, "Recalculate and Smooth Heading to CMG in CNV files. Files will be treated as individual.")

        ToolTip1.SetToolTip(T5B1, "Export PC File PUFI to Individual Trackplot file (1 PC File, 1 TP File)")
        ToolTip1.SetToolTip(T5B2, "Export Multiple PC Files' PUFI to Trackplot file (by each line)")
        ToolTip1.SetToolTip(T5B3, "Export Multiple PC Files' PUFI to AutoCAD SCR Script file (by each line)")
        ToolTip1.SetToolTip(T5B4, "Merge Multiple PC Files to PC0 Files by each fix and line")
        ToolTip1.SetToolTip(T5B5, "Merge Multiple PC Files to PC9 Files by each line")

        Baseform_LoadINI()
    End Sub

    Private Sub Baseform_LoadINI()
        'init some parameters
        ini_filepath = Path.GetTempPath
        If Len(ini_filepath) > 2 Then
            If Strings.Right(ini_filepath, 1) = "\" Then
                ini_filepath &= "Geo_Tools.ini"
            Else
                ini_filepath &= "\Geo_Tools.ini"
            End If
        Else
            ini_filepath = "Geo_Tools.ini"
        End If
        INI_Init(ini_filepath, ini_data)

        'use lastpath as temp string

        lastpath = INI_TryGetKey(ini_data, "Parameters", "ReCalc_CMG_When_Split")
        If Len(lastpath) > 0 Then
            T1G2CheckBox1.Checked = CBool(lastpath)
        Else
            T1G2CheckBox1.Checked = True
        End If


        lastpath = INI_TryGetKey(ini_data, "Parameters", "SmoothWindowSize")
        If Len(lastpath) > 0 Then
            T1G2TextBox1.Text = lastpath
            T2TextBox1.Text = lastpath
        End If

        lastpath = INI_TryGetKey(ini_data, "Parameters", "LastPath")
        If Len(lastpath) = 0 Then
            lastpath = "C:\"
        End If
    End Sub

    Private Sub Baseform_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'MsgBox("closing")
        INI_TryUpdateKey(ini_data, "Parameters", "LastPath", lastpath)
        INI_TryUpdateKey(ini_data, "Parameters", "SmoothWindowSize", T1G2TextBox1.Text)
        INI_TryUpdateKey(ini_data, "Parameters", "ReCalc_CMG_When_Split", T1G2CheckBox1.Checked.ToString)
        INI_Save(ini_filepath, ini_data)
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            System.Diagnostics.Process.Start("https://drive.google.com/open?id=1SjWL6970mbvCc-aSqMoVDI8DZkrbKub4")
        Catch ex As Exception
            MsgBox("Something wrong while opening the URL. Please manually type the address to your browser.")
        End Try
    End Sub
End Class
