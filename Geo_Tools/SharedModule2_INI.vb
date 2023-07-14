Imports IniParser
Imports IniParser.Model

Module Module_INI

    Sub INI_Init(ByVal ini_filename As String, ByRef ini_data As IniData)
        'load ini, make one if not exist
        Dim ini_parser = New FileIniDataParser()
        Dim fs As IO.FileStream

        If System.IO.File.Exists(ini_filename) Then
            'if ini file does exist, read it
            Try
                ini_data = ini_parser.ReadFile(ini_filename)
            Catch ex As Exception
                'if something wrong, overwrite it with an empty ini
                fs = IO.File.Create(ini_filename)
                fs.Close()
                MsgBox(ex.InnerException.ToString)
            End Try
        Else
            'if ini file doesn't exist, make an empty one
            Try
                fs = IO.File.Create(ini_filename)
                fs.Close()
            Catch ex As Exception
                MsgBox(ex.InnerException.ToString)
            End Try
        End If
    End Sub

    'Sub INI_Constructor(ini_data As IniData)
    '    Dim LastPath As String = "C:\test\"
    '    ini_data.Global.AddKey("LastRun", Format(Now, "yyyy-MM-dd HH:mm"))
    '    ini_data.Global.AddKey("Version", "2018-12-03")
    '    ini_data.Sections.AddSection("Parameters")
    '    ini_data.Sections("Parameters").AddKey("SmoothWindowSize", 10)
    '    ini_data.Sections("Parameters").AddKey("LastFilePath", LastPath)
    'End Sub

    Sub INI_Save(ByVal ini_filename As String, ByRef ini_data As IniData)
        Dim ini_parser = New FileIniDataParser()
        Try
            ini_parser.WriteFile(ini_filename, ini_data)
        Catch ex As Exception
            MsgBox(ex.InnerException.ToString)
        End Try
    End Sub

    Function INI_TryGetKey(ByRef ini_data As IniData, ByVal section As String, ByVal keyname As String) As String
        INI_TryGetKey = ""
        If Not (ini_data.TryGetKey(section & ini_data.SectionKeySeparator & keyname, INI_TryGetKey)) Then
            'make a key if not found
            ini_data.Sections.AddSection(section)
            ini_data.Sections(section).AddKey(keyname, "")
        End If
    End Function

    Sub INI_TryUpdateKey(ByRef ini_data As IniData, ByVal section As String, ByVal keyname As String, ByVal keyvalue As String)
        Try
            ini_data.Sections(section).GetKeyData(keyname).Value = keyvalue
        Catch ex As Exception
            MsgBox(ex.InnerException.ToString)
        End Try
    End Sub

End Module