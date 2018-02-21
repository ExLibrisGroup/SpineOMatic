Imports System.IO

Module SOMFunctions
    Sub writeFile(ByVal fpath As String, ByVal fstring As String, ByVal append As Boolean)

        'writes the "labelout.txt" file to path: c:\a_spine, or any other file/path
        'passed to the subroutine. "append" is true to append, false to replace file data.
        Dim sw As StreamWriter
        Dim i As Integer = 0
        Dim listrec As String = ""
        Try
            If (Not File.Exists(fpath)) Then
                sw = File.CreateText(fpath)
                sw.Close()
            End If

            sw = New StreamWriter(fpath, append)
            sw.WriteLine(fstring)
            sw.Close()

        Catch ex As IOException
            MsgBox("error writing to file: " & fpath & vbCrLf &
            ex.ToString)
        End Try
    End Sub
End Module