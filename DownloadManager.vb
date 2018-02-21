Imports System.IO
''' <summary>
''' Used to check for and download updated versions of SpineOMatic
''' </summary>
Public Class DownloadManager
    Private strUpdateBaseURL As String
    Private strCurrentVersion As String
    Private strMyPath As String
    Private strPreferredVersion As String = ""
    Private objWebClient As New Net.WebClient()
    Private strExeName As String

    ReadOnly Property PreferredVersion As String
        Get
            Return strPreferredVersion
        End Get
    End Property


    ''' <summary>
    ''' Create a new update checker
    ''' </summary>
    ''' <param name="UpdateBaseURL">Base URL to download updates from</param>
    ''' <param name="CurrentVersion">Version of the running SpineOMatic instance</param>
    ''' <param name="Path">Path to the directory containg the SpineOMatic executable</param>
    Public Sub New(ByVal UpdateBaseURL As String, ByVal CurrentVersion As String, ByVal Path As String)
        strUpdateBaseURL = UpdateBaseURL
        strCurrentVersion = CurrentVersion
        strMyPath = Path
    End Sub

    ''' <summary>
    ''' Checks for a new version of SpineOMatic
    ''' </summary>
    ''' <returns><c>True</c> if a new version exists,
    ''' <c>False</c> otherwise</returns>
    Public Function CheckForNewVersion() As Boolean
        Dim strVersionListFile As String = ""
        Dim arrVersionList As Array
        Dim i = 0
        Dim intVersionStart As Integer = 0, intVersionLength As Integer = 0
        Dim strVersionListURL As String = strUpdateBaseURL & "som_list.txt"

        strVersionListFile = objWebClient.DownloadString(strVersionListURL)

        'version list format:
        '*SpineLabeler-1_6;20120801 (* means this is the preferred version)
        'SpineLabeler-1_5;20120731  (program name-version_subversion ; date of release)
        '...etc
        arrVersionList = Split(strVersionListFile, vbCrLf)
        For i = 0 To arrVersionList.Length - 1
            If InStr(arrVersionList(i).ToString, "*") Then 'if preferred version is found
                strPreferredVersion = arrVersionList(i).ToString            'parse name and version number
                strExeName = strPreferredVersion.Substring(1, strPreferredVersion.IndexOf(";") - 1)
                intVersionStart = strPreferredVersion.IndexOf("-") + 1
                intVersionLength = strPreferredVersion.IndexOf(";") - intVersionStart
                strPreferredVersion = strPreferredVersion.Substring(intVersionStart, intVersionLength)
                strPreferredVersion = strPreferredVersion.Replace("_", ".")
                Exit For
            End If
        Next

        Return strPreferredVersion <> "" And strCurrentVersion <> strPreferredVersion
    End Function

    ''' <summary>
    ''' Downloads the newest version
    ''' </summary>
    Public Sub DownloadNewVersion()
        Dim strExeDownloadURL = strUpdateBaseURL & strExeName & ".exe"
        objWebClient.DownloadFile(strExeDownloadURL, strMyPath & strExeName & ".exe")
        RenameVersions()
    End Sub

    Public Sub DownloadJavaComponents(ByVal strJavaClass As String, ByVal strJAvaSDK As String, ByVal strJavaTest As String)
        If Not File.Exists(strMyPath & strJavaClass & ".class") Then
            DownloadFile(strJavaClass & ".class")
        End If
        If Not File.Exists(strMyPath & strJAvaSDK) Then
            DownloadFile(strJAvaSDK)
        End If
        If Not File.Exists(strMyPath & strJavaTest & ".class") Then
            DownloadFile(strJavaTest & ".class")
        End If
    End Sub

    Private Sub DownloadFile(ByVal fileName As String)
        Dim webrequest As String = strUpdateBaseURL & fileName
        Dim strErrorMessage As String = ""

        Try
            objWebClient.DownloadFile(webrequest, strMyPath & fileName)
        Catch ex As Exception
            If ex.Message.Contains("407") Then
                strErrorMessage = $"Your proxy server is not allowing you to connect to the BC server: {vbCrLf} {vbCrLf}" &
                $"{strUpdateBaseURL} {vbCrLf} {vbCrLf}" &
                "Ask your IT Networking office to allow access to ('whitelist') this server."
            Else
                strErrorMessage = $"{fileName} - Download error: {ex.Message}"
            End If
            MsgBox(strErrorMessage, MsgBoxStyle.Exclamation, "Download Error")
            Exit Sub
        End Try
    End Sub


    ''' <summary>
    ''' Renames old executable to backup name and new executable to SpineLabeler.exe
    ''' </summary>
    Private Sub RenameVersions()
        Dim strDownloadedExe = strMyPath & strExeName & ".exe"
        Dim strBackupVersionExe = "SpineLabeler-" & strCurrentVersion.Replace(".", "_") & ".exe"

        My.Computer.FileSystem.RenameFile(strMyPath & "SpineLabeler.exe", strBackupVersionExe)
        My.Computer.FileSystem.RenameFile(strDownloadedExe, "SpineLabeler.exe")
    End Sub
End Class