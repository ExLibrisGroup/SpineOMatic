Imports System.IO
Imports System.Management
Imports System.Drawing.Drawing2D
Imports System.Drawing.imaging
Imports System.Drawing.printing
Imports System.Threading
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Net
Imports System.Xml
Imports System.Web.Script.Serialization


Public Class Form1
    'v. 8.0: Major release
    '        Includes checking GitHub for new version and linking to GitHub wiki for documentation. Removed references to BC license and server.
    '
    'v. 7.0: Major release
    '       Includes option to use Alma's RESTful API URL to retrieve an item's XML file, as
    '       well as the deprecated Java SOAP APIs.  The RESTful XML file is converted into
    '       the SOAP XML format so that SpineOMatic can process the file without major 
    '       modification to the software.
    '
    'v. 6.12: Minor trial release
    '       Allows the ! (render as bar code) formatting code to be used in custom Pocket Labels.
    '       Will be distributed to one user who wanted the feature, and included in subsequent
    '       releases if it is acceptable.
    '       
    'v. 6.11: Minor release
    '       Clarified and changed LC/LC child lit./NLM and Dewey parsing behavior:
    '       Tweak and Test panel for the LC... parser: "Decimal" break description was changed to
    '       "Class Decimal".  Fixed a bug in "Break before decimal" that was causing a break
    '       after the decimal rather than before. An option was added to allow break both before
    '       or after the decimal.
    '       An option was added to the Dewey parser to allow breaking on the cutter decimal
    '       string after a specific number of characters following the decimal.
    '
    'v. 6.1: Minor release
    '       Fixed bug in "*" format code ("Suppress field display");
    '       Added check for unbalanced quotation marks in "Quoted text" format code.
    '
    'v. 6.0: Major release features:
    '       Holdings processor now lets user select no holdings, textual holdings parsed by SpineOMatic,
    '       or Ex Libris' Parsed Holdings fields.  If Parsed Holdings are specified but do not exist,
    '       SpineOMatic will parse the textual holdings.  The display will indicated which call number
    '       parser was used, and also which holdings parser was used.
    '       Asterisk (*) formatting code ("suppress" field display) suppresses display if the XML field
    '       is blank, or if the field is equal to any of three user-defined values.
    '       Tweak and Test SuDoc parser allows breaking on "Other" characters, with option
    '       to remove characters. (Behavior is consistent with LC, Dewey and Other parsers.)
    '       Dewey parser provides an option to print long numeric class numbers in groups of characters.
    '       New print option to send label text to a custom DOS batch file ("viados.bat"), which can
    '       print to legacy printers attached via LPT or COM ports, etc.
    '       Margins and line spacing can be entered in inches or centimeters.
    '       A decimal point or a comma can represent decimal fractions.
    '       Allows negative top margin and left margin settings.
    '       To insert XML field names into text boxes that allow it, items can be selected from
    '       a list of all XML fields (rather than typing the names).
    '
    'v. 5.21: Minor release to handle unencoded ampersands in the item's XML record.
    '       Added the tilde (~) character to stand for space characters in the Tweak and Test 
    '       Other Break text strings.
    '       Changed the "Hide cutter decimal" routine from removing all decimals to removing only
    '       the first character, if it is a decimal.
    '
    'v. 5.2: Minor release to fix a bug in the multi-label print to Desktop routine, which failed to
    '       change the print button from "Stop" to "Send to Desktop Printer".
    '
    'v. 5.1: Minor release to fix the Holdings parser, which was causing spaces between elements
    '       to be removed.
    '       A "Break on spaces" checkbox was added to the Holdings parser's Tweak and Test panel;
    '       Made improvements to the management of default settings for Spine, Custom, Custom/Flag slips 
    '       and Pocket Labels
    '       A "copy to clipboard" feature was added to send Report text and CurrentXML/settings.som
    '       text to the Windows clipboard.
    '
    'v. 5.0: Major release to add Pocket Label printing;
    '       Repairs errors due to unencoded angle bracket characters appearing in the data of
    '       returned XML files.
    '       Does not check the arc.bc.edu:8080 server at startup, but only when Check for
    '       Updates is clicked.  (Due to occasional arc crashes that prevent SpineOMatic from
    '       starting.)
    '       Allows any call_number_type to be handled by any of SpineOMatic's parsing routines. 
    '       Blank <call_number_type> can be converted to any specified type (0 - 8).
    '       Added option to the LC Tweak and Test panel to suppress the decimal that normally
    '       precedes the cutter.
    '       Added a Holdings parser to the Tweak and Test panels.
    '       Call number formats (Spine, Custom & Custom w/Flag Slips and Pocket Labels) each 
    '       have their own separate set of margin settings and other defaults.
    '       Added formatting characters "^" to suppress newline after field, "*" to
    '       suppresse display of a field if it is blank or zero, and "+" to look up <location_name>
    '       in the Label Prefixes table and use the label text (that allows line breaks via semicolons).  
    '       Increased maximum number of label copies from 5 to 99. 
    '       Added a "cancel print" option for Desktop printing, and added a warning for Batch 
    '       and FTP printing if more than 5 label copies are requested.
    '       Added keyboard shortcut CTRL p to trigger a manual print without having to use 
    '       the mouse.  
    '       Added a License Agreement that requires the user to either accept terms or cancel 
    '       installation on first use of software, change of version, or relocation to a different PC.
    '
    'v. 4.32: Minor release to fix wrapping (if wrapping was turned on for one field, it
    '       stayed on for other fields that did not specify wrapping).  Added a formatting
    '       code to add a text prefix to custom label fields, as well as to Spine fields
    '       "Include holdings" and "Include other value".  Double quotes around text cause
    '       text to be prefixed to the printed value. Eg: "copy: "<copy_id>
    '       Redid the fix (originally in v. 4.3) that was supposed to prevent loss of Custom
    '       fields upon saving.  
    '
    'v. 4.31: Minor release to fix a bug that prevented "Include holdings" from working.
    '       This is the first release to use two digits after the decimal of the release number.
    '
    'v. 4.3: Fixed a bug causing multi-cuttered LC call numbers to hide the decimal when
    '       breaking on cutter.  Fixed bug that caused Custom fields to be lost when
    '       user saved settings while Flag Slips checkbox was checked.
    '       Replaced code written to parse the Ex Libris XML file with VB.NET's XML parser.
    '       Also alerted user if errors were detected in
    '       user-specified XML fields, i.e., not found, extraneous characters, etc.  
    '       Checkbox added to either display error alert only, or to pop up a detailed message. 
    '       Added ability for user to add formatting characters to Custom fields:
    '       (%=parse call#, #=parse holdings, !=render as barcode, ~=add space) before entry. 
    '       Also allowed space (~) to be added before "Other" field in Spine label section.
    '       Added multi-label print capability, allowing label to be printed from 1 to 5 times.
    '       
    'v. 4.2: Fixed a bug to allow SpineOMatic to recognize international date settings. 
    '       Added Tweak and Test panel to allow user to modify the behaviors of
    '       SpineOMatic's parsing routines.
    '       Added a Dewey Decimal and an "Other" parser.
    '       Moved the Test Parsing section from Java Setup to
    '       the Tweak and Test Parsing panel.  Removed the portrait/landscape distinction
    '       when SpineOMatic parsed SuDoc numbers.  User can now set up parsing for one or
    '       the other.
    '       Added better checking for Java URL problems and credential issues.
    '       If the customer's PC cannot connect to BC servers due to blocking by their 
    '       proxy server, a message tells them to whitelist the BC servers.
    '
    'v. 4.1: Changes to wording and layout of Print Flag Slips checkbox and Label Printing
    '       Web Service Credentials. 
    '
    'v. 4.0: Removed need to provide a folder to receive Alma XML file.  The installation
    '       directory will be used by default;
    '       Imported graphic background for the About box that contains the BC seal graphic; 
    '       Java app class file and alma-sdk files are now automatically downloaded if needed, 
    '       without manual intervention;
    '       Added separate margin/orientation/maximum settings for Flag Slips.  Toggling the 
    '       "print flag slips" checkbox calls up Flag Slip settings or returns to standard settings;
    '       The Java application is now run as a process from within vb rather than as an
    '       external .bat file. Java installation is verified, and problems locating or 
    '       accessing java are reported to the user.
    '       A list of servers can be specified from which to obtain updates (i.e., the updatePath.
    '       If the default server fails, each server in the list is tried in turn to try to find
    '       a working server.  If none can be found, a "fail" message is displayed.
    '=======================================================================================
    'v. 3.3: Added SuDoc parsing for portrait and landscape modes; Changes to error checking
    '       and AboveCall#Text behavior; 
    '
    'v. 3.2: Added ability to put an additional field (e.g., <copy_id>) at end of spine label;
    '       Ensured label text in OutputBox does not end with unnecessary cr/lf;
    '       Fixed bug in Above Call# Text that produced incorrect matching.
    '       Limited User ID to 8 alpha characters.
    '
    'v. 3.1: Added Station name and User ID; Reports; Test Parsing; better LC/LC Children's lit/NLM
    '        call number parsing.
    '
    'v. 3.0: Cosmetic changes to admin panels; added "About" box with download/view
    '       of associated documentation. Added access to Alma Label Printing Web Service
    '       via desktop java app; Added option to use Ex Libris parsed call numbers.
    '=======================================================================================
    'v. 2.6: for "Custom" labels, text not enclosed in angle brackets (<...>) will print as-is on the label
    '       Bug fix: manual print button now checks line lengths against max. chars/line;
    'v. 2.5: adds textbox for url to Alma Label Printing Web Service
    'v. 2.4: adds barcode font dialog selection for use in flag slips;
    'v. 2.3: dlgSettings.UseEXDialog = True to enable print dialog selection in Windows 7
    'v. 2.2: corrects spacing & punctuation errors in incoming call numbers (for TML);
    'v. ...: 
    Dim somVersion As String = "8.1.1"
    Dim javaClassName As String = "almalabelu2" 'the java class name
    Dim javaSDKName As String = "alma-sdk.1.0.jar" 'the Ex Libris SDK for web services
    Dim javaTest As String = "javatest" 'java class that reports presence and version of java
    Dim mypath As String = "" 'path of startup directory will be used as mypath
    Dim servers As String = "arc.bc.edu:8080|libstaff.bc.edu:8080|mlib.bc.edu:8080"
    Dim lcxml As String = ""
    Dim issuexml As String = ""
    Dim locxml As String = ""
    Dim libxml As String = ""
    Dim otherxml As String = ""
    Dim titlexml As String = ""
    Dim libraryxml As String = ""
    Dim pixelsPerInchX As Integer = 0
    Dim pixelsPerInchY As Integer = 0
    Dim changeCount As Integer = 0
    Dim xmlReturned As String = ""
    Dim settings As String = ""
    Dim winFrom As Integer = 0
    Dim winTo As Integer = 0
    Dim wline As Array = Nothing
    Dim wlinesToPrint As Integer = 0
    Dim origText As String = ""
    Dim editText As String = ""
    Dim maxLines As Integer = 0
    Dim LABELS As Array
    Dim nxt As Integer = 0
    Dim horizPos As Integer = 0
    Dim fontname As String = ""
    Dim fontsize As Single = 0.0
    Dim fWeight As System.Drawing.FontStyle
    Dim bcWeight As System.Drawing.FontStyle
    Dim topMargin As Single = 0.0
    Dim leftMargin As Single = 0.0
    Dim lineSpacing As Single = 0.0
    Dim labelRows As Integer
    Dim labelCols As Integer
    Dim labelWidth As Single
    Dim labelHeight As Single
    Dim gapWidth As Single
    Dim gapHeight As Single
    Dim original_settings As String = ""
    Dim closing_settings As String = ""
    Dim saveTab As TabPage
    Dim lastxml As String = ""
    Dim ignoreChange As Boolean = True
    Dim ALTfile As String = ""
    Dim madeALTchanges As Boolean = False
    Dim statrec As String
    Dim lastbc As String = ""
    Dim cntype As String
    Dim almaReturnCode As String = ""
    Dim almaLibrary As String = ""
    Dim almaLocation As String = ""
    Dim usermessage As String = ""
    Dim settingsfound As Boolean = True
    Dim settingsLoaded As Boolean = False
    Dim settingsOpen As Boolean = False
    Dim logView As Boolean = False
    Dim flagSlipDefaults As String = ""
    Dim firstPage As Boolean = True
    Dim otherList As String = ""
    Dim xmlerr As String = ""
    Dim indenting As Boolean = False
    Dim wrapping As Boolean = False
    Dim totalLines As Integer = 0
    Dim labelCount As Integer = 0
    Dim needTypeCheck As Boolean = False 'alerts if call number parsers have been changed
    Dim xdoc As New System.Xml.XmlDocument
    Dim warranty_accepted As Boolean = True
    Dim spin As Integer = 1
    Dim stopPrinting As Boolean = False
    Dim spineDefaultLoaded As Boolean = False
    Dim customNonFlagDefaultLoaded As Boolean = False
    Dim customFlagDefaultLoaded As Boolean = False
    Dim pocketDefaultLoaded As Boolean = False
    Dim WithEvents client As New WebClient
    Dim licenseDeclined As Boolean = True
    Dim pcname As String = ""
    Dim spineVerticalLine As Boolean
    Dim nonFlagVerticalLine As Boolean
    Dim flagVerticalLine As Boolean
    Dim pocketVerticalLine As Boolean
    Dim usingDewey As Boolean = False
    Dim xtb As TextBox
    Dim xtbOrigColor As Color
    Private Const LB_SETTABSTOPS As Int32 = &HCB


    <DllImport("user32.dll")>
    Private Shared Function SendMessage(
       ByVal hWnd As IntPtr,
       ByVal wMsg As Int32,
       ByVal wParam As IntPtr,
       ByVal lParam As IntPtr) _
       As Int32
        'DLL import is used to set margins in the Reports ("StatsOut") textbox
    End Function
    Private Sub SetTabs()
        '{0, 65, 110, 165, 240, 255} (original settings)
        Dim ListBoxTabs() As Integer = {0, 60, 110, 180, 240, 255}
        Dim result As Integer
        Dim ptr As IntPtr
        Dim pinnedArray As GCHandle

        pinnedArray = GCHandle.Alloc(ListBoxTabs, GCHandleType.Pinned)
        ptr = pinnedArray.AddrOfPinnedObject()
        'Send LB_SETTABSTOPS message to TextBox.
        result = SendMessage(Me.statsOut.Handle, LB_SETTABSTOPS,
          New IntPtr(ListBoxTabs.Length), ptr)
        pinnedArray.Free()

        'Refresh the TextBox control.
        Me.statsOut.Refresh()
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If licenseDeclined Then Exit Sub
        Dim resp As String = ""
        writeStat("S") 'write "S" (scanned, not printed) to statrec, and write to stat file.
        If madeALTchanges = True Then
            Dim box = MessageBox.Show("Changes to your local label text file have not been saved." & vbCrLf & "Do you want to save them now?", "Save Settings", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If box = box.Yes Then
                btn_saveALT.PerformClick()
                MsgBox("Changes to your local label text file have been saved.", MsgBoxStyle.Information, "Settings Saved")
                madeALTchanges = False
            End If
        End If
        saveSettings("tostring") 'put current settings into the closing_settings string
        closing_settings = closing_settings.Replace(vbLf, "")
        original_settings = original_settings.Replace(vbLf, "")
        If original_settings <> closing_settings Then
            'Clipboard.SetText("orig:" & vbCrLf & original_settings & vbCrLf & "new:" & vbCrLf & closing_settings)
            Dim box = MessageBox.Show("Your settings have changed, but have not been saved." & vbCrLf &
            "Do you want to save them now?" & vbCrLf & vbCrLf &
            "(Click CANCEL to continue working.)", "Save Settings", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If box = box.Yes Then
                saveSettings("todisk")
                MsgBox("Your settings have been saved.", MsgBoxStyle.Information, "Settings Saved")
            Else
                If box = box.cancel Then
                    e.Cancel = True
                End If
            End If
        End If
    End Sub

    Private Sub Form1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If Asc(e.KeyChar) = 1 Then 'keychar 1 = "CTRL a"
            e.Handled = True
            If settingsOpen = False Then
                openSettings()
            Else
                'Me.CloseSettings_Click(Nothing, Nothing)
                CloseSettings()
            End If
        End If
        If Asc(e.KeyChar) = 16 Then 'CTRL p
            e.Handled = True
            ManualPrint.PerformClick()
        End If

    End Sub
    Private Sub NumericKeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles inMaxLines.KeyPress, inLineSpacing.KeyPress, TextBox23.KeyPress, TextBox22.KeyPress, TextBox21.KeyPress, TextBox20.KeyPress, TextBox19.KeyPress, inMaxChars.KeyPress, inLabelWidth.KeyPress, inLabelRows.KeyPress, inLabelHeight.KeyPress, inLabelCols.KeyPress, inGapWidth.KeyPress, inGapHeight.KeyPress, inFontSize.KeyPress, inBCFontSize.KeyPress, inStartCol.KeyPress, inStartRow.KeyPress, wrapWidth.KeyPress, plMin4.KeyPress, plMin3.KeyPress, plMin2.KeyPress, plMin1.KeyPress, plMax4.KeyPress, plMax3.KeyPress, plMax2.KeyPress, plMax1.KeyPress, plDistance.KeyPress, plLeftMargin.KeyPress, convertBlankTo.KeyPress, dosBlankLines.KeyPress, dosPlTabNum.KeyPress, dosPlColNum.KeyPress, appendAscii.KeyPress  'Handles TextBox.KeyPress
        Dim tb As TextBox = sender
        Dim dc As String = ""
        If decimalDOT.Checked Then dc = "." Else dc = ","
        If Not (Char.IsDigit(e.KeyChar) Or Char.IsControl(e.KeyChar) Or (e.KeyChar = dc And tb.Text.IndexOf(dc) < 0)) Then
            e.Handled = True
            Beep()
            Exit Sub
        End If
    End Sub
    Private Sub NegativeKeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles inTopMargin.KeyPress, inLeftMargin.KeyPress
        Dim tb As TextBox = sender
        Dim dc As String = ""
        If decimalDOT.Checked Then dc = "." Else dc = ","
        If Not (Char.IsDigit(e.KeyChar) Or Char.IsControl(e.KeyChar) Or (e.KeyChar = dc And tb.Text.IndexOf(dc) < 0) Or e.KeyChar = "-") Then
            e.Handled = True
            Beep()
        End If
    End Sub

    Private Sub NumericLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles inMaxLines.Leave, inTopMargin.Leave, inLineSpacing.Leave, inLeftMargin.Leave, TextBox23.Leave, TextBox22.Leave, TextBox21.Leave, TextBox20.Leave, TextBox19.Leave, inMaxChars.Leave, inLabelWidth.Leave, inLabelRows.Leave, inLabelHeight.Leave, inLabelCols.Leave, inGapWidth.Leave, inGapHeight.Leave, inFontSize.Leave, inBCFontSize.Leave, inStartCol.Leave, inStartRow.Leave, wrapWidth.Leave, plMin4.Leave, plMin3.Leave, plMin2.Leave, plMin1.Leave, plMax4.Leave, plMax3.Leave, plMax2.Leave, plMax1.Leave, plDistance.Leave, plLeftMargin.Leave, convertBlankTo.Leave, dosBlankLines.Leave, dosPlTabNum.Leave, dosPlColNum.Leave, appendAscii.Leave, deweydigitsperline.Leave, deweyDigitsToBreak.Leave  'Handles TextBox.KeyPress
        Dim tb As TextBox = sender
        If tb.Text.Length = 0 Then
            tb.Text = "0"
        End If
    End Sub
    Private Sub limitValues(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles deweydigitsperline.KeyPress, deweyDigitsToBreak.KeyPress
        Dim tb As TextBox = sender

        If Not "234567".Contains(e.KeyChar) Then
            e.Handled = True
            Beep()
        Else
            'deweydigitsperline.Text = e.KeyChar
            sender.Text = e.KeyChar
            e.Handled = True
        End If

    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        settingsfound = True
        Dim disclaimerInfo As String = ""
        Dim currentLicense As String = ""
        pcname = System.Net.Dns.GetHostName()
        Me.KeyPreview = True
        Me.Show()
        mypath = System.AppDomain.CurrentDomain.BaseDirectory()
        Me.Text = "SpineOMatic " & somVersion
        CloseSettings() 'close the settings panels
        Application.DoEvents()
        continueFormLoad()

    End Sub

    Private Sub continueFormLoad()
        licenseDeclined = False
        GetSettingsFile()
        'CloseSettings() Let the "GetSettingsFile" routine determine if settings exist or not.
        'If no settings file exists, panels should remain open
        Application.DoEvents()
        settingsLoaded = True
        original_settings = RichTextBox1.Text
        XMLPath.Text = mypath 'path is now always set to the installation directory, "mypath"
        Try
            FileSystemWatcher1.Path = XMLPath.Text
            FileSystemWatcher2.Path = XMLPath.Text

        Catch
            MsgBox("Directory to watch for incoming XML files is not valid." & vbCrLf &
            "Path = " & XMLPath.Text, MsgBoxStyle.Exclamation, "Invalid Path")
            FileSystemWatcher1.Path = ""
        End Try

        batchPreview.Text = GetBatch(batchNumber.Value)
        If batchPreview.Lines.Length > 0 Then
            'batchEntries.Text = batchPreview.Lines.Length - 1
            batchEntries.Text = countBatch()
        Else
            batchEntries.Text = "0"
        End If
        btnMonitor.Enabled = False

        createBatFiles()

        downloadAboveLcFile()
        Application.DoEvents()
        loadLabelText()
        lblStation.Text = station.Text
        usermessage = "Please enter your User ID in the 'User:' box above." & vbCrLf & vbCrLf &
            "The ID must be 8 characters or less." & vbCrLf & vbCrLf &
            "When done, press the ENTER key."
        If chkRequireUser.Checked Then
            usrname.Enabled = True
            OutputBox.Text = usermessage
            usrname.BackColor = Color.Yellow
            usrname.Focus()
        Else
            usrname.Text = "[none]"
            usrname.Enabled = False
            InputBox.Select()
            InputBox.Focus()
        End If

        Dim date1 As Date = Date.Now
        Dim dtNow As Date = Date.Now
        Dim dtFirstOfMonth As Date = dtNow.AddDays(-dtNow.Day + 1)
        fromScan.Format = DateTimePickerFormat.Short
        fromScan.CustomFormat = "MM/dd/yyyy"
        toScan.Format = DateTimePickerFormat.Short
        toScan.CustomFormat = "MM/dd/yyyy"
        'toScan.Value = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.CurrentCulture)
        'toScan.Value = date1.Month & "/" & date1.Day & "/" & date1.Year
        fromScan.Value = dtFirstOfMonth

        toScan.Value = date1.Date.ToString
        SetTabs() 'change tab settings of the Reports textbox.

        TabControl1.SelectedIndex = 1
        TabControl1.SelectedIndex = 0
        lbl_setclipboard.ForeColor = Color.MediumBlue
        InputBox.Focus()
    End Sub



    Private Function countBatch() As String
        Dim ln As Integer = -1
        Dim pos As Integer = 1
        Do
            pos = InStr(pos + 1, batchPreview.Text, "===============", CompareMethod.Text)
            ln = ln + 1
        Loop While pos <> 0
        Return CType(ln + 1, String)
    End Function
    Private Sub FileSystemWatcher1_Created(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs) Handles FileSystemWatcher1.Created
        'Watches a selected directory for the arrival of an Ex Libris' Alma item XML file
        'and generates a spine label when one arrives.
        '
        'If the file is being written into the directory by another application (i.e., a
        'Java program that retrieves the Alma file and writes in into the specified
        'directory), the FileSystemWatcher may fire before the file is completely written.
        'This routine waits until the file can be successfully read.  It tries up to 20
        'times, waiting 100ms between tries.
        Dim fileFound As Boolean = False
        Dim loopcnt As Integer
        loopcnt = 20

        Do While loopcnt > 0
            Try
                Dim tr As TextReader = New StreamReader(e.FullPath)
                xmlReturned = tr.ReadToEnd()
                tr.Close()
                fileFound = True
                Exit Do
            Catch ex As Exception
                loopcnt = loopcnt - 1
                Thread.Sleep(100)
            End Try
        Loop

        If fileFound Then
            InputBox.Text = e.Name.Replace(".xml", "")
            lastxml = e.FullPath

            'If xmlReturned.Contains("<bib_data link") Then
            '    xmlReturned = convertRESTfulXML()
            'End If

            getBarcodeFile()
            If AutoPrintBox.Checked Then
                ManualPrint.PerformClick()
            End If
        Else
            MsgBox("The complete Alma XML file did not arrive: " & e.FullPath, MsgBoxStyle.Exclamation, "File Incomplete")
        End If
    End Sub

    Private Function convertRESTfulXML() As String
        'Converts RESTful XML files into depricated SOAP format so that existing SpineOMatic code can 
        'be used with the new XML format.
        Dim doc As New XmlDocument
        Dim t As String = ""
        Dim cn As String = ""
        Dim ct As String = ""
        Dim n As String = ""
        Dim titl As String = ""
        Dim cntype As String = ""
        Dim e As Integer
        Dim mynode As XmlNode
        Dim anode As XmlNode
        Dim nl As XmlNodeList
        Dim pcn As String = "" 'parsed call number
        Dim pild As String = "" 'parsed issue level description
        Dim i As Integer = 0

        'The RESTful XML text is loaded into an XML document
        doc.LoadXml(xmlReturned)
        Dim elemList As XmlNodeList = doc.GetElementsByTagName("item_data") 'most user fields are under
        '<item_data>...</item_data>

        'Put all RESTful <item_data> fields into string "t"
        For e = 0 To elemList.Count - 1
            t = elemList(e).InnerXml & vbCrLf
        Next e

        'RESTful parsed_call_number and parsed_issue_level_description fields need to be modified 
        'in the final text.  Here is where we remove these fields from the text string "t":
        If t.Contains("<parsed_call_number") Then
            t = t.Substring(0, t.IndexOf("<parsed_call_number>")) & t.Substring(t.IndexOf("</parsed_call_number>") + 21)
        End If
        If t.Contains("<parsed_issue_level_description>") Then
            t = t.Substring(0, t.IndexOf("<parsed_issue_level_description>")) & t.Substring(t.IndexOf("</parsed_issue_level_description>") + 32)
        End If

        'These routines step through each element of the <parsed_call_number> and <parsed_issue_level_desctiption>
        'fields in the XML document, and a sequence number is added to the XML field names.
        'Ex:  <call_no>BX</call_no>  is changed to <call_no_1>BX</call_no_1>, etc...
        i = 1
        pcn = ""
        For Each mynode In doc.SelectNodes("/item/item_data/parsed_call_number/*")
            If IsNothing(mynode) Then Exit For
            pcn = pcn & "<call_no_" & i & ">" & mynode.InnerXml & "</call_no_" & i & ">" & vbCrLf
            i = i + 1
        Next
        pcn = "<parsed_call_number>" & vbCrLf & pcn & vbCrLf & "</parsed_call_number>"

        i = 1
        pild = ""
        For Each mynode In doc.SelectNodes("/item/item_data/parsed_issue_level_description/*")
            If IsNothing(mynode) Then Exit For
            pild = pild & "<issue_level_description_" & i & ">" & mynode.InnerXml & "</issue_level_description_" & i & ">" & vbCrLf
            i = i + 1
        Next
        pild = "<parsed_issue_level_description>" & vbCrLf & pild & vbCrLf & "</parsed_issue_level_description>"
        'the new fields are stored in variables 'pcn' (parsed call number) and 'pild' (parsed issue level description)
        'fields, and these modified fields are added back to string "t" later in the process.

        '<call_number> is not in the <item_data> section, so it's put in variable 'cn' and added to string 't' later
        nl = doc.GetElementsByTagName("call_number")
        cn = nl(0).OuterXml

        '<call_number_type> is not in <item_data> either, so it is extracted in 'ct', and later added to string 't'
        nl = doc.GetElementsByTagName("call_number_type")
        ct = "<call_number_type>" & nl(0).InnerXml & "</call_number_type>"

        '<title> is not in <item_data>, so it is extracted to 'nl' and added back to string 't' later.
        nl = doc.GetElementsByTagName("title")
        titl = nl(0).OuterXml

        '<library desc="O'Neill">ONL</library>' is changed into two fields:
        '1) <library_code>ONL</library_code>
        '2) <library_name>O'Neill</library_name>
        'liname and licode are added to string 't'
        Dim lid As XmlNodeList = doc.GetElementsByTagName("library")
        Dim licode As String = "<library_code>" & lid(0).InnerXml & "</library_code>"
        anode = doc.SelectSingleNode("//library")
        Dim liname As String = "<library_name>" & anode.Attributes(0).Value & "</library_name>"

        '<location desc="Offsite Collection (RM150 GOVD)">RM150_GOVD</location>  is changed into:
        '1) <location_name>Offsite Collection (RM150 GOVD)</location_name>
        '2) <location_code>RM150_GOVD</location_code>
        'lod and locode are added to string 't'
        Dim lod As XmlNodeList = doc.GetElementsByTagName("location")
        Dim locode As String = "<location_code>" & lod(0).InnerXml & "</location_code>"
        anode = doc.SelectSingleNode("//location")
        Dim loname As String = "<location_name>" & anode.Attributes(0).Value & "</location_name>"

        'all the new XML elements that were relocated or created are added back to string 't', and 
        'string 't' is inserted into an 'XmlShell' invisible text box, replacing the text "**XMLBODY**"
        'that identifies the spot where the new XML is to be interted.
        xmlReturned = xmlShell.Text.Replace("**XMLBODY**", cn & vbCrLf & ct & vbCrLf & titl & vbCrLf & licode & vbCrLf & liname & vbCrLf & loname & locode & vbCrLf & pcn & vbCrLf & pild & vbCrLf & t)

        'the new RESTful "<description>" field names are changed to <issue_level_description>, to
        'mimic the SOAP naming convention
        xmlReturned = xmlReturned.Replace("<description>", "<issue_level_description>")
        xmlReturned = xmlReturned.Replace("</description>", "</issue_level_description>")

        Return xmlReturned
    End Function

    Private Sub ManualPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ManualPrint.Click
        'manual print
        If ManualPrint.Text = "Stop Printing" Then
            stopPrinting = True
            SetPrintButtonText()
            Exit Sub
        End If
        Dim dayTime As String = ""
        Dim logentry As String = ""
        Dim barcodenum As String = ""
        Dim lencheck As Integer = 0
        Dim chkline As Array = Nothing
        Dim maxlines As Integer = CType(inMaxLines.Text, Integer)
        Dim linesOK As Boolean = True
        Dim maxchars As Integer = CType(inMaxChars.Text, Integer)
        Dim repeat As Integer = 0
        Dim i As Integer = 0
        Dim batchText As String = ""
        Dim labelin As String
        plDistance.BackColor = Color.White
        If chkUsePocketLabels.Checked Then
            If btnSL4.Checked Or btnSL6.Checked Or (btnPlCustom.Checked And PLcount.Value = 2) Then
                If CType(plDistance.Text, Single) = 0.0 Then
                    plDistance.BackColor = Color.Pink
                    MsgBox("When printing two pocket labels, you must specify a vertical distance" & vbCrLf &
                    "between the top lines of the two labels.", MsgBoxStyle.Exclamation, "No Distance Specified")
                    plDistance.Focus()
                    Exit Sub
                End If
            End If
        End If

        If Trim(OutputBox.Text) = "" Then
            Beep()
            InputBox.Focus()
            Exit Sub
        End If
        editText = OutputBox.Text
        barcodenum = InputBox.Text
        If (editText <> origText) And barcodenum <> "" And logEdits.Checked Then
            dayTime = DateTime.Now.ToString("ddd MMM d, yyyy HH:mm", CultureInfo.InvariantCulture)
            logentry = dayTime & vbTab & barcodenum & vbTab & origText.Replace(vbCrLf, "|") & vbTab & editText.Replace(vbCrLf, "|")
            writeFile(mypath & "changelog.txt", logentry, True)
            editText = ""
            origText = ""
        End If

        If OutputBox.Text.Contains("** ERROR **") Then
            Beep()
            MsgBox("Could not find this barcode number in Alma.", MsgBoxStyle.Exclamation, "Barcode Number Error")
            Exit Sub
        End If
        If OutputBox.Text.Contains("Java: **") Then
            Beep()
            MsgBox("Could not contact Alma.", MsgBoxStyle.Exclamation, "Connection Error")
        End If
        lencheck = checkLineLength()
        chkline = Split(OutputBox.Text, vbCrLf)
        If lencheck <> 99 Then
            'chkline = Split(OutputBox.Text, vbCrLf)
            MsgBox("Line #" & lencheck & ":" & vbCrLf & chkline(lencheck - 1) &
            vbCrLf & " contains more than " & maxchars & " characters.")
            Exit Sub
        End If
        If chkline.Length > maxlines Then
            Beep()
            MsgBox("Lable contains more than " & maxlines & " lines.", MsgBoxStyle.Exclamation, "Too Many Lines")
            Exit Sub
        End If
        If lblXMLWarn.Visible = True Then
            Beep()
            MsgBox("An XML <field> used in your settings is incorrect.", MsgBoxStyle.Exclamation, "XML Reference Error")
            Exit Sub
        End If

        '********************
        ' Use DOS Batch File
        '********************
        If useDOSBatch.Checked Then
            Dim txtout As String = ""
            Dim extraLines As Integer = CType(dosBlankLines.Text, Integer)
            Dim addcr As String = ""
            Dim tabpos As Integer = 0
            Dim k As Integer = 0
            Dim extraSpaces As Integer = CType(dosPlColNum.Text, Integer)
            Dim tabcount As Integer = CType(dosPlTabNum.Text, Integer)
            Dim addsp As String = ""
            Dim taray As Array
            Dim mg As String = ""

            If chkUsePocketLabels.Checked Then
                taray = packagePocket().Split(vbCrLf)
                txtout = ""
                For k = 0 To taray.Length - 1
                    tabpos = taray(k).replace(vbLf, "").IndexOf(vbTab)
                    If dosPlUseCol.Checked Then
                        addsp = New String(" ", extraSpaces - tabpos)
                        txtout = txtout & taray(k).Replace(vbTab, addsp) & vbCrLf
                    Else
                        If tabcount > 1 Then
                            addsp = New String(vbTab, tabcount)
                        Else
                            addsp = vbTab
                        End If
                        txtout = txtout & taray(k).replace(vbTab, addsp) & vbCrLf
                    End If
                Next
                txtout = txtout & terminator(taray.Length)
            Else
                txtout = txtout & OutputBox.Text & terminator(OutputBox.Lines.Length)
            End If

            viaDOS(txtout)
            Exit Sub
        End If

        If UseDesktop.Checked Then
            getPrintParams()
            PrintDocument2.PrinterSettings.PrinterName = inPrinterName.Text
            PrintDocument2.PrintController = New System.Drawing.Printing.StandardPrintController

            If chkUsePocketLabels.Checked Then
                labelin = packagePocket()
                labelin = labelin.Replace(vbCrLf, "|")
            Else
                labelin = OutputBox.Text.Replace(vbCrLf, "|")
            End If

            'For desktop printing, only one label will be in the "LABELS" array,
            'but the print routine always uses the LABELS array to get its input,
            'for single label printing and for multi-label batch printing.
            LABELS = labelin.Split(vbCrLf)

            repeat = CType(LabelRepeat.Value, Integer)
            If repeat > 1 And ManualPrint.Text <> "Stop Printing" Then
                ManualPrint.Text = "Stop Printing"
                printProgress.Visible = True
                Application.DoEvents()
            End If
            For i = 1 To repeat
                Application.DoEvents()
                If stopPrinting Then
                    SetPrintButtonText()
                    stopPrinting = False
                    printProgress.Visible = False
                    Exit For
                End If
                Try
                    printProgress.Text = "Printing " & i & " of" & repeat
                    Application.DoEvents()
                    PrintDocument2.Print()
                Catch ex As Exception
                    MsgBox("Printer settings are not correct." & vbCrLf &
                    "Make sure a valid printer has been selected, and try again." &
                    vbCrLf & ex.ToString, MsgBoxStyle.Exclamation, "Printer Selection Error")
                    Exit For
                End Try
            Next i
            '********************
            SetPrintButtonText()
            '********************
            printProgress.Visible = False
            nxt = 0
            writeStat("P") 'add P to end of statrec and write to statfile.
            OutputBox.Text = ""
            plOutput.Text = ""
            TempLabelBox.Text = ""
            InputBox.Text = ""
            InputBox.Focus()
        Else
            repeat = CType(LabelRepeat.Value, Integer)
            If repeat > 5 AndAlso MessageBox.Show(repeat & " labels will be printed.", "Confirm Multipl Label Request", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.Cancel Then Exit Sub
            If UseLaser.Checked Then 'SAVE TO BATCH
                If chkUsePocketLabels.Checked Then
                    batchText = packagePocket()
                Else
                    batchText = OutputBox.Text
                End If
                repeat = LabelRepeat.Value
                For i = 1 To repeat
                    sendToBatch2(batchText)
                Next i
                writeStat("B") 'add B to end of statrec and write to statfile
                OutputBox.Text = ""
                plOutput.Text = ""
                TempLabelBox.Text = ""
                InputBox.Text = ""
                InputBox.Focus()
            Else
                ftpPrint()
            End If
        End If
    End Sub
    Private Function terminator(ByVal actualLen As Integer) As String
        Dim trm As String = ""
        Dim desiredLen As Integer = CType(dosBlankLines.Text, Integer)
        If desiredLen > actualLen Then
            If dosAddLines.Checked Then
                trm = New String("^", desiredLen - actualLen).Replace("^", vbCrLf)
            Else
                trm = Chr(CType(appendAscii.Text, Integer))
            End If
        Else
            trm = ""
        End If
        Return trm
    End Function
    Private Sub SetPrintButtonText()
        ManualPrint.ForeColor = Color.Black
        If UseFTP.Checked Then ManualPrint.Text = "Send to FTP Printer"
        If UseLaser.Checked Then ManualPrint.Text = "Add to batch #" & batchNumber.Value.ToString
        If UseDesktop.Checked Then ManualPrint.Text = "Send to desktop printer"
        'If useDOSBatch.Checked Then ManualPrint.Text = "Run viados.bat"
    End Sub

    Private Function packagePocket() As String
        'combines spine label in OutputBox with pocket label in plOutput.
        'A tab character separates lines of spine text from a line of pocket label text.
        Dim k As Integer = 0
        Dim sp As String = ""
        Dim s As Integer = 0
        Dim p As Integer = 0
        Dim stext As String = ""
        Dim ptext As String = ""
        Do
            If s <= OutputBox.Lines.Length - 1 Then stext = OutputBox.Lines(s) : s = s + 1 Else stext = ""
            If p <= plOutput.Lines.Length - 1 Then ptext = plOutput.Lines(p) : p = p + 1 Else ptext = ""
            If stext = "" And ptext = "" Then Exit Do
            sp = sp & stext & vbTab & ptext & vbCrLf
            k = k + 1 : If k > 20 Then Return "*** Error preparing pocket labels for printing ***"
        Loop

        sp = sp.Substring(0, sp.Length - 2) 'remove final crlf

        Return sp
    End Function

    Private Sub btnPrintBatch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintBatch.Click
        Dim labelsIn As String
        Dim wrk As String
        labelsIn = GetBatch(batchNumber.Value)

        'modified------------------------------
        'LABELS = labelsIn.Split(vbCrLf)
        '--------------------------------------

        wrk = labelsIn.Replace(vbCrLf & "===============" & vbCrLf, "^")
        wrk = wrk.Replace(vbCrLf, "|")
        wrk = wrk.Replace("^", vbCrLf)
        'MsgBox("wrk:" & wrk)
        LABELS = wrk.Split(vbCrLf)

        If LABELS.Length = 0 Then
            MsgBox("There are no labels in batch number " & batchNumber.Value)
            Exit Sub
        End If

        nxt = 0
        getPrintParams()
        firstPage = True
        PrintDocument2.PrinterSettings.PrinterName = inPrinterName.Text
        PrintPreviewDialog1.Document = PrintDocument2
        Try
            PrintPreviewDialog1.ShowDialog()
        Catch
            MsgBox("The specified printer: " & vbCrLf & vbCrLf & inPrinterName.Text & vbCrLf & vbCrLf & "does not exist." _
            & vbCrLf & "Please select a different printer.", MsgBoxStyle.Exclamation, "Printer not found")
        End Try

    End Sub

    Private Sub getPrintParams()
        Dim cf As Single = 0.0
        If unitINCH.Checked Then
            cf = 100
        Else
            cf = 39.3701
        End If
        horizPos = 0
        fontname = inFontName.Text
        fontsize = CType(inFontSize.Text, Single)
        fWeight = FontStyle.Regular
        If inFontWeight.Checked Then fWeight = FontStyle.Bold
        maxLines = CType(inMaxLines.Text, Integer)
        topMargin = CType(inTopMargin.Text, Single) * cf '100
        leftMargin = CType(inLeftMargin.Text, Single) * cf '100
        lineSpacing = CType(inLineSpacing.Text, Single) * cf '100
        If UseLaser.Checked Then 'for batch printing, set params for multi-row & multi-column sheets
            labelRows = CType(inLabelRows.Text, Integer)
            labelCols = CType(inLabelCols.Text, Integer)
            labelWidth = CType(inLabelWidth.Text, Single) * cf '100
            labelHeight = CType(inLabelHeight.Text, Single) * cf '100
            gapWidth = CType(inGapWidth.Text, Single) * cf '100
            gapHeight = CType(inGapHeight.Text, Single) * cf '100
        Else ' for single-label desktop printing, rows & columns are always "1", with no gaps.
            labelRows = 1
            labelCols = 1
            labelWidth = 0
            labelHeight = 0
            gapWidth = 0
            gapHeight = 0
        End If
    End Sub
    Private Sub sendToBatch(ByVal labelout As String)
        'add a spine label or book flag to the end of the selected label batch
        labelout = labelout.Replace(vbCrLf, "|")

        'labelout = vbCrLf + "----------" & vbCrLf & labelout
        labelout = Mid$(labelout, 1, labelout.Length) 'remove final "|"
        writeFile(mypath & "labelbatch" & batchNumber.Value & ".txt", labelout, True)
        batchPreview.Text = batchPreview.Text + labelout
        batchEntries.Text = batchEntries.Text + 1
    End Sub

    Private Sub sendToBatch2(ByVal labelout As String)
        'NEW METHOD
        'add a spine label or book flag to the end of the selected label batch
        If CType(batchEntries.Text, Integer) > 0 Then
            labelout = vbCrLf & "===============" & vbCrLf & labelout
        End If
        writeFile(mypath & "labelbatch" & batchNumber.Value & ".txt", labelout, True)
        batchPreview.Text = batchPreview.Text + labelout
        batchEntries.Text = batchEntries.Text + 1

    End Sub

    Private Sub PrintDocument2_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument2.PrintPage
        Dim x As Single = 0.0
        Dim y As Single = 0.0
        Dim ypos As Single = 0.0
        Dim z As Integer = 0
        Dim t As Integer = 0
        Dim down As Integer = 0
        Dim across As Integer = 0
        Dim i As Integer = 0
        Dim pFont = New Font(fontname, fontsize, fWeight)
        If inBCFontWeight.Checked Then
            bcWeight = FontStyle.Bold
        Else
            bcWeight = FontStyle.Regular
        End If
        Dim bFont = New Font(inBCFontName.Text, CType(inBCFontSize.Text, Single), bcWeight)
        'Dim testlbl As String = "QA|76.73|.J38|H674|2006"
        Dim line As Array '= testlbl.Split("|")
        Dim test As String = ""
        Dim extraSpace As Integer = 0
        Dim bcwidth As Single = 0.0
        Dim cntrOffset As Integer = 0
        Dim nextLabel As String = ""
        Dim startRow As Integer = CType(inStartRow.Text, Integer)
        Dim startcol As Integer = CType(inStartCol.Text, Integer)

        Dim PlySave As Integer 'save starting point of call# label to use if a pocket label
        Dim PLx As Integer
        Dim PLy As Integer 'will subsequently be printed.
        Dim leftPocket As Integer = 0 'left margin of pocket label
        Dim plk As Integer = 0
        Dim plparams As Array
        Dim theseArePocketLabels As Boolean = False
        Dim cf As Single = 0.0
        If unitINCH.Checked Then
            cf = 100
        Else
            cf = 39.3701
        End If
        If chkUsePocketLabels.Checked Then
            plparams = PocketParams() 'called to get the labelCount (1 or 2) currently in effect
        End If
        If startRow > labelRows Then
            MsgBox("Starting row number is greater than the maximum rows.", MsgBoxStyle.Exclamation, "Starting Row Number Too Large")
            Exit Sub
        End If
        If startcol > labelCols Then
            MsgBox("Starting column number is greater than the maximum columns.", MsgBoxStyle.Exclamation, "Starting Column Number Too Large")
            Exit Sub
        End If
        ' values preceded by ! will display a barcode first, followed by the actual value on the next line
        ' values preceded by ~ will have extra vertical spacing
        If Not firstPage Then
            startRow = 1
            startcol = 1
        End If
        Do
            For down = startRow To labelRows
                y = topMargin + ((down - 1) * (labelHeight + gapHeight))

                For across = startcol To labelCols
                    nextLabel = getNextLabel() 'gets next label from the LABELS array
                    If nextLabel = "done" Then Exit Do
                    line = nextLabel.Split("|")

                    If line(0).contains(vbTab) Then 'pocket labels will always have a tab in the 1st line
                        theseArePocketLabels = True
                        plparams = PocketParams()
                        line = removePocket(line) 'split off the pocket label from the spine label
                    End If

                    x = leftMargin + ((across - 1) * (labelWidth + gapWidth))
                    PLx = x
                    ypos = y
                    PlySave = y 'for pocket labels, save starting y position
                    For z = 0 To line.Length() - 1
                        test = line(z)
                        'print barcode, if "!" precedes value
                        If Mid$(test, 1, 1) = "!" Then 'if this value should be printed as a barcode
                            test = Mid$(test, 2, test.Length - 1) 'remove the !
                            e.Graphics.DrawString("*" & test & "*", bFont, Brushes.Black, x, ypos)
                            e.Graphics.DrawString("*" & test & "*", bFont, Brushes.Black, x, ypos + CType(inBCFontSize.Text, Single))
                            ypos = ypos + lineSpacing + 20 'allow for extra vertical spacing
                            cntrOffset = 25
                        End If

                        If Mid$(test, 1, 1) = "~" Then 'if this value should be prededed by extra space
                            test = Mid$(test, 2, test.Length - 1) 'remove the ~
                            extraSpace = CType(inLineSpacing.Text, Single) * cf '* 100
                        End If

                        ypos = ypos + extraSpace
                        e.Graphics.DrawString(test, pFont, Brushes.Black, x + cntrOffset, ypos)
                        ypos = ypos + lineSpacing
                        extraSpace = 0
                        cntrOffset = 0
                    Next z

                    If UseLaser.Checked And FlagSlips.Enabled And FlagSlips.Checked And inLabelHeight.Text = "7.0" Then
                        'If chk_VerticalLine.Checked Then
                        e.Graphics.DrawLine(Pens.Black, x + labelWidth - 10, down * 20, x + labelWidth - 10, down * labelHeight)
                    End If

                    If theseArePocketLabels Then
                        Dim pltext As String = ""
                        Dim printbarcode As Boolean = False
                        PLy = PlySave 'start 1st pocket label at same top margin as the spine label
                        leftPocket = PLx + (CType(plLeftMargin.Text, Single) * cf) '* 100)
                        For i = 1 To labelCount
                            For plk = 0 To plOutput.Lines.Length - 1
                                pltext = plOutput.Lines(plk)
                                If Mid$(pltext, 1, 1) = "!" Then
                                    pltext = Mid$(pltext, 2, pltext.Length - 1) 'remove the !
                                    printbarcode = True
                                End If
                                e.Graphics.DrawString(pltext, pFont, Brushes.Black, leftPocket, PLy)
                                PLy = PLy + lineSpacing
                                If printbarcode Then
                                    e.Graphics.DrawString("*" & pltext & "*", bFont, Brushes.Black, leftPocket, PLy)
                                    e.Graphics.DrawString("*" & pltext & "*", bFont, Brushes.Black, leftPocket, PLy - 10 + CType(inBCFontSize.Text, Single))
                                    PLy = PLy + 30
                                    printbarcode = False
                                End If
                            Next plk
                            'if there is a 2nd pocket label, start it at the 1st top
                            ' margin plus the offset specified by the user ("plDistance")
                            PLy = PlySave + (CType(plDistance.Text, Single) * cf) '* 100)
                        Next i
                        plOutput.Text = ""
                    End If

                    If UseDesktop.Checked Then Exit Do

                Next across
                startcol = 1
            Next down
            startRow = 1
            If nxt <> LABELS.Length - 1 Then
                firstPage = False
                e.HasMorePages = True
            Else
                e.HasMorePages = False
                nxt = 0
            End If
            Exit Sub

        Loop While False
        nxt = 0
        e.HasMorePages = False
    End Sub

    Private Sub PrintDocument2_QueryPageSettings(ByVal sender As Object, ByVal e As System.Drawing.Printing.QueryPageSettingsEventArgs) Handles PrintDocument2.QueryPageSettings
        Dim pres As PrinterResolution
        'Dim c As Integer = 0
        'Dim i As Integer = 0
        'Dim lst As String = ""
        'Dim ps As New PaperSize("Custom", 100, 300)
        'e.PageSettings.PaperSize = ps
        If useLandscape.Checked Then
            e.PageSettings.Landscape = True
        Else
            e.PageSettings.Landscape = False
        End If
        pres = PrintDocument2.DefaultPageSettings.PrinterResolution
        'c = e.PageSettings.PrinterSettings.PaperSizes.Count
        'For i = 0 To c - 1
        '    lst = lst & e.PageSettings.PrinterSettings.PaperSizes(i).ToString() & vbCrLf
        'Next
        'MsgBox("Paper Sizes" & vbCrLf & lst)
    End Sub

    Private Function removePocket(ByVal line As Array) As Array
        'gets pocket label from batch file and separates the spine label part from the pocket label part.
        Dim k As Integer = 0
        Dim s As String = ""
        plOutput.Text = ""
        For k = 0 To line.Length - 1
            If line(k).contains(vbTab) Then
                plOutput.Text = plOutput.Text & line(k).substring(line(k).indexOf(vbTab) + 1) & vbCrLf
                line(k) = line(k).substring(0, line(k).indexOf(vbTab))
            Else
                plOutput.Text = plOutput.Text & line(k)
            End If
        Next
        s = plOutput.Text
        k = 0
        Try
            Do While s.Substring(s.Length - 2, 2) = vbCrLf 'trim all crlfs from end of pocket label
                s = s.Substring(0, s.Length - 2)
            Loop
        Catch ex As Exception
            ' Return Split("NOTPOCKET", "|")
        End Try
        plOutput.Text = s
        Return line
    End Function
    Private Function getNextLabel() As String
        Dim labelWork As String
        'nxt is defined globally
        If nxt > LABELS.Length - 1 Then
            nxt = 0
            Return "done"
        End If

        labelWork = LABELS(nxt).replace(vbLf, "")

        If Trim(labelWork) = "" Then
            nxt = 0
            Return "done"
        End If

        nxt = nxt + 1
        Return labelWork

    End Function

    Private Function GetBatch(ByVal bnum As Integer) As String
        Dim sw As StreamWriter
        Dim labelText As String = ""
        Dim wrk As String

        'If the labelBatch text file doesn't exist in executable's path, create it.
        If (Not File.Exists(mypath & "labelbatch" & bnum & ".txt")) Then
            sw = File.CreateText(mypath & "labelbatch" & bnum & ".txt")
            sw.Close()
        End If

        Dim tr As New StreamReader(mypath & "labelbatch" & bnum & ".txt")
        Dim i As Integer

        Try
            labelText = tr.ReadToEnd()
            tr.Close()
        Catch ex As Exception
            MsgBox("Can't read Label batch #" & bnum & vbCrLf &
            "Reason: " & ex.Message, MsgBoxStyle.Exclamation, "Batch File Error")
            labelText = "*ERROR READING BATCH #" & bnum
        End Try
        If labelText.Length <> 0 Then
            i = 1
            If Char.IsControl(labelText, 1) Then i = 3
            wrk = Mid$(labelText, i, labelText.Length) 'remove initial CR/LF and returned LF
            If wrk.Length > 1 Then
                wrk = Mid$(wrk, 1, wrk.Length - 1)
            End If
            If wrk.Contains("|") Then
                wrk = wrk.Replace(vbCrLf, vbCrLf & "===============" & vbCrLf).Replace("|", vbCrLf)
                writeFile(mypath & "labelbatch" & batchNumber.Value & ".txt", wrk, False)
            Else
                wrk = wrk.Replace(vbCrLf & vbCrLf & "=", vbCrLf & "=")
            End If
            Return wrk
        End If
        Return ""
    End Function

    Private Sub callAlma()

        Dim webClient As New System.Net.WebClient()
        Dim almaOK As Boolean = True
        Dim i As Integer = 0
        Dim addToHistory As Boolean
        Dim svcRequest As String = ""
        Dim fixedBarcode As String = ""
        Dim quot As String = """"
        lblXMLWarn.Visible = False
        xmlerr = ""
        If InputBox.Text = "" Then
            Beep()
            Exit Sub
        End If

        fixedBarcode = Replace(Trim(InputBox.Text), "+", "%2B")
        svcRequest = Trim(apiURL.Text) & Trim(apiMethod.Text.Replace("{item_barcode}", fixedBarcode)) & "&apikey=" & Trim(apiKey.Text)

        lastbc = InputBox.Text
        addToHistory = True

        If Trim(InputBox.Text) = "" Then addToHistory = False
        If HistoryList.Items.Contains(InputBox.Text) Then addToHistory = False
        If addToHistory Then
            If HistoryList.Items.Count = 0 Then
                HistoryList.Items.Add(InputBox.Text)
            Else
                If HistoryList.Items.Item(0) <> InputBox.Text Then
                    HistoryList.Items.Insert(0, InputBox.Text)
                End If
            End If

            If HistoryList.Items.Count = 6 Then
                HistoryList.Items.RemoveAt(5)
            End If
        End If


        'if XML is coming from a web call (via java servlet or RESTful call:
        webClient.Encoding = System.Text.Encoding.UTF8

        Try

            xmlReturned = webClient.DownloadString(svcRequest)

            If UseRestfulApi.Checked And dontConvert.Checked = False Then

                xmlReturned = xmlReturned.Replace("><", ">" & vbCrLf & "<")

                Dim convertedXML, e As String
                convertedXML = RESTfulToSOAP(xmlReturned)
                Dim pos, opn, clos As Integer

                pos = InStr(1, convertedXML, " desc=", CompareMethod.Text)
                Do While pos > 0
                    opn = InStrRev(convertedXML, "<", pos, CompareMethod.Text)
                    clos = InStr(opn, convertedXML, ">", CompareMethod.Text)
                    clos = InStr(clos + 1, convertedXML, ">", CompareMethod.Text)
                    e = Mid$(convertedXML, opn, clos - opn + 1)
                    convertedXML = convertedXML.Replace(e, deAttrib(e))
                    pos = InStr(clos + 1, convertedXML, " desc=", CompareMethod.Text)
                Loop

                convertedXML = convertedXML.Replace("call_number_type_code", "call_number_type")
                convertedXML = convertedXML.Replace("<description>", "<issue_level_description>")
                convertedXML = convertedXML.Replace("</description>", "</issue_level_description>")
                convertedXML = convertedXML.Replace("<enumeration_", "<enum_")
                convertedXML = convertedXML.Replace("</enumeration_", "</enum_")
                convertedXML = convertedXML.Replace("<chronology_", "<chron_")
                convertedXML = convertedXML.Replace("</chronology_", "</chron_")


                If convertedXML.Contains("<parsed_call_number>") Then
                    convertedXML = enumerate(convertedXML, "<parsed_call_number>", "</parsed_call_number>", "<call_no>")
                End If

                If convertedXML.Contains("<parsed_issue_level_description>") Then
                    convertedXML = enumerate(convertedXML, "<parsed_issue_level_description>", "</parsed_issue_level_description>", "<issue_level_description>")
                End If

                If Not convertedXML.Contains("<issue_level_description>") Then
                    convertedXML = convertedXML.Replace("</barcode>", "</barcode>" & vbCrLf & "<issue_level_description> </issue_level_description>")
                End If

                If Not convertedXML.Contains("<chron_") Then
                    Dim dummyChron = "</barcode>" & vbCrLf & "<chron_i> </chron_i>" & vbCrLf &
                    "<chron_j> </chron_j>" & vbCrLf & "<chron_k> </chron_k>" & vbCrLf &
                    "<chron_l> </chron_l>" & vbCrLf & "<chron_m> </chron_m>" & vbCrLf
                    convertedXML = convertedXML.Replace("</barcode>", dummyChron)
                End If

                xmlReturned = xmlShell.Text.Replace("**XMLBODY**", convertedXML)
            Else
                If dontConvert.Checked = True Then
                    TabControl1.SelectedIndex = 3
                End If
            End If
        Catch ex As Exception
            almaOK = False
            OutputBox.Text = "***** ERROR ***** " & vbCrLf & "Can't retrieve XML file." & vbCrLf & vbCrLf &
            "Error message: " & ex.Message
            'MsgBox("error: " & ex.Message)
        End Try
        If almaOK Then
            getBarcodeFile()
        End If
    End Sub
    Private Function enumerate(ByVal xml As String, ByVal xfrom As String, ByVal xto As String, ByVal breakon As String)
        Dim k As Integer = 0
        Dim parseString As String = ""
        Dim cline As Array
        Dim enumerated As String = ""
        parseString = BTween(xml, xfrom, xto)
        cline = Split(parseString, breakon)
        For k = 1 To cline.Length - 1
            enumerated = enumerated & (breakon & cline(k).ToString).Replace(">", "_" & k & ">")
        Next
        xml = xml.Replace(parseString, enumerated)
        Return xml
    End Function
    Private Function RESTfulToSOAP(ByVal restfulXml As String)
        'convert RESTful xml file returned by Alma into SOAP format
        Dim e As Integer
        Dim f As Integer
        Dim i As Integer
        Dim t As String = ""
        Dim rdoc As New XmlDocument
        Dim nodes As XmlNode
        Dim cnodes As XmlNode
        Dim element As String = ""
        Dim parts(2) As String
        parts(0) = "//bib_data"
        parts(1) = "//holding_data"
        parts(2) = "//item_data"

        rdoc.PreserveWhitespace = True
        rdoc.LoadXml(restfulXml)

        For i = 0 To parts.Length - 1

            nodes = rdoc.SelectSingleNode(parts(i))

            If nodes.HasChildNodes Then

                For e = 0 To nodes.ChildNodes.Count - 1

                    If nodes.ChildNodes(e).HasChildNodes Then
                        cnodes = nodes.ChildNodes(e)
                        If cnodes.ChildNodes.Count > 1 Then
                            t = t & "<" & cnodes.Name & ">" & vbCrLf
                            For f = 0 To cnodes.ChildNodes.Count - 1
                                t = t & cnodes.ChildNodes(f).OuterXml & vbCrLf
                            Next
                            t = t & "</" & cnodes.Name & ">" & vbCrLf
                        Else
                            element = nodes.ChildNodes(e).OuterXml
                            t = t & element & vbCrLf
                            'If element.Contains(" desc=") Then t = t & deAttrib(element) & vbCrLf
                        End If
                    End If
                Next
            End If
        Next i
        'writeFile(mypath & "converted.txt", t, False)
        Return t

    End Function

    Private Function deAttrib(ByVal e As String) As String
        Dim element As String
        Dim attrVal As String
        Dim codeval As String
        Dim quot As String = """"
        Dim XML1 As String
        Dim XML2 As String

        element = BTween(e, "<", " ")
        attrVal = BTween(e, quot, quot)
        codeval = BTween(e, ">", "<")

        XML1 = "<" & element & "_name>" & attrVal & "</" & element & "_name>"
        XML2 = "<" & element & "_code>" & codeval & "</" & element & "_code>"


        Return XML1 & vbCrLf & XML2
    End Function
    Private Function donothing(ByVal x As String)
        'MsgBox(RESTfulToSOAP(xmlReturned))

        'Dim doc As New XmlDocument
        'Dim t As String = ""
        'Dim cn As String = "" 'call number
        'Dim cntp As String = "" 'call number type
        'Dim n As String = ""
        'Dim titl As String = ""
        'Dim e As Integer
        'Dim Lnode As XmlNodeList
        'Dim anode As XmlNode

        'doc.LoadXml(xmlReturned)
        'Dim elemList As XmlNodeList = doc.GetElementsByTagName("item_data")
        'Dim c As XmlNodeList = doc.GetElementsByTagName("call_number")
        'cn = c(0).OuterXml

        'Lnode = doc.GetElementsByTagName("call_number_type")
        'cntp = "<call_number_type>" & Lnode(0).InnerXml & "</call_number_type>"

        'Dim ti As XmlNodeList = doc.GetElementsByTagName("title")
        'titl = ti(0).OuterXml

        'Dim lid As XmlNodeList = doc.GetElementsByTagName("library")
        'Dim licode As String = "<library_code>" & lid(0).InnerXml & "</library_code>"
        'anode = doc.SelectSingleNode("//library")
        'Dim liname As String = "<library_name>" & anode.Attributes(0).Value & "</library_name>"

        'Dim lod As XmlNodeList = doc.GetElementsByTagName("location")
        'Dim locode As String = "<location_code>" & lod(0).InnerXml & "</location_code>"
        'anode = doc.SelectSingleNode("//location")
        'Dim loname As String = "<location_name>" & anode.Attributes(0).Value & "</location_name>"

        'For e = 0 To elemList.Count - 1
        '    t = elemList(e).InnerXml & vbCrLf
        'Next e

        'xmlReturned = xmlShell.Text.Replace("**XMLBODY**", cn & vbCrLf & cntp & vbCrLf & titl & vbCrLf & licode & vbCrLf & liname & vbCrLf & loname & locode & vbCrLf & t)
        'xmlReturned = xmlReturned.Replace("<description>", "<issue_level_description>")
        'xmlReturned = xmlReturned.Replace("</description>", "</issue_level_description>")
        'xmlReturned = xmlReturned.Replace("enumeration_", "enum_")
        Return x
    End Function




    Private Sub getBarcodeFile()

        Dim xmltext As String
        Dim bcstart As Integer = 0
        Dim bcend As Integer = 0
        Dim current_date As String = Date.Now.Date
        Dim xmlerror As String = ""
        Dim er As String = ""
        Dim cdxml As String = "<current_date>" & current_date & "</current_date>"
        Dim xmlend As String = "</physical_item_display_for_printing>"
        Dim x As Array
        Dim linenum As Integer = 0
        Dim charpos As Integer = 0
        Dim coords As String
        Dim pos As Array
        Dim s As String = ""
        Dim i As Integer
        Dim loopcount As Integer = 0
        xmlReturned = xmlReturned.Replace(xmlend, cdxml & xmlend)

        loopcount = 0
        Do
            Try
                ' ANGLE BRACKET ERROR OCCURS HERE

                xdoc.LoadXml(xmlReturned.Replace("&", "|amp"))
                Exit Do
                '*********************************
            Catch ex As Exception 'if xdoc.LoadXml fails, it may be due to < character in the data.
                x = Split(xmlReturned, vbLf) 'need an array of lines to find line# of bad "<" character.
                s = ex.ToString
                If s.Contains("Name cannot begin with") Then 'If error is due to extraneous < character 
                    coords = Mid$(s, s.IndexOf("Line "), 24) 'finds location (line#, char pos. w/in line)
                    coords = coords.Replace("Line", "").Replace("position", "").Replace(".", "").Replace(" ", "")
                    pos = Split(coords, ",") 'comma separates line# from char position
                    linenum = pos(0) - 1
                    charpos = pos(1)
                    Mid$(x(linenum), charpos - 1, 1) = "`" 'replace offending < with accent mark
                    x(linenum) = x(linenum).replace("`", "|lt;") 'change accent to code representing < char. Later,
                    xmlReturned = ""                             'during XML display, |lt; will be replaced with <
                    For i = 0 To x.Length - 1
                        xmlReturned = xmlReturned & x(i) & vbLf
                    Next

                    loopcount = loopcount + 1
                    If loopcount > 10 Then
                        MsgBox("Too many tries to fix the XML record")
                        Exit Do
                    End If
                Else 'if this is not a "<" error, then give up.
                    MsgBox("SpineOMatic cannot process the XML due to invalid characters in the data." & vbCrLf &
                    "The error returned is: " & vbCrLf & ex.ToString)
                    Exit Sub
                End If
            End Try
        Loop While True


        xmltext = StripControlChars(xmlReturned, False)
        xmltext = xmltext.Replace("></", "> </")
        xmltext = xmltext.Replace("><", ">" & vbCrLf & "<")
        RichTextBox1.Clear()
        RichTextBox1.ForeColor = Color.Red
        RichTextBox1.Text = xmltext
        writeStat("S")
        If xmltext.IndexOf("</error>") - xmltext.IndexOf("<error>") > 8 Then
            OutputBox.Text = "***** ERROR ******" & vbCrLf & "can't find this barcode number in Alma"
            almaReturnCode = "Bad barcode"
            InputBox.SelectionStart = 0
            InputBox.SelectionLength = InputBox.Text.Length
            If File.Exists(lastxml) = True Then File.Delete(lastxml)
            cntype = "-"
            buildStatRec()
            Exit Sub
        End If

        If xmltext.IndexOf("</som_error>") - xmltext.IndexOf("<som_error>") > 8 Then
            er = xmltext.ToLower
            OutputBox.Text = "***** ERROR *****" & vbCrLf & "Can't connect to Alma"
            If er.IndexOf("unauthorized") > 0 Then
                OutputBox.Text = OutputBox.Text & vbCrLf & "Java: credentials unauthorized."
                almaReturnCode = "Bad credentials"
            Else
                If er.IndexOf("unknownhost") > 0 Then
                    OutputBox.Text = OutputBox.Text & vbCrLf & "Java: **Check Alma URL."
                    almaReturnCode = "Bad URL"
                Else
                    If er.IndexOf("unsupported endpoint") Or er.IndexOf("illegalargument") > 0 Then
                        OutputBox.Text = OutputBox.Text & vbCrLf & "Java: **Check Alma URL--It must start with http://"
                        almaReturnCode = "Bad URL"
                    End If
                End If
            End If
            cntype = "-"
            buildStatRec()
            InputBox.SelectionStart = 0
            InputBox.SelectionLength = InputBox.Text.Length
            If File.Exists(lastxml) = True Then File.Delete(lastxml)
            Exit Sub
        End If

        almaReturnCode = "OK"

        lcxml = xmlValue(inCallNumSource.Text).Replace(vbTab, "")

        If Not inIssueLevelSource.Text.Contains("parsed") Then
            issuexml = xmlValue(inIssueLevelSource.Text)
            If Trim(issuexml) <> "" Then
                issuexml = Custom2(inIssueLevelSource.Text)
            End If
        End If

        libxml = xmlValue(TextBox13.Text)
        locxml = xmlValue(TextBox14.Text)
        almaLibrary = libxml
        almaLocation = locxml
        cntype = xmlValue("<call_number_type>")
        If Trim(cntype) = "" Then cntype = convertBlankTo.Text
        testComboBox.Text = lcxml
        testCallNumType.Text = cntype

        buildStatRec()
        RichTextBox1.Enabled = False
        formatXML()
        RichTextBox1.Enabled = True
        If File.Exists(lastxml) = True Then File.Delete(lastxml)
        printCallNum()
    End Sub
    Private Sub getNodeList()
        batchPreview.Text = ""
        Dim nodelist As XmlNodeList
        Dim node As XmlNode
        nodelist = xdoc.SelectNodes("//*")
        MsgBox(nodelist.Count)
        For Each node In nodelist
            batchPreview.Text = batchPreview.Text & node.Name & vbCrLf
        Next

    End Sub
    Private Sub formatXML()
        'Makes the Current XML display look more like XML in the RichTextBox window.
        Dim pos As Integer = 0
        Dim tagstart As Integer = 0
        Dim tagend As Integer = 0
        Dim tagname As String = ""
        Dim nextclose As Integer = 0
        Dim nextend As Integer = 0
        Dim endname As String = ""
        Dim sp As Integer = 0
        Dim ep As Integer = 0
        Dim currentSelectionFont As Font

        Try
            With RichTextBox1
                While True
                    tagstart = .Find("<", tagstart, RichTextBoxFinds.MatchCase)
                    .SelectionStart = tagstart
                    .SelectionLength = 1
                    .SelectionColor = Color.Blue
                    If tagstart = -1 Then Exit While
                    tagend = .Find(">", tagstart, RichTextBoxFinds.MatchCase)
                    .SelectionStart = tagend
                    .SelectionLength = 1
                    .SelectionColor = Color.Blue
                    tagname = Mid$(.Text, tagstart + 1, tagend - tagstart + 1)
                    nextclose = .Find("</", tagend, RichTextBoxFinds.MatchCase)
                    If nextclose = -1 Then Exit While
                    .SelectionStart = nextclose + 1 'make the "/" blue
                    .SelectionLength = 1
                    .SelectionColor = Color.Blue
                    nextend = .Find(">", nextclose, RichTextBoxFinds.MatchCase)
                    endname = Mid$(.Text, nextclose + 1, nextend - nextclose + 1)
                    If tagname = endname.Replace("/", "") Then
                        'For RichTextBox.SelectionFont Property:
                        'If the current text selection has more than one font specified, the property SelectionFont is null,
                        'so we use a variable to hold the current one SelectionFont.
                        currentSelectionFont = RichTextBox1.SelectionFont
                        sp = tagend + 1
                        ep = nextclose - 1
                        .SelectionStart = sp
                        .SelectionLength = ep - sp + 1
                        .SelectionColor = Color.Black
                        .SelectionFont = New Font(currentSelectionFont, FontStyle.Bold)
                    End If
                    tagstart = tagstart + 1
                End While
            End With
        Catch ex As Exception
            If Not ex.ToString.Contains("Font prototype, FontStyle newStyle") Then
                MsgBox("error: " & ex.ToString)
            End If
        End Try
    End Sub
    Private Function StripControlChars(ByVal source As String, Optional ByVal KeepCRLF As _
Boolean = True) As String
        ' we use this to build the result
        Dim sb As New System.Text.StringBuilder(source.Length)
        Dim index As Integer
        For index = 0 To source.Length - 1
            If Not Char.IsControl(source, index) Then
                ' not a control char, so we can add to result
                sb.Append(source.Chars(index))
            ElseIf KeepCRLF AndAlso source.Substring(index,
            2) = ControlChars.CrLf Then
                ' it is a CRLF, and the user asked to keep it
                sb.Append(source.Chars(index))
            End If
        Next
        Return sb.ToString()
    End Function

    Private Sub printCallNum()
        Dim preLC As String = ""
        Dim mainLC As String = ""
        Dim issueLC As String = ""
        Dim otherLC As String = ""
        Dim taday As String = DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture)
        Dim checklines As Array = Nothing
        Dim maxChars As Integer = CType(inMaxChars.Text, Integer)
        Dim maxlines As Integer = CType(inMaxLines.Text, Integer)
        Dim charsOK As Boolean = True
        Dim linesOK As Boolean = True
        Dim i As Integer = 0
        Dim fontname As String = inFontName.Text
        Dim fontsize As Single = CType(inFontSize.Text, Single)
        Dim quot = """"
        Dim checkcr As String = ""
        Dim cr As String = ""
        Dim ttl As String = ""
        '**********************************************
        preLC = aboveCall(libxml, locxml) '& vbCrLf
        If preLC <> "" Then preLC = preLC & vbCrLf
        '**********************************************        
        preLC = preLC.Replace(";", vbCrLf)
        OutputBox.Font = New Font(fontname, fontsize)
        ttl = xmlValue("<title>")
        If ttl.Length > 70 Then
            ttl = ttl.Substring(0, 70) & "..."
        End If
        TextBox24.Text = ttl 'xmlValue("<title>")

        Label27.Text = xmlValue("<library_name>")
        Dim custom_field As String = ""
        Dim otherval As String = ""
        Dim parserName As String = ""

        If CustomLabel.Checked Then
            'create custom label
            parsedBy.BackColor = Color.White
            If FlagSlips.Checked Then
                parsedBy.ForeColor = Color.Green
                parsedBy.Text = "Label type: Custom/Flag Slips"
            Else
                parsedBy.ForeColor = Color.Green
                parsedBy.Text = "Label type: Custom"
            End If
            OutputBox.Text = ""
            OutputBox.Text = Custom2(CustomText.Text)
            If Mid$(OutputBox.Text, 1, 2) = vbCrLf Then
                OutputBox.Text = OutputBox.Text.Substring(2)
            End If
        Else
            If useExlibrisParsing.Checked = False Then
                parsedBy.ForeColor = Color.Blue
                parsedBy.BackColor = Color.White
                parserName = TabControl2.SelectedTab.Text.Replace("/Child.Lit/NLM", "...")
                parsedBy.Text = "Call# Parser: SpineOMatic " & parserName
                mainLC = parseLC(lcxml)
            Else
                parsedBy.ForeColor = Color.DarkRed
                parsedBy.BackColor = Color.White
                parsedBy.Text = "Call# Parser: Ex Libris"
                mainLC = getParsing2(parsingSource.Text)
            End If

            If chkIncludeHoldings.Checked Then 'if holdings are requested...
                If inIssueLevelSource.Text = "<parsed_issue_level_description>" Then
                    issueLC = vbCrLf & Trim(getParsing2("<parsed_issue_level_description>"))
                    holdingsBy.Text = "Holdings: Ex Libris parsed." : holdingsBy.ForeColor = Color.DarkRed
                    If issueLC = vbCrLf Then 'if no parsed holdings were found...
                        issuexml = xmlValue("<issue_level_description>")
                        issueLC = vbCrLf & Trim(parseIssue(issuexml)) 'use textual holdings
                        issueLC = issueLC.Replace("^", ":" & vbCrLf)
                        holdingsBy.Text = "Holdings: SpineOMatic, no Ex Libris" : holdingsBy.ForeColor = Color.DarkCyan
                        If issueLC = vbCrLf Then
                            holdingsBy.Text = "Holdings: None found." : holdingsBy.ForeColor = Color.Gray
                            issueLC = ""
                        End If
                    End If
                Else 'if requesting textual holdings
                    issueLC = vbCrLf & Trim(parseIssue(issuexml)) 'get textual holdings
                    issueLC = issueLC.Replace("^", ":" & vbCrLf)
                    holdingsBy.Text = "Holdings: SpineOMatic" : holdingsBy.ForeColor = Color.Blue
                    If issueLC = vbCrLf Then
                        holdingsBy.Text = "Holdings: none found." : holdingsBy.ForeColor = Color.Gray
                        issueLC = ""
                    End If
                End If
            Else
                holdingsBy.Text = "Holdings: Not requested." : holdingsBy.ForeColor = Color.Gray
                issueLC = ""
            End If

            If chkIncludeOther.Checked Then
                otherLC = Custom2(inOtherSource.Text)
            Else
                otherLC = ""
            End If

            checkcr = preLC & mainLC & issueLC
            OutputBox.Text = checkcr.Replace(vbCrLf & vbCrLf, vbCrLf) & otherLC
            origText = OutputBox.Text 'save text to later detect if it was manually changed
            If dontConvert.Checked Then OutputBox.Text = "** VIEW RESTful XML ONLY **" &
            vbCrLf & vbCrLf & "Uncheck 'Don't Convert XML' in the Alma Access panel " &
            "to restore normal operation."
        End If

        charsOK = True

        If dontConvert.Checked = False Then
            checklines = Split(OutputBox.Text, vbCrLf)
            Dim siz As Size = TextRenderer.MeasureText(OutputBox.Text, OutputBox.Font)
            If siz.Height > OutputBox.Height Then
                OutputBox.ScrollBars = ScrollBars.Vertical
            Else
                OutputBox.ScrollBars = ScrollBars.None
            End If
            If checklines.Length > maxlines Then
                Beep()
                MsgBox("Label contains more than " & maxlines & " lines.", MsgBoxStyle.Exclamation, "Too Many Lines")
                linesOK = False
            End If
            For i = 0 To checklines.Length - 1
                If checklines(i).length > maxChars Then
                    MsgBox("Line #" & i + 1 & ": " & checklines(i) &
                    vbCrLf & " contains more than " & maxChars & " characters.")
                    charsOK = False
                End If
            Next
        End If
        plOutput.Text = ""
        plOutput.Font = New Font(inFontName.Text, CType(inFontSize.Text, Single), FontStyle.Regular)

        If chkUsePocketLabels.Checked Then
            Try
                CreatePocketLabel()
            Catch ex As Exception
                MsgBox("User-defined pocket label error." & vbCrLf &
                "The maximum number of lines for XML field " & quot & ex.Message & quot &
" must not be less than the minimum, and/or must not be zero.", MsgBoxStyle.Exclamation, "Custom Label Line Space Error")
                Exit Sub
            End Try
        End If

        If AutoPrintBox.Checked And charsOK And linesOK Then
            editText = ""
            ManualPrint.PerformClick() 'send text to printer
        End If

    End Sub

    Private Function checkLineLength() As Integer
        Dim checklines As Array = Nothing
        Dim maxChars As Integer = CType(inMaxChars.Text, Integer)
        Dim i As Integer = 0
        Dim r As Integer = 99
        checklines = Split(OutputBox.Text, vbCrLf)
        For i = 0 To checklines.Length - 1
            If checklines(i).length > maxChars Then
                r = i + 1 'if a line exceeds the specified line length, the 1-based line number is returned
                Exit For
            End If
        Next
        Return r 'if r returns with 99, all lines are OK.
    End Function
    Private Function Custom2(ByVal fields As String) As String

        If Trim(fields) = "" Then
            Return ""
            Exit Function
        End If

        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim cf As String = ""
        Dim custom As Array
        Dim fmt As String = ""
        Dim fmtreturn As Array
        Dim freetext As String = ""
        Dim val = ""
        Dim lbl = ""
        Dim quot As String = """"
        Dim txt As String = ""
        Dim fldval As String = ""
        Dim s As String = vbCrLf
        Dim bg, nd As Integer
        Dim qlen As Integer = 0
        Dim charWrap As Integer = 0
        Dim s1, s2, s3 As String
        Dim stringLength As Integer = 0

        fields = fields.Replace("><", "/")

        custom = Split(fields, vbCrLf)
        'When Spine labels are selected, only the "Include other field" is sent through this routine.
        'If that field is wrapped, it must be wrapped to the Chars/Line value, not to a length.
        If Spine.Checked Then
            charWrap = CType(inMaxChars.Text, Integer)
        End If
        'getFmt returns an array in fmtreturn.  Fmtreturn(0) is the format string, Fmtreturn(1) is any quoted text

        Try
            For i = 0 To custom.Length - 1
                cf = custom(i)
                wrapping = False
                bg = cf.IndexOf("<") : nd = cf.IndexOf(">")
                If bg <> -1 And nd <> -1 And bg < nd Then 'if this is a <xml_value> in angle brackets,
                    fmtreturn = getFmt(Trim(cf.Substring(0, bg)))
                    fmt = fmtreturn(0)
                    cf = cf.Substring(bg, nd - bg + 1)
                    If fmt.Contains("*") Then
                        fldval = xmlValue(cf, True)
                    Else
                        fldval = xmlValue(cf)
                    End If
                    stringLength = 0
                    Dim stringLengthString As String : stringLengthString = vbNullString
                    For j = 0 To fmt.Length - 1
                        If "0123456789".Contains(fmt(j)) Then
                            stringLengthString += fmt.Substring(j, 1)
                        End If
                    Next
                    If stringLengthString IsNot vbNullString Then
                        stringLength = Convert.ToInt32(stringLengthString)
                    End If
                    If (fldval.Length > 0 And fldval.Length > stringLength) And stringLength > 0 Then
                        fldval = fldval.Substring(0, stringLength)
                    End If
                    txt = fmtreturn(1) & fldval
                Else 'if this is a free-text string
                    fmtreturn = getFmt(cf)
                    If fmtreturn(1) <> "" Then qlen = 2 'if quoted text exists, account for quote marks
                    txt = fmtreturn(1) & cf.Substring(fmtreturn(0).Length + fmtreturn(1).length + qlen)
                    fmt = fmtreturn(0)
                End If

                If Trim(txt) = "" Then Continue For
                '
                ' txt = Trim(getPrefix(fmt) & Trim(txt)) 'add any prefix to the value.
                If fmt <> "" Then

                    If fmt.Contains("*") Then
                        s1 = "" : s2 = "" : s3 = ""
                        If suppress1.Text = "" Then s1 = "@@" Else s1 = suppress1.Text
                        If suppress2.Text = "" Then s2 = "@@" Else s2 = suppress2.Text
                        If suppress3.Text = "" Then s3 = "@@" Else s3 = suppress3.Text
                        'If Trim(fldval.Replace("0", "")) = "" Then
                        If Trim(fldval.Replace(s1, "").Replace(s2, "").Replace(s3, "")) = "" Then
                            'Exit For
                            txt = ""
                        End If
                    End If

                    If fmt.Contains("~") Then lbl = lbl & vbCrLf
                    If fmt.Contains("!") Then lbl = lbl & "!"
                    If fmt.Contains("+") Then
                        lbl = lbl & aboveCall(libxml, locxml).Replace(";", vbCrLf)
                        txt = ""
                    End If

                    If fmt.Contains("=") Then wrapping = True

                    If fmt.Contains("%") Then 'if call number parse prefix,
                        val = wrap(parseLC(txt), charWrap) 'val is parsed call number.
                    Else
                        If fmt.Contains("#") Then '0therwise, if holdings-parse prefix,
                            val = wrap(parseIssue(txt).Replace("^", ":" & vbCrLf), charWrap) 'val is parsed holdings
                        Else
                            val = wrap(txt, charWrap) 'if neither, val is just the XML value
                        End If
                    End If

                    If fmt.Contains("^") Then
                        lbl = lbl & val & " " 'add the value to the outputbox, no line break.
                    Else
                        'If lbl <> "" Then lbl = lbl & val & vbCrLf 'add the value to the outputbox w/ linebreak.
                        If val <> "" Then lbl = lbl & val & vbCrLf 'add the value to the outputbox w/ linebreak.
                    End If

                Else 'if no formatting...
                    lbl = lbl & wrap(txt, charWrap) & vbCrLf
                End If
            Next

        Catch ex As Exception
            MsgBox("Error in Custom2(): " & ex.ToString)
        End Try
        If lbl.length > 2 Then
            lbl = vbCrLf & lbl.substring(0, lbl.length - 2) 'removes crlf from end of string
        End If
        Return lbl
    End Function
    Private Function getFmt(ByVal arg As String) As Array
        'loop thru free text arg, storing initial format chars (~ # % ! ^) in 'fmt'
        'Removes quoted "prefix" text that may contain these characters.
        Dim fmt As String = ""
        Dim quot As String = """"
        Dim a As Array
        Dim argnoquot As String = ""
        Dim c As String = ""
        Dim i As Integer = 0
        Dim qm As String = ""
        Dim qcount As Integer = 0
        If Trim(arg) = "" Then 'if arg is null,
            Return Split("" & "|" & "", "|") 'return two empty fields.
        End If

        If arg.Contains(quot) Then
            a = arg.Split(quot)
            qcount = a.Length - 1
            'Label27.Text = qcount
            If qcount Mod 2 <> 0 Then
                MsgBox("Unbalanced quotation marks in format string: " & arg, MsgBoxStyle.Exclamation, "Unbalanced quotes in format string")
                Return Split("" & "|" & "", "|") 'return two empty fields.
            End If
            argnoquot = a(0) & a(2)
            qm = a(1)
        Else
            argnoquot = arg
        End If

        For i = 1 To argnoquot.Length
            c = Mid$(argnoquot, i, 1)
            If "~=!#%^*+".Contains(c) Or "0123456789".Contains(c) Then
                fmt = fmt & c
            Else
                Exit For
            End If
        Next
        Return Split(fmt & "|" & qm, "|")
    End Function
    Private Function getPrefix(ByVal fmt As String) As String
        Dim quot As String = """"
        Dim a As Array
        If Not fmt.Contains(quot) Then Return ""
        a = fmt.Split(quot) 'splitting on quots puts the contents in element #1 (0 and 2 have the quots)
        Return a(1) 'return the text string between the quots.
    End Function
    Private Function CreateFlagSlip() As String
        Dim i As Integer = 0
        Dim custom_field As String = ""
        Dim custom As Array
        Dim lname As String = ""
        Dim cn As String = ""
        Dim itm As String = ""
        Dim bc As String = ""
        Dim titl As String = ""
        Dim mywidth As Single = 0.0
        Dim myfont As New Font(fontname, 10)
        custom = Split(CustomText.Text, vbCrLf)

        OutputBox.Font = myfont
        For i = 0 To custom.Length - 1
            custom_field = custom(i)

            If custom_field = "<call_number>" Then
                cn = parseLC(xmlValue(custom_field))
            End If
            If custom_field = "<issue_level_description>" Then
                itm = parseIssue(xmlValue(custom_field))
                itm = itm.Replace("^", ":") 'non-breaking colon returns as ^, and is replaced by :
                itm = wrap(itm)
            End If
            If custom_field = "<barcode>" Then
                bc = "!" & xmlValue(custom_field)
            End If
            If custom_field.Contains("<title>") Then
                titlexml = xmlValue(custom_field)
                titl = "~" & wrap(titlexml)
            End If
            If custom_field = "<location_name>" Then
                lname = "~" & xmlValue(custom_field)
                lname = wrap(lname)
            End If
        Next
        Return bc & vbCrLf & lname & vbCrLf & cn & vbCrLf & itm & vbCrLf & titl
    End Function
    Private Sub CreatePocketLabel()
        Dim plValues As Array
        Dim plParam As Array
        Dim plMin As Integer = 0
        Dim plMax As Integer = 0
        Dim i As Integer = 0
        Dim k As Integer = 0
        Dim temp As String = ""
        Dim formatChar As String = ""
        plParam = PocketParams()
        wrapping = True
        indenting = True


        For i = 0 To plParam.Length - 1
            plValues = Split(plParam(i), ";")
            If Trim(plValues(0)) = "" Then
                Continue For
            End If
            plMin = CType(plValues(1), Integer) : plMax = CType(plValues(2), Integer)
            If plMin > plMax Or plMax = 0 Then
                Throw New Exception(plValues(0)) 'send back the XML field name that has bad min/max lines.
                Exit Sub
            End If
            plWork.Text = ""

            If Mid$(plValues(0), 1, 1) = "!" Then
                formatChar = "!"
                plValues(0) = Mid$(plValues(0), 2, plValues(0).Length - 1) 'remove the !
            Else
                formatChar = ""
            End If

            temp = formatChar & xmlValue(plValues(0))
            If plValues(0).contains("title>") Then
                If temp.Substring(temp.Length - 2) = " /" Then
                    temp = temp.Substring(0, temp.Length - 2)
                End If
            End If
            ' plWork.Text = plWork.Text & wrap(xmlValue(plValues(0)))
            plWork.Text = plWork.Text & wrap(temp, 28)
            plWork.Text = plWork.Text & StrDup(5, vbCrLf)
            If plWork.Text = "" And plMin = 0 Then Continue For
            k = 0
            Do While k < plMin 'print minimum number of lines
                plOutput.Text = plOutput.Text & plWork.Lines(k).ToString & vbCrLf
                k = k + 1
            Loop

            Do While k < plMax And plWork.Lines(k) <> "" 'stop at maximum, or when no more lines.
                plOutput.Text = plOutput.Text & plWork.Lines(k).ToString & vbCrLf
                k = k + 1
            Loop

        Next
        plOutput.Text = plOutput.Text.Substring(0, plOutput.Text.Length - 2) 'remove final crlf
        wrapping = False
        indenting = False
    End Sub
    Private Function PocketParams() As Array
        Dim plParam As Array
        If btnSLB.Checked Then
            plParam = Split("<call_number>;1;1|<author>;1;2|<title>;1;2", "|")
            labelCount = 1
        Else
            If btnSL4.Checked Then
                plParam = Split("<call_number>;1;1|<author>;1;2|<title>;1;2", "|")
                labelCount = 2
            Else
                If btnSL6.Checked Then
                    plParam = Split("<call_number>;2;2|<author>;2;3|<title>;1;3", "|")
                    labelCount = 2
                Else
                    plParam = Split(plSrc1.Text & ";" & plMin1.Text & ";" & plMax1.Text & "|" &
                    plSrc2.Text & ";" & plMin2.Text & ";" & plMax2.Text & "|" &
                    plSrc3.Text & ";" & plMin3.Text & ";" & plMax3.Text & "|" &
                    plSrc4.Text & ";" & plMin4.Text & ";" & plMax4.Text, "|")
                    labelCount = PLcount.Value
                End If
            End If
        End If
        Return plParam
    End Function
    Private Function HexString(ByVal EvalString As String) As String
        Dim intStrLen As Integer
        Dim intLoop As Integer
        Dim strHex As String = ""

        EvalString = Trim(EvalString)
        intStrLen = Len(EvalString)
        For intLoop = 1 To intStrLen
            strHex = strHex & " " & Hex(Asc(Mid(EvalString, intLoop, 1)))
        Next
        HexString = strHex
    End Function
    Private Function wrap(ByVal textline As String, Optional ByVal useCharCount As Integer = 0) As String
        'wrap a string to the label width [if no useCharCount is passed to this routine]
        'wrap a string to a specified number of characters if useCharCount is provided]
        '***********
        If Not wrapping Then Return textline
        '***********
        If Trim(textline) = "" Then Return ""
        If textline.Length > 200 Then
            textline = Mid$(textline, 1, 200) & "..."
        End If
        Dim maxCharsPerLine As Integer = 0

        If useCharCount = 0 Then
            Dim chars As Integer = textline.Length
            Dim textwidth = Measure(textline, "t")
            Dim fitToWidth As Single = CType(wrapWidth.Text, Single) * 0.9 'use a 10% buffer
            If textwidth <= fitToWidth Then Return textline
            Dim cpi As Single = chars / textwidth
            maxCharsPerLine = CType(cpi * fitToWidth, Integer)

        Else
            maxCharsPerLine = useCharCount
            If textline.Length <= useCharCount Then Return textline
        End If

        Dim result As String = ""
        Dim base As Integer = 1
        Dim inc As Integer = 0
        Dim s As String = ""
        Dim lastSpace As Integer = 0
        Dim i As Integer = 0
        Dim indentChar As String = ""
        If indenting Then
            indentChar = " "
        Else
            indentChar = ""
        End If

        Try
            Do
                lastSpace = Mid$(textline, base, maxCharsPerLine).LastIndexOf(" ")
                If lastSpace = -1 Then
                    textline = textline.Insert(base + maxCharsPerLine - 1, "-|")
                    base = base + maxCharsPerLine + 2
                Else
                    Mid$(textline, base + lastSpace, 1) = "|"
                    base = base + lastSpace
                End If
                If textline.Length - base <= maxCharsPerLine Then Exit Do
            Loop While True
        Catch ex As Exception
            MsgBox("The font size (" & inFontSize.Text & ") is too large to properly fit in the label width (" & inLabelWidth.Text & ")")
            Return "[font too large]"
        End Try

        Return textline.Replace("|", vbCrLf & indentChar)
    End Function

    Private Function Measure(ByVal BannerText As String, ByVal fontType As String) As Single
        'fontType is "t" for measuring a text font, or bc for measuring a barcode font.
        Dim mywidth As Single
        Dim b As Bitmap
        Dim g As Graphics
        Dim stringSize As SizeF
        'Dim units As GraphicsUnit = GraphicsUnit.Inch
        Dim f As New Font(inFontName.Text, CType(inFontSize.Text, Single))
        Dim bc As New Font(inBCFontName.Text, CType(inBCFontSize.Text, Single))
        ' Compute the string dimensions in the given font
        b = New Bitmap(1, 1, PixelFormat.Format32bppArgb)
        g = Graphics.FromImage(b)
        g.PageUnit = GraphicsUnit.Inch

        If fontType = "t" Then
            stringSize = g.MeasureString(BannerText, f)
        Else
            stringSize = g.MeasureString(BannerText, bc)
        End If

        mywidth = stringSize.Width
        'Height = stringSize.Height
        g.Dispose()
        b.Dispose()
        Return mywidth
    End Function
    Function cleanup(ByVal callout As String) As String
        callout = callout.Replace(vbCrLf & " ", vbCrLf)
        callout = callout.Replace(vbCrLf & vbCrLf, vbCrLf)
        If callout.StartsWith(vbCrLf) Then
            callout = callout.Substring(1)
        End If
        '*** unfinished routine to limit labels to user-specified number of lines, and dump
        '*** remaining portion of call number on last line.
        'If usingDewey Then
        '    If deweyGroup3.Enabled And deweyGroup3.Checked Then
        '        Dim nlcount = Len(callout) - Len(callout.Replace("vbcrlf", ""))
        '        Dim maxlines As Integer = CType(deweyMaxLines.Text, Integer)
        '        If nlcount > maxlines Then
        '            Do While nlcount - maxlines <> 0

        '            Loop
        '        End If
        '    End If
        'End If
        Return callout
    End Function

    Private Function parseLC(ByVal cn As String) As String
        Dim callout As String = ""
        Dim cutter As String = ""
        Dim diagstring As String = ""
        Dim match As String = ""
        Dim pos As Integer = 1
        Dim i As Integer = 0
        If cn = "*" Then Return "" : Exit Function

        If Not cn.Contains("*") Then 'TEST call nums will start with *
            cntype = xmlValue("<call_number_type>") 'if not TEST, get call number type from XML
        Else
            cn = cn.Replace("*", "") 'if TEST, remove *
            'The TEST process will have already set the cntype, so don't get it from the XML record.
        End If

        If Trim(cntype) = "" Then cntype = convertBlankTo.Text
        '
        If DeweyType.Text.Contains(cntype) Then
            callout = parseDewey(cn)
            callout = cleanup(callout)
            Return callout
        Else
            If sudocType.Text.Contains(cntype) Then
                callout = parseSuDoc2(cn)
                callout = cleanup(callout)
                Return callout
            Else
                If otherType.Text.Contains(cntype) Then
                    callout = parseOther(cn)
                    callout = cleanup(callout)
                    Return callout
                Else
                    If Not lcType.Text.Contains(cntype) Then
                        MsgBox("Call number type " & cntype & " is not handeld by any parsing routine.")
                        Return ("Unhandled Call Number Type")
                    End If
                End If
            End If
        End If


        'If we arrive here, process LC:

        Do While Char.IsLetter(Mid$(cn, pos, 1)) 'loop through LC class letters
            callout = callout + Mid$(cn, pos, 1)
            pos = pos + 1
        Loop 'exit when we hit a non-letter

        If lcNumericBreak.Checked Then
            callout = callout + vbCrLf 'add a cr/lf
        End If
        'End If
        'loop through class numerics, including decimal if present.
        'loop while we see digits, or a decimal followed by a digit.
        If Mid$(cn, pos, 1) = " " Then
            pos = pos + 1
        End If

        Do While Char.IsDigit(Mid$(cn, pos, 1)) _
        Or (Mid$(cn, pos, 1) = "." And Char.IsDigit(Mid$(cn, pos + 1, 1)))
            callout = callout + Mid$(cn, pos, 1)
            pos = pos + 1
        Loop

        If lcDecimalBreakAfter.Checked Then
            callout = callout.Replace(".", "." & vbCrLf)
        End If

        If lcDecimalBreakB4.Checked Then
            callout = callout.Replace(".", vbCrLf & ".")
        End If

        If lcCutterBreak.Checked Then
            callout = callout + vbCrLf 'add a cr/lf after the class numerics
        Else
            callout = callout + " "
        End If

        cutter = cn.Substring(pos - 1) 'trim off the rest of the cn and put it in the cutter var.
        'if cutter starts with a space or punctuation other than a decimal, remove it
        If Mid$(cutter, 1, 1) = " " Or
        (Char.IsPunctuation(Mid$(cutter, 1, 1)) And Not Mid$(cutter, 1, 1) = ".") Then
            cutter = cutter.Substring(1)
        End If
        '*******
        If lcCutterBreak.Checked Then
            cutter = cutter.Replace(" .", vbCrLf & ".")
        Else
            cutter = cutter.Replace(" .", ".")
        End If

        If hideCutterDecimal.Checked Then
            'cutter = cutter.Replace(".", "")

            'If cutter begins with a decimal, remove it.
            'But make sure the cutter exists before referring to a substring...
            If cutter.Length > 1 AndAlso cutter.Substring(0, 1) = "." Then
                cutter = cutter.Substring(1)
            End If

        End If

        '*******
        'parseCaption2 determines which spaces in the cutter should cause a line break.
        'In ".H27 L43 2011", or ".H2345x 1997" or ".H143 1990", each space should cause a break.
        'In ".H27 v. 4", the space between "v." and "4" should not cause a break.

        cutter = parseCaption2(cutter) 'initially, all spaces are turned into "?"

        'then, each is examined to see if should cause a break or not.
        If lcInCutterBreak.Checked Then
            cutter = cutter.Replace(" ", vbCrLf) 'those that should break are turned back into spaces
        End If
        cutter = cutter.Replace("?", " ") 'those that should not break are left as ?,
        'and this line turns those back into ordinary spaces before printing occurs.

        callout = callout + cutter

        If Not lcOtherNoBreak.Checked Then
            lcBreak.Text = Trim(lcBreak.Text)
            For i = 0 To lcBreak.Lines.Length - 1
                match = Trim(lcBreak.Lines(i))
                If match <> "" And callout.Contains(match) Then
                    If lcOtherBreakB4.Checked Then
                        callout = callout.Replace(match, vbCrLf & match)
                    Else
                        callout = callout.Replace(match, match & vbCrLf)
                    End If
                    If lcRemoveAfter.Checked Then callout = callout.Replace(match, "")
                End If
            Next
        End If

        callout = cleanup(callout)

        Return callout
    End Function

    Private Function parseDewey(ByVal cn As String)
        Dim callout As String = ""
        Dim cutter As String = ""
        Dim afterdecimal As String = ""
        Dim pos As Integer = 0
        Dim p As Integer = 0
        Dim i As Integer = 0
        Dim c As Integer = 0
        Dim ch As String = ""
        Dim dcnt As Integer = 0
        cn = Trim(cn)
        'find cutter:
        'look for 1st digit
        For i = 0 To cn.Length - 1
            If Char.IsDigit(Mid$(cn, i + 1, 1)) Then
                Exit For
            End If
        Next

        'loop until space (digits after decimal may contain alpha chars.
        For c = i To cn.Length - 1
            If Mid$(cn, c + 1, 1) = " " Then
                Exit For
            End If
        Next
        callout = cn.Substring(0, c)
        cutter = cn.Substring(c)

        If deweyPrefixBreak.Checked Then
            If Char.IsLetter(Mid$(callout, 1, 1)) Then
                For i = 0 To callout.Length - 1
                    If Char.IsDigit(Mid$(callout, i + 1, 1)) Then
                        callout = callout.Substring(0, i).Replace(" ", vbCrLf) & vbCrLf & callout.Substring(i)
                        Exit For
                    End If
                Next
            End If
        End If

        If Not deweyDecimalNoBreak.Checked Then
            'look for a decimal that has a digit on either side
            pos = callout.IndexOf(".")
            If pos <> -1 Then
                If Char.IsDigit(callout.Substring(pos - 1, 1)) And Char.IsDigit(callout.Substring(pos + 1, 1)) Then
                    'replace this decimal (and not any others later in the call num) with a vertical bar
                    callout = callout.Substring(0, pos) & "|" & callout.Substring(pos + 1)
                    If deweyDecimalBreakb4.Checked Then
                        callout = callout.Replace("|", vbCrLf & ".") 'break before: make it new line then dot
                    Else
                        callout = callout.Replace("|", "." & vbCrLf) 'break after: make it dot then new line
                        If deweyGroup3.Checked Then
                            callout = groupSplit(callout)
                        End If
                    End If
                End If
            End If
        Else
            'Break on a specified number of digits past the decimal point
            pos = callout.IndexOf(".") 'find the class decimal

            If pos <> -1 Then
                If deweyDecBreak.Checked Then
                    'xxxxxxxxxxxxxxxxxxxx New Routine Goes Here xxxxxxxxxxxxxxxxxxxxxx
                    p = pos + 1
                    'count the number of digits after the class decimal
                    For i = p To callout.Length - 1
                        ch = callout.Substring(i, 1)
                        If Not Char.IsDigit(ch) Then Exit For
                        dcnt = dcnt + 1
                    Next
                    If dcnt > CType(deweyDigitsToBreak.Text, Integer) Then
                        callout = callout.Insert(pos + CType(deweyDigitsToBreak.Text, Integer) + 1, vbCrLf)
                    End If
                End If
            End If
        End If

        If deweyCutterBreak.Checked Then
            callout = callout & vbCrLf & cutter.Replace(". ", "|").Replace(" ", vbCrLf).Replace("|", ". ")
        Else
            callout = callout & " " & cutter
        End If


        If deweyOtherNoBreak.Checked Then Return callout
        If Trim(deweyCharBreak.Text) = "" Then Return callout
        Dim match As String = ""
        For i = 0 To deweyCharBreak.Lines.Length - 1
            match = Trim(deweyCharBreak.Lines(i))
            If match <> "" And callout.Contains(match) Then
                If deweyCharBreakb4.Checked Then
                    callout = callout.Replace(match, vbCrLf & match)
                Else
                    callout = callout.Replace(match, match & vbCrLf)
                End If
                If deweyRemoveAfter.Checked Then callout = callout.Replace(match, "")
            End If
        Next

        usingDewey = True
        callout = cleanup(callout)
        Return callout
    End Function
    Private Function groupSplit(ByVal cn As String) As String
        Dim pos As Integer
        Dim newcn As String = ""
        Dim j As Integer = 0
        Dim k As Integer = 0
        Dim brake As Integer = 0
        Dim cpl As Integer

        cpl = CType(deweydigitsperline.Text, Integer)
        pos = cn.IndexOf(".") + 2
        If pos = -1 Then Return cn
        pos = pos + 1
        newcn = cn.Substring(0, pos) & vbCrLf
        'pos points to first character after the decimal point
        j = pos
        Do While pos < cn.Length AndAlso Char.IsDigit(cn.Substring(pos))
            brake = brake + 1
            If brake > 20 Then Return "whoa"
            newcn = newcn & cn.Substring(pos, 1)
            j = j + 1
            k = k + 1
            If k Mod cpl = 0 And pos + 1 < cn.Length AndAlso Char.IsDigit(cn.Substring(pos + 1)) Then
                newcn = newcn & vbCrLf
                j = j + 2
            End If
            pos = pos + 1
        Loop
        newcn = newcn & cn.Substring(pos)

        Return newcn
    End Function
    Private Function parseSuDoc2(ByVal cn As String) As String
        Dim cpart As Array
        Dim pre, post As String
        Dim separator = ""
        Dim i As Integer = 0
        Dim callout As String = ""

        If sudocBreakB4Numerics.Checked Then
            For i = 0 To cn.Length - 1
                If Char.IsDigit(Mid$(cn, i + 1, 1)) Then
                    cn = cn.Substring(0, i) & vbCrLf & cn.Substring(i)
                    Exit For
                End If
            Next
        End If

        If sudocBreakBeforeColon.Checked Then
            separator = vbCrLf & ":"
        Else
            If sudocBreakAfterColon.Checked Then
                separator = ":" & vbCrLf
            Else
                separator = ":"
            End If
        End If

        cpart = Split(cn, ":")
        pre = cpart(0)
        post = ""
        If cpart.Length > 1 Then 'If colon exists...
            post = cpart(1)
            If sudocSlashBreak.Checked Then
                post = cpart(1).replace("/", vbCrLf)
            End If
            If sudocBreakDecimal.Checked Then
                pre = pre.Replace(".", vbCrLf & ".")
            End If

            If sudocDashBreak.Checked Then
                post = breakSpace(post)
                post = post.Replace("-", vbCrLf)
            End If
            post = post.Replace("?", vbCrLf)
        Else
            post = "" 'if no colon, "post" is null
            separator = ""
        End If

        callout = cleanup(pre & separator & post)

        If sudocOtherNoBreak.Checked Then Return callout
        If Trim(sudocCharBreak.Text) = "" Then Return callout
        Dim match As String = ""
        For i = 0 To sudocCharBreak.Lines.Length - 1
            match = Trim(sudocCharBreak.Lines(i))
            If match <> "" And callout.Contains(match) Then
                If sudocCharBreakB4.Checked Then
                    callout = callout.Replace(match, vbCrLf & match)
                Else
                    callout = callout.Replace(match, match & vbCrLf)
                End If
                If sudocRemoveAfter.Checked Then callout = callout.Replace(match, "")
            End If
        Next

        Return callout
    End Function
    Private Function parseOther(ByVal cn As String)
        Dim callout As String = ""
        Dim i As Integer = 0
        'Call number type 4 = "Shelving Control Number; 
        'type 8 is "Other".
        'type 7 is "source specified in subfield $2" - added by customer request
        'These simply break on spaces.
        'If cntype = "4" Or cntype = "8" Then
        callout = cn
        If otherAllSpaceBreak.Checked Then
            callout = cn.Replace(" ", vbCrLf)
        Else
            If otherFirstSpaceBreak.Checked Then
                i = callout.IndexOf(" ")
                If i <> -1 Then
                    Mid$(callout, i + 1, 1) = "|"
                    callout = callout.Replace("|", vbCrLf)
                End If
            End If
        End If


        If otherNumBreakb4.Checked Then
            For i = 0 To callout.Length - 1
                If IsNumeric(Mid$(callout, i + 1, 1)) Then
                    callout = callout.Substring(0, i) & vbCrLf & callout.Substring(i)
                    Exit For
                End If
            Next
        Else
            If otherNumBreakAfter.Checked Then
                For i = 0 To callout.Length - 1
                    If IsNumeric(Mid$(callout, i + 1, 1)) And Mid$(callout, i + 2, 1) = " " Then
                        Mid$(callout, i + 2, 1) = "|"
                        callout = callout.Replace("|", vbCrLf)
                    End If
                Next
            End If
        End If
        callout = cleanup(callout)

        If otherListNoBreak.Checked Then Return callout
        If Trim(otherBreak.Text) = "" Then Return callout
        Dim match As String = ""
        For i = 0 To otherBreak.Lines.Length - 1
            match = Trim(otherBreak.Lines(i))
            If match <> "" And callout.Contains(match) Then
                If otherListBreakb4.Checked Then
                    callout = callout.Replace(match, vbCrLf & match)
                Else
                    callout = callout.Replace(match, match & vbCrLf)
                End If
                If otherRemoveAfter.Checked Then callout = callout.Replace(match, "")
            End If
        Next
        callout = cleanup(callout)
        Return callout
        'End If
    End Function

    Private Function breakSpace(ByVal cn As String) As String
        'replace a breakable space with "?"
        Dim i As Integer = 1
        Dim done As Integer = 0
        cn = Trim(cn) 'make sure no spaces at beginning or end of string
        done = cn.Length
        Do
            If Char.IsDigit(Mid$(cn, i, 1)) And Mid$(cn, i + 1, 1) = " " And Char.IsLetter(Mid$(cn, i + 2, 1)) Then
                Mid$(cn, i + 1, 1) = "?"
            End If
            i = i + 1
        Loop While i < done - 2
        Return cn
    End Function
    Private Function getParsing2(ByVal xmlsource As String) As String
        Dim callout As String = ""
        Dim buildLC As String = ""
        Dim prefix As String = ""
        Dim i As Integer = 1
        If xmlsource = "<parsed_issue_level_description>" Then
            prefix = "<issue_level_description_"
        Else
            prefix = "<call_no_"
        End If

        Do
            'buildLC = xmlValue(xmlsource & "<call_no_" & i & ">")
            buildLC = xmlValue(xmlsource & prefix & i & ">")
            If buildLC = "" Then Exit Do
            If i <> 1 Then callout = callout & vbCrLf
            callout = callout & buildLC
            i = i + 1
        Loop
        Return callout.Replace(vbCrLf & vbCrLf, vbCrLf)
    End Function

    Private Function parseIssue(ByVal issueText As String) As String

        'only return holdings if <issue_level_description> is selected in the "Spine" label type.
        'I.e., ignore the "include holdings" checkbox if using "Custom" label format.

        If Not chkIncludeHoldings.Checked And Spine.Checked Then
            Return ""
        End If

        Dim issueWork As String = issueText
        Dim maxchars = CType(inMaxChars.Text, Integer)
        Dim op As Integer = 0
        Dim cp As Integer = 0
        Dim i As Integer = 0

        If ProtectColon.Checked And ProtectColon.Enabled Then
            'if a colon is found inside of parentheses, protect it by replacing it with
            'a caret (^), and then later replacing the caret with a colon plus a cr/lf.
            'The colon remains visible on the final label, but it can still cause a line
            'break (if "break on colon" is checked).
            Try
                op = InStr(cp + 1, issueWork, "(", CompareMethod.Text)
                cp = InStr(op + 1, issueWork, ")", CompareMethod.Text)
                For i = op To cp
                    If Mid$(issueWork, i, 1) = ":" Then Mid$(issueWork, i, 1) = "^"
                Next
            Catch
                'The routine above will throw an exception if there are no parentheses in
                'the string, so we are sent here, which does nothing. The error is simply
                'ignored.
            End Try

        End If

        'If the issue level description contains a colon, assume that no parsing is needed other than
        'replacing colons with cr/lf.
        'if no colons, assume older format material that needs to be parsed by the parseCaption
        'routine (i.e., break on space unless space immediately follows an alpha string.)

        If ColonBreak.Checked And issueWork.IndexOf(":") <> -1 Then
            issueWork = issueWork.Replace(" :", ":").Replace(": ", ":").Replace(":", vbCrLf)
            If ProtectColon.Checked Then
                issueWork = issueWork.Replace("^", ":" & vbCrLf)
            End If
        Else
            issueWork = issueWork
            issueWork = Trim(parseCaption2(issueWork))
            'breakable spaces are now " ", non-breakable spaces are "?"

            If spaceBreak.Checked Then
                issueWork = issueWork.Replace(" ", vbCrLf) 'replace breakable spaces with crlf
            End If

            issueWork = issueWork.Replace("?", " ") 'turn non-breakable spaces back into spaces
        End If

        If BreakParen.Checked And issueWork.Replace(vbCrLf, "").Length > maxchars Then
            issueWork = issueWork.Replace("(", vbCrLf + "(")
        End If

        If issueListNoBreak.Checked Then Return issueWork
        If Trim(issueBreak.Text) = "" Then Return issueWork

        Dim match As String = ""

        For i = 0 To issueBreak.Lines.Length - 1
            match = Trim(issueBreak.Lines(i))
            match = match.Replace("~", " ")
            If match <> "" And issueWork.Contains(match) Then
                If issueListBreakB4.Checked Then
                    issueWork = issueWork.Replace(match, vbCrLf & match)
                Else
                    issueWork = issueWork.Replace(match, match & vbCrLf)
                End If
                If issueRemoveAfter.Checked Then issueWork = issueWork.Replace(match, "")
            End If

        Next

        Return issueWork
    End Function

    Function parseCaption2(ByVal cutter As String) As String
        Dim c As String = cutter
        Dim i As Integer = 0
        Dim done As Integer = 0
        Dim pos As Integer = 0
        'replace every space in the cutter with a "?" (which ensures that all spaces will not
        'cause a break.)
        c = c.Replace(" ", "?")
        i = 1
        done = c.Length

        'Then examine the characters on both sides of each "?" to see if that space
        'should be breakable or not.
        Do

            If Char.IsDigit(Mid$(c, i, 1)) And Mid$(c, i + 1, 1) = "?" And Char.IsLetter(Mid$(c, i + 2, 1)) Then
                'if we find digit/?/letter, that ? should be replaced by a space, causing a break.
                'so in "H237?J14", the ? will be replaced by a breakable space
                Mid$(c, i + 1, 1) = " "
            End If
            'If Mid$(c, i, 1) = "x" And Mid$(c, i + 1, 1) = "?" And Char.IsDigit(Mid$(c, i + 2, 1)) Then
            If Char.IsLetter(Mid$(c, i, 1)) And Mid$(c, i + 1, 1) = "?" And IsYear(c, i + 2) = True Then
                'if "x?digit", the space should be breakable as well.
                'so in "H1234x?1927", the space will be breakable.
                'MODIFIED: Any letter (not just 'x') followed by a 4-digit date will break.
                Mid$(c, i + 1, 1) = " "
            End If

            If Char.IsDigit(Mid$(c, i, 1)) And Mid$(c, i + 1, 1) = "?" And Char.IsDigit(Mid$(c, i + 2, 1)) Then
                'if "digit?digit", the space should be breakable
                'so in "H27?2011", the space will be breakable.
                Mid$(c, i + 1, 1) = " "
            End If

            i = i + 1
        Loop While i < done - 2

        Return c
    End Function
    Private Function IsYear(ByVal c As String, ByVal i As Integer) As Boolean
        'c is the string containing the cutter
        'i is the position in c of the first character that might or might not be the first digit of a year.
        'If the four characters starting at position i are all numeric, IsYear returns True.
        If Char.IsDigit(Mid$(c, i, 1)) And Char.IsDigit(Mid$(c, i + 1, 1)) And Char.IsDigit(Mid$(c, i + 2, 1)) And Char.IsDigit(Mid$(c, i + 3, 1)) Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Function aboveCall(ByVal libname As String, ByVal liblocation As String) As String
        Dim labeltext As String = ""
        Dim listentry As String = ""
        Dim lookfor As String = libname & "+" & liblocation
        Dim len As Integer = lookfor.Length
        Dim testentry As String = ""
        Dim i As Integer = 0
        Dim itemindex As Integer = 0
        For i = 0 To altList.Items.Count - 1
            testentry = altList.Items.Item(i)
            If Mid$(testentry, 1, testentry.IndexOf("=")) = lookfor Then
                labeltext = testentry.Substring(testentry.IndexOf("=") + 1) + vbCrLf
                Exit For
            End If
        Next
        Return labeltext
    End Function

    Private Sub loadLabelText()
        ignoreChange = True
        Dim inrec As String = ""
        Dim itm As String = ""
        Dim pos As Integer = 0
        Dim tab1 As Integer = 0
        Dim tab2 As Integer = 0
        Dim library As String = ""
        Dim location As String = ""
        Dim labeltext As String = ""
        Dim sw As StreamWriter
        Dim savecombotext As String = ComboBox1.SelectedItem

        Dim sr As StreamReader
        Dim fs As FileStream

        If (Not File.Exists(mypath & ALTfile)) Then
            sw = File.CreateText(mypath & ALTfile)
            sw.Close()
        End If
        fs = New FileStream(mypath & ALTfile, FileMode.Open)
        sr = New StreamReader(fs)
        itm = sr.ReadLine()
        altList.Items.Clear()
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("All Libraries")

        Try
            While Not itm Is Nothing
                If itm.Length > 5 Then 'ignore blank/incomplete lines
                    If itm.Contains(vbTab) Then
                        tab1 = InStr(itm, vbTab)
                        tab2 = InStr(tab1 + 1, itm, vbTab, CompareMethod.Text)
                        library = Mid$(itm, 1, tab1 - 1)
                        location = Mid$(itm, tab1 + 1, tab2 - tab1 - 1)
                        labeltext = Mid$(itm, tab2 + 1, 99)
                        inrec = library & "+" & location & "=" & labeltext
                        If Trim(labeltext) = "" Then
                            itm = sr.ReadLine()
                            Continue While
                        End If
                    Else
                        inrec = Trim(itm)
                        pos = InStr(itm, "+", CompareMethod.Text)
                        library = Mid$(itm, 1, pos - 1)
                        pos = InStr(itm, "=", CompareMethod.Text)
                        If pos = inrec.Length Then
                            labeltext = ""
                        Else
                            labeltext = Mid$(inrec, pos + 1, inrec.Length - pos)
                        End If
                        If Trim(labeltext) = "" Then
                            itm = sr.ReadLine()
                            Continue While
                        End If
                    End If
                    If ComboBox1.Items.Contains(library) = False Then
                        ComboBox1.Items.Add(library)
                    End If

                    If ComboBox1.Text = "All Libraries" Or ComboBox1.Text = library Then
                        altList.Items.Add(inrec)
                    End If
                End If

                itm = sr.ReadLine()
            End While
        Catch ex As Exception
            MsgBox("The file '" & ALTfile & "' could not be found." &
            vbCrLf & "Reason: " & ex.Message)
        End Try

        sr.Close()
        fs.Close()
        ComboBox1.Text = savecombotext
        ignoreChange = False
    End Sub

    Private Sub writeFile(ByVal fpath As String, ByVal fstring As String, ByVal append As Boolean)

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

    'Private Sub GetVideoInfo()
    ' Dim query As New SelectQuery("Win32_DesktopMonitor")
    ' Dim searcher As New ManagementObjectSearcher(query)''

    'For Each envVar As ManagementBaseObject In searcher.Get()
    '   For Each obj As PropertyData In envVar.Properties
    'Console.WriteLine(obj.Name) 'get the name to pass in below
    '  Next
    ' pixelsPerInchX = envVar("PixelsPerXLogicalInch")
    'pixelsPerInchY = envVar("PixelsPerYLogicalInch")
    'Next
    'MsgBox("pixels/in X: " & pixelsPerInchX & " pixels/in Y: " & pixelsPerInchY)
    'End Sub

    Private Sub PrinterDialogButn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrinterDialogButn.Click
        Dim dlgSettings As New PrintDialog()
        dlgSettings.UseEXDialog = True
        dlgSettings.Document = PrintDocument1
        'dlgSettings.ShowDialog()
        If dlgSettings.ShowDialog() = Windows.Forms.DialogResult.OK Then
            inPrinterName.Text = dlgSettings.PrinterSettings.PrinterName
        End If
    End Sub

    Private Sub FontDialogButn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontDialogButn.Click
        Dim fntSettings As New FontDialog()
        'fntSettings.ShowDialog()
        If fntSettings.ShowDialog() = Windows.Forms.DialogResult.OK Then
            inFontName.Text = fntSettings.Font.Name
            inFontSize.Text = fntSettings.Font.Size
            inFontWeight.Checked = fntSettings.Font.Bold
        End If

    End Sub

    Private Sub ftpPrint()
        'Triggers print to FTP printer when labelout.txt is written to the installation
        'directory, alerting FileSystemWatcher2 to begin processing.
        Dim tline As Array = Split(OutputBox.Text, vbCrLf)
        Dim t As Integer = 0
        Dim labelcmds As String = ""
        Dim LMar As Integer = CType(TextBox19.Text, Integer)
        Dim Top As Integer = CType(TextBox20.Text, Integer)
        Dim Inc As Integer = CType(TextBox21.Text, Integer)
        Dim Fontnum As String = TextBox22.Text
        Dim Fontsize As String = TextBox23.Text
        Dim textout As String = ""
        Dim maxLines As Integer = CType(inMaxLines.Text, Integer)
        Dim labelTop As Integer = 0
        Dim lineCount As Integer = 0
        Dim curline As Integer = 0
        Dim lastLine As Integer = 0
        If Trim(tline(tline.Length - 1)) = "" Then
            lastLine = tline.Length - 2
        Else
            lastLine = tline.Length - 1
        End If
        Do
            labelTop = Top
            For t = curline To lastLine
                labelcmds = labelcmds & "T " & LMar.ToString & "," & labelTop.ToString & "," &
                "0," & Fontnum & "," & "pt" & Fontsize & ";" & tline(t)
                If t < tline.Length - 1 Then labelcmds = labelcmds & vbCrLf
                labelTop = labelTop + Inc
                lineCount = lineCount + 1
                If lineCount = maxLines Then
                    lineCount = 0
                    Exit For
                End If
            Next
            textout = textout & TextBox17.Text.Replace("#LABELTEXT#", labelcmds).Replace("#COUNT#", LabelRepeat.Value) & vbCrLf
            labelcmds = ""
            curline = t + 1
        Loop While curline <= lastLine
        writeFile(mypath & "labelout.txt", textout, False)
        changeCount = 0
        writeStat("F") 'Add "F" to statrec and write to statfile.
        OutputBox.Text = ""
        InputBox.Text = ""
        InputBox.Focus()
    End Sub

    Private Sub FileSystemWatcher2_Changed(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs) Handles FileSystemWatcher2.Changed
        '
        'Used with the FTP printer type only.  When the "labelout.txt" file is written
        '(this is the file that contains "JScript" label printing instructions for the industrial-style 
        'FTP-attached CAB printer), this filesystemwatcher's "changed" event is triggered.  This starts the 
        'ftpbat.bat batch file process which logs into the printer's FTP server and transfers
        'the contents of the labelout.txt file to the printer's "/execute" directory, causing the
        'printer to print the label.
        '
        'The "changed" event fires multiple times for a single file change. To avoid
        'multiple executions of this code, a SyncLock ensures that only one event at a 
        'time can enter the "start" routine. The "changeCount" is set to zero in the 
        'calling routine, and incremented every time the SyncLock is entered. In SyncLock,
        'only if changeCount is zero will the batch process start.

        Dim ftpBatchFilePath As String = mypath & "ftpbat.bat"
        Dim maxLines As Integer = CType(inMaxLines.Text, Integer)
        Dim startInfo As New ProcessStartInfo(ftpBatchFilePath)
        Dim p As Process
        'startInfo.WindowStyle = ProcessWindowStyle.Hidden 'makes sure entire process is invisible
        startInfo.Arguments = FTPip.Text 'the IP address is passed to ftpbat.bat as an argument
        Dim mlock As New Object
        SyncLock mlock
            If changeCount = 0 Then
                changeCount = changeCount + 1
                p = Process.Start(startInfo) 'tell Windows to start the batch file
            End If
        End SyncLock

    End Sub
    Private Sub viaDOS(ByVal labeltext As String)
        Dim dosPath As String = mypath & "viados.bat"
        Dim startInfo As New ProcessStartInfo(dosPath)
        Dim p As Process
        Dim p1 As String = ""
        Dim p2 As String = ""
        Dim quot As String = """"

        If hideDosWindow.Checked Then
            Dim batchtext As String = ""
            Dim sr As StreamReader
            Try
                sr = New StreamReader(mypath & "viados.bat")
                batchtext = sr.ReadToEnd()
                sr.Close()
            Catch ex As Exception
                MsgBox("No viados.bat file was found.", MsgBoxStyle.Exclamation, "No viados.bat file")
                Exit Sub
            End Try
            If batchtext.ToUpper.Contains("PAUSE") Or batchtext.ToUpper.Contains("PATHNAME") Then
                MsgBox("This batch file cannot in run in hidden mode because it contains 'pause' or 'PathName' commands " &
                "that will require a user response.", MsgBoxStyle.Exclamation, "Hidden Mode Error")
                Exit Sub
            End If
        End If

        writeFile(mypath & "label.txt", labeltext, False)
        Thread.Sleep(300)
        If Trim(dosParam1.Text).Contains("<") Then
            p1 = xmlValue(dosParam1.Text)
        Else
            p1 = dosParam1.Text
        End If
        If Trim(dosParam2.Text).Contains("<") Then
            p2 = xmlValue(dosParam2.Text)
        Else
            p2 = dosParam2.Text
        End If
        startInfo.Arguments = quot & p1 & quot & " " & quot & p2 & quot
        If hideDosWindow.Checked Then
            startInfo.WindowStyle = ProcessWindowStyle.Hidden
        End If
        Try
            p = Process.Start(startInfo)
        Catch ex As Exception
            MsgBox("The DOS Batch File could not be started." & vbCrLf & "Error: " & ex.Message, MsgBoxStyle.Exclamation, "Batch File Error")
        End Try

    End Sub
    Private Sub InputBox_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles InputBox.DragDrop
        Dim temptext As String = ""
        If e.Data.GetDataPresent(DataFormats.Text) Then
            temptext = e.Data.GetData(DataFormats.Text).ToString()
            If temptext.Length > 20 Then
                InputBox.Text = "???"
                Beep()
            Else
                InputBox.Focus()
                InputBox.Text = temptext
                AppActivate("SpineOMatic " & somVersion.Replace("/", "."))
                callAlma()
            End If
        End If
    End Sub

    Private Sub InputBox_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles InputBox.DragEnter
        If e.Data.GetDataPresent(DataFormats.Text) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub InputBox_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles InputBox.GotFocus
        InputBox.BackColor = Color.FromArgb(220, 255, 220)
    End Sub
    Private Sub InputBox_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles InputBox.LostFocus
        InputBox.BackColor = Color.FromArgb(255, 220, 220)
    End Sub
    Private Sub InputBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles InputBox.KeyPress

        If e.KeyChar = ControlChars.Cr Then
            e.Handled = True
            If InputBox.Text.ToLower = "admin" Then
                InputBox.Text = ""
                If settingsOpen Then CloseSettings() Else openSettings()
                updatePath.Visible = False
                InputBox.Focus()
            Else
                If InputBox.Text.ToLower = "adminsecure" Then
                    InputBox.Text = ""
                    openSettings()
                    InputBox.Focus()
                    updatePath.Visible = True
                Else
                    If InputBox.Text.ToLower = "adminl0gv1ew" Then
                        InputBox.Text = ""
                        openSettings()
                        InputBox.Focus()
                        OutputBox.Text = "Log View Enabled"
                        getWebLog()
                        byAlpha.Checked = True
                        chkGeoList.Visible = True
                        chkAddHostname.Visible = True
                    Else
                        If InputBox.Text.ToLower = "java" Then

                            OutputBox.Text = GetJavaVersionInfo()
                        Else
                            If usrname.Text = "" And chkRequireUser.Checked Then
                                OutputBox.Font = New Font("MS Sans Serif", 9, FontStyle.Regular)
                                OutputBox.Text = usermessage
                                usrname.BackColor = Color.Yellow
                                usrname.Focus()
                            Else
                                callAlma()
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ScanButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ScanButton.Click
        OutputBox.Text = ""
        plOutput.Text = ""
        lblXMLWarn.Visible = False
        HistoryList.Visible = False
        ArrowLabel.Text = "q"
        parsedBy.Text = ""
        holdingsBy.Text = ""
        InputBox.Focus()
        InputBox.Text = ""
    End Sub

    Private Sub AutoPrintBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoPrintBox.CheckedChanged
        If AutoPrintBox.Checked Then
            AutoPrintBox.Font = New Font(AutoPrintBox.Font, FontStyle.Bold)
            ReviewBox.Font = New Font(ReviewBox.Font, FontStyle.Regular)
        End If
        InputBox.Focus()
    End Sub

    Private Sub ReviewBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReviewBox.CheckedChanged
        If ReviewBox.Checked Then
            ReviewBox.Font = New Font(ReviewBox.Font, FontStyle.Bold)
            AutoPrintBox.Font = New Font(AutoPrintBox.Font, FontStyle.Regular)
        Else
            ReviewBox.Font = New Font(ReviewBox.Font, FontStyle.Regular)
            AutoPrintBox.Font = New Font(AutoPrintBox.Font, FontStyle.Bold)
        End If
        InputBox.Focus()

    End Sub

    Private Sub UseDesktop_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UseDesktop.CheckedChanged
        If UseDesktop.Checked Then
            refreshFTP()
            UseDesktop.Font = New Font(UseDesktop.Font, FontStyle.Bold)
            UseLaser.Font = New Font(UseLaser.Font, FontStyle.Regular)
            UseFTP.Font = New Font(UseFTP.Font, FontStyle.Regular)
            SheetSettings.Visible = False
            FTPGroup.Visible = False
            orientationPanel.Visible = True
            marginPanel.Visible = True
            UseDesktop.Size = New Point(360, 168)
            DesktopGroup.Visible = True
            ManualPrint.Text = "Send to desktop printer"
            AutoPrintBox.Text = "Auto Print"
            ToolTip1.SetToolTip(AutoPrintBox, "Print label without reviewing it")
            InputBox.Focus()
        End If
    End Sub
    Private Sub UseLaser_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UseLaser.CheckedChanged
        If UseLaser.Checked Then
            refreshFTP()
            UseLaser.Font = New Font(UseLaser.Font, FontStyle.Bold)
            UseDesktop.Font = New Font(UseDesktop.Font, FontStyle.Regular)
            UseFTP.Font = New Font(UseFTP.Font, FontStyle.Regular)
            FTPGroup.Visible = False
            DesktopGroup.Visible = True
            SheetSettings.Visible = True
            orientationPanel.Visible = True
            marginPanel.Visible = True
            AutoPrintBox.Text = "Auto Add"
            ToolTip1.SetToolTip(AutoPrintBox, "Add label to batch without reviewing it.")
            ManualPrint.Text = "Add to batch #" & batchNumber.Value.ToString()
            InputBox.Focus()
        End If
    End Sub
    Private Sub UseFTP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UseFTP.CheckedChanged
        If UseFTP.Checked Then

            UseFTP.Font = New Font(UseFTP.Font, FontStyle.Bold)
            UseLaser.Font = New Font(UseLaser.Font, FontStyle.Regular)
            UseDesktop.Font = New Font(UseDesktop.Font, FontStyle.Regular)
            DesktopGroup.Visible = False
            FTPGroup.Location = New Point(8, 28)
            FTPGroup.Size = New Point(400, 295) '400 x 280
            FTPGroup.Visible = True
            orientationPanel.Visible = False
            marginPanel.Visible = False
            AutoPrintBox.Text = "Auto Print"
            ToolTip1.SetToolTip(AutoPrintBox, "Print label without reviewing it.")
            refreshFTP()
            InputBox.Focus()
        End If
    End Sub
    Private Sub refreshFTP()
        If chkUsePocketLabels.Checked And UseFTP.Checked Then
            FTPwarning.Visible = True
            FTPwarning2.Visible = True
            ManualPrint.Text = "--- Not Supported ---"
            ManualPrint.Enabled = False
        Else
            FTPwarning.Visible = False
            FTPwarning2.Visible = False
            ManualPrint.Enabled = True
            ToolTip1.SetToolTip(ManualPrint, "Send label text to selected printer.")
            If UseFTP.Checked Then ManualPrint.Text = "Send to FTP Printer"
            If UseLaser.Checked Then ManualPrint.Text = "Add to batch #" & batchNumber.Value.ToString
            If UseDesktop.Checked Then ManualPrint.Text = "Send to desktop printer"
            If useDOSBatch.Checked Then
                ManualPrint.Text = "Run 'viados' batch file"
                ToolTip1.SetToolTip(ManualPrint, "Writes label text to 'label.txt' file, then starts 'viados.bat' batch file.")
            End If

        End If
        If Spine.Checked Then
            If radio_useSOMparsing.Checked Then
                spineType.Text = "Spine (SpineOMatic parsing)"
            Else
                spineType.Text = "Spine (Ex Libris parsing)"
            End If
        Else
            If CustomLabel.Checked Then
                spineType.Text = "Custom"
                If FlagSlips.Checked Then
                    spineType.Text = "Custom/Flag Slips"
                End If
            End If
        End If
    End Sub

    Private Sub openSettings()
        'expand user interface box to make administrative panels visible
        Me.Size = New Size(770, 453)
        settingsOpen = True
        lblToggleAdmin.Text = "t"
        InputBox.Focus()
    End Sub

    Private Sub saveSettings(ByVal savetype As String)
        TextBox24.ForeColor = Color.Red
        TextBox24.Text = "SAVING SETTINGS..."
        Application.DoEvents()
        Dim conlist As List(Of Control)
        Dim ctl As Control
        Dim val As String = ""
        Dim r As RadioButton
        Dim c As CheckBox
        Dim n As NumericUpDown
        Dim vline As String = "<spinelabel_settings>" & vbCrLf

        If pocketDefaultLoaded = True Then
            SaveDefaults(pocketDefaults)
        Else
            If spineDefaultLoaded = True Then
                SaveDefaults(spineDefaults)
            Else
                If customFlagDefaultLoaded = True Then
                    SaveDefaults(flagDefaults)
                    'wrapWidth.Text = flagWrapWidth.Text
                    flagWrapWidth.Text = wrapWidth.Text
                Else
                    SaveDefaults(nonFlagDefaults)
                    'wrapWidth.Text = nonFlagWrapWidth.Text
                    nonFlagWrapWidth.Text = wrapWidth.Text
                End If
            End If
        End If

        conlist = GetAllControls(Me)    ' function GetAllControls recursively loops through all 
        'controls on the form and returns a complete list.    

        For Each ctl In conlist

            If ctl.Tag <> Nothing Then  'for all controls that have a "tag" entry...
                If Mid$(ctl.Tag, 1, 5) = "radio" Then 'for radio button type controls
                    r = CType(ctl, RadioButton) 'save the radio button as either selected or not
                    If r.Checked Then val = "true" Else val = "false"
                End If
                If Mid$(ctl.Tag, 1, 5) = "check" Then 'for checkbox type controls
                    c = CType(ctl, CheckBox) 'save the checkbox as either checked or not
                    If c.Checked Then val = "true" Else val = "false"
                End If
                If Mid$(ctl.Tag, 1, 4) = "text" Then  'for all text type controls
                    val = ctl.Text 'save the text of .text attribute
                    val = val.Replace(vbCrLf, "~cr~") 'but change cr/lf to a code that can be 
                End If 'turned back into a cr/lf later

                If Mid$(ctl.Tag, 1, 4) = "texo" Then  'for all obscured text type controls
                    val = ctl.Text 'save the text of .text attribute
                    val = obscure(val)
                    val = val.Replace(vbCrLf, "~cr~") 'but change cr/lf to a code that can be 
                End If 'turned back into a cr/lf later

                If Mid$(ctl.Tag, 1, 5) = "numud" Then 'for the numeric up/down control
                    n = CType(ctl, NumericUpDown) 'save the numeric value 
                    val = n.Value
                End If 'Then put the saved value between <controlname>...</controlname> elements
                vline = vline & "<" & ctl.Tag & ">" & val & "</" & ctl.Tag & ">" & vbCrLf
            End If
        Next

        vline = vline & "</spinelabel_settings>"
        If savetype = "todisk" Then
            RichTextBox1.Clear()
            RichTextBox1.Text = vline
            formatXML()
            writeFile(mypath & "settings.som", vline, False)
            original_settings = vline.Replace(vbCrLf, "")
        Else
            closing_settings = vline.Replace(vbCrLf, "")
        End If
        TextBox24.Text = ""
    End Sub

    Private Sub SaveSettingsButn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveSettingsButn.Click
        saveSettings("todisk")
    End Sub
    Private Sub LoadSettingsButn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadSettingsButn.Click
        GetSettingsFile()
    End Sub
    Private Sub GetSettingsFile()
        TextBox24.ForeColor = Color.Red
        TextBox24.Text = "LOADING SETTINGS..."
        Application.DoEvents()
        Dim settingsFile As String = mypath & "settings.som"
        settingsfound = True
        Try
            Dim tr As TextReader = New StreamReader(settingsFile)
            RichTextBox1.Clear()
            RichTextBox1.ForeColor = Color.Red
            RichTextBox1.Text = tr.ReadToEnd
            tr.Close()
        Catch
            openSettings()
            MsgBox("Welcome to SpineOMatic.  To get started:" _
            & vbCrLf & vbCrLf & "* In the Print Setup panel, select a printer, and a text font." _
            & vbCrLf & "* In the Alma Access panel, enter your Ex Libris credentials." _
            & vbCrLf & "* Obtain an API key from the Ex Libris Developer Network," _
            & vbCrLf & "and fill out the 'RESTful API' section in the Alma Access panel." & vbCrLf _
            & vbCrLf & "There are lots of other options and settings available--" _
            & vbCrLf & "click the 'About' link to get the complete manual." _
            & vbCrLf & "The Quick Start section will help get you up and running.", MsgBoxStyle.Exclamation, "Getting Started")
            settingsfound = False
            If UseDesktop.Checked Then
                UseDesktop.Checked = False
                UseDesktop.Checked = True
            End If
            If UseLaser.Checked Then
                UseLaser.Checked = False
                UseLaser.Checked = True
            End If
            If UseFTP.Checked Then
                UseFTP.Checked = False
                UseFTP.Checked = True
            End If

            FlagSlips.Enabled = False
            CustomLabel.Checked = True
            LoadDefaults(nonFlagDefaults)
            Spine.Checked = True

            CustomText.Enabled = False


        End Try
        If settingsfound Then
            InstallSettings()
            formatXML()
            'each time settings are read from disk, they are saved for later comparison to
            'the settings that are in the form when the program closes. This can alert the user
            'that the current settings have not been saved.
            original_settings = RichTextBox1.Text
        End If
        TextBox24.Text = ""
        TextBox24.ForeColor = Color.Blue
    End Sub

    Private Sub InstallSettings()
        Dim conlist As List(Of Control)
        Dim ctl As Control
        Dim val As String = ""
        Dim t As TextBox
        Dim r As RadioButton
        Dim c As CheckBox
        Dim n As NumericUpDown
        conlist = GetAllControls(Me) 'GetAllControls recursively steps through all controls on 
        For Each ctl In conlist 'the form, and returns a list.
            If ctl.Tag <> Nothing Then 'if the control has a .tag attribute
                val = getval(ctl.Tag) 'get the .tag text
                If Mid$(ctl.Tag, 1, 5) = "radio" Then 'if the tag starts with "radio"
                    r = CType(ctl, RadioButton) 'then this is a radio button
                    If val <> "" Then 'if a value is present, set the .checked attribute as true or false
                        If val = "true" Then r.Checked = True Else r.Checked = False
                    End If
                End If
                If Mid$(ctl.Tag, 1, 5) = "check" Then 'if this is a checkbox
                    c = CType(ctl, CheckBox)
                    If val <> "" Then 'and a value is present, set the checkbox to checked or unchecked
                        If val = "true" Then c.Checked = True Else c.Checked = False
                    End If
                End If
                If Mid$(ctl.Tag, 1, 4) = "text" Then 'if this is a text value
                    t = CType(ctl, TextBox) 'set the textbox.text to the saved text
                    If val <> "" Then
                        t.Text = val.Replace("~cr~", vbCrLf) 'and turn codes back into real cr/lfs
                    End If
                End If

                If Mid$(ctl.Tag, 1, 4) = "texo" Then 'if this is an obscured text value
                    t = CType(ctl, TextBox) 'set the textbox.text to the saved text
                    If val <> "" Then
                        val = obscure(val)
                        t.Text = val.Replace("~cr~", vbCrLf) 'and turn codes back into real cr/lfs
                    End If
                End If

                If Mid$(ctl.Tag, 1, 5) = "numud" Then 'if this is a numeric up/down control
                    n = CType(ctl, NumericUpDown)
                    If val <> "" Then 'set the control's value to the number that was saved.
                        n.Value = CType(val, Integer)
                    End If
                End If
            End If
        Next
        'ProgressBar1.Visible = False
        'the following toggles make sure that the text on the manual print button matches
        'the actual setting of the printer radio button, or the viados selection.
        If UseDesktop.Checked Then
            UseDesktop.Checked = False
            UseDesktop.Checked = True
        End If
        If UseLaser.Checked Then
            UseLaser.Checked = False
            UseLaser.Checked = True
        End If
        If UseFTP.Checked Then
            UseFTP.Checked = False
            UseFTP.Checked = True
        End If
        If useDOSBatch.Checked Then
            useDOSBatch.Checked = False
            useDOSBatch.Checked = True
        End If

        If dosPlUseCol.Checked Then
            dosPlUseCol.Checked = False
            dosPlUseCol.Checked = True
        Else
            dosPlUseTab.Checked = False
            dosPlUseTab.Checked = True
        End If

        If CustomLabel.Checked Then
            FlagSlips.Enabled = True
        Else
            FlagSlips.Enabled = False
        End If
        If useExlibrisParsing.Checked Then
            parsingSource.Enabled = True
            inCallNumSource.Enabled = False
        Else
            parsingSource.Enabled = False
            inCallNumSource.Enabled = True
        End If
        If chkIncludeOther.Checked Then
            inOtherSource.Enabled = True
        Else
            inOtherSource.Enabled = False
        End If
        If radio_useLocal.Checked Then
            ALTfile = "myLabelText.txt"
        Else
            ALTfile = "aboveLabel.txt"
        End If

        If UseRestfulApi.Checked Then
            accessType.Text = "R"
            accessType.ForeColor = Color.Green
            ToolTip1.SetToolTip(accessType, "Using preferred RESTful API access to Alma.")
        Else
            accessType.Text = "J"
            accessType.ForeColor = Color.Red
            ToolTip1.SetToolTip(accessType, "Caution--Using deprecated Java access to Alma.")
        End If

        If chkUsePocketLabels.Checked Then 'pocket must be checked first, because spine or custom can also be checked,
            'but pocked labels predominate.
            LoadDefaults(pocketDefaults)
            'chk_VerticalLine.Checked = pocketVerticalLine
            pocketDefaultLoaded = True : spineDefaultLoaded = False : customFlagDefaultLoaded = False : customNonFlagDefaultLoaded = False
            PocketLabelPanel.Location = New Point(0, 30)
            PocketLabelPanel.Width = TabPage2.Size.Width
            PocketLabelPanel.Height = TabPage2.Size.Height - 30
            PocketLabelPanel.Visible = True
        Else
            If Spine.Checked Then
                LoadDefaults(spineDefaults)
                pocketDefaultLoaded = False : spineDefaultLoaded = True : customFlagDefaultLoaded = False : customNonFlagDefaultLoaded = False
            Else
                If CustomLabel.Checked Then
                    If FlagSlips.Checked Then
                        LoadDefaults(flagDefaults)
                        wrapWidth.Text = flagWrapWidth.Text
                        pocketDefaultLoaded = False : spineDefaultLoaded = False : customFlagDefaultLoaded = True : customNonFlagDefaultLoaded = False
                    Else
                        LoadDefaults(nonFlagDefaults)
                        wrapWidth.Text = nonFlagWrapWidth.Text
                        pocketDefaultLoaded = False : spineDefaultLoaded = False : customFlagDefaultLoaded = False : customNonFlagDefaultLoaded = True
                    End If
                End If
            End If
        End If

        If Not btnPlCustom.Checked Then
            userDefinedPanel.Enabled = False
        End If
        If deweyDecimalBreakAft.Checked Then
            deweyGroup3.Enabled = True
            deweydigitsperline.Enabled = True
            Label25.Enabled = True
        End If

        TabControl1.SelectedTab = TabPage2
        TabControl1.SelectedTab = TabPage1


    End Sub

    Public Function GetAllControls(ByVal parent As Control) As List(Of Control)
        'recursive routine to step through all contols on the form.  Necessary because
        'controls can contain controls which can contain controls, etc.
        Dim list As New List(Of Control)
        For Each ctl As Control In parent.Controls
            list.Add(ctl)
            If ctl.Controls.Count > 0 Then
                list.AddRange(GetAllControls(ctl))
            End If
        Next
        Return list 'returns a list of all controls on the form, no matter where they appear.
    End Function

    Private Function getval(ByVal tagname As String) As String
        'returns the value found between <tagname>...</tagname> in the Current XML richtextbox.
        Dim val As String = ""
        Dim tagstart As Integer = 0
        Dim opentag As String = ""
        Dim closetag As String = ""
        Dim datastart As Integer = 0
        Dim tagend As Integer = 0
        opentag = "<" & tagname & ">"
        closetag = "</" & tagname & ">"

        With RichTextBox1
            tagstart = .Find(opentag)
            If tagstart = -1 Then
                val = ""
                Return val
                Exit Function
            End If
            datastart = tagstart + opentag.Length
            tagend = .Find(closetag, tagstart + 1, RichTextBoxFinds.MatchCase)
            If datastart = tagend Then
                val = ""
            Else
                val = Mid$(.Text, datastart + 1, tagend - datastart)
            End If

        End With
        Return val
    End Function

    Private Sub ArrowLabel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ArrowLabel.Click
        If HistoryList.Visible = False Then
            HistoryList.Visible = True
            ArrowLabel.Text = "p"
        Else
            HistoryList.Visible = False
            ArrowLabel.Text = "q"
        End If
    End Sub

    Private Sub HistoryList_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles HistoryList.DoubleClick
        InputBox.Text = HistoryList.Items.Item(HistoryList.SelectedIndex)
        HistoryList.Visible = False
        ArrowLabel.Text = "q"
        OutputBox.Text = ""
        InputBox.Focus()
    End Sub

    Private Sub TempLabelBox_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TempLabelBox.MouseDown
        TempLabelBox.DoDragDrop(TempLabelBox.Text, DragDropEffects.Copy)
    End Sub

    Private Sub CreateTempLbl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CreateTempLbl.Click
        Dim tempnum As String = ""
        tempnum = "TX" & Date.Now.ToString("yyMMddHHmmssfff", CultureInfo.InvariantCulture)
        TempLabelBox.Text = tempnum
    End Sub

    Private Sub createTemp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles createTemp.CheckedChanged
        If createTemp.Checked Then
            CreateTempLbl.Visible = True
            TempLabelBox.Visible = True
        Else
            CreateTempLbl.Visible = False
            TempLabelBox.Visible = False
        End If
    End Sub

    Private Sub CheckForUpdates_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckForUpdates.Click
        Dim webClient As New System.Net.WebClient()

        Dim result As String = "The new SpineOMatic file was downloaded successfully."
        Dim msgStyle As Integer = Microsoft.VisualBasic.MsgBoxStyle.Information
        Dim getExe As Boolean = False
        Dim OKtoRename As Boolean = False
        'Dim webRequest As String = ""
        'Dim versionList As String = ""
        'Dim verlist As Array
        'Dim vstart As Integer = 0, vlen As Integer = 0
        'Dim pv As String = "" 'preferred version
        'Dim exename As String = ""
        'Dim i As Integer = 0

        Dim sLatestVersion As String
        Dim sDownloadURL As String

        Dim request As HttpWebRequest
        Dim response As HttpWebResponse
        Dim reader As StreamReader
        Dim json As String
        Dim jss As New JavaScriptSerializer()

        Try
            request = HttpWebRequest.Create("https://api.github.com/repos/ExLibrisGroup/SpineOMatic/releases/latest")
            request.UserAgent = "SpineOMatic"
            request.Method = "GET"
            response = request.GetResponse()
            reader = New StreamReader(response.GetResponseStream)
            json = reader.ReadToEnd()
            Dim dict As Dictionary(Of String, Object) = jss.Deserialize(Of Dictionary(Of String, Object))(json)
            sLatestVersion = dict("tag_name")
            Dim dict2 As Dictionary(Of String, Object) = dict("assets")(0)  'First asset is the binary; Second asset is the documentation.
            sDownloadURL = dict2("browser_download_url")
        Catch ex As Exception
            msgStyle = Microsoft.VisualBasic.MsgBoxStyle.Exclamation
            result = "ERROR--Unable to check for new version of SpineOMatic." _
            & vbCrLf & vbCrLf & ex.ToString()
            MsgBox(result, msgStyle, "SpineOMatic Download")
            Exit Sub
        End Try

        If sLatestVersion <> "" And sLatestVersion <> "v" & somVersion Then   'if the latest version is not what is running
            Dim box = MessageBox.Show("There is a newer version of SpineOMatic." & vbCrLf & "Do you want to download it now?", "Download Decision", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If box = box.No Then
                result = "A new version of SpineOMatic will not be downloaded at this time."
                getExe = False
            Else
                getExe = True
            End If
        Else
            MsgBox("Your version of SpineOMatic is current.")
            Exit Sub
        End If

        If getExe Then
            Try
                webClient.DownloadFile(sDownloadURL, mypath & "somtemp.exe")
                OKtoRename = True
            Catch ex As Exception
                msgStyle = Microsoft.VisualBasic.MsgBoxStyle.Exclamation
                result = "ERROR--Unable to download the new version of SpineOMatic." _
                & vbCrLf & vbCrLf & ex.ToString()
            End Try
        End If
        If OKtoRename Then
            My.Computer.FileSystem.RenameFile(mypath & "SpineLabeler.exe", "SpineLabeler-" & somVersion.Replace(".", "_") & ".exe")
            My.Computer.FileSystem.RenameFile(mypath & "somtemp.exe", "SpineLabeler.exe")
            result = "SpineOMatic version " & sLatestVersion & " has been downloded." & vbCrLf &
            "You must close and restart SpineOMatic to begin using the new version."
        End If

        MsgBox(result, msgStyle, "SpineOMatic Download")
    End Sub

    Private Sub batchNumber_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles batchNumber.ValueChanged
        If Not settingsLoaded Then Exit Sub
        Dim ln As Integer = 0
        ManualPrint.Text = "Add to batch #" & batchNumber.Value.ToString()

        batchPreview.Text = GetBatch(batchNumber.Value)
        If batchPreview.Lines.Length > 0 Then
            batchEntries.Text = countBatch()
        Else
            batchEntries.Text = "0"
        End If

        InputBox.Focus()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim box2 = MessageBox.Show("Click OK to empty batch file #" & batchNumber.Value & ", or " & vbCrLf &
        "click CANCEL to keep the file.", "Confirm File Clear", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
        If box2 = box2.OK Then
            writeFile(mypath & "labelbatch" & batchNumber.Value & ".txt", "", False)
            batchPreview.Text = GetBatch(batchNumber.Value)
            batchPreview.Text = ""
            If batchPreview.Lines.Length > 0 Then
                batchEntries.Text = countBatch()
            Else
                batchEntries.Text = "0"
            End If
        Else
            MsgBox("Label batch file #" & batchNumber.Value & " has not been cleared.")
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        If ignoreChange = False Then
            loadLabelText()
        End If
    End Sub

    Private Sub createBatFiles()
        'create a DOS batch file to ftp print commands to an FTP-style label printer
        'The file is re-created if it already exists, to ensure that the default path
        'is part of the command.

        'Also creates the "sendlabel.txt" file that logs in to the ftp printer and
        '"puts" the labelout.txt file into the printer.

        Dim batContents As String = ""
        batContents = "@echo OFF" & vbCrLf & "ftp -s:" & mypath & "sendlabel.txt %1"
        writeFile(mypath & "ftpbat.bat", batContents, False)

        batContents = FTPLogin.Text & vbCrLf & FTPPassword.Text & vbCrLf & "put " & mypath & "labelout.txt /execute/labelout.txt" & vbCrLf & "quit"
        writeFile(mypath & "sendlabel.txt", batContents, False)

        If Not File.Exists(mypath & "viados.bat") Then
            batContents = "REM Sample batch file displays label text" & vbCrLf & "REM in Windows Notepad" & vbCrLf &
            "start notepad.exe label.txt" & vbCrLf
            writeFile(mypath & "viados.bat", batContents, False)
        End If

        '
        'Create getalma.bat to run the java desktop app "almalabelu2.class"
        'Includes parameters: %1 = Alma url; %2 = Alma user name; %3 = Alma password;
        '%4 = institution code; %5 = barcode number; %6 = output directory for returned XML file
        '
        'batContents = "@echo OFF" & vbCrLf & _
        '"java -cp .;" & javaSDKName & " " & javaClassName & " %1 %2 %3 %4 %5 %6"
        'writeFile(mypath & "getalma.bat", batContents, False)

        'batContents = "java -cp .;" & javaSDKName & " " & javaClassName & " %1 %2 %3 %4 %5 %6" & vbCrLf & _
        '"pause"
        'writeFile(mypath & "getalmadiag.bat", batContents, False)

    End Sub
    Private Sub webDownload(ByVal fname As String, ByVal ftype As String, ByVal webPath As String, ByVal fdest As String)
        Dim webrequest As String = ""
        Dim webClient As New System.Net.WebClient
        Dim txt As String
        Dim success As Boolean = True
        Dim knt As Integer = 0
        webrequest = webPath & fname
        Try
            If ftype = "string" Then
                txt = webClient.DownloadString(webrequest)
                txt = txt.Replace(vbCrLf, vbLf).Replace(vbLf, vbCrLf)
                writeFile(fdest & fname, txt, False)
            Else
                webClient.DownloadFile(webrequest, fdest & fname)
            End If
        Catch ex As Exception
            success = False
            If ex.Message.Contains("407") Then
                MsgBox("Your proxy server is not allowing you to connect to the GitHub server:" & vbCrLf & vbCrLf &
                webPath & vbCrLf & vbCrLf &
                "Ask your IT Networking office to allow access to ('whitelist') this server.", MsgBoxStyle.Exclamation, "Proxy Server Block")
            End If
            MsgBox(fname & "- Download error: " & ex.Message)
            Exit Sub
        End Try

    End Sub

    Private Sub downloadAboveLcFile()

        Dim webRequest As String = ""
        Dim webClient As New System.Net.WebClient()
        Dim labelText As String = ""
        Dim msgStyle = Microsoft.VisualBasic.MsgBoxStyle.Exclamation
        Dim result As String = ""

        If radio_useLocal.Checked Then Exit Sub
        If altURL.Text = "" Then
            altURL.BackColor = Color.White
            Exit Sub
        End If
        If Mid$(altURL.Text, altURL.Text.Length, 1) <> "/" Then
            altURL.Text = altURL.Text & "/"
        End If
        webRequest = altURL.Text & "aboveLabel.txt"

        Try
            labelText = webClient.DownloadString(webRequest)
            labelText = labelText.Replace(vbCrLf, vbLf).Replace(vbLf, vbCrLf)

        Catch ex As Exception
            altURL.BackColor = Color.Pink
            msgStyle = Microsoft.VisualBasic.MsgBoxStyle.Exclamation
            MsgBox("ERROR -- The 'aboveLabel.txt' file could not be downloaded from " &
            webRequest &
            vbCrLf & "Reason: " & ex.Message, msgStyle, "AboveLabelText File Download")
            Exit Sub
        End Try
        altURL.BackColor = Color.White
        writeFile(mypath & "aboveLabel.txt", labelText, False)

    End Sub

    Private Sub SaveAboveLC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAboveLC.Click
        downloadAboveLcFile()
        loadLabelText()
    End Sub

    Private Sub CloseSettings() '_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseSettings.Click
        'reduce the size of the user interface box to hide the administrative panels
        Me.Size = New Size(233, 453)
        settingsOpen = False
        lblToggleAdmin.Text = "u"
        InputBox.Focus()
    End Sub

    Private Sub btnBCFontDialog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBCFontDialog.Click
        Dim fntSettings As New FontDialog()
        'fntSettings.ShowDialog()
        If fntSettings.ShowDialog() = Windows.Forms.DialogResult.OK Then
            inBCFontName.Text = fntSettings.Font.Name
            inBCFontSize.Text = fntSettings.Font.Size
            inBCFontWeight.Checked = fntSettings.Font.Bold
        End If
    End Sub

    Private Sub Label47_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label47.Click
        Dim picBox As New PictureBox
        Dim btnmsg As String
        picBox = PictureBox1
        Dim img As New Bitmap(picBox.Width, picBox.Height)
        Dim brush As New System.Drawing.Drawing2D.LinearGradientBrush(New PointF(0, 0), New PointF(img.Width, img.Height), System.Drawing.Color.Beige, System.Drawing.Color.Tan)
        'Dim gr As Graphics = Graphics.FromImage(img)
        Dim gr As Graphics = Graphics.FromImage(PictureBox1.BackgroundImage)
        'gr.FillRectangle(brush, New RectangleF(0, 0, img.Width, img.Height))
        'picBox.BackgroundImage = img
        Dim bcWeight = New FontStyle
        bcWeight = FontStyle.Bold
        Dim bFont As New Font("times new roman", CType(14.0, Single), FontStyle.Bold)
        Dim sFont As New Font("verdana", CType(10.0, Single), FontStyle.Bold)
        Dim tfont As New Font("verdana", CType(8.0, Single), FontStyle.Regular)
        Dim ufont As New Font("verdana", CType(12.0, Single), FontStyle.Bold)
        gr.DrawString("SpineOMatic", bFont, Brushes.Black, 5, 10)
        gr.DrawString("v" & somVersion & "", sFont, Brushes.Black, 120, 20)
        gr.DrawString("Works with Ex Libris' Alma to print spine labels,", tfont, Brushes.Black, 210, 20)
        gr.DrawString("flag slips, or other custom labels to a variety", tfont, Brushes.Black, 210, 32)
        gr.DrawString("of desktop or networked printers.", tfont, Brushes.Black, 210, 44)

        'gr.DrawString("Questions?", sFont, Brushes.Black, 5, 72)
        'gr.DrawString("Send them to: spineomatic-ggroup@bc.edu", tfont, Brushes.Black, 10, 90)

        gr.DrawString("X", ufont, Brushes.Red, img.Width - 20, 0)
        PictureBox1.Visible = True
        Application.DoEvents()
        btnmsg = "View the SpineOMatic " & somVersion & " Wiki"
        btnDocDownload.Text = btnmsg
        btnDocDownload.Visible = True
        btnDocDownload.BringToFront()
        saveTab = TabControl1.SelectedTab
        TabControl1.SelectedTab = TabPage4
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        PictureBox1.Visible = False
        btnDocDownload.Visible = False
        TabControl1.SelectedTab = saveTab
    End Sub

    Private Sub btnDocDownload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDocDownload.Click
        Dim webclient As New System.Net.WebClient()
        Try
            Process.Start("www.github.com/ExLibrisGroup/SpineOMatic/wiki")
        Catch ex As Exception
            MsgBox("Unable to open Spine-O-Matic Wiki" & vbCrLf _
            & "Error: " & ex.Message, MsgBoxStyle.Exclamation, "Connection Error")
        End Try
    End Sub


    Private Sub FTPInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FTPInfo.Click
        If FTPInfo.Text = "Info" Then
            FTPHelp.Visible = True
            FTPInfo.Text = "   X"
            FTPInfo.ForeColor = Color.Red
        Else
            FTPHelp.Visible = False
            FTPInfo.Text = "Info"
            FTPInfo.ForeColor = Color.Blue
        End If
    End Sub

    Private Sub useExlibrisParsing_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles useExlibrisParsing.CheckedChanged
        If useExlibrisParsing.Checked Then
            parsingSource.Enabled = True
            inCallNumSource.Enabled = False
        Else
            parsingSource.Enabled = False
            inCallNumSource.Enabled = True
        End If
    End Sub

    Private Sub chkIncludeHoldings_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeHoldings.CheckedChanged
        If chkIncludeHoldings.Checked Then
            inIssueLevelSource.Enabled = True
        Else
            inIssueLevelSource.Enabled = False
        End If
    End Sub
    Private Sub chkIncludeOther_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeOther.CheckedChanged
        If chkIncludeOther.Checked Then
            inOtherSource.Enabled = True
        Else
            inOtherSource.Enabled = False
        End If
    End Sub

    Private Sub XMLPath_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XMLPath.TextChanged
        If Trim(XMLPath.Text = "") Then
            btnMonitor.Enabled = False
        Else
            btnMonitor.Enabled = True
        End If
    End Sub

    Private Sub btnMonitor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMonitor.Click
        If Trim(XMLPath.Text) = "" Then Exit Sub
        If Directory.Exists(XMLPath.Text) Then
            FileSystemWatcher1.Path = XMLPath.Text
            MsgBox("The directory" & vbCrLf & XMLPath.Text & vbCrLf & "will be monitored for the arrival of new barcode XML files.", MsgBoxStyle.Information, "Directory Monitor")
            btnMonitor.Enabled = False
        Else
            MsgBox("The directory " & vbCrLf & XMLPath.Text _
            & vbCrLf & "does not exist." &
            vbCrLf & "Enter the name of an existing directory.",
            MsgBoxStyle.Exclamation, "Directory Monitor")
            XMLPath.Focus()
        End If
    End Sub

    Private Sub usrname_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles usrname.KeyPress
        If e.KeyChar = ControlChars.Cr Then
            e.Handled = True
            If usrname.Text = "" Then
                Beep()
                usrname.BackColor = Color.Pink
                Exit Sub
            End If
            usrname.BackColor = Color.White
            OutputBox.Text = ""
            usrname.Text = usrname.Text.ToUpper
            InputBox.Focus()
        End If
        If usrname.Text.Length > 7 Then 'if usrname 8 chars or more,
            If Asc(e.KeyChar) <> 8 Then 'allow backspace, even if username = 8 chars
                e.Handled = True 'if not backspace, ignore the character,
                Beep()           'and beep
            End If
        End If

    End Sub

    Private Function obscure(ByVal txtin As String) As String
        Dim res As String = ""
        Dim i, j As Integer
        Dim key As String = Chr(10) & Chr(11) & Chr(12) & Chr(13) & Chr(14) & Chr(15) & Chr(16) & Chr(17)
        For i = 1 To txtin.Length
            j = (key.Length - 1) Mod i
            res = res & Chr(Asc(Mid$(key, j + 1, 1)) Xor Asc(Mid$(txtin, i, 1)))
        Next i
        Return res
    End Function

    Private Sub btn_addALT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_addALT.Click
        If altText.Text <> "" Then
            If altText.Text.Contains("+") And altText.Text.Contains("=") Then
                altText.Text = altText.Text.Replace(" +", "+").Replace("+ ", "+").Replace(" =", "=").Replace("= ", "=")
                altList.Items.Add(altText.Text)
                altText.Text = ""
                madeALTchanges = True
                btn_saveALT.ForeColor = Color.Red
            Else
                Beep()
                MsgBox("Each entry must contain a plus sign (+) and an equal sign (=) following " & vbCrLf &
                "the Library name and the Location name, respectively.", MsgBoxStyle.Exclamation, "Missing + or = Signs")
                altText.Focus()
            End If
        End If
    End Sub

    Private Sub btn_cancelALT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cancelALT.Click
        altText.Text = ""
        altList.Enabled = True
        btn_changeALT.Enabled = False
        btn_deleteALT.Enabled = False
        'btn_cancelALT.Enabled = False
    End Sub

    Private Sub btn_deleteALT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_deleteALT.Click
        altList.Items.RemoveAt(altList.SelectedIndex)
        altText.Text = ""
        altList.Enabled = True
        madeALTchanges = True
        btn_changeALT.Enabled = False
        btn_deleteALT.Enabled = False
        'btn_cancelALT.Enabled = False
        btn_saveALT.ForeColor = Color.Red
    End Sub

    Private Sub btn_changeALT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_changeALT.Click
        Dim saveidx As Integer = altList.SelectedIndex
        If altText.Text <> "" Then
            If altText.Text.Contains("+") And altText.Text.Contains("=") Then
                altList.Enabled = True
                altList.Items.RemoveAt(saveidx)
                altList.Items.Insert(saveidx, altText.Text)
                altText.Text = ""
                madeALTchanges = True
                btn_changeALT.Enabled = False
                btn_deleteALT.Enabled = False
                'btn_cancelALT.Enabled = False
                btn_saveALT.ForeColor = Color.Red
            Else
                Beep()
                MsgBox("Each entry must contain a plus sign (+) and an equal sign (=) following " & vbCrLf &
                "the Library name and the Location name, respectively.", MsgBoxStyle.Exclamation, "Missing + or = Signs")
                altText.Focus()
            End If
        End If

    End Sub

    Private Sub altList_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles altList.DoubleClick
        If radio_useLocal.Checked = False Then Exit Sub
        btn_changeALT.Enabled = True
        btn_deleteALT.Enabled = True
        btn_cancelALT.Enabled = True
        altText.Text = altList.Items(altList.SelectedIndex)
        altList.Enabled = False
    End Sub

    Private Sub updateLabelText(ByVal fname As String)
        Dim filetext As String = ""
        Dim i As Integer = 0

        If altList.Items.Count <= 0 Then Exit Sub
        For i = 0 To altList.Items.Count - 1
            filetext = filetext & altList.Items.Item(i) & vbCrLf
        Next
        writeFile(fname, filetext, False)
        Application.DoEvents()
        loadLabelText()

    End Sub
    Private Sub hideALTedit()
        Label44.Visible = False
        Label54.Visible = False
        btn_addALT.Visible = False
        btn_changeALT.Visible = False
        btn_deleteALT.Visible = False
        btn_cancelALT.Visible = False
        btn_saveALT.Visible = False
        altText.Visible = False
    End Sub
    Private Sub showALTedit()
        Label44.Visible = True
        Label54.Visible = True
        btn_addALT.Visible = True
        btn_changeALT.Visible = True
        btn_deleteALT.Visible = True
        btn_cancelALT.Visible = True
        btn_saveALT.Visible = True
        altText.Visible = True
    End Sub

    Private Sub radio_useLocal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radio_useLocal.CheckedChanged
        If warranty_accepted = False Then Exit Sub
        If radio_useLocal.Checked = True Then
            altURL.BackColor = Color.White
            ALTfile = "myLabelText.txt"
            showALTedit()
        Else
            If madeALTchanges = True Then
                Dim box = MessageBox.Show("Changes to your local label text file have not been saved." & vbCrLf &
                "Do you want to save them now?", "Save Settings", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If box = box.Yes Then
                    btn_saveALT.PerformClick()
                    MsgBox("Changes to your local label text file have been saved.", MsgBoxStyle.Information, "Settings Saved")
                    madeALTchanges = False
                End If
            End If
            ALTfile = "aboveLabel.txt"
            hideALTedit()
        End If
        loadLabelText()
    End Sub

    Private Sub btn_saveALT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_saveALT.Click
        updateLabelText(mypath & ALTfile)
        btn_saveALT.ForeColor = Color.Black
        madeALTchanges = False
    End Sub

    Private Sub chkRequireUser_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRequireUser.CheckedChanged
        If chkRequireUser.Checked Then
            usrname.Enabled = True
            usrname.ForeColor = Color.Blue
            usrname.Text = ""
            OutputBox.Font = New Font("MS Sans Serif", 9, FontStyle.Regular)
            OutputBox.Text = usermessage
            usrname.BackColor = Color.Yellow
            usrname.Focus()
        Else
            usrname.Enabled = False
            usrname.ForeColor = Color.Gray
            usrname.BackColor = Color.White
            usrname.Text = "[none]"
            OutputBox.Text = ""
        End If
    End Sub

    Private Sub station_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles station.TextChanged
        lblStation.Text = station.Text
    End Sub

    Private Sub buildStatRec()

        Dim dt As String = DateTime.Now.ToString("yyyyMMdd", CultureInfo.CurrentCulture)
        'MsgBox("date: " & dt & " to short date string: " & DateTime.Now.ToShortDateString)
        Dim tm As String = DateTime.Now.ToString("HHmmss", CultureInfo.CurrentCulture)
        Dim parse As String = ""
        If useExlibrisParsing.Checked Then parse = "ExL" Else parse = "SoM"
        Try
            statrec = dt & vbTab &
            tm & vbTab &
            station.Text & vbTab &
            usrname.Text & vbTab &
            lastbc & vbTab &
            parse & vbTab &
            cntype & vbTab &
            almaReturnCode & vbTab &
            almaLibrary & vbTab &
            almaLocation
        Catch ex As Exception
            MsgBox("Error writing statistics record --" & ex.Message)
        End Try
        'MsgBox("almaLibrary: " & almaLibrary & " -- almaLocation: " & almaLocation)
        lastbc = ""
        parse = ""
        cntype = ""
        almaReturnCode = ""
        almaLibrary = ""
        almaLocation = ""
    End Sub

    Private Sub writeStat(ByVal result As String)
        'result = P if spine label has been printed, or "S" if the records has been scanned only
        If statrec <> "" Then
            statrec = statrec & vbTab & result
            'MsgBox("writing: " & statrec)
            writeFile(mypath & "activity_log.txt", statrec, True)
            statrec = ""
        End If
    End Sub

    Private Sub btnScan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnScan.Click
        Dim itm As String
        Dim tda As Array
        Dim tmp As Array
        Dim i As Integer
        Dim cnt(5000) As Integer : For i = 0 To 99 : cnt(i) = 0 : Next
        Dim idx As Integer
        Dim sr As StreamReader
        Dim fs As FileStream
        Dim lodate As String
        Dim hidate As String
        Dim hr As String = ""
        Dim dt As String = ""
        Dim dte As DateTime
        Dim fld As Integer = 0
        Dim arg As String = ""
        Dim tot As Integer = 0
        Dim thisrep As Integer = 0
        Dim colheading As String = ""
        Dim trimsort As Array
        Dim infile As String = "activity_log.txt"
        Dim geolist As String = ""
        Dim cntIP As Array
        Dim idxIP As Integer = 0
        Dim quot As String = """"
        Dim cidate As String = ""
        Dim latlong As String = ""

        If Not logView Then
            If (Not File.Exists(mypath & infile)) Then
                Dim sw As StreamWriter = File.CreateText(mypath & "activity_log.txt")
                sw.Close()
            End If

        Else
            infile = "somlog.log"

        End If

        STL.Items.Clear()

        fs = New FileStream(mypath & infile, FileMode.Open)
        sr = New StreamReader(fs)
        itm = sr.ReadLine()
        statsOut.Text = "          " & "SpineOMatic Labeling Activity Report -- From " & fromScan.Value & " to " & toScan.Value & " -- Station: " & station.Text & "~"

        dte = Convert.ToDateTime(fromScan.Value, CultureInfo.CurrentCulture)
        lodate = dte.ToString("yyyyMMdd", CultureInfo.CurrentCulture)
        dte = Convert.ToDateTime(toScan.Value, CultureInfo.CurrentCulture)
        hidate = dte.ToString("yyyyMMdd", CultureInfo.CurrentCulture)

        If radioByUser.Checked Then fld = 3 : colheading = radioByUser.Text.Replace("By ", "")
        If radioByLibrary.Checked Then fld = 8 : colheading = radioByLibrary.Text.Replace("By ", "")
        If radioByLocation.Checked Then fld = 9 : colheading = radioByLocation.Text.Replace("By ", "")
        If radioSearch.Checked Then
            If searchArg.Text = "" Then
                statsOut.Text = statsOut.Text & "No filter" & vbCrLf & vbCrLf
            Else
                statsOut.Text = statsOut.Text & "Filtered by: " & searchArg.Text & vbCrLf & vbCrLf
            End If
            If logView Then
                statsOut.Text = statsOut.Text & "  Date & Time" & vbTab & "IP" & vbTab & vbCrLf & "Hostname" & vbTab & "File" & vbCrLf
            Else
                statsOut.Text = statsOut.Text & "  Date & Time" & vbTab & "User" & vbTab & "Barcode" & vbTab & "Library" & vbTab & "Location" & vbCrLf
            End If
        End If
        Try
            Do While Not itm Is Nothing
                If itm.Length < 20 Then itm = sr.ReadLine() : Continue Do

                If Not logView Then
                    tda = Split(itm, vbTab)
                Else
                    tda = Split("" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab, vbTab)
                    tmp = Split(itm, vbTab)
                    tda(0) = tmp(0)
                    tda(1) = tmp(1)
                    tda(2) = "weblog"
                    tda(3) = tmp(2)
                    tda(7) = tmp(5)
                    tda(8) = tmp(3)
                    tda(9) = tmp(4)
                End If

                If tda(0) < lodate Then itm = sr.ReadLine() : Continue Do
                If Not logView And Not inclScanned.Checked Then
                    If tda(10) = "S" Then
                        itm = sr.ReadLine()
                        Continue Do
                    End If
                End If
                If tda(0) > hidate Then Exit Do
                tot = tot + 1
                If tda(7) <> "OK" Then tda(8) = "***" & tda(7) & "***" : tda(9) = "***" & tda(7) & "***"
                If radioSearch.Checked Then
                    dt = Mid$(tda(0), 5, 2) & "/" & Mid$(tda(0), 7, 2) & "/" & Mid$(tda(0), 1, 4)
                    cidate = Date.ParseExact(dt, "d", CultureInfo.InvariantCulture)
                    hr = Mid$(tda(1), 1, 2) & ":" & Mid$(tda(1), 3, 2) & ":" & Mid$(tda(1), 5, 2)
                    ' If logView And chkAddHostname.Checked Then tda(4) = resolveIP(tda(3))
                    arg = cidate & " " & hr & vbTab & tda(3) & vbTab & tda(4) & vbTab & tda(8) & vbTab & tda(9)
                    If searchArg.Text = "" Then
                        statsOut.Text = statsOut.Text & arg & vbCrLf
                        thisrep = thisrep + 1
                    Else
                        If arg.Contains(searchArg.Text) Then
                            statsOut.Text = statsOut.Text & arg & vbCrLf
                            thisrep = thisrep + 1
                        End If
                    End If
                Else
                    arg = tda(fld)
                    If STL.FindStringExact(arg) <> -1 Then
                        idx = STL.FindStringExact(arg)
                        cnt(idx) = cnt(idx) + 1
                    Else
                        idx = STL.Items.Add(arg)
                        cnt(idx) = 1
                    End If
                End If

                itm = sr.ReadLine()
            Loop
        Catch ex As Exception
            MsgBox("Error in reporting system -- " & ex.Message, MsgBoxStyle.Critical, "Program Error")
        End Try
        sr.Close()
        fs.Close()
        If Not radioSearch.Checked Then thisrep = STL.Items.Count
        If inclScanned.Checked Then
            statsOut.Text = statsOut.Text.Replace("~", vbCrLf & "          Total codes scanned and labels printed in date range: " & tot & " -- Lines in this report: " & thisrep & vbCrLf & vbCrLf)
        Else
            statsOut.Text = statsOut.Text.Replace("~", vbCrLf & "          Total labels printed in date range: " & tot & " -- Lines in this report: " & thisrep & vbCrLf & vbCrLf)
        End If
        If radioSearch.Checked Then Exit Sub

        statsOut.Text = statsOut.Text & vbTab & "Count" & vbTab & colheading & vbCrLf

        sortSTL.Items.Clear()
        sortSTL.Sorted = False
        For i = 0 To STL.Items.Count - 1
            If byCount.Checked Then
                sortSTL.Items.Add(String.Format("{0,6:######}", 999999 - cnt(i)) & "|" & String.Format("{0,6:######}", cnt(i)) & vbTab & STL.Items(i))
            Else
                sortSTL.Items.Add(Mid$(STL.Items(i) & Space(30), 1, 30) & "|" & String.Format("{0,6:######}", cnt(i)) & vbTab & STL.Items(i))
            End If
        Next

        ispList.Items.Clear()
        sortSTL.Sorted = True
        Dim LL As Array
        Dim locText As String = ""
        For i = 0 To sortSTL.Items.Count - 1
            locText = ""
            trimsort = Split(sortSTL.Items(i), "|")
            If chkAddHostname.Checked Then
                cntIP = Split(trimsort(1), vbTab)
                latlong = getLatLong(cntIP(1))
                If latlong = "** can't connect **" Then
                    MsgBox("Can't connect to IP resolver.")
                    Exit Sub
                Else
                    LL = Split(latlong, "|")
                    If LL.Length > 4 Then
                        locText = LL(2) & "/" & LL(3) & "/" & LL(4) & "/" & LL(5)
                    End If
                    Dim tloc As String = ""
                    tloc = LL(5) & " (" & LL(2) & "/" & LL(3) & "/" & LL(4) & ")"
                    If showIsp.Checked And Not ispList.Items.Contains(tloc) Then
                        ispList.Items.Add(tloc)
                    End If
                End If
                geolist = geolist & "field[" & i & "]=" & quot & latlong & quot & ";" & vbCrLf
            End If
            statsOut.Text = statsOut.Text & vbTab & trimsort(1) & vbTab & locText & vbCrLf
        Next

        If chkGeoList.Checked Then
            geolist = geolist & vbCrLf & "document.getElementById(" & quot & "usr" & quot & ").innerHTML = " & quot & "From " & fromScan.Value & " to " & toScan.Value & quot & ";"
            writeFile(mypath & "geoinfo.txt", geolist, False)
            Dim framein As String = ""
            Dim tr As TextReader = New StreamReader(mypath & "geoframe.txt")
            framein = tr.ReadToEnd()
            tr.Close()
            'MsgBox(framein)
            framein = framein.Replace("//insert_geolist", geolist)
            writeFile(mypath & "sommap.html", framein, False)
            Dim startInfo As New ProcessStartInfo(mypath & "sommap.html")
            Dim p As Process
            'startInfo.WindowStyle = ProcessWindowStyle.Hidden 'makes sure entire process is invisible
            'startInfo.Arguments = FTPip.Text 'the IP address is passed to ftpbat.bat as an argument
            p = Process.Start(startInfo) 'tell Windows to start the batch file
        End If
        If showIsp.Checked Then
            statsOut.Text = ""
            For i = 0 To ispList.Items.Count - 1
                statsOut.Text = statsOut.Text & ispList.Items(i) & vbCrLf
            Next
        End If
    End Sub


    Private Sub radio_useSystem_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radio_useSystem.CheckedChanged
        If settingsLoaded = False Then Exit Sub
        If radio_useSystem.Checked Then
            If altURL.Text = "" Then
                MsgBox("To use the system file, provide a URL to the web folder" & vbCrLf &
                "where the 'aboveLabel.txt' file can be found, and click 'Download'.", MsgBoxStyle.Information, "Use System File")
                altURL.BackColor = Color.Yellow
                altURL.Focus()
            End If
        End If
    End Sub


    Private Sub getWebLog()
        Dim webClient As New System.Net.WebClient
        Dim webRequest As String = ""
        Dim result As String = ""
        Try
            webRequest = updatePath.Text & "somlog.log"
            webClient.DownloadFile(webRequest, mypath & "somlog.log")
            OutputBox.Text = OutputBox.Text & vbCrLf & "Web log loaded"
            radioByUser.Text = "By IP"
            radioByLibrary.Text = "By Hostname"
            radioByLocation.Text = "By Filename"
            TabControl1.SelectedTab = TabPage6
            openSettings()
            logView = True
        Catch ex As Exception
            MsgBox("Web log download error: " & ex.Message, MsgBoxStyle.Exclamation, "Download Error")
            Exit Sub
        End Try
    End Sub
    Private Sub logByHostname()
        Dim fs As FileStream
        Dim sr As StreamReader
        Dim itm As String
        Dim i As Integer
        Dim cnt(100) As Integer : For i = 0 To 99 : cnt(i) = 0 : Next
        Dim dte As DateTime
        Dim arg As String
        Dim lodate As String
        Dim hidate As String
        Dim idx As Integer
        Dim fld, trimsort As Array

        statsOut.Text = "          " & "SpineOMatic Web Activity Log -- From " & fromScan.Value & " to " & toScan.Value & " -- Station: " & station.Text & "~" & vbCrLf

        If (Not File.Exists(mypath & "somlog.log")) Then
            MsgBox("No log file has been downloaded.")
            Exit Sub
        End If
        dte = Convert.ToDateTime(fromScan.Value, CultureInfo.InvariantCulture)
        lodate = dte.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
        dte = Convert.ToDateTime(toScan.Value, CultureInfo.InvariantCulture)
        hidate = dte.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
        STL.Items.Clear()

        fs = New FileStream(mypath & "somlog.log", FileMode.Open)
        sr = New StreamReader(fs)
        itm = sr.ReadLine()

        Do While Not itm Is Nothing
            fld = Split(itm, vbTab)
            If fld(0) < lodate Then itm = sr.ReadLine() : Continue Do
            If fld(0) > hidate Then Exit Do
            arg = fld(2) 'hostname
            If STL.FindStringExact(arg) <> -1 Then
                idx = STL.FindStringExact(arg)
                cnt(idx) = cnt(idx) + 1
            Else
                idx = STL.Items.Add(arg)
                cnt(idx) = 1
            End If

            itm = sr.ReadLine()
        Loop

        sr.Close()
        fs.Close()
        sortSTL.Items.Clear()
        sortSTL.Sorted = False
        For i = 0 To STL.Items.Count - 1
            If byCount.Checked Then
                sortSTL.Items.Add(String.Format("{0,6:######}", 999999 - cnt(i)) & "|" & String.Format("{0,6:######}", cnt(i)) & vbTab & STL.Items(i))
            Else
                sortSTL.Items.Add(Mid$(STL.Items(i) & Space(30), 1, 30) & "|" & String.Format("{0,6:######}", cnt(i)) & vbTab & STL.Items(i))
            End If
        Next

        sortSTL.Sorted = True
        For i = 0 To sortSTL.Items.Count - 1
            trimsort = Split(sortSTL.Items(i), "|")
            statsOut.Text = statsOut.Text & vbTab & trimsort(1) & vbCrLf
        Next
    End Sub

    Private Sub LoadDefaults(ByVal t As TextBox)
        Dim val As Array
        val = Split(t.Text, "|")
        inTopMargin.Text = val(0)
        inLeftMargin.Text = val(1)
        inLineSpacing.Text = val(2)
        inLabelWidth.Text = val(3)
        inGapWidth.Text = val(4)
        inLabelHeight.Text = val(5)
        inGapHeight.Text = val(6)
        inLabelRows.Text = val(7)
        inLabelCols.Text = val(8)
        inMaxLines.Text = val(9)
        inMaxChars.Text = val(10)
        usePortrait.Checked = CType(val(11), Boolean)
        useLandscape.Checked = Not CType(val(11), Boolean)
        If CustomLabel.Checked Then
            CustomText.Text = val(12)
            CustomText.Enabled = True
        Else
            CustomText.Enabled = False
        End If
    End Sub
    Private Sub SaveDefaults(ByVal t As TextBox)
        t.Text = inTopMargin.Text
        t.Text = t.Text & "|" & inLeftMargin.Text
        t.Text = t.Text & "|" & inLineSpacing.Text
        t.Text = t.Text & "|" & inLabelWidth.Text
        t.Text = t.Text & "|" & inGapWidth.Text
        t.Text = t.Text & "|" & inLabelHeight.Text
        t.Text = t.Text & "|" & inGapHeight.Text
        t.Text = t.Text & "|" & inLabelRows.Text
        t.Text = t.Text & "|" & inLabelCols.Text
        t.Text = t.Text & "|" & inMaxLines.Text
        t.Text = t.Text & "|" & inMaxChars.Text
        If usePortrait.Checked Then
            t.Text = t.Text & "|" & "true"
        Else
            t.Text = t.Text & "|" & "false"
        End If
        'If CustomLabel.Checked Then
        t.Text = t.Text & "|" & CustomText.Text
        'End If
    End Sub

    Public Function GetJavaVersionInfo() As String
        Dim ermsg As String = ""
        'Const quote As String = """"
        Dim mpath As String = mypath
        Dim noslash As String = Mid$(mypath, 1, mypath.Length - 1)

        Try
            Dim procStartInfo As New System.Diagnostics.ProcessStartInfo("java", "-cp " & """" & noslash & """" & " " & javaTest)

            procStartInfo.RedirectStandardOutput = True
            procStartInfo.RedirectStandardError = True
            procStartInfo.UseShellExecute = False
            procStartInfo.CreateNoWindow = True
            Dim proc As System.Diagnostics.Process = New Process()
            proc.StartInfo = procStartInfo
            proc.Start()
            'Return proc.StandardError.ReadToEnd()
            ermsg = proc.StandardError.ReadToEnd()
            If ermsg <> "" Then
                If ermsg.Contains("find or load") Then
                    Return "The Java test program (" & javaTest & ") was not downloaded." & vbCrLf & "Unable to perform test."
                Else
                    Return "Unknown error: " & ermsg
                End If
            Else
                Return proc.StandardOutput.ReadToEnd().Replace(vbCr, vbCrLf)
            End If
        Catch ex As Exception
            If ex.Message.Contains("cannot find the file") Then
                Return "Java is not installed, or is not accessible." & vbCrLf & vbCrLf &
                "Please report this problem to your local systems support staff."
            End If
        End Try
        Return "unknown error"
    End Function

    Private Sub tips_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tips.CheckedChanged
        If tips.Checked Then
            ToolTip1.Active = True
        Else
            ToolTip1.Active = False
        End If

    End Sub


    Private Function getLatLong2(ByVal ip As String) As String
        Dim webclient As New System.Net.WebClient
        Dim locxml As String = ""
        Dim hn As String = searchArg.Text
        'Dim host As IPHostEntry
        Dim xdoc As New System.Xml.XmlDocument
        Dim xfield As String = ""
        Static Dim cnissue As String = ""
        'host = Dns.GetHostEntry(hn)
        If cnissue = "x" Then Return "** can't connect **"
        Try
            locxml = webclient.DownloadString("http://freegeoip.net/xml/" & ip)
            xdoc.LoadXml(locxml)
            xfield = xdoc.SelectSingleNode("//Response/Latitude").InnerText
            xfield = xfield & "|" & xdoc.SelectSingleNode("//Response/Longitude").InnerText
            xfield = xfield & "|" & xdoc.SelectSingleNode("//Response/CountryName").InnerText
            xfield = xfield & "|" & xdoc.SelectSingleNode("//Response/RegionName").InnerText
            xfield = xfield & "|" & xdoc.SelectSingleNode("//Response/City").InnerText
        Catch ex As Exception
            'MsgBox("Unable to connect to server")
            cnissue = "x"
            Return "** can't connect **"
        End Try
        'MsgBox(host.HostName & vbCrLf & vbCrLf & locxml)
        Return xfield
    End Function
    Private Function getLatLong(ByVal ip As String) As String
        Dim webclient As New System.Net.WebClient
        Dim locxml As String = ""
        Dim hn As String = searchArg.Text
        'Dim host As IPHostEntry
        Dim xdoc As New System.Xml.XmlDocument
        Dim xfield As String = ""
        Static Dim cnissue As String = ""
        Dim ans As String = ""
        'host = Dns.GetHostEntry(hn)
        If cnissue = "x" Then Return "** can't connect **"
        Try
            locxml = webclient.DownloadString("http://www.telize.com/geoip/" & ip)
            locxml = locxml.Replace("{", "").Replace("}", ",")

            xfield = getjson("latitude", locxml) & "|" & getjson("longitude", locxml) &
            "|" & getjson("country", locxml) & "|" & getjson("region", locxml) & "|" &
            getjson("city", locxml) & "|" & getjson("isp", locxml)
        Catch ex As Exception
            MsgBox("Unable to connect to server" & vbCrLf & "Error: " & ex.Message)
            cnissue = "x"
            Return "** can't connect **"
        End Try
        'MsgBox(host.HostName & vbCrLf & vbCrLf & locxml)

        Return xfield
    End Function

    Private Function getjson(ByVal lookup As String, ByVal json As String)
        Dim quot = """"
        Dim pos As Integer = 0
        Dim ans As String = ""
        pos = InStr(json, lookup & quot, CompareMethod.Text)

        If pos <> 0 Then
            ans = BTween(json.Substring(pos), ":", ",")
            Return ans.Replace(quot, "")
        End If
        Return "unknown"
    End Function
    Private Function BTween(ByVal t As String, ByVal a As String, ByVal b As String) As String
        Dim pos1, pos2 As Integer
        Dim len_a As Integer = a.Length
        pos1 = t.IndexOf(a) + len_a '1
        pos2 = t.IndexOf(b, pos1)
        If pos1 = -1 Or pos2 = -1 Then Return "?"
        Return t.Substring(pos1, pos2 - pos1)
    End Function
    Private Function xmlValue(ByVal node_in, Optional okToSkip = False) As String
        Dim xfield As String = ""
        Dim prefix As String = ""
        Dim orig_node As String = node_in.ToString
        Dim node As String = ""
        Dim er As String = ""
        Dim text_node As String
        prefix = "//printout/section-01/physical_item_display_for_printing"
        text_node = orig_node
        node = prefix & orig_node.Replace("<", "/").Replace(">", "")
        Try
            xfield = xdoc.SelectSingleNode(node).InnerText
        Catch ex As Exception
            If Not node.Contains("parsed_") Then 'call_number") Then

                If ex.Message.Contains("Object reference not set") And okToSkip Then
                    'Return empty string; it is ok if this element doesn't exist in the XML.
                    Return ""
                End If
                Beep()
                If ex.Message.Contains("Object reference not set") Then
                    er = orig_node & " is not in the XML record."
                Else
                    If ex.Message.Contains("invalid token") Then
                        er = orig_node & " contains an unexpected character."
                    Else
                        er = "Error in " & orig_node & ": " & vbCrLf & vbCrLf & ex.Message
                    End If
                End If

                If chkXMLWarning.Checked Then
                    MsgBox(er & vbCrLf & vbCrLf &
                    "Check and correct the value.", MsgBoxStyle.Exclamation, "Invalid XML Field Name")
                Else
                    lblXMLWarn.Visible = True
                    xmlerr = xmlerr & er & vbCrLf
                End If
            End If
            xfield = ""
        End Try
        'MsgBox("xfield:" & xfield)
        'MsgBox("xfield returned:" & xfield.Replace("|amp", "&").Replace("|lt;", "<"))
        Return xfield.Replace("|amp", "&").Replace("&amp;", "&").Replace("|lt;", "<")
    End Function
    Private Sub getnodes()
        For Each node As XmlNode In xdoc.SelectNodes("/printout/section-01/physical_item_display_for_printing/*")

        Next
    End Sub
    Private Sub Spine_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Spine.CheckedChanged
        If Not settingsLoaded Then Exit Sub
        If Spine.Checked Then
            saveCurrentDefaults()
            LoadDefaults(spineDefaults)
            pocketDefaultLoaded = False
            spineDefaultLoaded = True
            customFlagDefaultLoaded = False
            customNonFlagDefaultLoaded = False
        End If
    End Sub

    Private Sub CustomLabel_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomLabel.CheckedChanged
        If Not settingsLoaded Then Exit Sub
        If CustomLabel.Checked Then
            FlagSlips.Enabled = True
            saveCurrentDefaults()
            If FlagSlips.Checked Then
                LoadDefaults(flagDefaults)
                wrapWidth.Text = flagWrapWidth.Text
                pocketDefaultLoaded = False
                spineDefaultLoaded = False
                customFlagDefaultLoaded = True
                customNonFlagDefaultLoaded = False
            Else
                LoadDefaults(nonFlagDefaults)
                wrapWidth.Text = nonFlagWrapWidth.Text
                pocketDefaultLoaded = False
                spineDefaultLoaded = False
                customFlagDefaultLoaded = False
                customNonFlagDefaultLoaded = True
            End If
        Else
            FlagSlips.Enabled = False
        End If
    End Sub

    Private Sub FlagSlips_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FlagSlips.CheckedChanged
        If Not settingsLoaded Then Exit Sub
        saveCurrentDefaults()
        If FlagSlips.Checked Then
            LoadDefaults(flagDefaults)
            wrapWidth.Text = flagWrapWidth.Text
            pocketDefaultLoaded = False
            spineDefaultLoaded = False
            customFlagDefaultLoaded = True
            customNonFlagDefaultLoaded = False
        Else
            LoadDefaults(nonFlagDefaults)
            wrapWidth.Text = nonFlagWrapWidth.Text
            pocketDefaultLoaded = False
            spineDefaultLoaded = False
            customFlagDefaultLoaded = False
            customNonFlagDefaultLoaded = True
        End If
    End Sub
    Private Sub chkUsePocketLabels_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUsePocketLabels.CheckedChanged
        If Not settingsLoaded Then Exit Sub
        If chkUsePocketLabels.Checked Then
            PocketLabelPanel.Location = New Point(0, 30)
            PocketLabelPanel.Width = TabPage2.Size.Width
            PocketLabelPanel.Height = TabPage2.Size.Height - 30
            PocketLabelPanel.Visible = True

            saveCurrentDefaults()
            LoadDefaults(pocketDefaults)
            pocketDefaultLoaded = True
            spineDefaultLoaded = False
            customFlagDefaultLoaded = False
            customNonFlagDefaultLoaded = False

        Else
            PocketLabelPanel.Visible = False
            'save pocket label defaults & load whatever was previously saved
            saveCurrentDefaults()
            If Spine.Checked Then
                LoadDefaults(spineDefaults)
                pocketDefaultLoaded = False
                spineDefaultLoaded = True
                customFlagDefaultLoaded = False
                customNonFlagDefaultLoaded = False
            Else
                CustomLabel.Checked = Not CustomLabel.Checked
                CustomLabel.Checked = Not CustomLabel.Checked
            End If
        End If
        refreshFTP()
        InputBox.Focus()
    End Sub

    Private Sub Label71_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label71.Click
        tweakParsingPanel.Visible = False
    End Sub

    Private Sub tweakTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tweakTest.Click
        tweakParsingPanel.Location = New Point(0, 0)
        tweakParsingPanel.Size = TabPage2.Size
        tweakParsingPanel.Visible = True
        testCallNumType.Text = "0"
    End Sub

    Private Sub saveCurrentDefaults()
        'save the margin defaults for the Call Number Format that is currently loaded
        If spineDefaultLoaded Then
            SaveDefaults(spineDefaults)
            'spineVerticalLine = chk_VerticalLine.Checked
        Else
            If customFlagDefaultLoaded Then
                SaveDefaults(flagDefaults)
                flagWrapWidth.Text = wrapWidth.Text
                'flagVerticalLine = chk_VerticalLine.Checked
            Else
                If customNonFlagDefaultLoaded Then
                    SaveDefaults(nonFlagDefaults)
                    nonFlagWrapWidth.Text = wrapWidth.Text
                    'nonFlagVerticalLine = chk_VerticalLine.Checked
                Else
                    SaveDefaults(pocketDefaults)
                    'pocketVerticalLine = chk_VerticalLine.Checked
                End If
            End If
        End If
    End Sub

    Private Sub btnTestParser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTestParser.Click
        If Not settingsLoaded Then Exit Sub
        Dim itm As String = ""
        itm = testComboBox.Text
        If testCallNumType.Text = "" Then
            MsgBox("You must specify a Call Number Type", MsgBoxStyle.Information, "No Call Number Type")
            Exit Sub
        End If
        cntype = testCallNumType.Text
        If itm.Contains("[LC]") Then itm = itm.Replace("[LC]", "")
        If itm.Contains("[LC child. lit]") Then itm = itm.Replace("[LC child. lit]", "")
        If itm.Contains("[NLM]") Then itm = itm.Replace("[NLM]", "")
        If itm.Contains("[Dewey]") Then itm = itm.Replace("[Dewey]", "")
        If itm.Contains("[SuDoc]") Then itm = itm.Replace("[SuDoc]", "")
        If itm.Contains("[Other]") Then itm = itm.Replace("[Other]", "")
        OutputBox.Font = New Font(inFontName.Text, CType(inFontSize.Text, Single), FontStyle.Regular)
        OutputBox.Text = parseLC("*" & itm)

    End Sub

    Private Sub testComboBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles testComboBox.SelectedIndexChanged
        Dim itm As String = ""
        itm = testComboBox.Items(testComboBox.SelectedIndex)
        If itm.Contains("[LC]") Then testCallNumType.Text = "0"
        If itm.Contains("[LC child. lit]") Then testCallNumType.Text = "0"
        If itm.Contains("[Dewey]") Then testCallNumType.Text = "1"
        If itm.Contains("[NLM]") Then testCallNumType.Text = "2"
        If itm.Contains("[SuDoc]") Then testCallNumType.Text = "3"
        If itm.Contains("[Other]") Then testCallNumType.Text = "4"
        OutputBox.Text = ""
        btnTestParser.Focus()
    End Sub

    Private Sub lcNumericBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lcNumericBreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub lcDecimalBreakB4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lcDecimalBreakB4.CheckedChanged
        btnTestParser.PerformClick()
    End Sub
    Private Sub lcNoDecimalBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lcNoDecimalBreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub lcCutterBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lcCutterBreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub lcInCutterBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lcInCutterBreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub hideCutterDecimal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles hideCutterDecimal.CheckedChanged
        btnTestParser.PerformClick()
    End Sub


    Private Sub testCallNumType_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles testCallNumType.TextChanged
        If Not settingsLoaded Then Exit Sub
        Dim tct As String = ""
        If Trim(testCallNumType.Text) = "" Then
            tct = convertBlankTo.Text '"blank"
            testCallNumType.Text = tct
        Else
            tct = testCallNumType.Text
        End If

        If lcType.Text.Contains(tct) Then
            TabControl2.SelectedTab = TabPage7
        Else
            If DeweyType.Text.Contains(tct) Then
                TabControl2.SelectedTab = TabPage8
            Else
                If sudocType.Text.Contains(tct) Then
                    TabControl2.SelectedTab = TabPage10
                Else
                    If otherType.Text.Contains(tct) Then
                        TabControl2.SelectedTab = TabPage11
                    Else
                        MsgBox("Call number type " & testCallNumType.Text & " is not handled by any parsing routine.")
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub sudocBreakDecimal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        btnTestParser.PerformClick()
    End Sub

    Private Sub sudocBreakBeforeColon_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        btnTestParser.PerformClick()
    End Sub

    Private Sub sudocBreakAfterColon_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        btnTestParser.PerformClick()
    End Sub

    Private Sub sudocSlashBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        btnTestParser.PerformClick()
    End Sub

    Private Sub sudocDashBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        btnTestParser.PerformClick()
    End Sub
    Private Sub sudocOtherNoBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sudocOtherNoBreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub sudocCharBreakB4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sudocCharBreakB4.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub sudocRemoveAfter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sudocRemoveAfter.CheckedChanged
        btnTestParser.PerformClick()
    End Sub
    Private Sub deweyDecimalBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deweyDecimalBreakb4.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub deweyCutterBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        btnTestParser.PerformClick()
    End Sub

    Private Sub deweyInCutterBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deweyCutterBreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub TabControl2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl2.SelectedIndexChanged
        Exit Sub
        If TabControl2.SelectedIndex = 0 Then
            testCallNumType.Text = "0"
        End If
        If TabControl2.SelectedIndex = 1 Then
            testCallNumType.Text = "1"
        End If
        If TabControl2.SelectedIndex = 2 Then
            testCallNumType.Text = "3"
        End If
        If TabControl2.SelectedIndex = 3 Then
            testCallNumType.Text = "4"
        End If
    End Sub

    Private Sub otherSpaceNoBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles otherSpaceNoBreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub otherAllSpaceBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles otherAllSpaceBreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub othernumnobreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles othernumnobreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub otherListNoBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles otherListNoBreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub otherListBreakb4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles otherListBreakb4.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub othernumbreakb4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles otherNumBreakb4.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub deweyDecimalNoBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deweyDecimalNoBreak.CheckedChanged
        'btnTestParser.PerformClick()
        If deweyDecimalNoBreak.Checked Then
            deweyDecBreak.Enabled = True
            deweyDigitsToBreak.Enabled = True
        Else
            deweyDecBreak.Enabled = False
            deweyDigitsToBreak.Enabled = False
        End If
        btnTestParser.PerformClick()
    End Sub

    Private Sub deweyCharBreakb4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deweyCharBreakb4.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub deweyOtherNoBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deweyOtherNoBreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub deweyPrefixBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deweyPrefixBreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub sudocNoDecimalBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sudocNoDecimalBreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub sudocNoColonBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sudocNoColonBreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub sudocBreakBeforeColon_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sudocBreakBeforeColon.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub sudocSlashNobreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sudocSlashNobreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub sudocDashNoBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sudocDashNoBreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub lcOtherNoBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lcOtherNoBreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub lcOtherBreakB4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lcOtherBreakB4.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub lcRemoveAfter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lcRemoveAfter.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub deweyRemoveAfter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deweyRemoveAfter.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub otherRemoveAfter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles otherRemoveAfter.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub sudocBreakB4Numerics_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sudocBreakB4Numerics.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub deweyGroup3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deweyGroup3.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub deweyDecBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deweyDecBreak.CheckedChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub LabelRepeat_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelRepeat.ValueChanged
        If LabelRepeat.Value > 1 And LabelRepeat.Value < 6 Then
            LabelRepeat.BackColor = Color.Pink
            LabelRepeat.ForeColor = Color.Black
        Else
            If LabelRepeat.Value > 5 Then
                LabelRepeat.BackColor = Color.Red
                LabelRepeat.ForeColor = Color.White
            Else
                LabelRepeat.BackColor = Color.PaleGreen
                LabelRepeat.ForeColor = Color.Black
            End If
        End If
        InputBox.Focus()
    End Sub

    Private Sub lblXMLWarn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblXMLWarn.Click
        Dim saveOutputBox As String = ""
        OutputBox.Text = xmlerr
    End Sub

    Private Sub lblToggleAdmin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblToggleAdmin.Click
        If lblToggleAdmin.Text = "u" Then
            openSettings()
        Else
            CloseSettings()
        End If
    End Sub

    Private Sub TabChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 0 Then
            If chkUsePocketLabels.Checked Then
                showLabelType.Text = "Pocket Labels"
            Else
                If Spine.Checked Then
                    showLabelType.Text = "Spine Labels"
                Else
                    If FlagSlips.Checked Then
                        showLabelType.Text = "Custom/Flag Slips"
                    Else
                        showLabelType.Text = "Custom Labels"
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub closeFormatInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles closeFormatInfo.Click
        formatInfoPanel.Visible = False
        formatInfoPanel.SendToBack()
    End Sub

    Private Sub showFormatInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles showFormatInfo.Click
        formatInfoPanel.Visible = True
        formatInfoPanel.BringToFront()
        formatInfoPanel.Focus()
    End Sub

    Private Sub btnPlCustom_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlCustom.CheckedChanged
        If btnPlCustom.Checked Then
            userDefinedPanel.Enabled = True
        Else
            userDefinedPanel.Enabled = False
        End If
    End Sub

    Private Function pingServer(ByVal server As String) As Boolean

        'MsgBox(TabControl2.SelectedTab.Text)
        'Dim p As New NetworkInformation.Ping()
        'Dim rtn As NetworkInformation.PingReply = p.Send("www.uno.edu", 3000)

        'If rtn.RoundtripTime = 0 Then
        '    MsgBox("server unavailable" & rtn.RoundtripTime & _
        '    vbCrLf & rtn.Status)

        'Else
        '    MsgBox("Time: " & rtn.RoundtripTime)
        'End If
        If My.Computer.Network.IsAvailable Then
            Try
                If My.Computer.Network.Ping(server, 5000) Then
                    pingServer = True
                Else
                    pingServer = False
                End If
            Catch ex As Exception
                pingServer = False
            End Try
        End If
    End Function

    Private Sub ColonBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ColonBreak.CheckedChanged
        If ColonBreak.Checked Then
            ProtectColon.Enabled = True
        Else
            ProtectColon.Enabled = False
        End If
    End Sub
    Private Function bcyaInfo() As String
        Dim strHostName As String = ""
        Dim strIPAddress As String = ""
        strHostName = System.Net.Dns.GetHostName()
        strIPAddress = System.Net.Dns.GetHostEntry(strHostName).AddressList(0).ToString()
        Return "Host Name: " & strHostName & "; IP Address: " & strIPAddress

    End Function
    Private Function ipconfig() As String
        Dim p As New Process
        Dim i As Integer = 0
        Dim ipinfo As String = ""
        With p.StartInfo
            .FileName = "ipconfig.exe"
            .CreateNoWindow = True
            .RedirectStandardOutput = True
            .RedirectStandardError = True
            .UseShellExecute = False
        End With
        Try
            p.Start()
            If p.WaitForExit(1500) Then
                ipconfig = p.StandardOutput.ReadToEnd
                configText.Text = ipconfig
                For i = 0 To configText.Lines.Length - 1
                    If configText.Lines(i).ToUpper.Contains("IPV4") Then
                        ipinfo = ipinfo & configText.Lines(i) & vbCrLf
                    End If
                Next
                ipconfig = ipinfo
            Else
                ipconfig = "No IPv4 Returned"
            End If
        Catch ex As Exception
            ipconfig = "Cannot run ipconfig"
        End Try
    End Function

    Private Sub plDistance_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles plDistance.TextChanged
        plDistance.BackColor = Color.White
    End Sub

    Private Sub Label62_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_setclipboard.Click
        If Trim(statsOut.Text) = "" Then
            MsgBox("There is no report text to copy to the clipboard.")
            Exit Sub
        End If

        Try
            Clipboard.SetText(statsOut.Text)
        Catch ex As Exception
            MsgBox("An error occurred when copying text to the Windows clipboard." & vbCrLf & vbCrLf &
            "Error: " & ex.Message, MsgBoxStyle.Exclamation, "Clipboard Error")
            Exit Sub
        End Try
        copyDone.Text = "...copied"
    End Sub

    Private Sub statsOut_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles statsOut.TextChanged
        copyDone.Text = ""
    End Sub

    Private Sub lbl_copyXMLtext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_copyXMLtext.Click
        Dim copyXmlHdr As String = ""
        Dim settingsText As String = ""

        If includeSettings.Checked Then
            Dim settingsFile As String = mypath & "settings.som"
            Try
                Dim tr As TextReader = New StreamReader(settingsFile)
                settingsText = tr.ReadToEnd
                settingsText = "*** Settings ***" & vbCrLf & vbCrLf & settingsText
                tr.Close()
            Catch somEx As Exception
                MsgBox("Error reading 'settings.som' file." & vbCrLf & vbCrLf &
                "Error: " & somEx.Message, MsgBoxStyle.Exclamation, "Settings File Read Error")
                settingsText = vbCrLf & vbCrLf & "************** Error reading settings file ****************"
            End Try
        Else
            settingsText = "Settings not requested."
        End If

        If Trim(RichTextBox1.Text) = "" Then
            MsgBox("There is no text to copy to the clipboard.")
            Exit Sub
        End If
        copyXmlHdr = bcyaInfo() & vbCrLf & Date.Now & vbCrLf & vbCrLf
        Try
            Clipboard.SetText(copyXmlHdr & RichTextBox1.Text.Replace(vbLf, vbCrLf) & vbCrLf & vbCrLf & settingsText)
        Catch ex As Exception
            MsgBox("An error occurred when copying text to the Windows clipboard." & vbCrLf & vbCrLf &
            "Error: " & ex.Message, MsgBoxStyle.Exclamation, "Clipboard Error")
            Exit Sub
        End Try
        xmlCopyDone.Text = "...copied"
    End Sub

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged
        xmlCopyDone.Text = ""
    End Sub

    Private Sub includeSettings_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles includeSettings.CheckedChanged
        xmlCopyDone.Text = ""
    End Sub

    Private Function getHostName(ByVal ip As String) As String
        Dim hn As String
        Try
            hn = System.Net.Dns.GetHostEntry(ip).HostName.ToString
        Catch ex As Exception
            hn = "(unknown host)"
        End Try
        Return hn
    End Function

    Private Function resolveIP(ByVal ip) As String
        'OutputBox.Text = getHostName(InputBox.Text)
        Dim iptext As String = ""
        Dim res As String = ""
        Dim pos As Integer = 0
        Dim pos2 As Integer = 0
        If File.Exists(mypath & "resolve.txt") Then
            File.Delete(mypath & "resolve.txt")
        End If
        Thread.Sleep(250)
        'lookup(InputBox.Text)
        lookup(ip)
        Thread.Sleep(550)
        iptext = IO.File.ReadAllText(mypath & "resolve.txt")
        pos = InStr(iptext, "Name:", CompareMethod.Text) + 5
        If pos < 6 Then
            res = "Host Unknown"
        Else
            pos2 = InStr(pos + 1, iptext, "Address", CompareMethod.Binary)
            res = Trim(Mid$(iptext, pos, pos2 - pos))
        End If
        'MsgBox(res)
        Return res
    End Function
    Private Sub lookup(ByVal ip As String)
        Dim procInfo As New ProcessStartInfo()
        'procInfo.WindowStyle = ProcessWindowStyle.Hidden
        procInfo.UseShellExecute = True
        procInfo.FileName = "c:\windows\system32\cmd.exe"
        'MsgBox("ip: " & ip)
        'MsgBox("/c nslookup " & ip & " > resolve.txt")

        procInfo.Arguments = "/c nslookup " & ip & " > resolve.txt"
        procInfo.WorkingDirectory = mypath

        procInfo.Verb = "runas"
        Process.Start(procInfo)
    End Sub

    Private Sub FTPLogin_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FTPLogin.TextChanged
        If settingsLoaded Then
            ftpArrow.Visible = True
            btn_ftpRegister.Visible = True
            ftpRegisterMsg.Visible = True
        End If
    End Sub

    Private Sub btn_ftpRegister_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ftpRegister.Click
        createBatFiles()
        ftpArrow.Visible = False
        btn_ftpRegister.Visible = False
        ftpRegisterMsg.Visible = False
    End Sub

    Private Sub FTPPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FTPPassword.TextChanged
        If settingsLoaded Then
            ftpArrow.Visible = True
            btn_ftpRegister.Visible = True
            ftpRegisterMsg.Visible = True
        End If
    End Sub

    Private Sub useDOSBatch_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles useDOSBatch.CheckStateChanged
        If useDOSBatch.Checked Then
            viaDOSPanel.Location = New Point(4, 16) 'DesktopGroup.Location
            viaDOSPanel.Size = New Size(408, 312) 'DesktopGroup.Size
            Label35.Visible = False
            UseDesktop.Visible = False
            UseLaser.Visible = False
            UseFTP.Visible = False
            orientationPanel.Visible = False
            marginPanel.Visible = False
            ManualPrint.Text = "Run 'viados' batch file"
            ToolTip1.SetToolTip(ManualPrint, "Writes label text to 'label.txt' file, then starts 'viados.bat' batch file.")
            ManualPrint.ForeColor = Color.Red
            viaDOSPanel.Visible = True
            batchDisplay.Text = getdosbatchfile("viados")
            batsignal.ForeColor = Color.Green
            LabelRepeat.Visible = False
        Else
            viaDOSPanel.Visible = False
            marginPanel.Visible = True
            orientationPanel.Visible = True
            Label35.Visible = False
            UseDesktop.Visible = True
            UseLaser.Visible = True
            UseFTP.Visible = True
            LabelRepeat.Visible = True
            SetPrintButtonText()
            refreshFTP()
        End If
        InputBox.Focus()
    End Sub
    Private Function getdosbatchfile(ByVal batfile)
        Dim dostext As String = ""
        Dim res As String = ""
        If Not File.Exists(mypath & batfile & ".bat") Then
            Return "" 'Create your own " & batfile & ".bat batch file."
        End If
        res = IO.File.ReadAllText(mypath & batfile & ".bat")
        res = res.Substring(0, res.Length - 2)
        Return res
    End Function

    Private Sub viadosSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles viadosSave.Click
        writeFile(mypath & "viados.bat", batchDisplay.Text, False)
        batsignal.ForeColor = Color.Green
    End Sub

    Private Sub batchDisplay_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles batchDisplay.TextChanged
        If batsignal.ForeColor = Color.Green Then
            batsignal.ForeColor = Color.Red
        End If
    End Sub

    Private Sub loadViados_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles loadViados.Click
        batchDisplay.Text = getdosbatchfile("viados")
        batsignal.ForeColor = Color.Green
    End Sub


    Private Sub unitINCH_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles unitINCH.CheckedChanged
        Dim ut As String = ""
        Dim clr As New Color
        If unitCM.Checked Then
            ut = "cm." : clr = Color.Maroon
        Else
            ut = "in." : clr = Color.Blue
        End If
        inUnits1.Text = ut : inUnits1.ForeColor = clr
        inUnits2.Text = ut : inUnits2.ForeColor = clr
        Label121.Text = ut : Label121.ForeColor = clr
        plUnits1.Text = ut : plUnits1.ForeColor = clr
        plUnits2.Text = ut : plUnits2.ForeColor = clr
        InputBox.Focus()
    End Sub

    Private Sub decimalDOT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles decimalDOT.CheckedChanged
        Dim pFrom As String = ""
        Dim pTo As String = ""

        If decimalDOT.Checked Then
            pFrom = "," : pTo = "."
        Else
            pFrom = "." : pTo = ","
        End If
        inTopMargin.Text = inTopMargin.Text.Replace(pFrom, pTo)
        inLeftMargin.Text = inLeftMargin.Text.Replace(pFrom, pTo)
        inLineSpacing.Text = inLineSpacing.Text.Replace(pFrom, pTo)
        inLabelHeight.Text = inLabelHeight.Text.Replace(pFrom, pTo)
        inLabelWidth.Text = inLabelWidth.Text.Replace(pFrom, pTo)
        inGapHeight.Text = inGapHeight.Text.Replace(pFrom, pTo)
        inGapWidth.Text = inGapWidth.Text.Replace(pFrom, pTo)
        plLeftMargin.Text = plLeftMargin.Text.Replace(pFrom, pTo)
        plDistance.Text = plDistance.Text.Replace(pFrom, pTo)
        wrapWidth.Text = wrapWidth.Text.Replace(pFrom, pTo)
        InputBox.Focus()
    End Sub

    Private Sub deweyDecimalBreakAft_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deweyDecimalBreakAft.CheckedChanged
        If deweyDecimalBreakAft.Checked Then
            deweyGroup3.Enabled = True
            deweydigitsperline.Enabled = True
            Label25.Enabled = True
        Else
            deweyGroup3.Enabled = False
            deweydigitsperline.Enabled = False
            Label25.Enabled = False
        End If
    End Sub

    Private Sub closeXbox_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles closeXbox.Click
        xboxPanel.Visible = False
        xtb.BackColor = xtbOrigColor
    End Sub

    Private Sub xmlLoad(ByVal sender As Object, ByVal e As System.EventArgs) Handles inOtherSource.DoubleClick, CustomText.DoubleClick, TextBox14.DoubleClick, TextBox13.DoubleClick, parsingSource.DoubleClick, inIssueLevelSource.DoubleClick, inCallNumSource.DoubleClick, plSrc1.DoubleClick, plSrc4.DoubleClick, plSrc3.DoubleClick, plSrc2.DoubleClick, dosParam1.DoubleClick, dosParam2.DoubleClick
        Dim node As XmlNode

        XBOX.Items.Clear()
        xboxPanel.Height = ManualPrint.Location.Y - 30
        xboxPanel.Width = TabControl1.Location.X - 4
        For Each node In xdoc.SelectNodes("/printout/section-01/physical_item_display_for_printing/*")
            XBOX.Items.Add("<" & node.Name & ">")
            If node.HasChildNodes Then
                For Each child As XmlNode In node
                    If child.Name <> "#text" Then XBOX.Items.Add("<" & node.Name & "><" & child.Name & ">")
                Next
            End If
        Next
        If XBOX.Items.Count <> 0 Then
            If Not xtb Is Nothing AndAlso xtb.BackColor <> Color.White Then
                xtb.BackColor = Color.White
            End If
            xtb = sender
            xtbOrigColor = xtb.BackColor
            xtb.BackColor = Color.AntiqueWhite
            xboxPanel.Visible = True
        Else
            MsgBox("XML fields can only be selected after a barcode has been scanned.", MsgBoxStyle.Information, "Select XML field")
        End If
    End Sub

    Private Sub XBOX_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles XBOX.DoubleClick
        Dim xf As String = ""
        xf = XBOX.SelectedItem
        If xtb.Multiline = False Then
            xtb.Text = xf
        Else
            If xtb.SelectionLength = 0 Then
                xtb.SelectedText = xf & vbCrLf
            Else
                xtb.SelectedText = xf
            End If
        End If
        xtb.BackColor = xtbOrigColor
        xboxPanel.Visible = False
    End Sub

    Private Sub Label126_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles viadosCaution.Click
        Dim hmsg As String = ""
        hmsg = "When running in hidden mode, the Command Prompt window will not be visible.  "
        hmsg = hmsg & "Make sure that there are no commands in the batch file (such as 'pause') "
        hmsg = hmsg & "that will require a keypress or other response from the user.  Otherwise, the "
        hmsg = hmsg & "batch process will never close." & vbCrLf & vbCrLf
        hmsg = hmsg & "If this happens, use the Task Manager to kill any stray 'cmd.exe' processes."

        MsgBox(hmsg, MsgBoxStyle.Information, "Run Hidden Warning")
    End Sub

    Private Sub hideDosWindow_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles hideDosWindow.CheckedChanged
        If hideDosWindow.Checked Then
            viadosCaution.Visible = True
        Else
            viadosCaution.Visible = False
        End If
    End Sub

    Private Sub chkAddHostname_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAddHostname.CheckedChanged
        If chkAddHostname.Checked Then
            chkGeoList.Enabled = True
        Else
            chkGeoList.Enabled = False
        End If
    End Sub

    Private Sub deweyDigitsToBreak_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deweyDigitsToBreak.TextChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub deweydigitsperline_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deweydigitsperline.TextChanged
        btnTestParser.PerformClick()
    End Sub

    Private Sub UseRestfulApi_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UseRestfulApi.CheckedChanged
        If UseRestfulApi.Checked Then
            accessType.Text = "R"
            accessType.ForeColor = Color.Green
            ToolTip1.SetToolTip(accessType, "Using preferred RESTful API access to Alma.")
        Else
            accessType.Text = "J"
            accessType.ForeColor = Color.Red
            ToolTip1.SetToolTip(accessType, "Caution--Using deprecated Java access to Alma.")
        End If
    End Sub

    Private Sub dontConvert_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dontConvert.CheckedChanged
        If dontConvert.Checked = True Then
            MsgBox("Caution--Check this box only to enable viewing the RESTful XML file in " &
            "the 'Current XML' panel. The RESTful XML must be converted to the older (Java) " &
            "format in order for SpineOMatic to function.", MsgBoxStyle.Information, "Warning")
        End If
    End Sub

    Private Sub UseServlet_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub UpdatePath_TextChanged(sender As Object, e As EventArgs) Handles updatePath.TextChanged

    End Sub
End Class
