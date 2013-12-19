Option Explicit On
Option Strict On

Imports System.IO
Imports Microsoft.Win32
Imports System.Text.RegularExpressions

Public Enum OS
    NOT_SUPPORTED = 0
    WINDOWS_2000 = 1
    WINDOWS_XP = 2
    WINDOWS_2003 = 3
    WINDOWS_VISTA = 4
    WINDOWS_7 = 5
    WINDOWS_8 = 6
End Enum

Public Enum GAME_ID
    CLASSIC = 0
    WINTER_ASSAULT = 1
    DARK_CRUSADE = 2
    SOULSTORM = 3
End Enum

Public Class fmMain
#Region "Declarations"
    '// Used by Dawn of War - Classic
    Public Const _GameKeyPattern_5 As String = "^[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]-[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]-[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]-[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]-[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]$"
    '// Used by Winter Assault, Dark Crusade and Soulstorm | Example: 34h3-3r5t-34z6-347h-54g4
    Public Const _GameKeyPattern_4 As String = "^[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]-[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]-[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]-[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]$"
#End Region

#Region "Load Event Handling"
    Private Sub fmMain_Load_Handler(sender As Object, e As EventArgs) Handles Me.Load
        '// Add a new line at the end of a existing logfile.
        Initialize_Log()
        '// Close log StreamWriter if form is closing.
        AddHandler Me.FormClosing, AddressOf Close_LogStream
        Log_Msg(PRÄFIX.INFO, "Application Startup - Initialized Logsystem")

        Log_Msg(PRÄFIX.INFO, "Application Startup - System informations - Windows: """ + [Enum].GetName(GetType(OS), GetOperatingSystem) + """")
        Log_Msg(PRÄFIX.INFO, "Application Startup - System informations - 64 Bit: """ + GetOS_ArchitectureAsString() + """")

        Log_Msg(PRÄFIX.INFO, "Application Startup - Loading Backgroundcolors")
        '// Transparent for each Status Icon's Background (Images are PNG)
        pbClassicStatus.BackColor = Color.Transparent
        pbWinterAssaultStatus.BackColor = Color.Transparent
        pbDarkCrusadeStatus.BackColor = Color.Transparent
        pbSoulstormStatus.BackColor = Color.Transparent

        Log_Msg(PRÄFIX.INFO, "Application Startup - Loading Soulstorm Installation Path")
        '// Query the existing data from the registry
        tbSoulstormInstallationDirectory.Text = GetRegInstallDirectory(_DBSoulstorm)
        If Regex.IsMatch(GetRegGameKey(_DBClassic), _GameKeyPattern_4) Then
            Log_Msg(PRÄFIX.INFO, "Application Startup - Regex Game Key - Is Match | Affected Game: Classic")
            '// Classic
            Dim _ClassicGameKeyParts As List(Of String) = GetGameKeyParts(GetRegGameKey(_DBClassic))
            tbClassicKey_1.Text = _ClassicGameKeyParts.Item(0)
            tbClassicKey_2.Text = _ClassicGameKeyParts.Item(1)
            tbClassicKey_3.Text = _ClassicGameKeyParts.Item(2)
            tbClassicKey_4.Text = _ClassicGameKeyParts.Item(3)
        End If
        If Regex.IsMatch(GetRegGameKey(_DBWinterAssault), _GameKeyPattern_5) Then
            Log_Msg(PRÄFIX.INFO, "Application Startup - Regex Game Key - Is Match | Affected Game: Winter Assault")
            '// Winter Assault
            Dim _WinterAssaultGameKeyParts As List(Of String) = GetGameKeyParts(GetRegGameKey(_DBWinterAssault))
            tbWinterAssaultKey_1.Text = _WinterAssaultGameKeyParts.Item(0)
            tbWinterAssaultKey_2.Text = _WinterAssaultGameKeyParts.Item(1)
            tbWinterAssaultKey_3.Text = _WinterAssaultGameKeyParts.Item(2)
            tbWinterAssaultKey_4.Text = _WinterAssaultGameKeyParts.Item(3)
            tbWinterAssaultKey_5.Text = _WinterAssaultGameKeyParts.Item(4)
        End If
        If Regex.IsMatch(GetRegGameKey(_DBDarkCrusade), _GameKeyPattern_5) Then
            Log_Msg(PRÄFIX.INFO, "Application Startup - Regex Game Key - Is Match | Affected Game: Dark Crusade")
            '// Dark Crusade
            Dim _DarkCrusadeGameKeyParts As List(Of String) = GetGameKeyParts(GetRegGameKey(_DBDarkCrusade))
            tbDarkCrusadeKey_1.Text = _DarkCrusadeGameKeyParts.Item(0)
            tbDarkCrusadeKey_2.Text = _DarkCrusadeGameKeyParts.Item(1)
            tbDarkCrusadeKey_3.Text = _DarkCrusadeGameKeyParts.Item(2)
            tbDarkCrusadeKey_4.Text = _DarkCrusadeGameKeyParts.Item(3)
            tbDarkCrusadeKey_5.Text = _DarkCrusadeGameKeyParts.Item(4)
        End If
        If Regex.IsMatch(GetRegGameKey(_DBSoulstorm), _GameKeyPattern_5) Then
            Log_Msg(PRÄFIX.INFO, "Application Startup - Regex Game Key - Is Match | Affected Game: Soulstorm")
            '// Soulstorm
            Dim _SoulstormGameKeyParts As List(Of String) = GetGameKeyParts(GetRegGameKey(_DBSoulstorm))
            tbSoulstormKey_1.Text = _SoulstormGameKeyParts.Item(0)
            tbSoulstormKey_2.Text = _SoulstormGameKeyParts.Item(1)
            tbSoulstormKey_3.Text = _SoulstormGameKeyParts.Item(2)
            tbSoulstormKey_4.Text = _SoulstormGameKeyParts.Item(3)
            tbSoulstormKey_5.Text = _SoulstormGameKeyParts.Item(4)
        End If
    End Sub

    ''' <summary>Split a String on each '-' to get all single parts of a whole game key.</summary>
    Private Function GetGameKeyParts(_FullGameKey As String) As List(Of String)
        Log_Msg(PRÄFIX.INFO, "Functions - GetGameKeyParts - Split Key | Key: """ + _FullGameKey.Substring(0, _FullGameKey.LastIndexOf("-")) + "-XXXX""")
        Dim _GameKeyParts As New List(Of String)
        _GameKeyParts.AddRange(Regex.Split(_FullGameKey.ToUpper, "-"))
        Return _GameKeyParts
    End Function
#End Region
    
    Private Sub Unlock()
        '// Check if entered game keys are valid and operation system is supported / Soulstorm.exe path is correct.
        If Regex.IsMatch(GetCompleteGameKey(GAME_ID.CLASSIC), _GameKeyPattern_4) AndAlso _
           Regex.IsMatch(GetCompleteGameKey(GAME_ID.WINTER_ASSAULT), _GameKeyPattern_5) AndAlso _
           Regex.IsMatch(GetCompleteGameKey(GAME_ID.DARK_CRUSADE), _GameKeyPattern_5) AndAlso _
           Regex.IsMatch(GetCompleteGameKey(GAME_ID.SOULSTORM), _GameKeyPattern_5) Then

            If Not GetOperatingSystem() = OS.NOT_SUPPORTED Then
                If Not IsMatchSoulstormEXE(tbSoulstormInstallationDirectory.Text) Then '// NOT ENTFERNEN
                    Dim _Unlocker As New Cls_RaceUnlocker(GetCompleteGameKey(GAME_ID.CLASSIC), _
                                                          GetCompleteGameKey(GAME_ID.WINTER_ASSAULT), _
                                                          GetCompleteGameKey(GAME_ID.DARK_CRUSADE), _
                                                          GetCompleteGameKey(GAME_ID.SOULSTORM), _
                                                          tbSoulstormInstallationDirectory.Text)
                    _Unlocker.Unlock_Process_Start()
                End If
            Else
                Log_Msg(PRÄFIX.EXCEPTION, "PreUnlock Process - Status - Operation system not supported")
                MessageBox.Show("Your operation system is not supported yet!", "Wrong OS.", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            Log_Msg(PRÄFIX.WARNING, "PreUnlock Process - Status - Wrong CD-Keys found")
            MessageBox.Show("Please enter a only valid CD-Keys!", "Wrong gamekey syntax.", MessageBoxButtons.OK, MessageBoxIcon.Hand)
        End If

    End Sub

#Region "Soulstorm Path Selection/Check"
    Private Sub ChooseSoulstormPath()
        Dim _FolderDialog As FolderBrowserDialog = New FolderBrowserDialog With {.Description = "Dawn of War - Soulstorm, directory.", _
                                                                                 .ShowNewFolderButton = False}

        If _FolderDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            IsMatchSoulstormEXE(_FolderDialog.SelectedPath)
        End If
    End Sub

    ''' <summary>Check if in the SoulstormFolder the Soulstorm.exe exists. If yes, check for updates for the Game.</summary>
    Private Function IsMatchSoulstormEXE(_SoulstormFolderPath As String) As Boolean
        If File.Exists(_SoulstormFolderPath + "\Soulstorm.exe") Then
            If FileVersionInfo.GetVersionInfo(_SoulstormFolderPath + "\Soulstorm.exe").FileVersion = "1, 4, 0, 0" Then
                Log_Msg(PRÄFIX.INFO, "Functions - IsMatchSoulstormEXE - Soulstorm is UpToDate")
                tbSoulstormInstallationDirectory.Text = PathShorten(_SoulstormFolderPath, 340, tbSoulstormInstallationDirectory.Font)
            Else
                Log_Msg(PRÄFIX.INFO, "Functions - IsMatchSoulstormEXE - Soulstorm Update available")
                Select Case MessageBox.Show("You should update your Soulstorm installation. Download now?" + vbCrLf + vbCrLf + _
                                            "Patch: SS_DE_1.20_Patch.zip | 111 MiB" + vbCrLf + _
                                            "SHA1: fb26609a168b489d3fcd5aba6581b2154d9872de" + vbCrLf + vbCrLf + _
                                            "Note: includes patch 1.1 and 1.2 in german.", _
                                            "Patch(s) available! | Version: 1.4.0.0 | Current: " + _
                                            FileVersionInfo.GetVersionInfo(_SoulstormFolderPath).FileVersion.Replace(" ", "").Replace(",", "."), MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                    Case Windows.Forms.DialogResult.Yes
                        Process.Start("http://fire-emerald.com/custom/patches/ss_de_1.20_patch.zip")
                End Select
            End If
            Return True
        End If
        Log_Msg(PRÄFIX.WARNING, "Functions - IsMatchSoulstormEXE - No Soulstorm.exe found. | Directory: """ + _SoulstormFolderPath + "\Soulstorm.exe""")
        Select Case MessageBox.Show("Please check the Installation Path. The 'Soulstorm.exe' couldn't found!" + vbCrLf + _
                                    "Selected: """ + tbSoulstormInstallationDirectory.Text + """", "Soulstorm.exe not found.", MessageBoxButtons.OKCancel, MessageBoxIcon.Hand)
            Case Windows.Forms.DialogResult.OK
                ChooseSoulstormPath()
        End Select
        Return False
    End Function
#End Region

#Region "Functions"
    ''' <summary>Get a shorted file/directory path if the path is to long.</summary>
    ''' <param name="_Path">The path which should be returned truncated.</param>
    ''' <param name="_Length">The Length in pixel which shouldn't be exceeded. </param>
    ''' <param name="_TextFont">Used font.</param>
    Private Function PathShorten(_Path As String, _Length As Integer, _TextFont As Font) As String
        Log_Msg(PRÄFIX.INFO, "Functions - PathShorten - Cut | Path: """ + _Path + """ | Length: """ + _Length.ToString + """ | Font: """ + _TextFont.ToString)
        Dim PathParts() As String = Split(_Path, "\")
        Dim PathBuild As New System.Text.StringBuilder(_Path.Length)
        Dim LastPart As String = PathParts(PathParts.Length - 1)
        Dim PrevPath As String = ""

        'Erst prüfen ob der komplette String evtl. bereits kürzer als die Maximallänge ist
        If TextRenderer.MeasureText(_Path, _TextFont).Width < _Length Then
            Return _Path
        End If

        For i As Integer = 0 To PathParts.Length - 1
            PathBuild.Append(PathParts(i) & "\")
            If TextRenderer.MeasureText(PathBuild.ToString & "...\" & LastPart, _TextFont).Width >= _Length Then
                Return PrevPath
            Else
                PrevPath = PathBuild.ToString & "...\" & LastPart
            End If
        Next
        Return PrevPath
    End Function

    ''' <summary>Merges all 4/5 TextBoxes of a Game Key together and add '-' between them.</summary>
    Private Function GetCompleteGameKey(_GameId As GAME_ID) As String
        Log_Msg(PRÄFIX.INFO, "Functions - GetCompleteGameKey - Merge | GameID: """ + [Enum].GetName(GetType(GAME_ID), _GameId) + """")
        Select Case _GameId
            Case GAME_ID.CLASSIC : Return tbClassicKey_1.Text + "-" + tbClassicKey_2.Text + "-" + tbClassicKey_3.Text + "-" + tbClassicKey_4.Text
            Case GAME_ID.WINTER_ASSAULT : Return tbWinterAssaultKey_1.Text + "-" + tbWinterAssaultKey_2.Text + "-" + tbWinterAssaultKey_3.Text + "-" + tbWinterAssaultKey_4.Text + "-" + tbWinterAssaultKey_5.Text
            Case GAME_ID.DARK_CRUSADE : Return tbDarkCrusadeKey_1.Text + "-" + tbDarkCrusadeKey_2.Text + "-" + tbDarkCrusadeKey_3.Text + "-" + tbDarkCrusadeKey_4.Text + "-" + tbDarkCrusadeKey_5.Text
            Case GAME_ID.SOULSTORM : Return tbSoulstormKey_1.Text + "-" + tbSoulstormKey_2.Text + "-" + tbSoulstormKey_3.Text + "-" + tbSoulstormKey_4.Text + "-" + tbSoulstormKey_5.Text
            Case Else : Return ""
        End Select
    End Function
#End Region

#Region "OS System"
    ''' <summary>Check if System is 64 bit YES/NO.</summary>
    Private Function GetOS_ArchitectureAsString() As String
        Return Environment.Is64BitOperatingSystem.ToString.ToUpper.Replace("TRUE", "YES").Replace("FALSE", "NO")
    End Function

    ''' <summary>Get the current operation system.</summary>
    Private Function GetOperatingSystem() As OS
        Log_Msg(PRÄFIX.INFO, "Functions - GetOperationSystem")
        Select Case Environment.OSVersion.Platform
            Case PlatformID.Win32NT
                Select Case Environment.OSVersion.Version.Major
                    Case 5
                        Select Case Environment.OSVersion.Version.Minor
                            Case 0 : Return OS.WINDOWS_2000  '// Windows 2000
                            Case 1 : Return OS.WINDOWS_XP    '// Windows XP
                            Case 2 : Return OS.WINDOWS_2003  '// Windows 2003
                        End Select
                    Case 6
                        Select Case Environment.OSVersion.Version.Minor
                            Case 0 : Return OS.WINDOWS_VISTA '// Windows Vista/Windows 2008 Server
                            Case 1 : Return OS.WINDOWS_7     '// Windows 7
                            Case 2 : Return OS.WINDOWS_8     '// Windows 8
                        End Select
                End Select
        End Select
        Return OS.NOT_SUPPORTED '// Nicht unterstützt.
    End Function
#End Region

#Region "Window Movement"
    Dim _MovedWhileDown As Boolean = False
    Dim _MouseDown As Boolean = False

    Dim mouseOffset As Point

    Private Sub Me_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Me.MouseDown
        _MouseDown = True
        mouseOffset = New Point(-e.X, -e.Y)
    End Sub
    Private Sub Me_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Me.MouseMove
        If _MouseDown Then
            _MovedWhileDown = True
            If e.Button = MouseButtons.Left Then
                Dim mousePos = Control.MousePosition
                mousePos.Offset(mouseOffset.X, mouseOffset.Y)
                Location = mousePos
            End If
        Else
            _MovedWhileDown = False
        End If
    End Sub
    Private Sub Me_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Me.MouseUp
        If Not _MovedWhileDown Then
            Log_Msg(PRÄFIX.INFO, "Window Movement - Form MouseUp - Application Exit")
            Application.Exit()
        Else
            _MouseDown = False
            _MovedWhileDown = False
        End If
    End Sub
#End Region

#Region "Button Management"
#Region "Close Button"
    Private Sub pbClose_MouseDown(sender As Object, e As MouseEventArgs) Handles pbClose.MouseDown
        pbClose.Image = My.Resources.close_mouse_down
    End Sub

    Private Sub pbClose_MouseEnter(sender As Object, e As EventArgs) Handles pbClose.MouseEnter
        pbClose.Image = My.Resources.close_mouse_over
    End Sub

    Private Sub pbClose_MouseLeave(sender As Object, e As EventArgs) Handles pbClose.MouseLeave
        pbClose.Image = My.Resources.close
    End Sub

    Private Sub pbClose_MouseUp(sender As Object, e As MouseEventArgs) Handles pbClose.MouseUp
        pbClose.Image = My.Resources.close
        Log_Msg(PRÄFIX.INFO, "Button Management - Close MouseUp - Application Exit")
        Application.Exit()
    End Sub
#End Region
#Region "Minimize Button"
    Private Sub pbMinimize_MouseDown(sender As Object, e As MouseEventArgs) Handles pbMinimize.MouseDown
        pbMinimize.Image = My.Resources.minimize_mouse_down
    End Sub
    Private Sub pbMinimize_MouseEnter(sender As Object, e As EventArgs) Handles pbMinimize.MouseEnter
        pbMinimize.Image = My.Resources.minimize_mouse_over
    End Sub
    Private Sub pbMinimize_MouseLeave(sender As Object, e As EventArgs) Handles pbMinimize.MouseLeave
        pbMinimize.Image = My.Resources.minimize
    End Sub
    Private Sub pbMinimize_MouseUp(sender As Object, e As MouseEventArgs) Handles pbMinimize.MouseUp
        pbMinimize.Image = My.Resources.close
        Log_Msg(PRÄFIX.INFO, "Button Management - Minimize MouseUp - Minimize")
        Me.WindowState = FormWindowState.Minimized
    End Sub
#End Region
#Region "Unlock Button"
    Private Sub pbUnlock_MouseDown(sender As Object, e As MouseEventArgs) Handles pbUnlock.MouseDown
        pbUnlock.Image = My.Resources.unlock_mouse_down
    End Sub
    Private Sub pbUnlock_MouseEnter(sender As Object, e As EventArgs) Handles pbUnlock.MouseEnter
        pbUnlock.Image = My.Resources.unlock_mouse_over
    End Sub
    Private Sub pbUnlock_MouseLeave(sender As Object, e As EventArgs) Handles pbUnlock.MouseLeave
        pbUnlock.Image = My.Resources.unlock
    End Sub
    Private Sub pbUnlock_MouseUp(sender As Object, e As MouseEventArgs) Handles pbUnlock.MouseUp
        pbUnlock.Image = My.Resources.unlock
        Log_Msg(PRÄFIX.INFO, "Button Management - Unlock MouseUp - Start unlock process")
        Unlock()
    End Sub
#End Region
#Region "Choose Button"
    Private Sub pbChoose_MouseDown(sender As Object, e As MouseEventArgs) Handles pbChoose.MouseDown
        pbChoose.Image = My.Resources.choose_mouse_down
    End Sub
    Private Sub pbChoose_MouseEnter(sender As Object, e As EventArgs) Handles pbChoose.MouseEnter
        pbChoose.Image = My.Resources.choose_mouse_over
    End Sub
    Private Sub pbChoose_MouseLeave(sender As Object, e As EventArgs) Handles pbChoose.MouseLeave
        pbChoose.Image = My.Resources.choose
    End Sub
    Private Sub pbChoose_MouseUp(sender As Object, e As MouseEventArgs) Handles pbChoose.MouseUp
        pbChoose.Image = My.Resources.choose
        Log_Msg(PRÄFIX.INFO, "Button Management - Choose MouseUp - Start select soulstorm path")
        ChooseSoulstormPath()
    End Sub
#End Region
#End Region

#Region "GameKeyCheck Management"
    Private Sub ValidateKey(_tb As Object, e As EventArgs) Handles _
        tbClassicKey_1.TextChanged, tbClassicKey_2.TextChanged, tbClassicKey_3.TextChanged, tbClassicKey_4.TextChanged, _
        tbWinterAssaultKey_1.TextChanged, tbWinterAssaultKey_2.TextChanged, tbWinterAssaultKey_3.TextChanged, tbWinterAssaultKey_4.TextChanged, tbWinterAssaultKey_5.TextChanged, _
        tbDarkCrusadeKey_1.TextChanged, tbDarkCrusadeKey_2.TextChanged, tbDarkCrusadeKey_3.TextChanged, tbDarkCrusadeKey_4.TextChanged, tbDarkCrusadeKey_5.TextChanged, _
        tbSoulstormKey_1.TextChanged, tbSoulstormKey_2.TextChanged, tbSoulstormKey_3.TextChanged, tbSoulstormKey_4.TextChanged, tbSoulstormKey_5.TextChanged

        '// Switch the current TextBox if it has 4 characters and check  if the entered game key is valid.
        If _tb Is tbClassicKey_1 Or _tb Is tbClassicKey_2 Or _tb Is tbClassicKey_3 Or _tb Is tbClassicKey_4 Then
            If Regex.IsMatch(GetCompleteGameKey(GAME_ID.CLASSIC), _GameKeyPattern_4) Then
                pbClassicStatus.Image = My.Resources.hacken1
            ElseIf Not tbClassicKey_1.Text = "" Or Not tbClassicKey_2.Text = "" Or Not tbClassicKey_3.Text = "" Or Not tbClassicKey_4.Text = "" Then
                pbClassicStatus.Image = My.Resources.kreuz1
            Else
                pbClassicStatus.Image = Nothing
            End If

            SwitchTextBox(CType(_tb, TextBox))
        ElseIf _tb Is tbWinterAssaultKey_1 Or _tb Is tbWinterAssaultKey_2 Or _tb Is tbWinterAssaultKey_3 Or _tb Is tbWinterAssaultKey_4 Or _tb Is tbWinterAssaultKey_5 Then
            If Regex.IsMatch(GetCompleteGameKey(GAME_ID.WINTER_ASSAULT), _GameKeyPattern_5) Then
                pbWinterAssaultStatus.Image = My.Resources.hacken1
            ElseIf Not tbWinterAssaultKey_1.Text = "" Or Not tbWinterAssaultKey_2.Text = "" Or Not tbWinterAssaultKey_3.Text = "" Or Not tbWinterAssaultKey_4.Text = "" Or Not tbWinterAssaultKey_5.Text = "" Then
                pbWinterAssaultStatus.Image = My.Resources.kreuz1
            Else
                pbWinterAssaultStatus.Image = Nothing
            End If

            SwitchTextBox(CType(_tb, TextBox))
        ElseIf _tb Is tbDarkCrusadeKey_1 Or _tb Is tbDarkCrusadeKey_2 Or _tb Is tbDarkCrusadeKey_3 Or _tb Is tbDarkCrusadeKey_4 Or _tb Is tbDarkCrusadeKey_5 Then
            If Regex.IsMatch(GetCompleteGameKey(GAME_ID.DARK_CRUSADE), _GameKeyPattern_5) Then
                pbDarkCrusadeStatus.Image = My.Resources.hacken1
            ElseIf Not tbDarkCrusadeKey_1.Text = "" Or Not tbDarkCrusadeKey_2.Text = "" Or Not tbDarkCrusadeKey_3.Text = "" Or Not tbDarkCrusadeKey_4.Text = "" Or Not tbDarkCrusadeKey_5.Text = "" Then
                pbDarkCrusadeStatus.Image = My.Resources.kreuz1
            Else
                pbDarkCrusadeStatus.Image = Nothing
            End If

            SwitchTextBox(CType(_tb, TextBox))
        ElseIf _tb Is tbSoulstormKey_1 Or _tb Is tbSoulstormKey_2 Or _tb Is tbSoulstormKey_3 Or _tb Is tbSoulstormKey_4 Or _tb Is tbSoulstormKey_5 Then
            If Regex.IsMatch(GetCompleteGameKey(GAME_ID.SOULSTORM), _GameKeyPattern_5) Then
                pbSoulstormStatus.Image = My.Resources.hacken1
            ElseIf Not tbSoulstormKey_1.Text = "" Or Not tbSoulstormKey_2.Text = "" Or Not tbSoulstormKey_3.Text = "" Or Not tbSoulstormKey_4.Text = "" Or Not tbSoulstormKey_5.Text = "" Then
                pbSoulstormStatus.Image = My.Resources.kreuz1
            Else
                pbSoulstormStatus.Image = Nothing
            End If

            SwitchTextBox(CType(_tb, TextBox))
        End If
    End Sub
#End Region

#Region "Switch Management"
    '// Switch the given TextBox, if it has 4 characters.
    Private Sub SwitchTextBox(_tb As TextBox)
        If _tb.Text.Length = 4 Then
            Select Case True
                Case _tb Is tbClassicKey_1 : tbClassicKey_2.Focus()
                Case _tb Is tbClassicKey_2 : tbClassicKey_3.Focus()
                Case _tb Is tbClassicKey_3 : tbClassicKey_4.Focus()
                Case _tb Is tbClassicKey_4 : tbWinterAssaultKey_1.Focus()
                Case _tb Is tbWinterAssaultKey_1 : tbWinterAssaultKey_2.Focus()
                Case _tb Is tbWinterAssaultKey_2 : tbWinterAssaultKey_3.Focus()
                Case _tb Is tbWinterAssaultKey_3 : tbWinterAssaultKey_4.Focus()
                Case _tb Is tbWinterAssaultKey_4 : tbWinterAssaultKey_5.Focus()
                Case _tb Is tbWinterAssaultKey_5 : tbDarkCrusadeKey_1.Focus()
                Case _tb Is tbDarkCrusadeKey_1 : tbDarkCrusadeKey_2.Focus()
                Case _tb Is tbDarkCrusadeKey_2 : tbDarkCrusadeKey_3.Focus()
                Case _tb Is tbDarkCrusadeKey_3 : tbDarkCrusadeKey_4.Focus()
                Case _tb Is tbDarkCrusadeKey_4 : tbDarkCrusadeKey_5.Focus()
                Case _tb Is tbDarkCrusadeKey_5 : tbSoulstormKey_1.Focus()
                Case _tb Is tbSoulstormKey_1 : tbSoulstormKey_2.Focus()
                Case _tb Is tbSoulstormKey_2 : tbSoulstormKey_3.Focus()
                Case _tb Is tbSoulstormKey_3 : tbSoulstormKey_4.Focus()
                Case _tb Is tbSoulstormKey_4 : tbSoulstormKey_5.Focus()
            End Select
        End If
    End Sub
#End Region

End Class
