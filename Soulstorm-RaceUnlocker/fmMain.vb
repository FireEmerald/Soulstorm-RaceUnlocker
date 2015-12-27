
'* Copyright (C) 2013-2015 FireEmerald <https://github.com/FireEmerald>
'* Copyright (C) 2008-2009 n0|Belial2003 <http://dow.4players.de/forum/index.php?page=User&userID=10286&s=4d85aca336eaa03924c488f8e7e6ed7cd7389caa>
'*
'* Project: Soulstorm - Race Unlocker
'*
'* Requires: .NET Framework 4 or higher, because of the RegistryKey.OpenBaseKey Method.
'*
'* This program is free software; you can redistribute it and/or modify it
'* under the terms of the GNU General Public License as published by the
'* Free Software Foundation; either version 2 of the License, or (at your
'* option) any later version.
'*
'* This program is distributed in the hope that it will be useful, but WITHOUT
'* ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'* FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'* more details.
'*
'* You should have received a copy of the GNU General Public License along
'* with this program. If not, see <http://www.gnu.org/licenses/>.

Imports System.IO
Imports System.Management
Imports System.Text.RegularExpressions
Imports System.Security.Cryptography

Public Enum GAME_ID
    NOT_SET = 0
    CLASSIC = 1
    WINTER_ASSAULT = 2
    DARK_CRUSADE = 3
    SOULSTORM = 4
End Enum

Public Class fmMain
#Region "Declarations"
    '// Link to the Soulstorm patches.
    Private Const SOULSTORM_PATCH_LINK As String = "http://www.patches-scrolls.de/patch/4741/7/"
    '// Complete Soulstorm folder path
    Private _SoulstormFolderPath As String = ""

    Private Structure SoulstormExeHash
        Dim Hash, VersionString As String
    End Structure

    Private ReadOnly _Soulstorm_Exe_1_0 As New SoulstormExeHash With {.Hash = "A0714DAEC81095D4AC4B2CB4BF11D2CFAC32E261", .VersionString = "DOW-Engine 1.0.9409, Dawn of War: Soulstorm 1.0, Build 9409"}
    Private ReadOnly _Soulstorm_Exe_1_1 As New SoulstormExeHash With {.Hash = "075BB04276A5C5A285BA25FCD38B25F00BA3477B", .VersionString = "DOW-Engine 1.1.100, Dawn of War: Soulstorm 1.0, Build 100"}
    Private ReadOnly _Soulstorm_Exe_1_2 As New SoulstormExeHash With {.Hash = "36B6D805E452541764873ED6D1EAD07A7C025C64", .VersionString = "DOW-Engine 1.2.120, Dawn of War: Soulstorm 1.0, Build 120"}
    Private ReadOnly _KnownSoulstormHashes As New List(Of SoulstormExeHash)({_Soulstorm_Exe_1_0, _Soulstorm_Exe_1_1, _Soulstorm_Exe_1_2})

    '// RegEx Patterns:
    '// - used for Dawn of War - Classic:
    Public Const _GameKeyPattern_5 As String = "^[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]-[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]-[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]-[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]-[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]$"
    '// - used for Winter Assault, Dark Crusade and Soulstorm: (Example: 34h3-3r5t-34z6-347h-54g4)
    Public Const _GameKeyPattern_4 As String = "^[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]-[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]-[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]-[0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z][0-9,A-Z,a-z]$"
#End Region

#Region "Load Event Handling"
    Private Sub fmMain_Load_Handler(sender As Object, e As EventArgs) Handles Me.Load
        '// Adds a new line at the end of an existing log file.
        Initialize_Log()
        Log_Msg(PREFIX.INFO, "Application Startup - Initialized Log system")

        Log_Msg(PREFIX.INFO, "Application Startup - RaceUnlocker - Version: '{0}'", New Object() {My.Application.Info.Version.ToString})

        Log_Msg(PREFIX.INFO, "Application Startup - System informations - Windows | Version: '{0}' | FriendlyName: '{1}'", New Object() {Environment.OSVersion.VersionString, GetOS_FriendlyName()})
        Log_Msg(PREFIX.INFO, "Application Startup - System informations - 64 Bit: '{0}'", New Object() {GetOS_ArchitectureAsString()})

        Log_Msg(PREFIX.INFO, "Application Startup - System informations - All 32bit Registry keys are redirected to 'HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node'")
        Log_Msg(PREFIX.INFO, "Application Startup - System informations - All 64bit Registry keys stored in 'HKEY_LOCAL_MACHINE\SOFTWARE'")
        Log_Msg(PREFIX.INFO, "Application Startup - System informations - All Dawn of War (Soulstorm or older) Applications are 32bit, so they are always at the 32bit directory.")

        Log_Msg(PREFIX.INFO, "Application Startup - Loading Background colors")
        '// Transparent for each Status Icon's Background (Images are PNG)
        pbClassicStatus.BackColor = Color.Transparent
        pbWinterAssaultStatus.BackColor = Color.Transparent
        pbDarkCrusadeStatus.BackColor = Color.Transparent
        pbSoulstormStatus.BackColor = Color.Transparent

        Log_Msg(PREFIX.INFO, "Application Startup - Loading Soulstorm Installation Path")
        '// Load the existing data from the registry
        Dim _RegSoulstormFolderPath As String = GetRegInstallDirectory(_DBSoulstorm)
        tbSoulstormInstallationDirectory.Text = PathShorten(_RegSoulstormFolderPath, 340, tbSoulstormInstallationDirectory.Font)
        '// Overrides the path set before by TextChanged event.
        _SoulstormFolderPath = _RegSoulstormFolderPath

        Log_Msg(PREFIX.INFO, "Application Startup - Soulstorm Installation Path - RegPath: '" & _SoulstormFolderPath & "' | RegPathResized: '" & tbSoulstormInstallationDirectory.Text & "'")

        If Regex.IsMatch(GetRegGameKey(_DBClassic), _GameKeyPattern_4) Then
            Log_Msg(PREFIX.INFO, "Application Startup - RegEx Game Key - Is Match | Affected Game: Classic")
            '// Classic
            Dim _ClassicGameKeyParts As List(Of String) = GetGameKeyParts(GetRegGameKey(_DBClassic))
            tbClassicKey_1.Text = _ClassicGameKeyParts.Item(0)
            tbClassicKey_2.Text = _ClassicGameKeyParts.Item(1)
            tbClassicKey_3.Text = _ClassicGameKeyParts.Item(2)
            tbClassicKey_4.Text = _ClassicGameKeyParts.Item(3)
        End If
        If Regex.IsMatch(GetRegGameKey(_DBWinterAssault), _GameKeyPattern_5) Then
            Log_Msg(PREFIX.INFO, "Application Startup - RegEx Game Key - Is Match | Affected Game: Winter Assault")
            '// Winter Assault
            Dim _WinterAssaultGameKeyParts As List(Of String) = GetGameKeyParts(GetRegGameKey(_DBWinterAssault))
            tbWinterAssaultKey_1.Text = _WinterAssaultGameKeyParts.Item(0)
            tbWinterAssaultKey_2.Text = _WinterAssaultGameKeyParts.Item(1)
            tbWinterAssaultKey_3.Text = _WinterAssaultGameKeyParts.Item(2)
            tbWinterAssaultKey_4.Text = _WinterAssaultGameKeyParts.Item(3)
            tbWinterAssaultKey_5.Text = _WinterAssaultGameKeyParts.Item(4)
        End If
        If Regex.IsMatch(GetRegGameKey(_DBDarkCrusade), _GameKeyPattern_5) Then
            Log_Msg(PREFIX.INFO, "Application Startup - RegEx Game Key - Is Match | Affected Game: Dark Crusade")
            '// Dark Crusade
            Dim _DarkCrusadeGameKeyParts As List(Of String) = GetGameKeyParts(GetRegGameKey(_DBDarkCrusade))
            tbDarkCrusadeKey_1.Text = _DarkCrusadeGameKeyParts.Item(0)
            tbDarkCrusadeKey_2.Text = _DarkCrusadeGameKeyParts.Item(1)
            tbDarkCrusadeKey_3.Text = _DarkCrusadeGameKeyParts.Item(2)
            tbDarkCrusadeKey_4.Text = _DarkCrusadeGameKeyParts.Item(3)
            tbDarkCrusadeKey_5.Text = _DarkCrusadeGameKeyParts.Item(4)
        End If
        If Regex.IsMatch(GetRegGameKey(_DBSoulstorm), _GameKeyPattern_5) Then
            Log_Msg(PREFIX.INFO, "Application Startup - RegEx Game Key - Is Match | Affected Game: Soulstorm")
            '// Soulstorm
            Dim _SoulstormGameKeyParts As List(Of String) = GetGameKeyParts(GetRegGameKey(_DBSoulstorm))
            tbSoulstormKey_1.Text = _SoulstormGameKeyParts.Item(0)
            tbSoulstormKey_2.Text = _SoulstormGameKeyParts.Item(1)
            tbSoulstormKey_3.Text = _SoulstormGameKeyParts.Item(2)
            tbSoulstormKey_4.Text = _SoulstormGameKeyParts.Item(3)
            tbSoulstormKey_5.Text = _SoulstormGameKeyParts.Item(4)
        End If
    End Sub

    ''' <summary>Splits a String on each '-' to get all single parts of a whole game key.</summary>
    Private Function GetGameKeyParts(_FullGameKey As String) As List(Of String)
        Log_Msg(PREFIX.INFO, "Functions - GetGameKeyParts - Split Key | Key: '" & _FullGameKey.Substring(0, _FullGameKey.LastIndexOf("-")) & "-XXXX'")
        Dim _GameKeyParts As New List(Of String)
        _GameKeyParts.AddRange(Regex.Split(_FullGameKey.ToUpper, "-"))
        Return _GameKeyParts
    End Function
#End Region

    Private Sub Unlock()
        '// Check if the entered game keys are valid and the 'Soulstorm.exe' path is correct.
        If Regex.IsMatch(GetCompleteGameKey(GAME_ID.CLASSIC), _GameKeyPattern_4) AndAlso _
           Regex.IsMatch(GetCompleteGameKey(GAME_ID.WINTER_ASSAULT), _GameKeyPattern_5) AndAlso _
           Regex.IsMatch(GetCompleteGameKey(GAME_ID.DARK_CRUSADE), _GameKeyPattern_5) AndAlso _
           Regex.IsMatch(GetCompleteGameKey(GAME_ID.SOULSTORM), _GameKeyPattern_5) Then

            If IsMatchSoulstormEXE(_SoulstormFolderPath) Then
                '// Start Unlock Process. First the registry unlock, then the *.exe part.
                Dim _Unlocker As New Cls_RaceUnlocker(GetCompleteGameKey(GAME_ID.CLASSIC), _
                                                      GetCompleteGameKey(GAME_ID.WINTER_ASSAULT), _
                                                      GetCompleteGameKey(GAME_ID.DARK_CRUSADE), _
                                                      GetCompleteGameKey(GAME_ID.SOULSTORM), _
                                                      _SoulstormFolderPath)
                _Unlocker.Unlock_Registry()

                If _Unlocker.GetRegistryUnlockStatus = Cls_RaceUnlocker.UNLOCK_REGISTRY_SUCCESSFULLY Then
                    _Unlocker.Unlock_Exe()

                    If _Unlocker.GetExeUnlockStatus = Cls_RaceUnlocker.UNLOCK_EXE_SUCCESSFULLY Then
                        MessageBox.Show("Unlock process successfully completed.", "All Races unlocked!", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        '// Error while unlocking the executable.
                        MessageBox.Show(_Unlocker.GetExeUnlockStatus, "Unlock process failed (Executable)", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If
                Else
                    '// Error while unlocking the registry.
                    MessageBox.Show(_Unlocker.GetRegistryUnlockStatus, "Unlock process failed (Registry)", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If

                '// Show the log file if the user wants to see it.
                Select Case MessageBox.Show("Would you like to see the log file?", "Process informations", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    Case Windows.Forms.DialogResult.Yes
                        Try
                            Process.Start(GetFullLogfilePath)
                        Catch ex As Exception
                            MessageBox.Show("The log file couldn't be found!" & vbCrLf & _
                                            vbCrLf & _
                                            ex.Message, "Log file not found", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                End Select
                '// Unlock process END
            End If
        Else
            Log_Msg(PREFIX.WARNING, "PreUnlock Process - Status - Wrong CD-Keys found")
            MessageBox.Show("Please enter a only valid CD-Keys!", "Serial Number syntax", MessageBoxButtons.OK, MessageBoxIcon.Hand)
        End If

    End Sub

#Region "Soulstorm Path Selection/Check"
    Private Sub ChooseSoulstormPath()
        Dim _FolderDialog As FolderBrowserDialog = New FolderBrowserDialog With {.Description = "Select your Dawn of War - Soulstorm directory:", _
                                                                                 .ShowNewFolderButton = False}

        If _FolderDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            IsMatchSoulstormEXE(_FolderDialog.SelectedPath)
        End If
    End Sub

    ''' <summary>Check if the SoulstormFolder contains a Soulstorm.exe. If so, we compare the hash of the current Soulstorm.exe with our stored hashes.</summary>
    Private Function IsMatchSoulstormEXE(_TempSoulstormFolderPath As String) As Boolean
        Dim _fiSoulstormExe As New FileInfo(Path.Combine(_TempSoulstormFolderPath, "Soulstorm.exe"))

        If _fiSoulstormExe.Exists Then

            Dim _SoulstormExeHash As String = SHA1_CalculateHash(_fiSoulstormExe)
            If IsNothing(_KnownSoulstormHashes.FirstOrDefault(Function(v) String.Equals(v.Hash, _SoulstormExeHash))) OrElse Not String.Equals(_Soulstorm_Exe_1_2.Hash, _SoulstormExeHash) Then
                Log_Msg(PREFIX.INFO, "Functions - IsMatchSoulstormEXE - Soulstorm Update available | Current Hash of the Soulstorm.exe: {0}", New Object() {_SoulstormExeHash})
                Select Case MessageBox.Show("Seems like your Soulstorm installation isn't up-to date." & vbCrLf & _
                                            "Would you like to visit the download site now?" & vbCrLf & _
                                            "(Multi-language)", "Patch(es) available! | Version: 1.2.0", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                    Case Windows.Forms.DialogResult.Yes
                        Try
                            Process.Start(SOULSTORM_PATCH_LINK)
                        Catch ex As Exception
                            MessageBox.Show("A error occurred, while accessing the website." & vbCrLf & _
                                            "Visit the page manually: '" & SOULSTORM_PATCH_LINK & "'", "Connection error.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End Try
                        Return False
                End Select
            Else
                Log_Msg(PREFIX.INFO, "Functions - IsMatchSoulstormEXE - Soulstorm is UpToDate")
                tbSoulstormInstallationDirectory.Text = PathShorten(_SoulstormFolderPath, 340, tbSoulstormInstallationDirectory.Font)
                '// Save the complete path. Overrides the path set before by TextChanged event.
                _SoulstormFolderPath = _fiSoulstormExe.DirectoryName
            End If
            Return True
        End If
        Log_Msg(PREFIX.WARNING, "Functions - IsMatchSoulstormEXE - No Soulstorm.exe found. | Directory: '" & _fiSoulstormExe.FullName & "'")
        Select Case MessageBox.Show("Please check the installation path. The 'Soulstorm.exe' couldn't be found!" & vbCrLf & _
                                    "'" & _fiSoulstormExe.FullName & "'", "Soulstorm.exe not found", MessageBoxButtons.OKCancel, MessageBoxIcon.Hand)
            Case Windows.Forms.DialogResult.OK
                ChooseSoulstormPath()
        End Select
        Return False
    End Function
#End Region

#Region "Additional Methods"

    ''' <summary>
    ''' Calculates the SHA-1 hash for a given file.
    ''' </summary>
    ''' <param name="FileToHash">FileInfo of the input file.</param>
    ''' <returns>Returns the SHA-1 Hash as upper case hex digits.</returns>
    ''' <remarks>Make sure, that the file exists.</remarks>
    Private Function SHA1_CalculateHash(FileToHash As FileInfo) As String
        Const BLOCKSIZE As Integer = 256 * 256

        Using _SHA1 As New SHA1CryptoServiceProvider, _fs As New FileStream(FileToHash.FullName, FileMode.Open, FileAccess.Read, FileShare.Read)
            '// Get file size
            Dim _fileSize As Long = _fs.Length

            '// Declare buffer + other vars
            Dim _readBuffer(BLOCKSIZE - 1) As Byte, _readBytes As Integer
            Dim _transformBuffer As Byte(), _transformBytes As Integer, _transformBytesTotal As Long = 0
            '// Read first block
            _readBytes = _fs.Read(_readBuffer, 0, BLOCKSIZE)
            '// Read + transform block wise
            While _readBytes > 0
                '// Save last read
                _transformBuffer = _readBuffer
                _transformBytes = _readBytes
                '// Read buffer
                _readBuffer = New Byte(BLOCKSIZE - 1) {}
                _readBytes = _fs.Read(_readBuffer, 0, BLOCKSIZE)
                '// Transform
                Select Case _readBytes
                    Case 0 : _SHA1.TransformFinalBlock(_transformBuffer, 0, _transformBytes)
                    Case Else : _SHA1.TransformBlock(_transformBuffer, 0, _transformBytes, _transformBuffer, 0)
                End Select
                '// Inform about progress here
                _transformBytesTotal += _transformBytes
            End While
            '// All done
            Return HexStringFromBytes(_SHA1.Hash)
        End Using
    End Function

    ''' <summary>
    ''' Converts an array of bytes to a string of hex digits.
    ''' </summary>
    ''' <param name="bytes">Array of bytes.</param>
    ''' <returns>String of upper case hex digits.</returns>
    Public Shared Function HexStringFromBytes(bytes As Byte()) As String
        Dim _sb = New System.Text.StringBuilder()
        For Each b As Byte In bytes
            _sb.Append(b.ToString("x2"))
        Next
        Return _sb.ToString.ToUpper
    End Function

    ''' <summary>Get a shorted file/directory path if the path is to long.</summary>
    ''' <param name="_Path">The path which should be returned truncated.</param>
    ''' <param name="_Length">The Length in pixel which shouldn't be exceeded. </param>
    ''' <param name="_TextFont">Used font.</param>
    Private Function PathShorten(_Path As String, _Length As Integer, _TextFont As Font) As String
        Log_Msg(PREFIX.INFO, "Functions - PathShorten - Cut | Path: '" & _Path & "' | Length: '" & _Length.ToString & "' | Font: '" & _TextFont.ToString & "'")
        Dim PathParts() As String = Split(_Path, "\")
        Dim PathBuild As New System.Text.StringBuilder(_Path.Length)
        Dim LastPart As String = PathParts(PathParts.Length - 1)
        Dim PrevPath As String = ""

        '// We check first, if the complete String is shorter then the max length.
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
        Log_Msg(PREFIX.INFO, "Functions - GetCompleteGameKey - Merge | GameID: '" & [Enum].GetName(GetType(GAME_ID), _GameId) & "'")
        Select Case _GameId
            Case GAME_ID.CLASSIC : Return tbClassicKey_1.Text & "-" & tbClassicKey_2.Text & "-" & tbClassicKey_3.Text & "-" & tbClassicKey_4.Text
            Case GAME_ID.WINTER_ASSAULT : Return tbWinterAssaultKey_1.Text & "-" & tbWinterAssaultKey_2.Text & "-" & tbWinterAssaultKey_3.Text & "-" & tbWinterAssaultKey_4.Text & "-" & tbWinterAssaultKey_5.Text
            Case GAME_ID.DARK_CRUSADE : Return tbDarkCrusadeKey_1.Text & "-" & tbDarkCrusadeKey_2.Text & "-" & tbDarkCrusadeKey_3.Text & "-" & tbDarkCrusadeKey_4.Text & "-" & tbDarkCrusadeKey_5.Text
            Case GAME_ID.SOULSTORM : Return tbSoulstormKey_1.Text & "-" & tbSoulstormKey_2.Text & "-" & tbSoulstormKey_3.Text & "-" & tbSoulstormKey_4.Text & "-" & tbSoulstormKey_5.Text
            Case Else : Return ""
        End Select
    End Function
#End Region

#Region "OS System"
    ''' <summary>Check if System is 64 bit YES/NO.</summary>
    Private Function GetOS_ArchitectureAsString() As String
        Return Environment.Is64BitOperatingSystem.ToString.ToUpper.Replace("TRUE", "YES").Replace("FALSE", "NO")
    End Function

    ''' <summary>Returns the operating system version or the String 'Unknown'.</summary>
    Private Function GetOS_FriendlyName() As String
        Dim _name As Object = (From x In New ManagementObjectSearcher("SELECT Caption FROM Win32_OperatingSystem").Get.Cast(Of ManagementObject)() _
                               Select x.GetPropertyValue("Caption")).FirstOrDefault()
        Return If(_name IsNot Nothing, _name.ToString.Trim(), "Unknown")
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
            '// Short click (without moving the mouse) will close the form.
            'Log_Msg(PREFIX.INFO, "Window Movement - Form MouseUp - Application Exit")
            'Application.Exit()
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
        Log_Msg(PREFIX.INFO, "Button Management - Close MouseUp - Application Exit")
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
        Log_Msg(PREFIX.INFO, "Button Management - Minimize MouseUp - Minimize")
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
        Log_Msg(PREFIX.INFO, "Button Management - Unlock MouseUp - Start unlock process")
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
        Log_Msg(PREFIX.INFO, "Button Management - Choose MouseUp - Start select soulstorm path")
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

        '// Switch the current TextBox if it contains 4 characters and check if the entered game key is valid.
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
    '// Switches the given TextBox, if it contains 4 characters.
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

#Region "KeyDown Clipboard Import System"
    Private Sub ClassicClipboardImport(sender As Object, e As KeyEventArgs) Handles _
        tbClassicKey_1.KeyDown, _
        tbClassicKey_2.KeyDown, _
        tbClassicKey_3.KeyDown, _
        tbClassicKey_4.KeyDown

        If e.Control AndAlso e.KeyCode = Keys.V Then ClipboardImportHandler(GAME_ID.CLASSIC)
    End Sub
    Private Sub WinterAssaultClipboardImport(sender As Object, e As KeyEventArgs) Handles _
        tbWinterAssaultKey_1.KeyDown, _
        tbWinterAssaultKey_2.KeyDown, _
        tbWinterAssaultKey_3.KeyDown, _
        tbWinterAssaultKey_4.KeyDown, _
        tbWinterAssaultKey_5.KeyDown

        If e.Control AndAlso e.KeyCode = Keys.V Then ClipboardImportHandler(GAME_ID.WINTER_ASSAULT)
    End Sub
    Private Sub DarkCrusadeClipboardImport(sender As Object, e As KeyEventArgs) Handles _
        tbDarkCrusadeKey_1.KeyDown, _
        tbDarkCrusadeKey_2.KeyDown, _
        tbDarkCrusadeKey_3.KeyDown, _
        tbDarkCrusadeKey_4.KeyDown, _
        tbDarkCrusadeKey_5.KeyDown

        If e.Control AndAlso e.KeyCode = Keys.V Then ClipboardImportHandler(GAME_ID.DARK_CRUSADE)
    End Sub
    Private Sub SoulstormClipboardImport(sender As Object, e As KeyEventArgs) Handles _
        tbSoulstormKey_1.KeyDown, _
        tbSoulstormKey_2.KeyDown, _
        tbSoulstormKey_3.KeyDown, _
        tbSoulstormKey_4.KeyDown, _
        tbSoulstormKey_5.KeyDown

        If e.Control AndAlso e.KeyCode = Keys.V Then ClipboardImportHandler(GAME_ID.SOULSTORM)
    End Sub

    Private Sub ClipboardImportHandler(_GameID As GAME_ID)
        If Clipboard.ContainsText Then
            If Regex.IsMatch(Clipboard.GetText, _GameKeyPattern_4) Then
                Dim _KeyParts() As String = Regex.Split(Clipboard.GetText, "-")
                tbClassicKey_1.Text = _KeyParts(0)
                tbClassicKey_2.Text = _KeyParts(1)
                tbClassicKey_3.Text = _KeyParts(2)
                tbClassicKey_4.Text = _KeyParts(3)
            ElseIf Regex.IsMatch(Clipboard.GetText, _GameKeyPattern_5) AndAlso Not _GameID = GAME_ID.CLASSIC Then
                Dim _KeyParts() As String = Regex.Split(Clipboard.GetText, "-")
                Select Case _GameID
                    Case GAME_ID.WINTER_ASSAULT
                        tbWinterAssaultKey_1.Text = _KeyParts(0)
                        tbWinterAssaultKey_2.Text = _KeyParts(1)
                        tbWinterAssaultKey_3.Text = _KeyParts(2)
                        tbWinterAssaultKey_4.Text = _KeyParts(3)
                        tbWinterAssaultKey_5.Text = _KeyParts(4)
                    Case GAME_ID.DARK_CRUSADE
                        tbDarkCrusadeKey_1.Text = _KeyParts(0)
                        tbDarkCrusadeKey_2.Text = _KeyParts(1)
                        tbDarkCrusadeKey_3.Text = _KeyParts(2)
                        tbDarkCrusadeKey_4.Text = _KeyParts(3)
                        tbDarkCrusadeKey_5.Text = _KeyParts(4)
                    Case GAME_ID.SOULSTORM
                        tbSoulstormKey_1.Text = _KeyParts(0)
                        tbSoulstormKey_2.Text = _KeyParts(1)
                        tbSoulstormKey_3.Text = _KeyParts(2)
                        tbSoulstormKey_4.Text = _KeyParts(3)
                        tbSoulstormKey_5.Text = _KeyParts(4)
                End Select
            Else
                MessageBox.Show("Supported import syntax:" & vbCrLf & _
                                "XXXX-XXXX-XXXX-XXXX (Classic)" & vbCrLf & _
                                "XXXX-XXXX-XXXX-XXXX-XXXX (WA, DC, SS)", "Clipboard content", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub
#End Region

#Region "TextChanged Event - SoulstormPath"
    Private Sub tbSoulstormInstallationDirectory_TextChanged(sender As Object, e As EventArgs) Handles tbSoulstormInstallationDirectory.TextChanged
        _SoulstormFolderPath = tbSoulstormInstallationDirectory.Text
    End Sub
#End Region
End Class
