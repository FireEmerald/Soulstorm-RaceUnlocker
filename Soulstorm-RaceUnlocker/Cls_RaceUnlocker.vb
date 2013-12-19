Option Explicit On
Option Strict On

Imports System.IO

Public Class Cls_RaceUnlocker

#Region "Declarations"
    '// Sub Directorys of the Game *.exe files.
    Private _SubDirClassic As String = "Unlocker\Classic and Winter Assault"
    Private _SubDirDarkCrusade As String = "Unlocker\Dark Crusade"

    '// Set by SubNew
    Private _ClassicGameKey, _
            _WinterAssaultGameKey, _
            _DarkCrusadeGameKey, _
            _SoulstormGameKey, _
            _SoulstormInstallationPath As String

    '// All already installed games, but not needed anymore
    Dim _Installed_Games As New List(Of GAME_ID)
    '// All games with a registry installation path to a directory with no <GAME>.exe
    Dim _Wrong_Installed_Games As New List(Of GAME_ID)
    '// All not installed games.
    Dim _Not_Installed_Games As New List(Of GAME_ID)
#End Region

    Sub New(ClassicKey As String, WinterAssaultKey As String, DarkCrusadeKey As String, SoulstormKey As String, SoulstormInstallationPath As String)
        Log_Msg(PRÄFIX.INFO, "Unlock - Initialize")
        _ClassicGameKey = ClassicKey
        _WinterAssaultGameKey = WinterAssaultKey
        _DarkCrusadeGameKey = DarkCrusadeKey
        _SoulstormGameKey = SoulstormKey
        _SoulstormInstallationPath = SoulstormInstallationPath
    End Sub

    Public Sub Unlock_Process_Start()
        Log_Msg(PRÄFIX.INFO, "Unlock - Start - Classic: """ + _ClassicGameKey.Substring(0, _ClassicGameKey.LastIndexOf("-")) + "-XXXX"" | " + _
                                              "Winter Assault: """ + _WinterAssaultGameKey.Substring(0, _WinterAssaultGameKey.LastIndexOf("-")) + "-XXXX"" | " + _
                                              "Dark Crusade: """ + _DarkCrusadeGameKey.Substring(0, _DarkCrusadeGameKey.LastIndexOf("-")) + "-XXXX"" | " + _
                                              "Soulstorm: " + _SoulstormGameKey.Substring(0, _SoulstormGameKey.LastIndexOf("-")) + "-XXXX"" | " + _
                                              "Soulstorm InstallLocation: """ + _SoulstormInstallationPath + """")

        Dim _PermissionTestResult As String = RegistryPermissionTest()
        If _PermissionTestResult = "" Then
            '// Check if Classic/Winter Assault and Dark Crusade has a registry entry for the InstallLocation
            CheckIfInstalled(_DBClassic)
            CheckIfInstalled(_DBWinterAssault)
            CheckIfInstalled(_DBDarkCrusade)

            '// Delete the old THQ folder
            DeleteRegTHQ()

            '// Create a new one with the directorys
            CreateRegDirectory("THQ")

            CreateRegDirectory(_DBClassic.RegGameSubDirectory, "THQ")
            CreateRegDirectory(_DBDarkCrusade.RegGameSubDirectory, "THQ")
            CreateRegDirectory(_DBSoulstorm.RegGameSubDirectory, "THQ")
            CreateRegDirectory("1.00.0000", "THQ\" + _DBSoulstorm.RegGameSubDirectory)

            '// Insert the keys / installation directorys
            '// Classic & Winter Assault | Key, Key, InstallDir
            CreateRegKey(_DBClassic.RegSerialNumberKeyName, _ClassicGameKey, _DBClassic.RegGameSubDirectory)
            CreateRegKey(_DBWinterAssault.RegSerialNumberKeyName, _WinterAssaultGameKey, _DBClassic.RegGameSubDirectory)

            CreateRegKey(_DBClassic.RegInstallLocKeyName, _SoulstormInstallationPath + "\" + _SubDirClassic + "\", _DBClassic.RegGameSubDirectory)

            '// Dark Crusade | Key, InstallDir
            CreateRegKey(_DBDarkCrusade.RegSerialNumberKeyName, _DarkCrusadeGameKey, _DBDarkCrusade.RegGameSubDirectory)

            CreateRegKey(_DBDarkCrusade.RegInstallLocKeyName, _SoulstormInstallationPath + "\" + _SubDirDarkCrusade, _DBDarkCrusade.RegGameSubDirectory)

            '// Soulstorm with all AddOns | Key, Key, Key, Key, InstallDir
            '// W40KCDKEY = Classic
            CreateRegKey("W40K" + _DBClassic.RegSerialNumberKeyName, _ClassicGameKey, _DBSoulstorm.RegGameSubDirectory)
            '// WXPCDKEY = Winter Assault
            CreateRegKey(_DBWinterAssault.RegSerialNumberKeyName.Remove(0, 6) + "CDKEY", _WinterAssaultGameKey, _DBSoulstorm.RegGameSubDirectory)
            '// DXP2CDKEY = Dark Crusade
            CreateRegKey("DXP2" + _DBDarkCrusade.RegSerialNumberKeyName, _DarkCrusadeGameKey, _DBSoulstorm.RegGameSubDirectory)
            '// CDKEY = Soulstorm
            CreateRegKey(_DBSoulstorm.RegSerialNumberKeyName, _SoulstormGameKey, _DBSoulstorm.RegGameSubDirectory)

            CreateRegKey(_DBSoulstorm.RegInstallLocKeyName, _SoulstormInstallationPath, _DBSoulstorm.RegGameSubDirectory)
        Else
            Log_Msg(PRÄFIX.EXCEPTION, "Unlock - Permission Test - Failed: """ + _PermissionTestResult + """")
            MessageBox.Show(_PermissionTestResult, "Registry error.", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End If
    End Sub

    ''' <summary>Check if a Game has a registry entry for the installation path. If yes, it will be added to the List Of X.</summary>
    Private Sub CheckIfInstalled(_Game As GameData)
        If Not GetRegInstallDirectory(_Game) = "" Then
            If IsInstalled(_Game) Then
                '// In Registry and installed
                _Installed_Games.Add(_Game.ID)
            Else
                '// In Registry but not Installed
                _Wrong_Installed_Games.Add(_Game.ID)
            End If
        Else
            '// No InstallLocation in the registry.
            _Not_Installed_Games.Add(_Game.ID)
        End If
    End Sub

    ''' <summary></summary>
    ''' <param name="_Game"></param>
    ''' <returns></returns>
    Private Function IsInstalled(_Game As GameData) As Boolean
        If File.Exists(GetRegInstallDirectory(_Game) + "\" + _Game.ExeName) Then Return True
        Return False
    End Function

#Region "Old"
    'Private ClassicKey As String, WAKey As String, DCKey As String, SSKey As String, userRoot As String
    ''Set-Methods  START: used for saving entered Keys into private vars of this class
    'Public Sub SetClassicKey(classicKey__1 As [String])
    '    ClassicKey = classicKey__1
    'End Sub
    'Public Sub SetWAKey(waKey__1 As [String])
    '    WAKey = waKey__1
    'End Sub
    'Public Sub SetDCKey(dcKey__1 As [String])
    '    DCKey = dcKey__1
    'End Sub
    'Public Sub SetuserRoot(userroot__1 As [String])
    '    userRoot = userroot__1
    'End Sub
    ''Set-Methods END

    'Private Sub Determine3264bitOS()
    '    Dim SOFTWARE_KEY As String = "Software"
    '    Dim COMPANY_NAME As String = "Wow6432Node"
    '    Dim win3264 As RegistryKey = Registry.LocalMachine.OpenSubKey(SOFTWARE_KEY, False).OpenSubKey(COMPANY_NAME, False)
    '    If win3264 IsNot Nothing Then
    '        SetuserRoot("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\THQ\")
    '    Else
    '        SetuserRoot("HKEY_LOCAL_MACHINE\SOFTWARE\THQ\")
    '    End If
    'End Sub

    ''Get-Methods START: used in RegWriter() to get desired values of vars and registry entries
    'Private Function getClassicKey() As [String]
    '    Return ClassicKey
    'End Function
    'Private Function getWAKey() As [String]
    '    Return WAKey
    'End Function
    'Private Function getDCKey() As [String]
    '    Return DCKey
    'End Function
    'Private Function getSSKey() As [String]
    '    SSKey = DirectCast(Registry.GetValue(getuserRoot() & "Dawn of War - Soulstorm", "CDKEY", RegistryValueKind.[String]), String)
    '    'TODO SSKey auslesen (32/64bit anpassung)
    '    Return SSKey
    'End Function
    'Private Function getuserRoot() As [String]
    '    Return userRoot
    'End Function
    'Private Function getInstallLocation() As [String]
    '    Dim IL As [String] = DirectCast(Registry.GetValue(getuserRoot() & "Dawn of War - Soulstorm", "InstallLocation", RegistryValueKind.[String]), String)

    '    Return IL
    'End Function
    ''Get-Methods END

    ''RegWriter START: writes Registry entries and also creates the fake exe files (should have been sepereated, bad style, i know, and now STFU
    'Public Sub RegWriter()
    '    Determine3264bitOS()

    '    Dim DoW As String = userRoot & "Dawn of War"
    '    'registry path to DoW+WA
    '    Dim DoWDC As String = userRoot & "Dawn of War - Dark Crusade"
    '    'registry path to DC
    '    Dim DoWSS As String = userRoot & "Dawn of War - Soulstorm"
    '    'registry path to SS
    '    'Classic+WA Stuff START: Writes the keys of Classic and WA into the registry + creates the fake exe files
    '    If getClassicKey() IsNot Nothing Then
    '        'Writes following registry entries into Classic registry dir: Classic CD-Key, InstallLocation of SS
    '        Registry.SetValue(DoW, "CDKEY", getClassicKey(), RegistryValueKind.[String])
    '        Registry.SetValue(DoW, "InstallLocation", getInstallLocation() & "\", RegistryValueKind.[String])
    '        'the Backslash at the end of the path here is a MUST, verification won't work otherwise => Reric make bug ;)
    '        'Create fake Exe of DoW Classic START
    '        If Not File.Exists(getInstallLocation() & "\W40k.exe") Then
    '            'aint I a smart-ass? *g* fake exe like DCUnlock used, don't work, so this was the easiest solution i could come up with
    '            File.Copy(getInstallLocation() & "\GraphicsConfig.exe", getInstallLocation() & "\W40k.exe")
    '            'Create fake Exe of DoW Classic END

    '        End If
    '    End If


    '    If getWAKey() IsNot Nothing Then
    '        Registry.SetValue(DoW, "CDKEY_WXP", getWAKey(), RegistryValueKind.[String])

    '        If Not File.Exists(getInstallLocation() & "\W40kWA.exe") Then
    '            File.Copy(getInstallLocation() & "\GraphicsConfig.exe", getInstallLocation() & "\W40kWA.exe")
    '        End If
    '    End If
    '    'Classic+WA Stuff END

    '    'DC Stuff START: Writes the key of DC into the registry + creates the fake exe file
    '    If getDCKey() IsNot Nothing Then
    '        Registry.SetValue(DoWDC, "CDKEY", getDCKey(), RegistryValueKind.[String])
    '        Registry.SetValue(DoWDC, "InstallLocation", getInstallLocation(), RegistryValueKind.[String])

    '        If Not File.Exists(getInstallLocation() & "\DarkCrusade.exe") Then
    '            File.Copy(getInstallLocation() & "\GraphicsConfig.exe", getInstallLocation() & "\DarkCrusade.exe")




    '        End If
    '    End If
    '    'DC Stuff END

    '    'SS Stuff START: Writes the keys of Classic, WA and DC into the registry of Soulstorm
    '    If getSSKey() IsNot Nothing Then

    '        Registry.SetValue(DoWSS, "W40KCDKEY", getClassicKey(), RegistryValueKind.[String])
    '        Registry.SetValue(DoWSS, "WXPCDKEY", getWAKey(), RegistryValueKind.[String])

    '        Registry.SetValue(DoWSS, "DXP2CDKEY", getDCKey(), RegistryValueKind.[String])
    '    End If
    '    'SS Stuff END
    'End Sub
    ''RegWriter END
#End Region
End Class
