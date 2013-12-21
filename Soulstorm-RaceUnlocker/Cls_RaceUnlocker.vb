
'* Copyright (C) 2013 FireEmerald <https://github.com/FireEmerald>
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

Option Explicit On
Option Strict On

Imports System.IO

Public Class Cls_RaceUnlocker

#Region "Declarations"
    '// Sub Directorys of the Game *.exe files.
    Private _SubDirClassic As String = "Unlocker\Classic and Winter Assault"
    Private _SubDirDarkCrusade As String = "Unlocker\Dark Crusade"

    '// Status of the registry and exe unlock process.
    Private _RegistryUnlockStatus As String = ""
    Private _ExeUnlockStatus As String = ""

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

#Region "Propertys"
    Public ReadOnly Property GetRegistryUnlockStatus As String
        Get
            Return _RegistryUnlockStatus
        End Get
    End Property
    Public ReadOnly Property GetExeUnlockStatus As String
        Get
            Return _ExeUnlockStatus
        End Get
    End Property
#End Region

#Region "Registry part"
    ''' <summary>Edit the registry entrys of the games.</summary>
    Public Sub Unlock_Registry()
        Log_Msg(PRÄFIX.INFO, "Unlock - Start - Classic: """ + _ClassicGameKey.Substring(0, _ClassicGameKey.LastIndexOf("-")) + "-XXXX"" | " + _
                                              "Winter Assault: """ + _WinterAssaultGameKey.Substring(0, _WinterAssaultGameKey.LastIndexOf("-")) + "-XXXX"" | " + _
                                              "Dark Crusade: """ + _DarkCrusadeGameKey.Substring(0, _DarkCrusadeGameKey.LastIndexOf("-")) + "-XXXX"" | " + _
                                              "Soulstorm: " + _SoulstormGameKey.Substring(0, _SoulstormGameKey.LastIndexOf("-")) + "-XXXX"" | " + _
                                              "Soulstorm InstallLocation: """ + _SoulstormInstallationPath + """")

        _RegistryUnlockStatus = ""
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

            _RegistryUnlockStatus = "done"
        Else
            Log_Msg(PRÄFIX.EXCEPTION, "Unlock - Permission Test - Failed | Exception Msg: """ + _PermissionTestResult.Substring(_PermissionTestResult.LastIndexOf("%ex%")).Replace("%ex%", "") + """")
            _RegistryUnlockStatus = _PermissionTestResult.Substring(0, _PermissionTestResult.IndexOf("%ex%"))
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

    ''' <summary>Returns True, if the registry path of the game is correct (*.exe exists).</summary>
    Private Function IsInstalled(_Game As GameData) As Boolean
        If File.Exists(GetRegInstallDirectory(_Game) + "\" + _Game.ExeName) Then Return True
        Return False
    End Function
#End Region

#Region "Exe part"
    ''' <summary>Dublicate the GraphicsConfig.exe of Soulstorm and rename them like the addons.</summary>
    Public Sub Unlock_Exe()
        _ExeUnlockStatus = "no_error"
        '// Exe unlock Process
        If Directory.Exists(_SoulstormInstallationPath) Then
            Log_Msg(PRÄFIX.INFO, "Unlock - Exe Unlock - Directory Exists | Path: """ + _SoulstormInstallationPath + """")
            '// ...\Dawn of War - Soulstorm\Unlocker
            CreateDirectory(_SoulstormInstallationPath + "\Unlocker")
            '// ...\Dawn of War - Soulstorm\Unlocker\Classic and Winter Assault
            CreateDirectory(_SoulstormInstallationPath + "\" + _SubDirClassic)
            CreateApplication(_SoulstormInstallationPath + "\" + _SubDirClassic, _DBClassic)
            CreateApplication(_SoulstormInstallationPath + "\" + _SubDirClassic, _DBWinterAssault)
            '// ...\Dawn of War - Soulstorm\Unlocker\Dark Crusade
            CreateDirectory(_SoulstormInstallationPath + "\" + _SubDirDarkCrusade)
            CreateApplication(_SoulstormInstallationPath + "\" + _SubDirDarkCrusade, _DBDarkCrusade)
        Else
            Log_Msg(PRÄFIX.EXCEPTION, "Unlock - DirectoryCheck - Not Found | Directory: """ + _SoulstormInstallationPath + """")
            _ExeUnlockStatus = "The path to your Dawn of War Soulstorm directory does not exist."
        End If
    End Sub

    ''' <summary>Create a directory with the given path. If the directory already exists, True will be returned.</summary>
    Private Function CreateDirectory(_Path As String) As Boolean
        Try
            Directory.CreateDirectory(_Path)
        Catch ex As Exception
            If IsNothing(_Path) Then _Path = "NOTHING"

            Log_Msg(PRÄFIX.EXCEPTION, "Unlock - CreateDirectory - Exeption | Path: """ + _Path + """")
            Log_Msg(PRÄFIX.EXCEPTION, "Unlock - CreateDirectory - Exeption | Message: """ + ex.ToString + """")
            _ExeUnlockStatus = "A error occured while creating a new directory in your Dawn of War directory." + vbCrLf + vbCrLf + _
                               "Check the logfile for more informations."
            Return False
        End Try
        Log_Msg(PRÄFIX.INFO, "Unlock - CreateDirectory - Successful | Path: """ + _Path + """")
        Return True
    End Function

    ''' <summary>Creates the game *.exe in the given path with the given game.</summary>
    Private Function CreateApplication(_Path As String, _Game As GameData) As Boolean
        Try
            File.WriteAllBytes(_Path + "\" + _Game.ExeName, My.Resources.GraphicsConfig)
        Catch ex As Exception
            If IsNothing(_Path) Then _Path = "NOTHING"
            If IsNothing(_Game) Then _Game = New GameData With {.ExeName = "NOTHING", .ID = GAME_ID.NOT_SET}

            Log_Msg(PRÄFIX.EXCEPTION, "Unlock - CreateApplication - Exception | Path/Game: """ + _Path + """\""" + _Game.ExeName + """")
            Log_Msg(PRÄFIX.EXCEPTION, "Unlock - CreateApplication - Exception | Message: """ + ex.ToString + """")
            _ExeUnlockStatus = "A error occured while copying the fake exe files in your Dawn of War Soulstorm directory." + vbCrLf + vbCrLf + _
                               "Check the logfile for more informations."
            Return False
        End Try
        Log_Msg(PRÄFIX.INFO, "Unlock - CreateApplication - Successful | Path: """ + _Path + """\""" + _Game.ExeName + """")
        Return True
    End Function
#End Region
End Class
