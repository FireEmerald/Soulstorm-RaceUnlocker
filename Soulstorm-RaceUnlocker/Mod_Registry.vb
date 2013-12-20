
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

Imports Microsoft.Win32
Imports System.Text.RegularExpressions

Public Structure GameData
    ''' <summary>Doesn't contain an Integer. Only for [ENUM].GetName(.GetType(GAME_ID), >Value.</summary>
    Dim ID As GAME_ID
    ''' <summary>The name of the GameEXE. (*.exe)</summary>
    Dim ExeName As String
    ''' <summary>The name of the Registry Directory of a game. (Dawn of War)</summary>
    Dim RegGameSubDirectory As String
    ''' <summary>The name of the entry of the gamekey. (CDKEY)</summary>
    Dim RegSerialNumberKeyName As String
    ''' <summary>The name of the entry of the InstallLocation. (InstallLocation)</summary>
    Dim RegInstallLocKeyName As String
End Structure

Module Mod_Registry

#Region "Declarations"
    '// Architecture-based THQ path (64/32 bit)
    '// All 32bit Registrykeys are redirected to 'HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node'
    '// All 64bit Registrykeys stored in 'HKEY_LOCAL_MACHINE\SOFTWARE'
    '// All Dawn of War (Soulstorm or older) Applications are 32bit, so they are always at the 32bit directory.
    Private Const _THQ_SubKey As String = "SOFTWARE\THQ"

    Private _BaseKey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)

    '// Default infos about the games.
    Public ReadOnly _DBClassic As New GameData With {.RegGameSubDirectory = "Dawn of War", .ExeName = "W40k.exe", .ID = GAME_ID.CLASSIC, _
                                                     .RegSerialNumberKeyName = "CDKEY", .RegInstallLocKeyName = "InstallLocation"}
    Public ReadOnly _DBWinterAssault As New GameData With {.RegGameSubDirectory = "Dawn of War", .ExeName = "W40kWA.exe", .ID = GAME_ID.WINTER_ASSAULT, _
                                                           .RegSerialNumberKeyName = "CDKEY_WXP", .RegInstallLocKeyName = "InstallLocation"}
    Public ReadOnly _DBDarkCrusade As New GameData With {.RegGameSubDirectory = "Dawn of War - Dark Crusade", .ExeName = "DarkCrusade.exe", .ID = GAME_ID.DARK_CRUSADE, _
                                                         .RegSerialNumberKeyName = "CDKEY", .RegInstallLocKeyName = "InstallLocation"}
    Public ReadOnly _DBSoulstorm As New GameData With {.RegGameSubDirectory = "Dawn of War - Soulstorm", .ExeName = "Soulstorm.exe", .ID = GAME_ID.SOULSTORM, _
                                                       .RegSerialNumberKeyName = "CDKEY", .RegInstallLocKeyName = "InstallLocation"}
#End Region

#Region "Program-specific functions"
    ''' <summary>Returns the GameKey of a Game. If the entry was found "" will be returned.</summary>
    Public Function GetRegGameKey(_Game As GameData) As String
        Log_Msg(PRÄFIX.INFO, "Registry - GetRegGameKey - (Prepare) Get Key | Directory: (" + _BaseKey.ToString + ") """ + _THQ_SubKey + "\" + _Game.RegGameSubDirectory + """ | Type: """ + _Game.RegSerialNumberKeyName + """")
        Return RegReadKey(_BaseKey, _THQ_SubKey + "\" + _Game.RegGameSubDirectory, _Game.RegSerialNumberKeyName)
    End Function

    ''' <summary>Returns the InstallLocation of a Game. If the entry was found "" will be returned.</summary>
    Public Function GetRegInstallDirectory(_Game As GameData) As String
        Log_Msg(PRÄFIX.INFO, "Registry - GetRegInstallDirectory - (Prepare) Get InstallLocation | Directory: (" + _BaseKey.ToString + ") """ + _THQ_SubKey + "\" + _Game.RegGameSubDirectory + """ | Type: """ + _Game.RegSerialNumberKeyName + """")
        Return RegReadKey(_BaseKey, _THQ_SubKey + "\" + _Game.RegGameSubDirectory, _Game.RegInstallLocKeyName)
    End Function

    ''' <summary>Creates a new registry key in the given "SOFTWARE\THQ\" >SubDirectory. If the key exists, he will be overwritten.</summary>
    Public Function CreateRegKey(_KeyName As String, _NewValue As String, Optional _SubDirectory As String = "") As Boolean
        '// If the NewValue is a Serialnumber, mask it.
        Dim _LogNewValue As String = _NewValue
        If Regex.IsMatch(_NewValue, fmMain._GameKeyPattern_4) Or Regex.IsMatch(_NewValue, fmMain._GameKeyPattern_5) Then _LogNewValue = _NewValue.Substring(0, _NewValue.LastIndexOf("-")) + "-XXXX"
        Log_Msg(PRÄFIX.INFO, "Registry - CreateRegKey - (Prepare) Create New  | KeyName: """ + _KeyName + """ | With Value: """ + _LogNewValue + """ | Sub Directory: (" + _BaseKey.ToString + "\" + _THQ_SubKey + ") """ + _SubDirectory + """")

        If _SubDirectory = "" Then Return RegCreateKey(_BaseKey, _THQ_SubKey, _KeyName, _NewValue)
        Return RegCreateKey(_BaseKey, _THQ_SubKey + "\" + _SubDirectory, _KeyName, _NewValue)
    End Function

    ''' <summary>Creates a SubDirectory in "SOFTWARE\" >ExistingSubDirectory \ >NewSubDirectory.</summary>
    Public Function CreateRegDirectory(_NewSubDirectory As String, Optional _ExistingSubDirectory As String = "") As Boolean
        Log_Msg(PRÄFIX.INFO, "Registry - CreateRegDirectory - (Prepare) Create New | NewSubDirectory: """ + _NewSubDirectory + """ | ExistingSubDirectory: (SOFTWARE) """ + _ExistingSubDirectory + """")

        If _ExistingSubDirectory = "" Then Return RegCreateDirectory(_BaseKey, "SOFTWARE", _NewSubDirectory)
        Return RegCreateDirectory(_BaseKey, "SOFTWARE" + "\" + _ExistingSubDirectory, _NewSubDirectory)
    End Function

    ''' <summary>Deletes the SOFTWARE\THQ directory with all sub directorys/keys.</summary>
    Public Function DeleteRegTHQ() As Boolean
        Log_Msg(PRÄFIX.INFO, "Registry - DeleteRegTHQ - (Prepare) Delete the THQ registry directory | Directory: """ + _BaseKey.ToString + "\" + _THQ_SubKey + """")
        Return RegDeleteDirectory(_BaseKey, _THQ_SubKey)
    End Function
#End Region

#Region "Functions"
    ''' <summary>Reads the value of a registry key. If the key doesn't exist, a empty string will be returned.</summary>
    ''' <param name="_RegBaseKey">The base directory which contains the sub directory.</param>
    ''' <param name="_RegSubDirectory">The sub directory in the base directory which should contain the key.</param>
    ''' <param name="_RegKeyName">Name of the registry key.</param>
    ''' <returns>Returns the registry value. If no value was found, a empty string is returned.</returns>
    Private Function RegReadKey(_RegBaseKey As RegistryKey, _RegSubDirectory As String, _RegKeyName As String) As String
        Dim _RegValue As String
        Try
            _RegValue = _RegBaseKey.OpenSubKey(_RegSubDirectory).GetValue(_RegKeyName, "").ToString
        Catch ex As Exception
            Dim _Directory As String = "NOTHING"
            Try
                _Directory = _RegBaseKey.OpenSubKey(_RegSubDirectory).ToString
            Catch
            Finally
                Dim _Base As String = "NOTHING"
                If Not IsNothing(_RegBaseKey) Then _Base = _RegBaseKey.ToString()
                If IsNothing(_RegSubDirectory) Then _RegSubDirectory = "NOTHING"
                If IsNothing(_RegKeyName) Then _RegKeyName = "NOTHING"

                Log_Msg(PRÄFIX.EXCEPTION, "Registry - RegRead - Exception occured | @RegDirectory: (Base: """ + _Base + """ | Sub: """ + _RegSubDirectory + """) """ + _Directory.ToString + """ @RegKeyName: """ + _RegKeyName + """")
                Log_Msg(PRÄFIX.EXCEPTION, "Registry - RegRead - Exception occured | Exception Msg: """ + ex.ToString + """")
            End Try
            Return ""
        End Try
        Return _RegValue
    End Function

    ''' <summary>Creates a new registry key. If the key already exists, the key will be overwritten.</summary>
    ''' <param name="_RegBaseKey">The base directory which contains the sub directory.</param>
    ''' <param name="_RegSubDirectory">The sub directory in the base directory which should contain the key.</param>
    ''' <param name="_RegKeyName">Name of the registry key.</param>
    ''' <param name="_NewValue">Value for the registry key.</param>
    ''' <param name="_Type">Type for the registry key.</param>
    ''' <returns>Returns True if the registry key was successfully created/overwritten.</returns>
    Private Function RegCreateKey(_RegBaseKey As RegistryKey, _RegSubDirectory As String, _RegKeyName As String, _NewValue As String, Optional _Type As RegistryValueKind = RegistryValueKind.String) As Boolean
        Try
            _RegBaseKey.OpenSubKey(_RegSubDirectory, True).SetValue(_RegKeyName, _NewValue, _Type)
        Catch ex As Exception
            Dim _Directory As String = "NOTHING"
            Try
                '// "_Directory" still has the "NOTHING" value, if this fails. So we don't need to check if _Directory is Nothing.
                _Directory = _RegBaseKey.OpenSubKey(_RegSubDirectory).ToString
            Catch '// If we don't catch the exception, the application crashs. So we need it...
            Finally
                Dim _Base As String = "NOTHING"
                If Not IsNothing(_RegBaseKey) Then _Base = _RegBaseKey.ToString
                If IsNothing(_RegSubDirectory) Then _RegSubDirectory = "NOTHING"
                If IsNothing(_RegKeyName) Then _RegKeyName = "NOTHING"
                If IsNothing(_NewValue) Then _NewValue = "NOTHING"

                Log_Msg(PRÄFIX.EXCEPTION, "Registry - RegSet - Exception occured | @RegDirectory: (Base: """ + _Base + """ | Sub: """ + _RegSubDirectory + """) """ + _Directory + """ | @RegKeyName: """ + _RegKeyName + """ | @NewValue: """ + _NewValue + """")
                Log_Msg(PRÄFIX.EXCEPTION, "Registry - RegSet - Exception occured | Exception Msg: """ + ex.ToString + """")
            End Try
            Return False
        End Try
        Return True
    End Function

    ''' <summary>Creates a SubDirectory in the given BaseDirectory depending on the SubDirectory.</summary>
    ''' <param name="_RegBaseKey">The base directory which contains the sub directory.</param>
    ''' <param name="_RegSubDirectory">The sub directory in the base directory.</param>
    ''' <param name="_NewSubDirectory">The name of the new sub directory.</param>
    ''' <returns>Returns True if the directory was successfully created.</returns>
    Private Function RegCreateDirectory(_RegBaseKey As RegistryKey, _RegSubDirectory As String, _NewSubDirectory As String) As Boolean
        Try
            _RegBaseKey.OpenSubKey(_RegSubDirectory, True).CreateSubKey(_NewSubDirectory)
        Catch ex As Exception
            Dim _Directory As String = "NOTHING"
            Try
                _Directory = _RegBaseKey.OpenSubKey(_RegSubDirectory).ToString
            Catch
            Finally
                Dim _Base As String = "NOTHING"
                If Not IsNothing(_RegBaseKey) Then _Base = _RegBaseKey.ToString
                If IsNothing(_RegSubDirectory) Then _RegSubDirectory = "NOTHING"
                If IsNothing(_NewSubDirectory) Then _NewSubDirectory = "NOTHING"

                Log_Msg(PRÄFIX.EXCEPTION, "Registry - RegCreateDirectory - Exception occured | @RegDirectory: (Base: """ + _Base + """ | Sub: """ + _RegSubDirectory + """) """ + _Directory.ToString + """ | @NewSubDirectory: """ + _NewSubDirectory + """")
                Log_Msg(PRÄFIX.EXCEPTION, "Registry - RegCreateDirectory - Exception occured | Exception Msg: """ + ex.ToString + """")
            End Try
            Return False
        End Try
        Return True
    End Function

    ''' <summary>Deletes a complete directory. Including all sub keys/directorys.</summary>
    ''' <param name="_RegBaseDirectory">The directory which contains the sub directory which should be deleted.</param>
    ''' <param name="_DelDirectory">The sub directory of the base directory, which should be deleted.</param>
    ''' <returns>Returns True if deleted successfully. Otherwise you will get False.</returns>
    Private Function RegDeleteDirectory(_RegBaseDirectory As RegistryKey, _DelDirectory As String) As Boolean
        Try
            _RegBaseDirectory.DeleteSubKeyTree(_DelDirectory)
        Catch ex As Exception
            Dim _Directory As String = "NOTHING"
            If Not IsNothing(_RegBaseDirectory) Then _Directory = _RegBaseDirectory.ToString
            If IsNothing(_DelDirectory) Then _DelDirectory = "NOTHING"

            Log_Msg(PRÄFIX.EXCEPTION, "Registry - RegDeleteDirectory - Exception occured | @RegBaseDirectory: """ + _Directory.ToString + """ | @RegDelDirectory: """ + _DelDirectory + """")
            Log_Msg(PRÄFIX.EXCEPTION, "Registry - RegDeleteDirectory - Exception occured | Exception Msg: """ + ex.ToString + """")
            Return False
        End Try
        Return True
    End Function
#End Region

#Region "Registry Permission Test"
    Public Function RegistryPermissionTest() As String
        Dim _TestKey_Name As String = "What does the Khorne Berzerker say"
        Dim _TestKey_Value As String = "BLOOD FOR THE BLOOD GOD"
        Dim _TestKey_Directory As String = "CHAOS"

        Log_Msg(PRÄFIX.INFO, "Registry - RegistryPermissionTest - Start")
        Log_Msg(PRÄFIX.INFO, "Registry - RegistryPermissionTest - Create TestKey, directory: """ + _BaseKey.ToString + "\SOFTWARE\" + _TestKey_Directory + """ | Key: """ + _TestKey_Name + """ | Value: """ + _TestKey_Value + """")
        Try
            '// Create a new registry directory with a new key.
            _BaseKey.OpenSubKey("SOFTWARE", True).CreateSubKey(_TestKey_Directory).SetValue(_TestKey_Name, _TestKey_Value)
            '// Exception handling ...
        Catch ex As Security.SecurityException : Return "You don't have the required permissions to create or modify registry keys." + GetException(ex)
        Catch ex As UnauthorizedAccessException : Return "The RegistryKey is read-only, and thus cannot be written to; for example, it is a root-level node." + GetException(ex)
        Catch ex As ArgumentException : Return """KeyName"" does not begin with a valid registry root or ""KeyName"" is longer than the maximum length allowed (255 characters)." + GetException(ex)
        Catch ex As Exception : Return "Something went wrong while accessing your registry (write)." + GetException(ex)
        End Try

        Log_Msg(PRÄFIX.INFO, "Registry - RegistryPermissionTest - Read TestKey value, from directory: """ + _BaseKey.ToString + "\SOFTWARE\" + _TestKey_Directory + """ | Key: """ + _TestKey_Name + """")
        Dim _RegValue As String = ""
        Try
            '// Try to read the new registered keyValue.
            _RegValue = _BaseKey.OpenSubKey("SOFTWARE\" + _TestKey_Directory).GetValue(_TestKey_Name).ToString
            '// Exception handling ...
        Catch ex As Security.SecurityException : Return "You don't have the required permissions to read from the registry." + GetException(ex)
        Catch ex As ArgumentException : Return "The keyName doesn't begin with a valid registry root." + GetException(ex)
        Catch ex As Exception : Return "Something went wrong while accessing your registry (read)." + GetException(ex)
        End Try

        If Not _RegValue = _TestKey_Value Then Return "Something went wrong while accessing your registry (check)."

        Log_Msg(PRÄFIX.INFO, "Registry - RegistryPermissionTest - Delete TestKey directory: """ + _BaseKey.ToString + "\SOFTWARE\" + _TestKey_Directory + """")
        Try
            '// Delete the registered key and the directory.
            _BaseKey.OpenSubKey("SOFTWAREX", True).DeleteSubKeyTree(_TestKey_Directory)
            '// Exception handling ...
        Catch ex As Security.SecurityException : Return "You don't have the required permissions to delete a registry key." + GetException(ex)
        Catch ex As UnauthorizedAccessException : Return "You don't have the necessary registry rights." + GetException(ex)
        Catch ex As ArgumentNullException : Return "subkey is Nothing." + GetException(ex)
        Catch ex As ArgumentException : Return "Deletion of a root hive is attempted or subkey does not specify a valid registry subkey." + GetException(ex)
        Catch ex As Exception : Return "Something went wrong while accessing your registry (delete)." + GetException(ex)
        End Try

        Return ""
    End Function

    Private Function GetException(ex As Exception) As String
        Return vbCrLf + vbCrLf + "Please login as administrator and start the Application with:" + vbCrLf + _
                                 """Right click"" -> ""Start as administrator"" !" + _
                                 "%ex%" + ex.ToString
    End Function
#End Region
End Module
