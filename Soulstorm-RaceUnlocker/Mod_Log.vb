
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

Imports System.Text
Imports System.IO

Public Enum PRÄFIX
    INFO = 0
    WARNING = 1
    EXCEPTION = 2
End Enum

Module Mod_Log
#Region "Declarations"
    '// The complete logfile name including the type.
    Private _LogfileName As String = "Race Unlocker Log.log"
    '// Contains the complete log.
    Private _Log As New StringBuilder
#End Region

#Region "Propertys"
    ''' <summary>Returns the whole Log.</summary>
    Public ReadOnly Property GetLog As String
        Get
            Return _Log.ToString
        End Get
    End Property
    ''' <summary>Returns the name of the logfile including the type.</summary>
    Public ReadOnly Property GetLogfileName As String
        Get
            Return _LogfileName
        End Get
    End Property
#End Region
    
    Public Sub Initialize_Log()
        Dim _StrWrt As New StreamWriter(My.Computer.FileSystem.SpecialDirectories.Desktop + "\" + _LogfileName, True, Encoding.ASCII) With {.AutoFlush = True}
        If File.Exists(My.Computer.FileSystem.SpecialDirectories.Desktop + "\" + _LogfileName) Then
            _StrWrt.WriteLine("")
        End If
        _StrWrt.Dispose()
        _StrWrt.Close()
    End Sub

    ''' <summary>Adds a string to the log followed by a new line.</summary>
    Public Sub Log_Msg(_Präfix As PRÄFIX, _Message As String)
        Dim _StrWrt As New StreamWriter(My.Computer.FileSystem.SpecialDirectories.Desktop + "\" + _LogfileName, True, Encoding.ASCII) With {.AutoFlush = True}
        _Log.AppendLine(GetSyntax(_Präfix) + _Message)
        _StrWrt.WriteLine(GetSyntax(_Präfix) + _Message)
        _StrWrt.Close()
        _StrWrt.Dispose()
    End Sub

    ''' <summary>Adds a string to the log. NO new line!</summary>
    Public Sub Log_Msg_Clean(_Message As String)
        Dim _StrWrt As New StreamWriter(My.Computer.FileSystem.SpecialDirectories.Desktop + "\" + _LogfileName, True, Encoding.ASCII) With {.AutoFlush = True}
        _Log.Append(_Message)
        _StrWrt.Write(_Message)
        _StrWrt.Close()
        _StrWrt.Dispose()
    End Sub

    ''' <summary>Returns the header präfix for the logsystem.</summary>
    Private Function GetSyntax(_Präfix As PRÄFIX) As String
        Return DateTime.Now.ToString("| yyyy.dd.MM | HH:mm ss | ") + [Enum].GetName(GetType(PRÄFIX), _Präfix) + " | "
    End Function
End Module
