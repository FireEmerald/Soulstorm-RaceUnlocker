
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

Option Explicit On
Option Strict On

Imports System.Text
Imports System.IO

Public Enum PREFIX
    INFO = 0
    WARNING = 1
    EXCEPTION = 2
End Enum

Module Mod_Log

    Private _FiLogfile As FileInfo       '// FileInfo of the log file.
    Private _LogBuilder As StringBuilder '// Stores the complete log file.

#Region "Properties"
    ''' <summary>Returns the whole Log.</summary>
    Public ReadOnly Property GetLog As String
        Get
            Return _LogBuilder.ToString
        End Get
    End Property
    ''' <summary>Returns the full path to the log file including it's type.</summary>
    Public ReadOnly Property GetFullLogfilePath As String
        Get
            Return _FiLogfile.FullName
        End Get
    End Property
#End Region

    Public Sub Initialize_Log()
        _LogBuilder = New StringBuilder
        _FiLogfile = New FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Race Unlocker Log.log"))

        Dim _strWrt As New StreamWriter(_fiLogfile.FullName, True, Encoding.ASCII) With {.AutoFlush = True}
        If _FiLogfile.Exists Then
            _strWrt.WriteLine("")
        End If
        _strWrt.Dispose()
        _strWrt.Close()
    End Sub

    ''' <summary>Adds a string to the log, followed by a new line.</summary>
    Public Sub Log_Msg(_Prefix As PREFIX, _Message As String, Optional _Args As Object() = Nothing)
        Dim _strWrt As New StreamWriter(_FiLogfile.FullName, True, Encoding.ASCII) With {.AutoFlush = True}
        If IsNothing(_Args) Then
            _LogBuilder.AppendLine(GetSyntax(_Prefix) & _Message)
            _strWrt.WriteLine(GetSyntax(_Prefix) & _Message)
        Else
            _LogBuilder.AppendLine(GetSyntax(_Prefix) & String.Format(_Message, _Args))
            _strWrt.WriteLine(GetSyntax(_Prefix) & String.Format(_Message, _Args))
        End If
        _strWrt.Close()
        _strWrt.Dispose()
    End Sub

    ''' <summary>Returns the header prefix for the log system.</summary>
    Private Function GetSyntax(_Prefix As PREFIX) As String
        Return DateTime.Now.ToString("| yyyy.dd.MM | HH:mm ss | ") & [Enum].GetName(GetType(PREFIX), _Prefix) & " | "
    End Function
End Module
