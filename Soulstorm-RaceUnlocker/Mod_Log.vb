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
    Private _LogfileName As String = "Race Unlocker Log.log"
    '// Contains the complete log.
    Private _Log As New StringBuilder
    '// Writes each line to a logfile on the desktop.

#End Region

    ''' <summary>Returns the whole Log.</summary>
    Public ReadOnly Property GetLog As String
        Get
            Return _Log.ToString
        End Get
    End Property

    Public Sub Close_LogStream(sender As Object, e As EventArgs)
        '_StrWrt.Close()
    End Sub

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
