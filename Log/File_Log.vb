
Imports System.IO 'file io

Public Class File_Log
    Inherits Log

    Private _Path As String
    Private _IsAppending As Boolean = True
    Private _AppendIndex As Integer = 0 'where last append stopped

    Private _FStream As FileStream
    Private _FSReader As StreamReader
    Private _FSWriter As StreamWriter

    Public Property Path() As String
        Get
            Return _Path
        End Get

        Set(ByVal value As String)
            _Path = value
        End Set
    End Property

    Public Property IsAppending() As Boolean
        Get
            Return _IsAppending
        End Get

        Set(ByVal value As Boolean)
            _AppendIndex = 0
            _IsAppending = value
        End Set
    End Property

    Public Function Load() As Integer
        Try
            Dim iCount As Integer = 0
            _FStream = New FileStream(_Path, FileMode.Open, FileAccess.Read)
            _FSReader = New StreamReader(_FStream)

            Do Until _FSReader.EndOfStream = True
                Add(_FSReader.ReadLine())
                iCount += 1
            Loop

            _FSReader.Close()
            _FStream.Close()
            Return iCount
        Catch ex As Exception
            Return -1
        End Try
    End Function

    Public Function Save() As Integer
        Try
            If _IsAppending Then
                _FStream = New FileStream(_Path, FileMode.Append, FileAccess.Write)
            Else
                _FStream = New FileStream(_Path, FileMode.Create, FileAccess.Write)
            End If
            _FSWriter = New StreamWriter(_FStream)

            For i As Integer = _AppendIndex To Count - 1
                _FSWriter.WriteLine(Item(i))
                _AppendIndex += 1
            Next

            _FSWriter.Close()
            _FStream.Close()
            Return _AppendIndex
        Catch ex As Exception
            Return -1
        End Try
    End Function

    Sub New(Optional ByVal sPath As String = "", Optional ByVal bIsAppending As Boolean = True)
        Path = sPath
        IsAppending = bIsAppending
    End Sub
End Class
