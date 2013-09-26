

Imports System.Data.OleDb

Public Class Access_DB
    Inherits DataBox

    Private _Provider As String = "Microsoft.ACE.OLEDB.12.0"
    Private _DataSource As String = ""
    Private _Password As String = ""

    Public Const Version As String = "0.1.0" 'major.minor.build

    ''' <summary>
    ''' Database provider
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Provider() As String
        Get
            Return _Provider
        End Get

        Set(ByVal value As String)
            _Provider = value
        End Set
    End Property

    ''' <summary>
    ''' Database source
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DataSource() As String
        Get
            Return _DataSource
        End Get

        Set(ByVal value As String)
            _DataSource = value
        End Set
    End Property

    ''' <summary>
    ''' Database password, if any
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Password() As String
        Get
            Return _Password
        End Get

        Set(ByVal value As String)
            _Password = value
        End Set
    End Property

    Public Overrides ReadOnly Property ConnectionString() As String
        Get
            Dim _ConnectionString As String = ""
            _ConnectionString = String.Format("Provider={0}; Data Source={1}; ", _Provider, _DataSource)
            If _Password.Length <> 0 Then
                _ConnectionString &= String.Format("Jet OLEDB:Database Password={0};", _Password)
            Else
                _ConnectionString &= String.Format("Persist Security Info=False;", _Password)
            End If
            If _Constants.Length <> 0 Then
                _ConnectionString &= _Constants
            End If
            Return _ConnectionString
        End Get
    End Property

    Public Overrides Function Open() As Boolean
        _Connection = New OleDbConnection(ConnectionString)
        Try
            If Not Log Is Nothing Then
                Log.Update(Now, "Open - " & DataSource)
            End If
            _Connection.Open()
        Catch ex As Exception
        End Try
        Return IsOpen()
        Return False
    End Function

    Public Overrides Function ReadRaw(ByVal strSQL As String) As System.Data.Common.DbDataReader
        If strSQL.Length = 0 Or IsOpen() = False Then
            Return Nothing
        Else
            _Command = New OleDbCommand(strSQL, _Connection)
            Try
                _Reader = _Command.ExecuteReader
            Catch ex As Exception
                Return Nothing
            End Try
            Return _Reader
        End If
    End Function

    Public Overrides Function Read(ByVal strSQL As String) As System.Collections.ArrayList
        Dim oDataReader As OleDb.OleDbDataReader
        If Not Log Is Nothing Then
            Log.Update(Now, strSQL)
        End If
        oDataReader = ReadRaw(strSQL)
        Dim AllRows As New ArrayList
        If Not oDataReader Is Nothing Then
            Do While oDataReader.Read()
                Dim i As Integer, AllColumns As New ArrayList
                For i = 0 To oDataReader.FieldCount - 1
                    AllColumns.Add(oDataReader(i))
                Next
                AllRows.Add(AllColumns)
            Loop
            Return AllRows
        Else
            Dim NoRows As New ArrayList
            Return NoRows
        End If
    End Function

    Public Overrides Function Write(ByVal strSQL As String) As Integer
        If strSQL.Length = 0 Or IsOpen() = False Then
            Return -1
        Else
            Dim iRows As Integer
            _Command = New OleDbCommand(strSQL, _Connection)
            Try
                If Not Log Is Nothing Then
                    Log.Update(Now, strSQL)
                End If
                iRows = _Command.ExecuteNonQuery
            Catch ex As Exception
                Return -1
            End Try
            Return iRows
        End If
    End Function

    Public Overrides Function Close() As Boolean
        If Not Log Is Nothing Then
            Log.Update(Now, "Close - " & DataSource)
        End If
        _Connection.Close()
        Return IsClose()
    End Function
End Class