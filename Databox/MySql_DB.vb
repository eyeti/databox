

Imports MySql.Data.MySqlClient

Public Class MySql_DB
    Inherits DataBox

    Private _Server As String = "localhost"
    Private _Port As Integer = 3306
    Private _Database As String = ""
    Private _UserId As String = ""
    Private _Password As String = ""

    Public Const Version As String = "0.1.0" 'major.minor.build

    ''' <summary>
    ''' Database server
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Server() As String
        Get
            Return _Server
        End Get

        Set(ByVal value As String)
            _Server = value
        End Set
    End Property

    ''' <summary>
    ''' Database server port, default is 3306
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Port() As Integer
        Get
            Return _Port
        End Get

        Set(ByVal value As Integer)
            _Port = value
        End Set
    End Property

    ''' <summary>
    ''' Database name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Database() As String
        Get
            Return _Database
        End Get

        Set(ByVal value As String)
            _Database = value
        End Set
    End Property

    ''' <summary>
    ''' Database user id
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property UserId() As String
        Get
            Return _UserId
        End Get

        Set(ByVal value As String)
            _UserId = value
        End Set
    End Property

    ''' <summary>
    ''' Database password
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
            _ConnectionString = String.Format("Server={0}; Port={1}; Database={2}; Uid= {3}; Pwd={4};", _Server, _Port, _Database, _UserId, _Password)
            If _Constants.Length <> 0 Then
                _ConnectionString &= _Constants
            End If
            Return _ConnectionString
        End Get
    End Property

    Public Overrides Function Open() As Boolean
        _Connection = New MySqlConnection(ConnectionString)
        Try
            If Not Log Is Nothing Then
                Log.Update(Now, "Open - " & Database & "@" & Server)
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
            _Command = New MySqlCommand(strSQL, _Connection)
            Try
                _Reader = _Command.ExecuteReader
            Catch ex As Exception
                Return Nothing
            End Try
            Return _Reader
        End If
    End Function

    Public Overrides Function Read(ByVal strSQL As String) As System.Collections.ArrayList
        Dim oDataReader As MySqlDataReader
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
            _Command = New MySqlCommand(strSQL, _Connection)
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
            Log.Update(Now, "Close - " & Database & "@" & Server)
        End If
        _Connection.Close()
        Return IsClose()
    End Function
End Class