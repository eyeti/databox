

Imports System.Data.SqlClient

Public Class MsSql_DB
    Inherits DataBox

    Private _DataSource As String
    Private _IntialCatalog As String
    Private _UseWindowsID As Boolean
    Private _UserId As String
    Private _Password As String

    Public Const Version As String = "0.1.0" 'major.minor.build

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
    ''' Database initial catalog / name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property IntialCatalog() As String
        Get
            Return _IntialCatalog
        End Get

        Set(ByVal value As String)
            _IntialCatalog = value
        End Set
    End Property

    ''' <summary>
    ''' Use user's windows login details?
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property UseWindowsID() As Boolean
        Get
            Return _UseWindowsID
        End Get

        Set(ByVal value As Boolean)
            _UseWindowsID = value
        End Set
    End Property

    ''' <summary>
    ''' Database username
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
            _ConnectionString = String.Format("Data Source={0}; Initial Catalog={1}; ", _DataSource, _IntialCatalog)
            If UseWindowsID Then
                _ConnectionString &= "Integrated Security=True;"
            Else
                _ConnectionString &= String.Format("User Id={0}; Password={1}; ", _UserId, _Password)
            End If
            If _Constants.Length <> 0 Then
                _ConnectionString &= _Constants
            End If
            Return _ConnectionString
        End Get
    End Property

    Public Overrides Function Open() As Boolean
        _Connection = New SqlClient.SqlConnection(ConnectionString)
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
            _Command = New SqlClient.SqlCommand(strSQL, _Connection)
            Try
                _Reader = _Command.ExecuteReader
            Catch ex As Exception
                Return Nothing
            End Try
            Return _Reader
        End If
    End Function

    Public Overrides Function Read(ByVal strSQL As String) As System.Collections.ArrayList
        Dim oDataReader As SqlClient.SqlDataReader
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
            _Command = New SqlClient.SqlCommand(strSQL, _Connection)
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
