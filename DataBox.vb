

Imports System.Data.Common 'database

''' <summary>
''' Database Input/Output operations
''' </summary>
''' <remarks></remarks>
Public MustInherit Class DataBox
    Private _Log As Log = Nothing

    Protected _Constants As String = ""
    Protected _Connection As DbConnection
    Protected _Command As DbCommand
    Protected _Reader As DbDataReader

    ''' <summary>
    ''' Database connection string
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public MustOverride ReadOnly Property ConnectionString() As String

    ''' <summary>
    ''' Connection string constants to be appended
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ConnectionConstants() As String
        Get
            Return _Constants
        End Get
        Set(ByVal value As String)
            _Constants = value
        End Set
    End Property

    ''' <summary>
    ''' Database log
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Log() As Log
        Get
            Return _Log
        End Get
        Set(ByVal value As Log)
            _Log = value
        End Set
    End Property

    ''' <summary>
    ''' Open the database for read/write access and return success
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public MustOverride Function Open() As Boolean

    ''' <summary>
    ''' Checks if the database is already open
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsOpen() As Boolean
        Try
            If _Connection.State = ConnectionState.Open Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Read from the database and return a reader
    ''' </summary>
    ''' <param name="strSQL">SQL statement used in read</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public MustOverride Function ReadRaw(ByVal strSQL As String) As DbDataReader

    ''' <summary>
    ''' Read from the database and return an arraylist of values
    ''' </summary>
    ''' <param name="strSQL">SQL statement used in read</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public MustOverride Function Read(ByVal strSQL As String) As ArrayList

    ''' <summary>
    ''' Write to the database and return records affected
    ''' </summary>
    ''' <param name="strSQL">SQL statement used in write</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' 
    Public MustOverride Function Write(ByVal strSQL As String) As Integer

    ''' <summary>
    ''' Close the database and return success
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public MustOverride Function Close() As Boolean

    ''' <summary>
    ''' Check if the database is closed
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsClose() As Boolean
        If _Connection.State = ConnectionState.Closed Then
            Return True
        Else
            Return False
        End If
    End Function
End Class