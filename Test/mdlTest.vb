

Imports DataBox

Module mdlTest

    Sub Main()
        '===========================================================
        'general tests
        '===========================================================

        'Dim DBLog As New File_Log("db.log", True)

        Dim DBAccess As New Access_DB
        With DBAccess
            .DataSource = My.Computer.FileSystem.SpecialDirectories.Desktop & "\databoxes\access_db.accdb"
            '.Log = DBLog
        End With
        Console.WriteLine(String.Format("{0}: {1}: {2}", "Access_DB", Access_DB.Version, AboutAccess(DBAccess)))

        Dim DBExcel As New Excel_DB
        With DBExcel
            .DataSource = My.Computer.FileSystem.SpecialDirectories.Desktop & "\databoxes\excel_db.xlsx"
            .ConnectionConstants = "Extended Properties=""Excel 12.0 Xml;HDR=No;Imex=1"";"
            '.Log = DBLog
        End With
        Console.WriteLine(String.Format("{0}: {1}: {2}", "Excel_DB", Excel_DB.Version, AboutExcel(DBExcel)))

        'mssql
        'don't have an mssql installation currently so...

        Dim DBMySql As New MySql_DB
        With DBMySql
            .Database = "mysql_db"
            .UserId = "root"
            .Password = "some_password"
            .ConnectionConstants = "allow zero datetime=yes;"
            '.Log = DBLog
        End With
        Console.WriteLine(String.Format("{0}: {1}: {2}", "MySql_DB", MySql_DB.Version, AboutMySql(DBMySql)))

        Dim DBSqlite As New Sqlite_DB
        With DBSqlite
            .DataSource = My.Computer.FileSystem.SpecialDirectories.Desktop & "\databoxes\sqlite_db.s3db"
            .ConnectionConstants = "version=3; failifmissing=true;"
            '.Log = DBLog
        End With
        Console.WriteLine(String.Format("{0}: {1}: {2}", "Sqlite_DB", Sqlite_DB.Version, AboutSqlite(DBSqlite)))

        '===========================================================
        'Create, Read, Delete, Update tests
        '===========================================================


        Console.ReadKey()
    End Sub

#Region "Access_DB"
    Function AboutAccess(ByRef oObj As Access_DB) As String
        If Not oObj Is Nothing Then
            Dim strVersion() As String = Access_DB.Version.Split(".")
            If strVersion.Length <> 3 Then Return "Error! Check that version is 'major.minor.build'"
            Dim iMajor, iMinor, iBuild As Integer
            iMajor = strVersion(0)
            iMinor = strVersion(1)
            iBuild = strVersion(2)
            If oObj.Open = True Then
                Dim arrlstAbout As ArrayList = oObj.Read(String.Format("SELECT about, dependency FROM tblchangelog WHERE major={0} AND minor={1} AND build={2}", iMajor, iMinor, iBuild))
                oObj.Close()
                If arrlstAbout.Count <> 0 Then
                    If Not oObj.Log Is Nothing Then
                        CType(oObj.Log, File_Log).Save()
                    End If
                    Return arrlstAbout(0)(0) & " | " & arrlstAbout(0)(1)
                Else
                    Return "Error! Check that database is not empty"
                End If
            Else
                Return "Error! Check that database exists"
            End If
        End If
        Return "Error! Check that database is supplied"
    End Function

    Function ChangelogAccess(ByRef oObj As Access_DB) As Dictionary(Of String, String)
        If Not oObj Is Nothing Then
            Dim strVersion() As String = Access_DB.Version.Split(".")
            If strVersion.Length <> 3 Then Return Nothing
            Dim iMajor, iMinor, iBuild As Integer
            iMajor = strVersion(0)
            iMinor = strVersion(1)
            iBuild = strVersion(2)
            If oObj.Open = True Then
                Dim arrlstAbout As ArrayList = oObj.Read(String.Format("SELECT major, minor, build, changelog FROM tblchangelog WHERE major={0} AND minor={1} AND build={2}", iMajor, iMinor, iBuild))
                oObj.Close()
                If arrlstAbout.Count <> 0 Then
                    Dim dictReturned As New Dictionary(Of String, String)
                    For Each SingleRow In arrlstAbout
                        'major.minor.build | changelog
                        dictReturned.Add(SingleRow(0) & "." & SingleRow(1) & "." & SingleRow(2), SingleRow(3))
                    Next
                    Return dictReturned
                End If
            End If
        End If
        Return Nothing
    End Function
#End Region

#Region "Excel_DB"
    Function AboutExcel(ByRef oObj As Excel_DB) As String
        If Not oObj Is Nothing Then
            Dim strVersion() As String = Excel_DB.Version.Split(".")
            If strVersion.Length <> 3 Then Return "Error! Check that version is 'major.minor.build'"
            Dim iMajor, iMinor, iBuild As Integer
            iMajor = strVersion(0)
            iMinor = strVersion(1)
            iBuild = strVersion(2)
            If oObj.Open = True Then
                 Dim arrlstAbout As ArrayList = oObj.Read(String.Format("SELECT * FROM [tblchangelog$]"))
                oObj.Close()
                If arrlstAbout.Count <> 0 Then
                    If Not oObj.Log Is Nothing Then
                        CType(oObj.Log, File_Log).Save()
                    End If
                    'Return arrlstAbout(0)(0) & " | " & arrlstAbout(0)(1) 'dont know WHERE in excel so return last row
                    Dim iCount As Integer = arrlstAbout.Count - 1
                    Return arrlstAbout(iCount)(3) & " | " & arrlstAbout(iCount)(4)
                Else
                    Return "Error! Check that database is not empty"
                End If
            Else
                Return "Error! Check that database exists"
            End If
        End If
        Return "Error! Check that database is supplied"
    End Function

    Function ChangelogExcel(ByRef oObj As Excel_DB) As Dictionary(Of String, String)
        If Not oObj Is Nothing Then
            Dim strVersion() As String = Excel_DB.Version.Split(".")
            If strVersion.Length <> 3 Then Return Nothing
            Dim iMajor, iMinor, iBuild As Integer
            iMajor = strVersion(0)
            iMinor = strVersion(1)
            iBuild = strVersion(2)
            If oObj.Open = True Then
                Dim arrlstAbout As ArrayList = oObj.Read(String.Format("SELECT major, minor, build, changelog FROM [tblchangelog$] WHERE major={0} AND minor={1} AND build={2}", iMajor, iMinor, iBuild))
                oObj.Close()
                If arrlstAbout.Count <> 0 Then
                    Dim dictReturned As New Dictionary(Of String, String)
                    For Each SingleRow In arrlstAbout
                        'major.minor.build | changelog
                        dictReturned.Add(SingleRow(0) & "." & SingleRow(1) & "." & SingleRow(2), SingleRow(3))
                    Next
                    Return dictReturned
                End If
            End If
        End If
        Return Nothing
    End Function
#End Region

#Region "MsSql_DB"
    Function AboutMsSql(ByRef oObj As MsSql_DB) As String
        If Not oObj Is Nothing Then
            Dim strVersion() As String = MsSql_DB.Version.Split(".")
            If strVersion.Length <> 3 Then Return "Error! Check that version is 'major.minor.build'"
            Dim iMajor, iMinor, iBuild As Integer
            iMajor = strVersion(0)
            iMinor = strVersion(1)
            iBuild = strVersion(2)
            If oObj.Open = True Then
                Dim arrlstAbout As ArrayList = oObj.Read(String.Format("SELECT about, dependency FROM tblchangelog WHERE major={0} AND minor={1} AND build={2}", iMajor, iMinor, iBuild))
                oObj.Close()
                If arrlstAbout.Count <> 0 Then
                    If Not oObj.Log Is Nothing Then
                        CType(oObj.Log, File_Log).Save()
                    End If
                    Return arrlstAbout(0)(0) & " | " & arrlstAbout(0)(1)
                Else
                    Return "Error! Check that database is not empty"
                End If
            Else
                Return "Error! Check that database exists"
            End If
        End If
        Return "Error! Check that database is supplied"
    End Function

    Function ChangelogMsSql(ByRef oObj As MsSql_DB) As Dictionary(Of String, String)
        If Not oObj Is Nothing Then
            Dim strVersion() As String = MsSql_DB.Version.Split(".")
            If strVersion.Length <> 3 Then Return Nothing
            Dim iMajor, iMinor, iBuild As Integer
            iMajor = strVersion(0)
            iMinor = strVersion(1)
            iBuild = strVersion(2)
            If oObj.Open = True Then
                Dim arrlstAbout As ArrayList = oObj.Read(String.Format("SELECT major, minor, build, changelog FROM tblchangelog WHERE major={0} AND minor={1} AND build={2}", iMajor, iMinor, iBuild))
                oObj.Close()
                If arrlstAbout.Count <> 0 Then
                    Dim dictReturned As New Dictionary(Of String, String)
                    For Each SingleRow In arrlstAbout
                        'major.minor.build | changelog
                        dictReturned.Add(SingleRow(0) & "." & SingleRow(1) & "." & SingleRow(2), SingleRow(3))
                    Next
                    Return dictReturned
                End If
            End If
        End If
        Return Nothing
    End Function
#End Region

#Region "MySql_DB"
    Function AboutMySql(ByRef oObj As MySql_DB) As String
        If Not oObj Is Nothing Then
            Dim strVersion() As String = MySql_DB.Version.Split(".")
            If strVersion.Length <> 3 Then Return "Error! Check that version is 'major.minor.build'"
            Dim iMajor, iMinor, iBuild As Integer
            iMajor = strVersion(0)
            iMinor = strVersion(1)
            iBuild = strVersion(2)
            If oObj.Open = True Then
                Dim arrlstAbout As ArrayList = oObj.Read(String.Format("SELECT about, dependency FROM tblchangelog WHERE major={0} AND minor={1} AND build={2}", iMajor, iMinor, iBuild))
                oObj.Close()
                If arrlstAbout.Count <> 0 Then
                    If Not oObj.Log Is Nothing Then
                        CType(oObj.Log, File_Log).Save()
                    End If
                    Return arrlstAbout(0)(0) & " | " & arrlstAbout(0)(1)
                Else
                    Return "Error! Check that database is not empty"
                End If
            Else
                Return "Error! Check that database exists"
            End If
        End If
        Return "Error! Check that database is supplied"
    End Function

    Function ChangelogMySql(ByRef oObj As MySql_DB) As Dictionary(Of String, String)
        If Not oObj Is Nothing Then
            Dim strVersion() As String = MySql_DB.Version.Split(".")
            If strVersion.Length <> 3 Then Return Nothing
            Dim iMajor, iMinor, iBuild As Integer
            iMajor = strVersion(0)
            iMinor = strVersion(1)
            iBuild = strVersion(2)
            If oObj.Open = True Then
                Dim arrlstAbout As ArrayList = oObj.Read(String.Format("SELECT major, minor, build, changelog FROM tblchangelog WHERE major={0} AND minor={1} AND build={2}", iMajor, iMinor, iBuild))
                oObj.Close()
                If arrlstAbout.Count <> 0 Then
                    Dim dictReturned As New Dictionary(Of String, String)
                    For Each SingleRow In arrlstAbout
                        'major.minor.build | changelog
                        dictReturned.Add(SingleRow(0) & "." & SingleRow(1) & "." & SingleRow(2), SingleRow(3))
                    Next
                    Return dictReturned
                End If
            End If
        End If
        Return Nothing
    End Function
#End Region

#Region "Sqlite_DB"
    Function AboutSqlite(ByRef oObj As Sqlite_DB) As String
        If Not oObj Is Nothing Then
            Dim strVersion() As String = Sqlite_DB.Version.Split(".")
            If strVersion.Length <> 3 Then Return "Error! Check that version is 'major.minor.build'"
            Dim iMajor, iMinor, iBuild As Integer
            iMajor = strVersion(0)
            iMinor = strVersion(1)
            iBuild = strVersion(2)
            If oObj.Open = True Then
                Dim arrlstAbout As ArrayList = oObj.Read(String.Format("SELECT [about], [dependency] FROM tblchangelog WHERE [major]={0} AND [minor]={1} AND [build]={2}", iMajor, iMinor, iBuild))
                oObj.Close()
                If arrlstAbout.Count <> 0 Then
                    If Not oObj.Log Is Nothing Then
                        CType(oObj.Log, File_Log).Save()
                    End If
                    Return arrlstAbout(0)(0) & " | " & arrlstAbout(0)(1)
                Else
                    Return "Error! Check that database is not empty"
                End If
            Else
                Return "Error! Check that database exists"
            End If
        End If
        Return "Error! Check that database is supplied"
    End Function

    Function ChangelogSqlite(ByRef oObj As Sqlite_DB) As Dictionary(Of String, String)
        If Not oObj Is Nothing Then
            Dim strVersion() As String = Sqlite_DB.Version.Split(".")
            If strVersion.Length <> 3 Then Return Nothing
            Dim iMajor, iMinor, iBuild As Integer
            iMajor = strVersion(0)
            iMinor = strVersion(1)
            iBuild = strVersion(2)
            If oObj.Open = True Then
                Dim arrlstAbout As ArrayList = oObj.Read(String.Format("SELECT [major], [minor], [build], [changelog] FROM tblchangelog WHERE [major]={0} AND [minor]={1} AND [build]={2}", iMajor, iMinor, iBuild))
                oObj.Close()
                If arrlstAbout.Count <> 0 Then
                    Dim dictReturned As New Dictionary(Of String, String)
                    For Each SingleRow In arrlstAbout
                        'major.minor.build | changelog
                        dictReturned.Add(SingleRow(0) & "." & SingleRow(1) & "." & SingleRow(2), SingleRow(3))
                    Next
                    Return dictReturned
                End If
            End If
        End If
        Return Nothing
    End Function
#End Region
End Module