
''' <summary>
''' Database log
''' </summary>
''' <remarks></remarks>
Public MustInherit Class Log
    Inherits List(Of String)
    Private Const Seperator As Char = "|"

    Public Function Update(ByVal xTimestamp As Date, ByVal sAction As String) As Integer
        Add(xTimestamp.ToString & Seperator & sAction)
        Return Me.Count
    End Function

End Class





