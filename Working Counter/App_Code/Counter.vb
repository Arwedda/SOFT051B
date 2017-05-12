Public Class Counter
    Private mCount As Long

    Public Function GetCount() As Long
        GetCount = mCount
    End Function

    Public Sub Reset()
        mCount = 0
    End Sub

    Public Sub Up()
        mCount = mCount + 1
    End Sub

    Public Sub Down()
        mCount = mCount - 1
    End Sub

    Public Sub Rangecheck()
        If mCount > 10 Then
            mCount = 10
        ElseIf mCount < 0 Then
            Reset()
        End If
    End Sub
End Class