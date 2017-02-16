    Public Function VLookupAll(ByVal Target As String, Key() As String, SourceArray() As String, ByVal NumRow As Long) As String

        Dim i As Long
        Dim RecArray() As String

        For i = 1 To NumRow
            If Key(i) = Target Then
                VLookupAll = SourceArray(i)
                Exit Function
            End If
        Next i

    End Function
