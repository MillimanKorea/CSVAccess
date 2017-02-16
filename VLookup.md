    Public Function VLookup(ByVal Target As String, Key() As String, SourceArray() As String, ByVal FieldNum As Long, ByVal NumRow As Long) As String

        Dim i As Long
        Dim RecArray() As String

        For i = 1 To NumRow
            If Key(i) = Target Then
                RecArray = Split(SourceArray(i), ",")
                VLookup = RecArray(FieldNum - 1)
                Exit Function
            End If
        Next i

    End Function
