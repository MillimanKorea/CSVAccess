    Public Function HLookupAll(ByVal Target As String, Key As String, SourceArray() As String, ByVal NumRow As Long) As String

        Dim i As Long
        Dim RecArray() As String
        Dim TargetColNum As Long
        Dim TempStr As String

        '첫번째 라인 Split
        RecArray = Split(Key, ",")

        'Target 에 해당하는 컬럼의 인덱스 찾기
        For i = 0 To UBound(RecArray())
            If RecArray(i) = Target Then
                TargetColNum = i + 1
                Exit For
            End If
        Next i

        For i = 1 To NumRow
            RecArray = Split(SourceArray(i), ",")
            TempStr = TempStr & RecArray(TargetColNum - 1) & ","
        Next i

        HLookupAll = TempStr

    End Function
