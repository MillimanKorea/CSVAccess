    Public Sub CSVImport(ByVal CSVFileName As String, _
                         ByRef ResultArray() As String, _
                         ByRef Field As String, _
                         ByRef Key() As String, _
                         ByRef ColKey() As Long, _
                         ByRef AttrCol() As String, _
                         ByRef NumRow As Long, _
                         ByRef NumCol As Long)

        Dim S As String
        Dim fnr As Long
        Dim RecArray() As String
        Dim RecCount As Long
        Dim i As Long
        Dim j As Long
        Dim Temp As Double

        'file number setting
        fnr = FreeFile()

        'file open
        Open EB_Path(CSVFileName) For Input As fnr

        For i = 1 To MaxKeyNum
            Temp = Temp + ColKey(i)
        Next i
        If Temp = 0 Then ColKey(1) = 1

        '파일 끝까지 반복해서 읽어들이기
        Do While Not EOF(fnr)
            '한줄씩 읽어들여서 S 에 저장
            Line Input #fnr, S

            '읽어들이는 레코드 카운트
            RecCount = RecCount + 1

            '배열 S 에 저장된 내용을 comma 기준으로 분리
            RecArray = Split(S, ",")

            '컬럼 변수명 저장 후 Field 배열로 반환
            If RecCount = 1 Then
                Field = S
            End If

            '정의된 Field 갯수를 NumCol 에 저장 후 반환
            If RecCount = 2 Then
                NumCol = UBound(RecArray) + 1
                For i = 1 To NumCol
                    AttrCol(i) = Left(RecArray(i - 1), 1)
                Next i
            End If

            '데이터 파일 정보를 담고 있는 처음 3라인을 읽은 이후, 즉, 데이터 값 처리 부분
            If RecCount > 3 Then
                NumRow = NumRow + 1
                'Split 처리 안한 상태로 바로 반환
                ResultArray(NumRow) = S

                'Key 배열 조합해서 생성 - MaxKeyNum 만큼 반복
                For i = 1 To MaxKeyNum
                    For j = 1 To NumCol
                        If ColKey(i) = j And i = 1 Then
                            Key(NumRow) = RecArray(i - 1)
                        ElseIf ColKey(i) = j And ColKey(i) <> 0 Then
                            Key(NumRow) = Key(NumRow) & "_" & RecArray(i - 1)
                        End If
                    Next j
                Next i
            End If

        Loop

        'file close
        Close fnr

    End Sub
