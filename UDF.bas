Attribute VB_Name = "UDF"
Option Explicit


''IGN{
Function EB_Path(S As String) As String
    EB_Path = ActiveWorkbook.Path & "\" & S
End Function
''IGN}


Public Function JoinArray(ByRef SourceArray() As String, ByVal NumRow As Long) As String

    Dim i As Long
    Dim TempStr As String

    For i = 1 To NumRow
        If TempStr <> "" Then
            TempStr = TempStr & "," & SourceArray(i)
        Else
            TempStr = SourceArray(i)
        End If
    Next i
    
    Debug.Print TempStr
    JoinArray = TempStr

End Function

Public Function Lookup(ByVal TargetRow As String, SourceArray() As String, ByVal TargetCol As String) As String

    Dim i As Long
    Dim RecArray() As String, DataArray() As String
    Dim TargetColNum As Long
    Dim TempStr As String: TempStr = ""
    Dim Adj As Long: Adj = 10
    Dim NumRow As Long
    
    NumRow = CLng(SourceArray(2))
    
    '첫번째 라인 Split
    RecArray = Split(SourceArray(0), ",")
    
    'TargetCol 에 해당하는 컬럼의 인덱스 찾기
    For i = 0 To UBound(RecArray())
        If RecArray(i) = TargetCol Then
            TargetColNum = i + 1
            Exit For
        End If
    Next i
    'TargetCol 에 해당되는 필드가 존재하지 않는 경우 error message 표시 후 종료
    If i > UBound(RecArray()) Then
        Debug.Print TargetCol & " 에 해당하는 필드가 존재하지 않습니다."
        Exit Function
    End If
    
    For i = 1 + Adj To NumRow + Adj
        If SourceArray(i) <> "" Then
            RecArray = Split(SourceArray(i), "|")
            If RecArray(0) = TargetRow Then
                RecArray = Split(RecArray(1), ",")
                Lookup = RecArray(TargetColNum - 1)
                Exit For
            End If
        End If
    Next i
    'TargetRow 에 해당되는 필드가 존재하지 않는 경우 error message 표시
    If i > NumRow + Adj Then
        Debug.Print TargetRow & " 에 해당하는 레코드가 존재하지 않습니다."
    End If
    
End Function


Public Function VLookup(ByVal Target As String, SourceArray() As String, ByVal FieldNum As Long) As String
    
    Dim i As Long
    Dim RecArray() As String
    Dim Adj As Long: Adj = 10
    Dim Key() As String
    Dim NumRow As Long, NumCol As Long
    
    NumRow = CLng(SourceArray(2))
    NumCol = CLng(SourceArray(3))
    If FieldNum > NumCol Then
        Debug.Print "배열의 컬럼 갯수보다 큰 필드번호가 입력되었습니다."
        Exit Function
    End If
    
    For i = 1 + Adj To NumRow + Adj
        If SourceArray(i) <> "" Then
            Key = Split(SourceArray(i), "|")
            If Key(0) = Target Then
                RecArray = Split(SourceArray(i), ",")
                VLookup = RecArray(FieldNum - 1)
                Exit Function
            End If
        End If
    Next i
    
    If i > NumRow + Adj Then
        Debug.Print Target & " 에 해당하는 레코드가 존재하지 않습니다."
    End If
    
End Function



Public Function VLookupAll(ByVal Target As String, SourceArray() As String) As String
    
    Dim i As Long
    Dim Adj As Long: Adj = 10
    Dim Key() As String
    Dim NumRow As Long
    
    NumRow = CLng(SourceArray(2))
    
    For i = 1 + Adj To NumRow + Adj
        If SourceArray(i) <> "" Then
            Key = Split(SourceArray(i), "|")
            If Key(0) = Target Then
                VLookupAll = Key(1)
                Exit Function
            End If
        End If
    Next i
    'Target 에 해당되는 필드가 존재하지 않는 경우 error message 표시
    If i > NumRow + Adj Then
        Debug.Print Target & " 에 해당하는 레코드가 존재하지 않습니다."
    End If
    
End Function



Public Function HLookupAll(ByVal Target As String, SourceArray() As String) As String

    Dim i As Long
    Dim RecArray() As String, DataArray() As String
    Dim TargetColNum As Long
    Dim TempStr As String: TempStr = ""
    Dim Adj As Long: Adj = 10
    Dim NumRow As Long
    
    NumRow = CLng(SourceArray(2))
    
    '첫번째 라인 Split
    RecArray = Split(SourceArray(0), ",")
    
    'Target 에 해당하는 컬럼의 인덱스 찾기
    For i = 0 To UBound(RecArray())
        If RecArray(i) = Target Then
            TargetColNum = i + 1
            Exit For
        End If
    Next i
    'Target 에 해당되는 필드가 존재하지 않는 경우 error message 표시 후 종료
    If i > UBound(RecArray()) Then
        Debug.Print Target & " 에 해당하는 필드가 존재하지 않습니다."
        Exit Function
    End If
    
    For i = 1 + Adj To NumRow + Adj
        If SourceArray(i) <> "" Then
            RecArray = Split(SourceArray(i), "|")
            RecArray = Split(RecArray(1), ",")
            TempStr = TempStr & RecArray(TargetColNum - 1) & ","
        End If
    Next i
    
    HLookupAll = Left(TempStr, Len(TempStr) - 1)

End Function



Public Sub CSVImport(ByVal CSVFileName As String, _
                     ByRef ResultArray() As String, _
                     ByRef KeyColStr As String)

    Dim S As String
    Dim fnr As Long
    Dim RecArray() As String
    Dim RecCount As Long
    Dim i As Long
    Dim j As Long
    Dim Temp As Double
    Dim NumRow As Long, NumCol As Long
    Dim ColKey() As String
    
    
    'file number setting
    fnr = FreeFile()
    
    'file open
    Open EB_Path(CSVFileName) For Input As fnr
    
    If KeyColStr = "" Then KeyColStr = "1"
    ColKey = Split(KeyColStr, ",")
    
    '데이터는 index 11 부터 넣기 시작함
    '앞쪽 0~10 까지의 11개 공간은 배열에 대한 정보를 넣는 공간으로 사용됨
    NumRow = 10
                
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
            ResultArray(0) = S
        End If
        
        '정의된 Field 갯수를 NumCol 에 저장 후 반환
        If RecCount = 2 Then
            NumCol = UBound(RecArray) + 1
            For i = 1 To NumCol
                If i = 1 Then
                    ResultArray(1) = Left(RecArray(i - 1), 1)
                Else
                    ResultArray(1) = ResultArray(1) & "," & Left(RecArray(i - 1), 1)
                End If
            Next i
        End If

        '데이터 파일 정보를 담고 있는 처음 3라인을 읽은 이후, 즉, 데이터 값 처리 부분
        If RecCount > 3 Then
            NumRow = NumRow + 1
            
            'Key 배열 조합해서 생성 - MaxKeyNum 만큼 반복
            For i = 0 To UBound(ColKey)
                For j = 1 To NumCol
                    If CLng(ColKey(i)) = j And i = 0 Then
                        ResultArray(NumRow) = RecArray(i)
                    ElseIf CLng(ColKey(i)) = j And CLng(ColKey(i)) <> 0 Then
                        ResultArray(NumRow) = ResultArray(NumRow) & "_" & RecArray(i)
                    End If
                Next j
            Next i
        
            '데이터 저장
            ResultArray(NumRow) = ResultArray(NumRow) & "|" & S
            
        End If

    Loop
   
    'file close
    Close fnr
    
    ResultArray(2) = NumRow
    ResultArray(3) = NumCol
    
End Sub

