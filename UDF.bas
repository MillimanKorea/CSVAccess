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
    Dim Adj As Long: Adj = 0
    Dim NumRow As Long
    Dim FlagSort As String

    NumRow = CLng(BST_NORow(SourceArray))

    '첫번째 라인 Split
    RecArray = Split(SourceArray(0), "|")
    '"Sorted" or "NotSorted"
    FlagSort = RecArray(4)
    RecArray = Split(RecArray(0), ",")

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

    If FlagSort = "NotSorted" Then
        '순차검색
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
    Else
        '이진검색
        i = BinSearch(SourceArray, TargetRow, 1, NumRow)
        RecArray = Split(SourceArray(i), "|")
        RecArray = Split(RecArray(1), ",")
        Lookup = RecArray(TargetColNum - 1)
    End If

End Function


Public Function LookupNum(ByVal TargetRowNum As Long, SourceArray() As String, ByVal TargetCol As String) As String

    Dim i As Long
    Dim RecArray() As String, DataArray() As String
    Dim TargetColNum As Long
    Dim TempStr As String: TempStr = ""
    Dim Adj As Long: Adj = 0
    Dim NumRow As Long

    NumRow = CLng(BST_NORow(SourceArray))

    '첫번째 라인 Split
    RecArray = Split(SourceArray(0), "|")
    RecArray = Split(RecArray(0), ",")

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

    RecArray = Split(SourceArray(TargetRowNum + Adj), "|")
    RecArray = Split(RecArray(1), ",")
    LookupNum = RecArray(TargetColNum - 1)

    'TargetRowNum 에 해당되는 필드가 존재하지 않는 경우 error message 표시
    If i > NumRow + Adj Then
        Debug.Print TargetRowNum & " 번째 레코드가 존재하지 않습니다."
    End If

End Function


Public Function VLookup(ByVal TargetRow As String, SourceArray() As String, ByVal FieldNum As Long) As String
    
    Dim i As Long
    Dim RecArray() As String
    Dim Adj As Long: Adj = 0
    Dim NumRow As Long, NumCol As Long
    Dim FlagSort As String
    
    NumRow = CLng(BST_NORow(SourceArray))
    NumCol = CLng(BST_NOCol(SourceArray))
    If FieldNum > NumCol Then
        Debug.Print "배열의 컬럼 갯수보다 큰 필드번호가 입력되었습니다."
        Exit Function
    End If
    
    '첫번째 라인 Split
    RecArray = Split(SourceArray(0), "|")
    '"Sorted" or "NotSorted"
    FlagSort = RecArray(4)
    
    If FlagSort = "NotSorted" Then
        '순차검색
        For i = 1 + Adj To NumRow + Adj
            If SourceArray(i) <> "" Then
                RecArray = Split(SourceArray(i), "|")
                If RecArray(0) = TargetRow Then
                    RecArray = Split(RecArray(1), ",")
                    VLookup = RecArray(FieldNum - 1)
                    Exit Function
                End If
            End If
        Next i
        
        If i > NumRow + Adj Then
            Debug.Print TargetRow & " 에 해당하는 레코드가 존재하지 않습니다."
        End If
    Else
        '이진검색
        i = BinSearch(SourceArray, TargetRow, 1, NumRow)
        RecArray = Split(SourceArray(i), "|")
        RecArray = Split(RecArray(1), ",")
        VLookup = RecArray(FieldNum - 1)
    End If
    
End Function



Public Function VLookupAll(ByVal Target As String, SourceArray() As String) As String
    
    Dim i As Long
    Dim Adj As Long: Adj = 0
    Dim RecArray() As String
    Dim NumRow As Long
    Dim FlagSort As String
    
    NumRow = CLng(BST_NORow(SourceArray))
    
    '첫번째 라인 Split
    RecArray = Split(SourceArray(0), "|")
    '"Sorted" or "NotSorted"
    FlagSort = RecArray(4)
    
    If FlagSort = "NotSorted" Then
        '순차검색
        For i = 1 + Adj To NumRow + Adj
            If SourceArray(i) <> "" Then
                RecArray = Split(SourceArray(i), "|")
                If RecArray(0) = Target Then
                    VLookupAll = RecArray(1)
                    Exit Function
                End If
            End If
        Next i
        'Target 에 해당되는 필드가 존재하지 않는 경우 error message 표시
        If i > NumRow + Adj Then
            Debug.Print Target & " 에 해당하는 레코드가 존재하지 않습니다."
        End If
    Else
        '이진검색
        i = BinSearch(SourceArray, Target, 1, NumRow)
        RecArray = Split(SourceArray(i), "|")
        VLookupAll = RecArray(1)
    End If
    
End Function



Public Function HLookupAll(ByVal Target As String, SourceArray() As String) As String

    Dim i As Long
    Dim RecArray() As String, DataArray() As String
    Dim TargetColNum As Long
    Dim TempStr As String: TempStr = ""
    Dim Adj As Long: Adj = 0
    Dim NumRow As Long
    
    NumRow = CLng(BST_NORow(SourceArray))
    
    '첫번째 라인 Split
    RecArray = Split(SourceArray(0), "|")
    RecArray = Split(RecArray(0), ",")
    
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



Public Sub CSVImport(ByVal CSVFileName As String, ByRef CSVArray() As String, ByRef KeyColStr As String, ByVal SortFlag As Integer)

    Dim S As String
    Dim fnr As Long
    Dim RecArray() As String
    Dim RecCount As Long
    Dim i As Long
    Dim j As Long
    Dim Temp As Double
    Dim NumRow As Long, NumCol As Long
    Dim ColKey() As String
    Dim CSVHeader(4) As String
    
    
    'file number setting
    fnr = FreeFile()
    
    'file open
    Open EB_Path(CSVFileName) For Input As fnr
    
    If KeyColStr = "" Then KeyColStr = "1"
    ColKey = Split(KeyColStr, ",")
    
    'Index 0 에 배열에 대한 정보를 넣고("|" separator 로 구분) Index 부터 Key 를 포함한 데이터를 순서대로 저장함
    NumRow = 0
                
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
            CSVHeader(0) = S
        End If
        
        '정의된 Field 갯수를 NumCol 에 저장 후 반환
        If RecCount = 2 Then
            NumCol = UBound(RecArray) + 1
            For i = 1 To NumCol
                If i = 1 Then
                    CSVHeader(1) = Left(RecArray(i - 1), 1)
                Else
                    CSVHeader(1) = CSVHeader(1) & "," & Left(RecArray(i - 1), 1)
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
                        CSVArray(NumRow) = RecArray(i)
                    ElseIf CLng(ColKey(i)) = j And CLng(ColKey(i)) <> 0 Then
                        CSVArray(NumRow) = CSVArray(NumRow) & "_" & RecArray(i)
                    End If
                Next j
            Next i
        
            '데이터 저장
            CSVArray(NumRow) = CSVArray(NumRow) & "|" & S
            
        End If

    Loop
   
    'file close
    Close fnr
    
    CSVHeader(2) = NumRow
    CSVHeader(3) = NumCol
    
    Select Case SortFlag
        Case 1
            CSVHeader(4) = "Sorted"
            Call QuickSort(CSVArray, 1, NumRow)
        Case 2
            CSVHeader(4) = "Sorted"
            Call HeapSort(CSVArray, NumRow)
        Case 3
            CSVHeader(4) = "Sorted"
            Call InsertionSort(CSVArray, NumRow)
        Case Else
            CSVHeader(4) = "NotSorted"
    End Select
    
    CSVArray(0) = CSVHeader(0) & "|" & _
                  CSVHeader(1) & "|" & _
                  CSVHeader(2) & "|" & _
                  CSVHeader(3) & "|" & _
                  CSVHeader(4)

End Sub


Public Sub QuickSort(ByRef SrcArray() As String, ByVal min As Long, ByVal max As Long)
    
    Dim med_value As String
    Dim hi As Long
    Dim lo As Long
    Dim i As Long
    Dim j As Long, k As Long
        
    If max <= min Then Exit Sub
    
    i = Int((max - min + 1) * 0.5 + min)
    med_value = SrcArray(i)
    SrcArray(i) = SrcArray(min)
    
    lo = min
    hi = max
    
    For j = 1 To max
        
        For k = hi To lo Step -1
            If SrcArray(k) < med_value Or k <= lo Then
                hi = k
                Exit For
            End If
        Next k
'        Do While List(hi) >= med_value
'            hi = hi - 1
'            If hi <= lo Then Exit For
'        Loop
     
        If hi <= lo Then
            SrcArray(lo) = med_value
            Exit For
        End If
    
        SrcArray(lo) = SrcArray(hi)
        lo = lo + 1
     
        For k = lo To hi
            If SrcArray(lo) >= med_value Or lo >= hi Then
                lo = k
                Exit For
            End If
        Next k
'        Do While List(lo) < med_value
'            lo = lo + 1
'            If lo >= hi Then Exit For
'        Loop
     
        If lo >= hi Then
            lo = hi
            SrcArray(hi) = med_value
            Exit For
        End If
     
        ' Swap the lo and hi values.
        SrcArray(hi) = SrcArray(lo)
    Next j
    
    Call QuickSort(SrcArray(), min, lo - 1)
    Call QuickSort(SrcArray(), lo + 1, max)

End Sub


Function BinSearch(ByRef SrcArray() As String, ByVal Target As String, ByVal min As Long, ByVal max As Long) As Long
    
    Dim low As Long: low = min
    Dim high As Long: high = max
    Dim i As Long
    Dim j As Long
    Dim SrcRec() As String
    Dim SrcData As String
    
    For j = min To max
        i = (low + high) / 2
        SrcRec = Split(SrcArray(i), "|")
        SrcData = SrcRec(0)
        If Target = SrcData Then
            BinSearch = i
            Exit For
        ElseIf Target < SrcData Then
            high = (i - 1)
        Else
            low = (i + 1)
        End If
        If low > high Then
            Exit For
        End If
    Next j
'    Do While low <= high
'        i = (low + high) / 2
'        If Target = List(i) Then
'            BinSearch = i
'            Exit For
'        ElseIf Target < List(i) Then
'            high = (i - 1)
'        Else
'            low = (i + 1)
'        End If
'    Loop
    
    If BinSearch = 0 Then
        BinSearch = 0
    End If

End Function



Public Function BST_NORow(ByRef SourceArray() As String) As Long

    Dim TempArray() As String
    
    TempArray = Split(SourceArray(0), "|")
    BST_NORow = CLng(TempArray(2))
    
End Function


Public Function BST_NOCol(ByRef SourceArray() As String) As Long

    Dim TempArray() As String
    
    TempArray = Split(SourceArray(0), "|")
    BST_NOCol = CLng(TempArray(3))
    
End Function


Public Function Dec2Str(ByRef SourceDec As Long, ByVal Dec As Long) As String

    Dim i As Integer
    Dim Div As Double
    Dim Result As String
    
    Result = ""
    
    For i = 0 To Dec - 1
        Div = SourceDec / (10 ^ i)
        If i = 0 And SourceDec = 0 Then
            Result = Result
        ElseIf Div < 1 Then
            Result = Result & "0"
        End If
    Next i
    
    Result = Result & SourceDec
    Dec2Str = Result
    
End Function


Sub InsertionSort(ByRef SrcArray() As String, ByVal max As Long)

    Dim lngCounter1 As Long
    Dim lngCounter2 As Long
    Dim varTemp As String

    For lngCounter1 = 1 To max
        
        varTemp = SrcArray(lngCounter1)
        
        For lngCounter2 = lngCounter1 To 1 Step -1

            If SrcArray(lngCounter2 - 1) > varTemp Then
                SrcArray(lngCounter2) = SrcArray(lngCounter2 - 1)
            Else
                Exit For
            End If

        Next lngCounter2

        SrcArray(lngCounter2) = varTemp

    Next lngCounter1

End Sub


Public Sub HeapSort(ByRef SrcArray() As String, ByVal max As Long)
 
    Dim j As String, the_end As Long
    Dim Ascending As Boolean: Ascending = True
 
    Call heapify(SrcArray, max, Ascending)
 
    the_end = max
    Do While the_end >= 1
        j = SrcArray(the_end)
        SrcArray(the_end) = SrcArray(1)
        SrcArray(1) = j

        the_end = the_end - 1

        Call siftDown(SrcArray, 1, the_end, Ascending)
    Loop
 
End Sub


Public Sub heapify(ByRef list() As String, ByVal count As Long, ByVal Ascending As Boolean)
 
    Dim start As Long

    'start = (count - 2) / 2
    start = count / 2
      
    Do While start > 0
        'Call siftDown(list, start, count - 1, Ascending)
        Call siftDown(list, start, count, Ascending)
        start = start - 1
    Loop

End Sub


Public Sub siftDown(ByRef list() As String, ByVal the_start As Long, ByVal the_end As Long, ByVal Ascending As Boolean)
 
    Dim k As String, root As Long, child As Long, swap As Long
    root = the_start
 
    Do While root * 2 <= the_end
        child = root * 2
        swap = root
  
        If Ascending Then
            If list(swap) < list(child) Then
                swap = child
            End If
  
            If child + 1 <= the_end And list(swap) < list(child + 1) Then
                swap = child + 1
            End If
        Else
  
            If list(child) < list(swap) Then
                swap = child
            End If
   
            If child + 1 <= the_end And list(child + 1) < list(swap) Then
                swap = child + 1
            End If
        End If
          
        If swap <> root Then
            k = list(root)
            list(root) = list(swap)
            list(swap) = k
            root = swap
        Else
            Exit Sub
        End If
 
    Loop
 
End Sub

