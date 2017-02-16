Option Explicit

Public Const MaxColNum As Long = 300
Public Const MaxKeyNum As Long = 5

''START(1,1)
Public Sub Main()

    Dim TotResultArray(10000) As String                 'CSV 전체 내용 저장하는 배열
    Dim FieldRec As String                              'CSV 중 Field 이름 저장하는 변수(CSV 첫번째 줄)
    Dim KeyArray(10000) As String                       'CSV 중 Key 정보 저장하는 배열
    Dim ColAttr(MaxColNum) As String                    '컬럼별 속성 정보(String or Double) 저장하는 배열
    Dim i As Long, j As Long
    Dim JoinStr As String
    Dim TempStr As String
    Dim No_Row As Long, No_Col As Long                  'CSV 의 행/열 갯수
    Dim ResultData() As String
    Dim KeyCol(MaxKeyNum) As Long                       'Key 로 지정할 수 있는 컬럼은 최대 5개
    Dim InputFileName As String
    Dim TargetKey As String
    Dim TargetField As String


'====================================================================================================================================================
'사용자가 입력해줘야 하는 파라미터 값
'====================================================================================================================================================
    KeyCol(1) = 1                                       '첫번째 키의 컬럼 번호
    KeyCol(2) = 2                                       '두번째 키의 컬럼 번호
    KeyCol(3) = 3                                       '세번째 키의 컬럼 번호
    KeyCol(4) = 4                                       '네번째 키의 컬럼 번호
    KeyCol(5) = 5                                       '다섯번째 키의 컬럼 번호

    InputFileName = "load_comm.csv"                     'CSV 파일명
    TargetKey = "P029107001B_0_1_0_0"                   '검색 대상이 되는 Key 값
    TargetField = "Alp_Ini_GP"                          '검색 대상이 되는 Field 이름
'====================================================================================================================================================


    Call CSVImport(InputFileName, TotResultArray(), FieldRec, KeyArray(), KeyCol(), ColAttr(), No_Row, No_Col)
'====================================================================================================================================================
    'CSVImport Sub
    'CSV data file 의 내용을 TotResultArray 배열에 저장. String 속성의 배열임
    '1st 파라미터(InputFileName) : 대상 CSV 파일명
    '2nd 파라미터(TotResultArray()) : CSV 내용을 반환받기 위한 배열
    '3rd 파라미터(FieldRec) : CSV 내용 중 각 Field 명을 반환받기 위한 변수. CSV 의 첫번째 줄에 해당함
    '4th 파라미터(KeyArray()) : CSV 내용 중 Key Field 내용을 반환받기 위한 변수. 현재는 무조건 첫번째 컬럼
    '5th 파라미터(KeyCol()) : CSV 내용 중 Key Field 의 위치를 담고 있는 변수. 특별히 지정하지 않으면 첫번째 컬럼만 Key Field 로 인식. 0 이면 적용하지 않음
    '6th 파라미터(ColAttr()) : 컬럼별 속성("S"tring or "D"ouble) 반환
    '7th 파라미터(No_Row) : CSV 데이터의 행 갯수 반환
    '7th 파라미터(No_Col) : CSV 데이터의 열 갯수 반환
    '[주의] ByRef 로 처리되는 배열변수는 반드시 0 부터 인덱스 정의 필요. 그렇지 않으면 하나씩 뒤로 밀리게 됨
'====================================================================================================================================================

'    JoinStr = JoinArray(TotResultArray(), No_Row)
'    Debug.Print JoinStr

    For j = 6 To 29
        'Debug.Print VLookup(TargetKey, KeyArray(), TotResultArray(), j, No_Row)
    Next j

    'Debug.Print "VLookupAll", VLookupAll(TargetKey, KeyArray(), TotResultArray(), No_Row)
    Debug.Print "HLookupAll", HLookupAll(TargetField, FieldRec, TotResultArray(), No_Row)

'====================================================================================================================================================
    'VLookup Function : 반환값은 반드시 Double type
    '1st 파라미터(TargetKey) : 찾고자 하는 Key 값
    '2nd 파라미터(KeyArray()) : Source Array 중 Key 배열
    '3rd 파라미터(TotResultArray()) : Source Array 전체
    '4th 파라미터(j) : 값을 가지고 오는 대상 컬럼 순번
    '5th 파라미터(No_Row) : 배열 행 갯수
'====================================================================================================================================================

'====================================================================================================================================================
    'VLookupAll Function - VLookup 과 유사하나 TargetKey 에 해당하는 행 값 전체를 반환받음
    '반환값은 반드시 String Type. comma 로 구분되어있기 때문에 반드시 Split 처리해줘야 함
    '1st 파라미터(TargetKey) : 찾고자 하는 Key 값
    '2nd 파라미터(KeyArray()) : Source Array 중 Key 배열
    '3rd 파라미터(TotResultArray()) : Source Array 전체
    '4th 파라미터(No_Row) : 배열 행 갯수
'====================================================================================================================================================

'====================================================================================================================================================
    'HLookupAll Function - VLookup 과 유사하나 TargetField 에 해당하는 열 값 전체를 반환받음
    '반환값은 반드시 String Type. comma 로 구분되어있기 때문에 반드시 Split 처리해줘야 함
    '1st 파라미터(TargetField) : 찾고자 하는 Key 값
    '2nd 파라미터(FieldRec) : CSV 파일 내용 중 필드의 변수명이 들어있는 첫번째 행의 내용
    '3rd 파라미터(TotResultArray()) : Source Array 전체
    '4th 파라미터(No_Row) : 배열 행 갯수
'====================================================================================================================================================


'    TempStr = VLookupAll("1010101011M", KeyArray(), TotResultArray(), No_Row)
'    ResultData = Split(TempStr, ",")

'    For j = LBound(ResultData) To UBound(ResultData)
'        Debug.Print LBound(ResultData), UBound(ResultData), j, ResultData(j)
'    Next j
'
'    Debug.Print TempDbl

    'Debug.Print No_Row, No_Col


End Sub
