Option Explicit

Public Const MaxColNum As Long = 300
Public Const MaxKeyNum As Long = 5

''START(1,1)
Public Sub Main()
    
    Dim TotResultArray(10000) As String                 'CSV 전체 내용 저장하는 배열
    Dim i As Long, j As Long
    Dim JoinStr As String
    Dim TempStr As String
    Dim ResultData() As String
    Dim KeyCol As String                                '키로 사용될 컬럼의 번호를 콤마 구분으로 입력
    Dim InputFileName As String                         'CSV 파일 이름(경로 포함)
    Dim TargetKey As String
    Dim TargetField As String

'====================================================================================================================================================
'사용자가 입력해줘야 하는 파라미터 값
'====================================================================================================================================================
    KeyCol = "1,2,3,4,5"                                '첫번째부터 다섯번째 컬럼을 순서대로 조합해서 키로 사용함
    
    InputFileName = "load_comm.csv"                     'CSV 파일명
    TargetKey = "P029107001B_0_1_0_0"                   '검색 대상이 되는 Key 값
    TargetField = "Alp_Ini_GP"                          '검색 대상이 되는 Field 이름
'====================================================================================================================================================
    
    
    Call CSVImport(InputFileName, TotResultArray(), KeyCol)
'====================================================================================================================================================
    'CSVImport Sub
    'CSV data file 의 내용을 TotResultArray 배열에 저장. String 속성의 배열임
    '1st 파라미터(InputFileName) : 대상 CSV 파일명
    '2nd 파라미터(TotResultArray()) : CSV 내용을 반환받기 위한 배열
        'index 0 : 필드 이름
        'index 1 : 필드 속성
        'index 11~ : 데이터 레코드
    '3rd 파라미터(KeyCol) : CSV 내용 중 Key Field 의 위치를 담고 있는 변수. 콤마로 구분.
    '[주의] ByRef 로 처리되는 배열변수는 반드시 0 부터 인덱스 정의 필요. 그렇지 않으면 하나씩 뒤로 밀리게 됨
'====================================================================================================================================================
    
'    JoinStr = JoinArray(TotResultArray(), TSize(0))
'    Debug.Print JoinStr
    
    Debug.Print "# of Rows: ", TotResultArray(2)
    Debug.Print "# of Cols: ", TotResultArray(3)
    
    For j = 6 To 29
        Debug.Print j, "VLookup", VLookup(TargetKey, TotResultArray(), j)
    Next j
    
    Debug.Print "VLookupAll", VLookupAll(TargetKey, TotResultArray())
    Debug.Print "HLookupAll", HLookupAll(TargetField, TotResultArray())

'====================================================================================================================================================
    'VLookup Function : 반환값은 반드시 Double type
    '1st 파라미터(TargetKey) : 찾고자 하는 Key 값
    '2nd 파라미터(TotResultArray()) : Source Array 전체
    '3rd 파라미터(j) : 값을 가지고 오는 대상 컬럼 순번
'====================================================================================================================================================
    
'====================================================================================================================================================
    'VLookupAll Function - VLookup 과 유사하나 TargetKey 에 해당하는 행 값 전체를 반환받음
    '반환값은 반드시 String Type. comma 로 구분되어있기 때문에 반드시 Split 처리해줘야 함
    '1st 파라미터(TargetKey) : 찾고자 하는 Key 값
    '2nd 파라미터(TotResultArray()) : Source Array 전체
'====================================================================================================================================================
    
'====================================================================================================================================================
    'HLookupAll Function - VLookup 과 유사하나 TargetField 에 해당하는 열 값 전체를 반환받음
    '반환값은 반드시 String Type. comma 로 구분되어있기 때문에 반드시 Split 처리해줘야 함
    '1st 파라미터(TargetField) : 찾고자 하는 Key 값
    '2nd 파라미터(TotResultArray()) : Source Array 전체
'====================================================================================================================================================
    
    
'    TempStr = VLookupAll("1010101011M", TotResultArray())
'    ResultData = Split(TempStr, ",")

'    For j = LBound(ResultData) To UBound(ResultData)
'        Debug.Print LBound(ResultData), UBound(ResultData), j, ResultData(j)
'    Next j
'
'    Debug.Print TempDbl
    
    'Debug.Print No_Row, No_Col
    

End Sub


