Attribute VB_Name = "CSV"
Option Explicit

Public Const MaxColNum As Long = 300
Public Const MaxKeyNum As Long = 5

''START(1,1)
Public Sub Main()
    
    Dim ResultArray(10000) As String                 'CSV 전체 내용 저장하는 배열
    Dim i As Long, j As Long
    Dim JoinStr As String
    Dim TempStr As String
    Dim ResultData() As String
    Dim KeyCol As String                                '키로 사용될 컬럼의 번호를 콤마 구분으로 입력
    Dim InputFileName As String                         'CSV 파일 이름(경로 포함)
    Dim TargetKey As String
    'Dim TargetField As String
    Dim TargetField(1 To 29) As String
    

'====================================================================================================================================================
'사용자가 입력해줘야 하는 파라미터 값
'====================================================================================================================================================
    KeyCol = "1,2,3,4,5"                                '첫번째부터 다섯번째 컬럼을 순서대로 조합해서 키로 사용함
    
    InputFileName = "load_comm.csv"                     'CSV 파일명
    TargetKey = "P029107001B_0_1_0_0"                   '검색 대상이 되는 Key 값
    'TargetField = "Alp_Ini_GP"                          '검색 대상이 되는 Field 이름
    
    TargetField(1) = "Plan_Code"
    TargetField(2) = "BT"
    TargetField(3) = "PT"
    TargetField(4) = "GP"
    TargetField(5) = "Pay_Mode"
    TargetField(6) = "Alp_Fix"
    TargetField(7) = "Alp_FA"
    TargetField(8) = "Alp_Ini_GP"
    TargetField(9) = "Alp_GP"
    TargetField(10) = "Alp_STD_Annual_Yr"
    TargetField(11) = "Alp_Amort_Yr"
    TargetField(12) = "Alp_Round_Digit"
    TargetField(13) = "Beta_Fix"
    TargetField(14) = "Beta_Fix_AT"
    TargetField(15) = "Beta_FA"
    TargetField(16) = "Beta_GP"
    TargetField(17) = "Beta_GP_AT"
    TargetField(18) = "Betadash_Fix"
    TargetField(19) = "Betadash_FA"
    TargetField(20) = "Betadash_GP"
    TargetField(21) = "Beta_GP_Ann"
    TargetField(22) = "Betadash_Res"
    TargetField(23) = "Betadash_FA_Ann"
    TargetField(24) = "Betadash_Ann"
    TargetField(25) = "Betadash_PU_Mode"
    TargetField(26) = "Beta_Min_Fixed"
    TargetField(27) = "Beta_Min_Prem"
    TargetField(28) = "Gamma"
    TargetField(29) = "Alp_GP2"

'====================================================================================================================================================
    
    
    Call CSVImport(InputFileName, ResultArray(), KeyCol, 0)
'====================================================================================================================================================
    'CSVImport Sub
    'CSV data file 의 내용을 ResultArray 배열에 저장. String 속성의 배열임
    '1st 파라미터(InputFileName) : 대상 CSV 파일명
    '2nd 파라미터(ResultArray()) : CSV 내용을 반환받기 위한 배열
        'index 0 : 필드 이름
        'index 1 : 필드 속성
        'index 11~ : 데이터 레코드
    '3rd 파라미터(KeyCol) : CSV 내용 중 Key Field 의 위치를 담고 있는 변수. 콤마로 구분.
    '4th 파라미터(SortFlag) : 배열의 Sorting 여부. 0 이면 정렬하지 않음. 1이면 정렬(QuickSort)
    '[주의] ByRef 로 처리되는 배열변수는 반드시 0 부터 인덱스 정의 필요. 그렇지 않으면 하나씩 뒤로 밀리게 됨
'====================================================================================================================================================
    
'    Call QuickSort(ResultArray(), 11, 11 + CLng(ResultArray(2)))

    
'    JoinStr = JoinArray(ResultArray(), TSize(0))
'    Debug.Print JoinStr
    
    Debug.Print "# of Rows: ", ResultArray(2)
    Debug.Print "# of Cols: ", ResultArray(3)
    
    Debug.Print "VLookupAll", VLookupAll(TargetKey, ResultArray())
    Debug.Print "HLookupAll", HLookupAll(TargetField(8), ResultArray())

    For j = 1 To 29
        Debug.Print j, "VLookup", VLookup(TargetKey, ResultArray(), j)
        Debug.Print j, "Lookup", Lookup(TargetKey, ResultArray(), TargetField(j))
    Next j

'====================================================================================================================================================
    'VLookup Function : 반환값은 반드시 Double type
    '1st 파라미터(TargetKey) : 찾고자 하는 Key 값
    '2nd 파라미터(ResultArray()) : Source Array 전체
    '3rd 파라미터(j) : 값을 가지고 오는 대상 컬럼 순번
'====================================================================================================================================================
    
'====================================================================================================================================================
    'Lookup Function : 반환값은 반드시 Double type
    '1st 파라미터(TargetKey) : 찾고자 하는 Key 값
    '2nd 파라미터(ResultArray()) : Source Array 전체
    '3rd 파라미터(FieldName) : 값을 가지고 오는 대상 컬럼 이름
'====================================================================================================================================================

'====================================================================================================================================================
    'VLookupAll Function - VLookup 과 유사하나 TargetKey 에 해당하는 행 값 전체를 반환받음
    '반환값은 반드시 String Type. comma 로 구분되어있기 때문에 반드시 Split 처리해줘야 함
    '1st 파라미터(TargetKey) : 찾고자 하는 Key 값
    '2nd 파라미터(ResultArray()) : Source Array 전체
'====================================================================================================================================================
    
'====================================================================================================================================================
    'HLookupAll Function - VLookup 과 유사하나 TargetField 에 해당하는 열 값 전체를 반환받음
    '반환값은 반드시 String Type. comma 로 구분되어있기 때문에 반드시 Split 처리해줘야 함
    '1st 파라미터(TargetField) : 찾고자 하는 Key 값
    '2nd 파라미터(ResultArray()) : Source Array 전체
'====================================================================================================================================================
    
    
'    TempStr = VLookupAll("1010101011M", ResultArray())
'    ResultData = Split(TempStr, ",")

'    For j = LBound(ResultData) To UBound(ResultData)
'        Debug.Print LBound(ResultData), UBound(ResultData), j, ResultData(j)
'    Next j
'
'    Debug.Print TempDbl
    
    'Debug.Print No_Row, No_Col
    

End Sub


