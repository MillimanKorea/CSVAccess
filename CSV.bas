Attribute VB_Name = "CSV"
Option Explicit

Public Const MaxColNum As Long = 300
Public Const MaxKeyNum As Long = 5

''START(1,1)
Public Sub Main()
    
    Dim ResultArray(10000) As String                 'CSV ��ü ���� �����ϴ� �迭
    Dim i As Long, j As Long
    Dim JoinStr As String
    Dim TempStr As String
    Dim ResultData() As String
    Dim KeyCol As String                                'Ű�� ���� �÷��� ��ȣ�� �޸� �������� �Է�
    Dim InputFileName As String                         'CSV ���� �̸�(��� ����)
    Dim TargetKey As String
    'Dim TargetField As String
    Dim TargetField(1 To 29) As String
    

'====================================================================================================================================================
'����ڰ� �Է������ �ϴ� �Ķ���� ��
'====================================================================================================================================================
    KeyCol = "1,2,3,4,5"                                'ù��°���� �ټ���° �÷��� ������� �����ؼ� Ű�� �����
    
    InputFileName = "load_comm.csv"                     'CSV ���ϸ�
    TargetKey = "P029107001B_0_1_0_0"                   '�˻� ����� �Ǵ� Key ��
    'TargetField = "Alp_Ini_GP"                          '�˻� ����� �Ǵ� Field �̸�
    
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
    'CSV data file �� ������ ResultArray �迭�� ����. String �Ӽ��� �迭��
    '1st �Ķ����(InputFileName) : ��� CSV ���ϸ�
    '2nd �Ķ����(ResultArray()) : CSV ������ ��ȯ�ޱ� ���� �迭
        'index 0 : �ʵ� �̸�
        'index 1 : �ʵ� �Ӽ�
        'index 11~ : ������ ���ڵ�
    '3rd �Ķ����(KeyCol) : CSV ���� �� Key Field �� ��ġ�� ��� �ִ� ����. �޸��� ����.
    '4th �Ķ����(SortFlag) : �迭�� Sorting ����. 0 �̸� �������� ����. 1�̸� ����(QuickSort)
    '[����] ByRef �� ó���Ǵ� �迭������ �ݵ�� 0 ���� �ε��� ���� �ʿ�. �׷��� ������ �ϳ��� �ڷ� �и��� ��
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
    'VLookup Function : ��ȯ���� �ݵ�� Double type
    '1st �Ķ����(TargetKey) : ã���� �ϴ� Key ��
    '2nd �Ķ����(ResultArray()) : Source Array ��ü
    '3rd �Ķ����(j) : ���� ������ ���� ��� �÷� ����
'====================================================================================================================================================
    
'====================================================================================================================================================
    'Lookup Function : ��ȯ���� �ݵ�� Double type
    '1st �Ķ����(TargetKey) : ã���� �ϴ� Key ��
    '2nd �Ķ����(ResultArray()) : Source Array ��ü
    '3rd �Ķ����(FieldName) : ���� ������ ���� ��� �÷� �̸�
'====================================================================================================================================================

'====================================================================================================================================================
    'VLookupAll Function - VLookup �� �����ϳ� TargetKey �� �ش��ϴ� �� �� ��ü�� ��ȯ����
    '��ȯ���� �ݵ�� String Type. comma �� ���еǾ��ֱ� ������ �ݵ�� Split ó������� ��
    '1st �Ķ����(TargetKey) : ã���� �ϴ� Key ��
    '2nd �Ķ����(ResultArray()) : Source Array ��ü
'====================================================================================================================================================
    
'====================================================================================================================================================
    'HLookupAll Function - VLookup �� �����ϳ� TargetField �� �ش��ϴ� �� �� ��ü�� ��ȯ����
    '��ȯ���� �ݵ�� String Type. comma �� ���еǾ��ֱ� ������ �ݵ�� Split ó������� ��
    '1st �Ķ����(TargetField) : ã���� �ϴ� Key ��
    '2nd �Ķ����(ResultArray()) : Source Array ��ü
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


