# CSV Access

AlphaBooster 에서 CSV file 을 직접 읽어서 Array 저장 후 Lookup 하기 위한 VBA Code

# 작동 프로세스 및 구조
# ![CSV Access Process](https://raw.githubusercontent.com/MillimanKorea/CSVAccess/master/CSVAccess.png)

# [UDF.BAS](https://github.com/MillimanKorea/VBAUtils/blob/master/UDF.bas)


1. **Sub CSVImport(CSVFileName, Array(), KeyCol)**
 + CSV data file 의 내용을 TotResultArray 배열에 저장. String 속성의 배열임
 + 1st 파라미터(CSVFileName) : 대상 CSV 파일명
 + 2nd 파라미터(Array()) : CSV 내용을 반환받기 위한 배열
 + 3rd 파라미터(KeyCol) : CSV 내용 중 Key Field 의 위치를 담고 있는 변수. 콤마로 구분

2. **Function VLookup(TargetKey, Array(), j)**
 + VLookup Function : 반환값은 반드시 Double type
 + 1st 파라미터(TargetKey) : 찾고자 하는 Key 값
 + 2nd 파라미터(Array()) : Source Array 전체
 + 3rd 파라미터(j) : 값을 가지고 오는 대상 컬럼 순번

3. **Function VLookupAll(TargetKey, Array())**
 + VLookupAll Function - VLookup 과 유사하나 TargetKey 에 해당하는 행 값 전체를 반환받음
 + 반환값은 반드시 String Type. comma 로 구분되어있기 때문에 반드시 Split 처리해줘야 함
 + 1st 파라미터(TargetKey) : 찾고자 하는 Key 값
 + 2nd 파라미터(Array()) : Source Array 전체

4. **Function HLookupAll(TargetField, Array())**
 + HLookupAll Function - VLookup 과 유사하나 TargetField 에 해당하는 열 값 전체를 반환받음
 + 반환값은 반드시 String Type. comma 로 구분되어있기 때문에 반드시 Split 처리해줘야 함
 + 1st 파라미터(TargetField) : 찾고자 하는 Key 값
 + 2nd 파라미터(Array()) : Source Array 전체
 

# [CSV.BAS](https://github.com/MillimanKorea/VBAUtils/blob/master/CSV.bas)
UDF.BAS 의 sub 및 function 을 이용해서 CSV File Access 를 하는 Sample Code
