# VBA Utils

AlphaBooster 에서 CSV file 을 직접 읽어서 Array 저장 후 Lookup 하기 위한 Code


# [UDF.BAS](https://github.com/MillimanKorea/VBAUtils/blob/master/UDF.bas)

1. **Function VLookup**
+ VLookup Function : 반환값은 반드시 Double type
+ 1st 파라미터(TargetKey) : 찾고자 하는 Key 값
+ 2nd 파라미터(KeyArray()) : Source Array 중 Key 배열
+ 3rd 파라미터(TotResultArray()) : Source Array 전체
+ 4th 파라미터(j) : 값을 가지고 오는 대상 컬럼 순번
+ 5th 파라미터(No_Row) : 배열 행 갯수

2. **Function VLookupAll**
+ VLookupAll Function - VLookup 과 유사하나 TargetKey 에 해당하는 행 값 전체를 반환받음
+ 반환값은 반드시 String Type. comma 로 구분되어있기 때문에 반드시 Split 처리해줘야 함
+ 1st 파라미터(TargetKey) : 찾고자 하는 Key 값
+ 2nd 파라미터(KeyArray()) : Source Array 중 Key 배열
+ 3rd 파라미터(TotResultArray()) : Source Array 전체
+ 4th 파라미터(No_Row) : 배열 행 갯수

3. **Function HLookupAll**
+ HLookupAll Function - VLookup 과 유사하나 TargetField 에 해당하는 열 값 전체를 반환받음
+ 반환값은 반드시 String Type. comma 로 구분되어있기 때문에 반드시 Split 처리해줘야 함
+ 1st 파라미터(TargetField) : 찾고자 하는 Key 값
+ 2nd 파라미터(FieldRec) : CSV 파일 내용 중 필드의 변수명이 들어있는 첫번째 행의 내용
+ 3rd 파라미터(TotResultArray()) : Source Array 전체
+ 4th 파라미터(No_Row) : 배열 행 갯수

4. **Sub CSVImport**
+ CSV data file 의 내용을 TotResultArray 배열에 저장. String 속성의 배열임
+ 1st 파라미터(InputFileName) : 대상 CSV 파일명
+ 2nd 파라미터(TotResultArray()) : CSV 내용을 반환받기 위한 배열
+ 3rd 파라미터(FieldRec) : CSV 내용 중 각 Field 명을 반환받기 위한 변수. CSV 의 첫번째 줄에 해당함
+ 4th 파라미터(KeyArray()) : CSV 내용 중 Key Field 내용을 반환받기 위한 변수. 현재는 무조건 첫번째 컬럼
+ 5th 파라미터(KeyCol()) : CSV 내용 중 Key Field 의 위치를 담고 있는 변수. 특별히 지정하지 않으면 첫번째 컬럼만 Key Field 로 인식. 0 이면 적용하지 않음
+ 6th 파라미터(ColAttr()) : 컬럼별 속성("S"tring or "D"ouble) 반환
+ 7th 파라미터(No_Row) : CSV 데이터의 행 갯수 반환
+ 8th 파라미터(No_Col) : CSV 데이터의 열 갯수 반환
+ [주의] ByRef 로 처리되는 배열변수는 반드시 0 부터 인덱스 정의 필요. 그렇지 않으면 하나씩 뒤로 밀리게 됨


# [CSV.BAS](https://github.com/MillimanKorea/VBAUtils/blob/master/CSV.bas)
UDF.BAS 의 sub 및 function 을 이용해서 CSV File Access 를 하는 Sample Code
