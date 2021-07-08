import numpy as np
import pandas as pd
import openpyxl

print('테이블 명세서 코드')

#엑셀 파일 불러오기
#데이터 처리
dbInfo = pd.read_excel('saleviscaretable.xlsx')
print(dbInfo.head())
## 반복문을 위한 테이블명을 조사 중복을 없앤다
dbInfo['테이블명'].duplicated()
## true == 중복, false == 최초의 값, 즉, false만 뽑으면 각 테이블명만 담아낼수 있다..

dbTableNames = dbInfo['테이블명'].drop_duplicates()


arrTableNames = np.array(dbTableNames.values.tolist())
print("------어레이에 담긴 테이블명들 확인-------")
print(arrTableNames)
print("----------------------------------------")
##테이블명의 인덱스 넘버
tableNameIndex = dbTableNames.index
#테이블명의 인덱스로 해당 테이블명의 논리테이블명을 가져온다
arrTableInfo =[]
for i in tableNameIndex:
    x = dbInfo.at[i, '논리테이블명']
    arrTableInfo.append(x)


# Workbook == 엑셀파일전체 AND Worksheet == 엑셀시트하나하나 
# spreadsheet 라는 변수에 .active 사용시 활성화된 시트를 불러온다. 초기엔 하나뿐..
# 시트이름 바꾸기 ex) sheet2.title = '수집 데이터'

# 원하는 셀에 값 입력하기.. 예) A1, B1, C1, D1 == '테이블 이름'

## 데이터프레임처럼 pd.read_excel 로도 엑셀파일을 불러올 수 있지만, 
## Workbook 의 메소드로도 기존 엑셀파일을 로드 할수 있다.
## 예) openpyxl.load_workbook(파일명) ==> wb=openpyxl.load_workbook('test.xlsx')
## 엑셀파일 저장은 Workbook.save(파일명) 으로 가능하다.
## 예) wb.save('test.xlsx')

## 엑셀 파일 만들기



def createWB(a):
    wb = openpyxl.Workbook()
    print(str(a) + "번째 엑셀시트 생성")
    sheets = ['sheet%d' % a]
    for x in sheets:

        x = wb.create_sheet('sheet' + str(a))
        x.title = '테이블명세서' + str(a)
        print(str(a) + "번째 엑셀시트의 헤드 작성")
        printTableHead(x)
        print(str(a) + "번째 엑셀시트의 값 작성")
        printTableValues(x,a)
    
    ##엑셀파일 저장
    wb.save(filename="test.xlsx")
    

## 테이블명세서의 헤더 만드는 함수.

def printTableHead(x):

    x.merge_cells('A1:D1') 
    x['A1']= '테이블 이름'
    x.merge_cells('E1:J1')
    x.merge_cells('A2:D2') 
    x['A2']= '테이블 설명'
    x.merge_cells('E2:J2')
    x.merge_cells('A3:D3') 
    x['A3']= 'PRIMARY KEY'
    x.merge_cells('E3:J3') 
    x.merge_cells('A4:D4') 
    x['A4']= 'FOREIGN KEY'
    x.merge_cells('A5:D5') 
    x['A5']= 'INDEX'
    x.merge_cells('A6:D6') 
    x['A6']= 'UNIQUE INDEX'
    x.merge_cells('E4:J4')
    x.merge_cells('E5:J5') 
    x.merge_cells('E6:J6') 
    x['A7']= 'NO'
    x['B7']= 'PK'
    x['C7']= 'AI'
    x['D7']= 'FK'
    x['E7']= 'NULL'
    x['F7']= '컬럼 이름'
    x['G7']= 'TYPE'
    x['H7']= 'DEFAULT'
    x['I7']= '설명'
    x['J7']= '참조 테이블'

##테이블의 값 넣는 함수
def printTableValues(x,a):
    print(str(a) + "값 insert 시작")

    i = arrTableNames[a]
    j = arrTableInfo[a]
    #테이블명 입력
    x['E1'] = i
    #논리테이블명입력
    x['E2'] = j
    ## 각 테이블명의 각각의 데이터프레임 새로 생성
    tableDetailByName = dbInfo.loc[dbInfo['테이블명'] == arrTableNames[a]]
    #PK조사
    thisPK = tableDetailByName.loc[tableDetailByName['KEY']=='PRI']
    PKString = thisPK['컬럼명'].values
    yStr=''
    if not PKString.all():
        print("PK 컬럼 없음")
    else:
        lenPK = len(PKString)
        for y in range(lenPK):
            yStr += thisPK['컬럼명'].values[y]

        #PK정보입력
        x['E3']= yStr
    #FK조사
    thisFK = tableDetailByName.loc[tableDetailByName['KEY']=='MUL']
    FKString = thisFK['컬럼명'].values
    #FK정보입력
    zStr=''
    if not FKString.all():
        print("FK컬럼 없음")
    else:
        lenFK = len(FKString)
        for z in range(lenFK):
            zStr += thisFK['컬럼명'].values[z]
        x['E4']= zStr

    #컬럼의 갯수조사 인덱스 설정
    nn = tableDetailByName.index +1
    #컬럼 인덱스 입력
    for i, value in enumerate(nn):
        x.cell(row=i+8, column=1, value=value)
    #컬럼명 조사
    columnNames = tableDetailByName['컬럼명']
    cn = columnNames.values
    #컬럼명 입력
    for i, value in enumerate(cn):
        x.cell(row=i+8, column=6, value=value)
    #데이터 형 조사
    dataTypes = tableDetailByName['데이터 길이']
    dt = dataTypes.values
    #데이터 형 입력
    for i, value in enumerate(dt):
        x.cell(row=i+8, column=7, value=value)
    #컬럼설명 조사
    columnDescs = tableDetailByName['컬럼설명']
    cd = columnDescs.values
    #컬럼설명 입력
    for i, value in enumerate(cd):
        x.cell(row=i+8, column=9, value=value)
    #눌 허용 조사
    isNull = tableDetailByName['NULL허용']
    isn = isNull.values
    # 눌 허용값 입력
    for i, value in enumerate(isn):
        x.cell(row=i+8, column=5, value=value)
    #디폴트값 조사
    isDefault = tableDetailByName['디폴트값']
    isd = isDefault.values
    isdS = isd.astype(str)

    #디폴트값 입력
    for i, value in enumerate(isdS):
        x.cell(row=i+8, column=8, value=value)
    
    ## PK와 FK 컬럼에 'Y' 표기
    ## .values 또는 to._numpy()  numpy배열로 변환함!
    ## 인덱싱으로 특정값을 변경한다. (Indexing: a[a < 0] = 0) 
    isThisPK = tableDetailByName['KEY'].to_numpy()
    isThisPK[isThisPK=='PRI']='Y'
    for i, value in enumerate(isThisPK):
        x.cell(row=i+8, column=2, value=value)
    
    isThisFK = tableDetailByName['KEY'].values
    isThisFK[isThisFK=='MUL'] = 'Y'
    isThisFK[isThisFK!='MUL'] = ''
    for i, value in enumerate(isThisFK):
        x.cell(row=i+8, column=4, value=value)
    ## 데이터프레임에 아예 변화를 주기때문에 엑셀화 할때 제일 마지막단계에서 실행해주고 다음
    ## 테이블로 넘어간다..


#테이블명의 length
lenarr = len(arrTableNames)


#for loop 을 이용하여 함수호출!

print("테이블명세서 작업 시작")
for a in range(lenarr):
    
    createWB(a)
            
