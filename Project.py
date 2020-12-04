#프로그래밍언어 프로그램 개발
import openpyxl
import random

# 안내문장 출력 함수
def sentence_print():
    print ('[내 집 마련을 위한 아파트 실거래가 조회 프로그램]\n')
    print ('1. 우리동네 검색')
    print ('2. 아파트명 검색')
    print ('3. 실거래가 검색')
    print ('4. 복합 조건 검색')
    print ('5. 랜덤 매물 추천')
    print ('6. 프로그램 종료')
    print (' * 원하는 항목의 숫자를 입력하여 주세요.\n')

# 복합조건 안내문장 출력 함수
def compound_sentence_print():
    print ('[복합 조건 선택]\n')
    print ('1. 동네 & 아파트명 검색')
    print ('2. 동네 & 실거래가 검색')
    print ('3. 아파트명 & 실거래가 검색')
    print (' * 원하는 항목의 숫자를 입력하여 주세요.\n')

# 랜덤항목 안내문장 출력 함수
def random_sentence_print():
    print ('[랜덤 조건 선택]\n')
    print ('1. 동네 랜덤 검색')
    print ('2. 아파트명 랜덤 검색')
    print ('3. 실거래가 랜덤 검색')
    print (' * 원하는 항목의 숫자를 입력하여 주세요.\n')

# 아파트 매물 정보 출력 함수
def houseinfo_print(*format):
    print ('\n[아파트 매물 정보]\n')
    arr = format                            
    aptprice = str(arr[0][6] * 0.0001) + "억원"
    print ('- 동네: ' + arr[0][0])
    print ('- 아파트명: ' + arr[0][2])
    print ('- 층: ' + arr[0][7])
    print ('- 면적: ' + arr[0][3])
    print ('- 가격: ' + aptprice)
    print ('- 계약년월: ' + arr[0][4])
    print ('- 건축년도: ' + arr[0][8] +'\n')
    return arr

# 랜덤기능에 대한 조건별 함수
def random_function(inputnumber):
    # 조회조건에 대한 입력숫자를 받아 대상건을 조회할 때 매물건수를 집계하여 랜덤범위에서 상한수치로 사용한다. 
    # (리스트의 시작 순서에 맞게 매물건수를 -1 진행, 
    #  또한 매물건수가 1건일 땐 랜덤범위는 0,0으로 실행이 안되기 때문에 강제 0으로 지정)
    copydict={}    #조회대상에 대한 복수의 리스트 데이터를 저장할 딕셔너리
    copyformat=[]  #딕셔내리 내 데이터를 변환저장 하기 위한 리스트
    if inputnumber == 1:
        search = input('검색할 동을 입력해주세요. ex: 역삼동, 방배동\n')
        apt_cnt=0  #아파트매물 건수 초기화
        for format in data:
          if search in format[0]:
            copydict[apt_cnt] = {apt_cnt: format[:]} #딕셔너리 내 리스트 복사
            apt_cnt=apt_cnt+1

        if apt_cnt == 0:
           print('조회 결과가 없습니다.\n')
        elif apt_cnt == 1:
           i = 0  #매물건수 1건일 때 리스트 첫번째 데이터만 나올 수 있게 강제 지정 list[0]
        else:
           i = random.randrange(0, apt_cnt-1) #순서가 0부터 시작되기에 -1 진행, ex: 아파트 매물이 3개일 때 순서는 0~2로 저장
        
        copyformat= list(copydict[i].values())       
        print('★총 매물 개수: ' + str(apt_cnt) + '★')
        houseinfo_print(copyformat[0])

    elif inputnumber == 2:
        search = input('검색할 아파트명을 입력해주세요. ex: 레미안, 청솔\n')
        apt_cnt=0  #아파트매물 건수 초기화
        for format in data:
          if search in format[2]:
            copydict[apt_cnt] = {apt_cnt: format[:]}
            apt_cnt=apt_cnt+1
            
        if apt_cnt == 0:
           print('조회 결과가 없습니다.\n')
        elif apt_cnt == 1:
           i = 0  #매물건수 1건일 때 리스트 첫번째 데이터만 나올 수 있게 강제 지정 list[0]
        else:
           i = random.randrange(0, apt_cnt-1) #순서가 0부터 시작되기에 -1 진행, ex: 아파트 매물이 3개일 때 순서는 0~2로 저장
        
        copyformat= list(copydict[i].values())       
        print('★총 매물 개수: ' + str(apt_cnt) + '★')
        houseinfo_print(copyformat[0])

    elif inputnumber ==3:
        search = input('검색할 실거래가 미만을 입력해주세요. ex: 200000000, 300000000\n')
        apt_cnt=0  #아파트매물 건수 초기화
        for format in data:
          if int(search) > int(format[6])*10000:  
            copydict[apt_cnt] = {apt_cnt: format[:]}
            apt_cnt=apt_cnt+1
                
        if apt_cnt == 0:
           print('조회 결과가 없습니다.\n')
        elif apt_cnt == 1:
           i = 0  #매물건수 1건일 때 리스트 첫번째 데이터만 나올 수 있게 강제 지정 list[0]
        else:
           i = random.randrange(0, apt_cnt-1) #순서가 0부터 시작되기에 -1 진행, ex: 아파트 매물이 3개일 때 순서는 0~2로 저장
        
        copyformat= list(copydict[i].values())       
        print('★총 매물 개수: ' + str(apt_cnt) + '★')
        houseinfo_print(copyformat[0])

# 엑셀파일 읽어오기 및 데이터 저장
filename = "Aptprice.xlsx" #파일명
filedata = openpyxl.load_workbook(filename) #엑셀파일 로드
detaildata = filedata.worksheets[0] #엑셀파일 내 시트

data = []
for row in detaildata.rows:
    data.append([
        row[0].value, #시군구
        row[1].value, #번지
        row[2].value, #아파트명
        row[3].value, #면적
        row[4].value, #계약년월
        row[5].value, #계약일
        row[6].value, #가격
        row[7].value, #층
        row[8].value, #건축년도
        row[9].value  #도로명
    ])

#메인 절차 진행
while True:
  try:
    sentence_print()
    inputnumber = input('입력란: ')
    if inputnumber == '1': #우리동네 검색
      search = input('검색할 동을 입력해주세요. ex: 역삼동, 방배동\n')
      for format in data:
        if search in format[0]:
            houseinfo_print(format)

    elif inputnumber == '2': #아파트명 검색
      search = input('검색할 아파트명을 입력해주세요. ex: 레미안, 청솔\n')
      for format in data:
         if search in format[2]:
            houseinfo_print(format)

    elif inputnumber == '3': #실거래가 검색
      search = input('검색할 실거래가 미만을 입력해주세요. ex: 200000000, 300000000\n')
      for format in data:
         if int(search) > int(format[6])*10000:  
            houseinfo_print(format)

    elif inputnumber == '4': #복합 조건 검색                    
      compound_sentence_print()
      compoundinputnumber = input('입력란: ')   
      if compoundinputnumber == '1':   #1-동네 & 아파트명 검색
        search1 = input('검색할 동을 입력해주세요. ex: 역삼동, 방배동\n')
        search2 = input('검색할 아파트명을 입력해주세요. ex: 레미안, 청솔\n')
        for format in data:
          if search1 in format[0] and search2 in format[2]:
            houseinfo_print(format)
      elif compoundinputnumber == '2': #2-동네 & 실거래가 검색
        search1 = input('검색할 동을 입력해주세요. ex: 역삼동, 방배동\n')
        search2 = input('검색할 실거래가 미만을 입력해주세요. ex: 200000000, 300000000\n')
        for format in data:
          if search1 in format[0] and int(search2) > int(format[6])*10000:
            houseinfo_print(format)
      elif compoundinputnumber == '3': #3-아파트명 & 실거래가 검색                 
        search1 = input('검색할 아파트명을 입력해주세요. ex: 레미안, 청솔\n')
        search2 = input('검색할 실거래가 미만을 입력해주세요. ex: 200000000, 300000000\n')
        for format in data:
          if search1 in format[2] and int(search2) > int(format[6])*10000:
            houseinfo_print(format)
      else:
        print('조건에 없는 숫자입니다!\n')
        print('*******************************\n')

    elif inputnumber == '5': #랜덤 매물 추천
      random_sentence_print()
      randominputnumber = input('입력란: ')   
      if randominputnumber == '1':   #1-동네 랜덤 검색
        random_function(1)
      elif randominputnumber == '2': #2-아파트명 랜덤 검색
        random_function(2)
      elif randominputnumber == '3': #3-실거래가 랜덤 검색
        random_function(3)
      else:
        print('조건에 없는 숫자입니다!\n')
        print('*******************************\n')

    elif inputnumber == '6': #프로그램 종료
      print ('종료 완료')
      break

    else:
      print('목록에 해당하는 숫자를 눌러주세요!\n')
      print('*******************************\n')
  except:
    exit      