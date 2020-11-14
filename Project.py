#프로그래밍언어 프로그램 개발
import openpyxl

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
    

filename = "abc.xlsx"
filedata = openpyxl.load_workbook(filename)
detaildata = filedata.worksheets[0]

data = []
for row in detaildata.rows:
    data.append([
        row[0].value,
        row[1].value,
        row[2].value,
        row[3].value,
        row[4].value,
        row[5].value,
        row[6].value,
        row[7].value,
        row[8].value,
        row[9].value
    ])

print(data)


while True:
    sentence_print()
    inputnumber = input('입력란: ')
    if inputnumber == '1':
      search = input('검색할 동을 입력해주세요. ex: 역삼동, 방배동\n')

    elif inputnumber == '2':
      search = input('검색할 아파트명을 입력해주세요. ex: 레미안, 청솔\n')

    elif inputnumber == '3':
      search = input('검색할 실거래가 미만을 입력해주세요. ex: 200000000, 300000000\n')

    elif inputnumber == '4':                    
      search = input('검색할 아파트명을 입력해주세요. ex: 레미안, 청솔\n')

    elif inputnumber == '5':                    
      search = input('구상중\n')        

    elif inputnumber == '6':
      print ('종료 완료')
      break
    else:
      print('목록에 해당하는 숫자를 눌러주세요!\n')
      print('*******************************\n')