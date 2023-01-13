from distutils.command.build_scripts import first_line_re
import openpyxl
import random
import math

import base64
## Crypto는 pip install pycryptodome
from Crypto import Random
from Crypto.Cipher import AES
import random


BS = 16
pad = lambda s: s + (BS - len(s.encode('utf-8')) % BS) * chr(BS - len(s.encode('utf-8')) % BS)
unpad = lambda s : s[:-ord(s[len(s)-1:])]

class AESCipher:
    def __init__(self, key):
        self.key = bytes.fromhex(key.zfill(32))
    
    # 이름처럼 다른 사람이랑 같을 수도 있는 걸 
    # 아웃풋을 다르게 보이게 하기 위함
    def encrypt_with_iv(self, raw, iv):
        raw = pad(raw)
        cipher = AES.new(self.key, AES.MODE_CBC, iv)
        return base64.b64encode(cipher.encrypt(raw.encode('utf-8')))
    
    def decrypt_with_iv(self, enc, iv):
        enc = base64.b64decode(enc)
        cipher = AES.new(self.key, AES.MODE_CBC, iv)
        return cipher.decrypt(enc).decode('utf-8')
    
    # ID 처럼 고유 값을 암호화할 경우 쓸 함수
    def encrypt_without_iv(self, raw):
        raw = pad(raw)
        cipher = AES.new(self.key, AES.MODE_ECB)
        return base64.b64encode(cipher.encrypt( raw.encode('utf-8')))
    
    def decrypt_without_iv(self, enc):
        enc = base64.b64decode(enc)
        cipher = AES.new(self.key, AES.MODE_ECB)
        return cipher.decrypt(enc).decode('utf-8')

## 엑셀 여는 코드
wb = openpyxl.load_workbook('./data.xlsx')
first_sheet = wb['본선Data']

print("phase 1")
# ## 주소 변경(ex) 서울특별시 성북구 동소문로24가길 16-4 => 서울특별시 성북구 )
## 집 주소랑 사업체 주소 둘 다 처리
for i in range(100000):
    old_address = str(first_sheet['H' + str(i + 2)].value)
    old_address = old_address.split(' ')
    first_sheet['H' + str(i + 2)].value = old_address[0] + ' ' + old_address[1]
    
    old_address = str(first_sheet['I' + str(i + 2)].value)
    old_address = old_address.split(' ')
    first_sheet['I' + str(i + 2)].value = old_address[0] + ' ' + old_address[1]

print("phase 2")
## 나이 0~3이면 내림 4~6은 5로 7~9는 올림
## 만약 만나이 조정 결과가 85 이상이면 85으로 바꿈
## 그것에 맞추어서 생년월일의 년도 갱신
for i in range(100000):
    degree = int(first_sheet['F' + str(i + 2)].value)
    if degree == 8 or degree == 9:
        first_sheet['F' + str(i + 2)].value = random.randrange(8, 10)
        

print("phase 3")
## 나이 0~3이면 내림 4~6은 5로 7~9는 올림
## 만약 만나이 조정 결과가 85 이상이면 85으로 바꿈
## 그것에 맞추어서 생년월일의 년도 갱신
for i in range(100000):
    age = int(first_sheet['D' + str(i + 2)].value)
    if age % 10 < 4:
        age = (age // 10) * 10
    elif age % 10 > 6:
        age = (age // 10 + 1) * 10
    else:
        age = ((age // 10) * 10) + 5
        
    if age > 85:
        age = 85
    year = 2022 - age
    date = str(first_sheet['J' + str(i + 2)].value).split('-')
    month = date[1]
    day = date[2]
    
    first_sheet['D' + str(i + 2)].value = age
    first_sheet['J' + str(i + 2)].value = str(year) + '-' + date[1] + '-' + date[2]

print("phase 4")
## 이름은 iv를 일련번호로 넣어서 사용한 AES-CBC로 암호화
## 키는 원래 무작위로 넣어야 하지만 여기서는 편의로
## key1 = 'fcb467442aed13872e3d7e33d29448c2' 일련 번호를 iv로 만들때 사용 설정함
## key2 = '682cbeeb75ad969468c938e71c7a1dbe' 이름을 암호화 할때 키로 설정
key1 = 'fcb467442aed13872e3d7e33d29448c2'
key2 = '682cbeeb75ad969468c938e71c7a1dbe'
for i in range(100000):
    
    ID = str(first_sheet['A' + str(i + 2)].value)
    name = str(first_sheet['B' + str(i + 2)].value)
    ## IV 없이 돌리는 코드
    encrypted_ID = AESCipher(key1).encrypt_without_iv(ID)

    ## IV를 만들어내는 코드
    hex_encrypted = bytes.hex(encrypted_ID)
    if len(hex_encrypted) < 32:
        hex_encrypted = hex_encrypted.zfill(32)
    else:
        hex_encrypted = hex_encrypted[0:32]
    iv = bytes.fromhex(hex_encrypted)
    ## IV 있이 돌리는 코드
    encrypted_name = AESCipher(key2).encrypt_with_iv(name, iv)

    ## bytes를 utf-8 인코딩으로 string으로 돌리는 방법
    first_sheet['B' + str(i + 2)].value = str(encrypted_name, 'utf-8')

print("phase 5")
## 결과 저장
#wb.save('./out.xlsx')

