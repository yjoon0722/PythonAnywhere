import telegram
import os
import datetime
import time

# TimeZone 설정
os.environ["TZ"] = "Asia/Seoul"
time.tzset()

# 텔레그램 설정
token = "1996603464:AAEiX7uT2pzz_jFczCftsR2bN7VH1V9eiQw"
bot = telegram.Bot(token=token)
chat_id = "-1001375693771"

# 채팅창정보 알아내는 코드
# updates = bot.getUpdates()
# print(updates)
# for i in updates:
#     print(i)

# 봇이 메세지 보내는 코드
# bot.send_message(chat_id = '-527683268', text="테스트.")


week_days = ['월', '화', '수', '목', '금', '토', '일']

# 창고 이름
def get_warehouse_name(key):
    return '고려포장' if '한진' in key else '고려포장' if '건영' in key else '준테크' if 'CJ' in key else ''

# 택배 송장 번호
def get_invoice_number(invoice_number, delivery_message):
    if invoice_number is None or invoice_number == '':
        return delivery_message
    return invoice_number.replace("-","")

# 택배사 이름
def get_carrier_name(key):
    return '한진택배' if '한진' in key else '건영택배' if '건영' in key else 'CJ택배' if 'CJ' in key else ''

# 택배사 코드
def get_carrier_id(key):
    return 'kr.hanjin' if '한진' in key else 'kr.kunyoung' if '건영' in key else 'kr.cjlogistics' if 'CJ' in key else ''

# 택배사 배송 조회 URL
def get_carrier_url(key, invoice):
    invoice_number = invoice.replace("-","")
    if not invoice_number or not invoice_number.isdigit():
        return ""
    if '한진' in key:
        base_url = 'http://www.hanjinexpress.hanjin.net/customer/hddcw18_ms.tracking?w_num='
    elif '건영' in key:
        base_url = 'https://www.kunyoung.com/goods/goods_02.php?mulno='
    elif 'CJ' in key:
        base_url = 'http://nplus.doortodoor.co.kr/web/detail.jsp?slipno='
    else:
        base_url = None
    return base_url + invoice_number if base_url is not None else ""
