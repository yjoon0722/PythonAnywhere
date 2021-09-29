import imaplib
import imapclient
import pyzmail
import pprint
import datetime
import time
import telegram
import os
from common import *

todayDate = datetime.datetime.now().date()
currentTime = datetime.datetime.fromtimestamp(time.time()).strftime("%Y.%m.%d %H:%M")

if todayDate.weekday() == 5 or todayDate.weekday() == 6 :
    pass
else :
    # 발주서 저장 path
    shipmentOrderPath = "/home/intosharp/ReceiveData/{}/0_Send/".format(str(todayDate).replace("-",""))
    # 송장번호 저장 path
    courierInvoicePath = "/home/intosharp/ReceiveData/{}/1_Receive/".format(str(todayDate).replace("-",""))

######################################################################################################################################

    def createFolder(directory):
        try:
            if not os.path.exists(directory):
                os.makedirs(directory)
        except Exception as e:
            print('Error: ',e)

    createFolder(shipmentOrderPath)
    createFolder(courierInvoicePath)

######################################################################################################################################

    # imap 서버 연결
    imap_obj = imapclient.IMAPClient('imap.naver.com',ssl=True)

    # imap 서버 로그인
    imap_obj.login('id','password')

    # imap 서버 메일 폴더 리스트 출력
    # pprint.pprint(imap_obj.list_folders())

    # imap 서버 폴더 선택
    # (Naver기준) Inbox = 받은메일함 / Sent Messages = 보낸메일함 / Drafts = 임시보관함
    # imap_obj.select_folder("Sent Messages", readonly=True) # 혹시 모를 파일 손상 대비 readonly

    # 선택한 imap 서버 폴더에서 이메일의 UUID값 검색
    # UIDs = imap_obj.search(["FROM","no_reply@email.apple.com"]) # 발신자로 검색
    # UIDs = imap_obj.search([u"ON", sendFileDate]) # 특정 날짜로 검색

    print("================ 다운로드 시작 ================")
    print(currentTime)
    try:
        # 발주서 발송 메일 검색
        imap_obj.select_folder("Sent Messages", readonly=True)
        sendUIDs = imap_obj.search([u"ON", todayDate])

        for uid in sendUIDs:
            raw_msg = imap_obj.fetch(uid,['BODY[]'])
            msg = pyzmail.PyzMessage.factory(raw_msg[uid][b'BODY[]'])
            for mp in msg.mailparts:
                if mp.filename != None and mp.filename.find('xlsx') != -1 and mp.filename.find('발주서') != -1 and mp.filename.find('용달') == -1:
                    print(mp.filename,len(mp.get_payload()))
                    # ** 파일 다운로드시 같은 이름의 파일이 있을경우 자동으로 덮어씌움 **
                    file = open(shipmentOrderPath + mp.filename, "wb")
                    file.write(mp.get_payload())
                    file.close()

        # 고려포장 받은 메일 검색 (당일)
        imap_obj.select_folder("영업관리/고려포장", readonly=True)
        receiveUIDs_today = imap_obj.search([u"ON", todayDate])

        for uid in receiveUIDs_today:
            raw_msg = imap_obj.fetch(uid,['BODY[]'])
            msg = pyzmail.PyzMessage.factory(raw_msg[uid][b'BODY[]'])
            for mp in msg.mailparts:
                # 파일이 없지않고, xlsx파일이며, 발송한 날짜의 고려포장파일
                if mp.filename != None and mp.filename.find('xlsx') != -1 and mp.filename.find('발주서') != -1 and mp.filename.find(str(todayDate).replace("-","")) != -1:
                    print(mp.filename,len(mp.get_payload()))
                    file = open(courierInvoicePath + mp.filename, "wb")
                    file.write(mp.get_payload())
                    file.close()

        # 준테크 받은 메일 검색 (당일)
        imap_obj.select_folder("영업관리/준테크", readonly=True)
        receiveUIDs_joontech_today = imap_obj.search([u"ON", todayDate])

        for uid in receiveUIDs_joontech_today:
            raw_msg = imap_obj.fetch(uid,['BODY[]'])
            msg = pyzmail.PyzMessage.factory(raw_msg[uid][b'BODY[]'])
            for mp in msg.mailparts:
                # 파일이 없지않고, xlsx파일이며, 준테크파일이고, 메일내용에 오늘 날짜가 포함
                if mp.filename != None and mp.filename.find('xlsx') != -1 and mp.filename.find('송장번호') != -1 and msg.text_part.get_payload().decode(msg.text_part.charset).find("{}월{}일".format(todayDate.month,todayDate.day)) != -1:
                    print(mp.filename,len(mp.get_payload()))
                    file = open(courierInvoicePath + mp.filename, "wb")
                    file.write(mp.get_payload())
                    file.close()

        print(currentTime)
        print("================ 다운로드 끝 ================")

    except Exception as e:
        bot.send_message(chat_id = chat_id, text="[{}]\n파일 다운로드 중 오류가 발생했습니다\nError : \n{}".format(currentTime,e))
        print("Error : ",e)
    finally:
        # imap 서버 로그아웃
        imap_obj.logout()

######################################################################################################################################
