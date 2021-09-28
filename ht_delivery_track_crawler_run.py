# -*- coding: utf8 -*-

# pip install requests
# pip install beautifulsoup4
# pip install firebase-admin

from datetime import datetime

import firebase_admin
import requests
from bs4 import BeautifulSoup
from firebase_admin import credentials, firestore
from common import *

week_days = ['월', '화', '수', '목', '금', '토', '일']

def getDateTimeString(date):
    dateString = date.strftime('%Y-%m-%d')
    timeString = date.strftime('%H:%M')
    return f"{dateString} ({week_days[date.weekday()]}) {timeString}"

# 한진택배 조회 
def getHanJinTrack(trackId):
    tracker = {}
    try:
        url = f'http://www.hanjinexpress.hanjin.net/customer/hddcw18_ms.tracking?w_num={trackId}'
        response = requests.get(url)
        html = response.text
        soup = BeautifulSoup(html, 'html.parser')
        list = soup.findAll(attrs={'bgcolor' : 'white'})
        progresses = []
        for item in list:
            tds = item.select("td")
            rows = {}
            rows['status'] = parseHanJinStatus(tds[0].text)
            times = tds[1].text.replace('\xa0\xa0','/').split('/')
            rows['time'] = getDateTimeString(datetime.datetime.strptime(f'{times[0]}-{times[1]}-{times[2]} {times[3]}:00', '%Y-%m-%d %H:%M:%S'))
            rows['location'] = tds[2].text
            progresses.append(rows)
        tracker['progresses'] = progresses
        if progresses:
            status = next((item for item in progresses if item['status']['id'] == 'delivered'), None)
            if not status:
                status = progresses[-1]
        else:
            title = soup.select("TITLE")[0].text
            if title == '조회에러':
                status = {'status' : parseHanJinStatus('송장오류'), 'time': '', 'location': ''}
            else:
                status = {'status' : parseHanJinStatus(''), 'time': '', 'location': ''}
                
        print(f"한진택배: {trackId} - {status['status']['text']}")
        tracker['status'] = status
    except Exception as e:
        print('Error getHanJinTrack: ', e)
    return tracker

def parseHanJinStatus(status): 
    return {
        '송장오류': { 'id': 'track_error',          'text': '송장오류' },
        '미배달':  { 'id': 'track_error',          'text': '미배달' },

        '화물접수': { 'id': 'at_pickup',            'text': '상품인수' },
        
        '화물입고': { 'id': 'in_transit',           'text': '화물입고' },
        '화물출발': { 'id': 'in_transit',           'text': '화물출발' },
        '화물도착': { 'id': 'in_transit',           'text': '화물도착' },
        '셔틀하차': { 'id': 'in_transit',           'text': '셔틀하차' },

        '배송출발': { 'id': 'out_for_delivery',     'text': '배송출발' },
        '배송완료': { 'id': 'delivered',            'text': '배송완료' },

    }.get(status, { 'id': 'information_received', 'text': '상품준비중' })

# CJ대한통운
def getCJlogisticsTrack(trackId):
    tracker = {}
    try:
        url = f'http://nplus.doortodoor.co.kr/web/detail.jsp?slipno={trackId}'
        response = requests.get(url)
        html = response.text        
        soup = BeautifulSoup(html, 'html.parser')
        tables = soup.select('body > center > table')[-1]
        list = tables.findAll(attrs={'bgcolor' : '#F6F6F6'})

        progresses = []
        
        for item in list:
            td = item.select("td")
            if td:
                rows = {}
                rows['status'] = parseCJlogisticsStatus(td[5].text.strip())
                rows['time'] = getDateTimeString(datetime.datetime.strptime(f'{td[0].text.strip()} {td[1].text.strip()}', '%Y-%m-%d %H:%M:%S'))
                rows['location'] = td[3].text
                progresses.append(rows)
        
        tracker['progresses'] = progresses

        if progresses:
            status = next((item for item in progresses if item['status']['id'] == 'delivered'), None)
            if not status:
                status = progresses[0]
        else:
            # if '미등록운송장' in html:
            #     status = parseHanJinStatus('송장오류')
            # else:
            #     status = parseHanJinStatus('')
            if '미등록운송장' in html:
                status = {'status' : parseCJlogisticsStatus('송장오류'), 'time': '', 'location': ''} 
            else:
                status = {'status' : parseCJlogisticsStatus(''), 'time': '', 'location': ''}

        print(f"CJ대한통운: {trackId} - {status['status']['text']}")
        tracker['status'] = status
    except Exception as e:
        print('Error getCJlogisticsTrack: ', e)
    return tracker

def parseCJlogisticsStatus(status): 
    return {
        '송장오류': { 'id': 'track_error',          'text': '송장오류' },
        '미배달':  { 'id': 'track_error',          'text': '미배달' },

        '집화처리': { 'id': 'at_pickup',            'text': '상품인수' },
        
        '간선하차': { 'id': 'in_transit',           'text': '간선하차' },
        '간선상차': { 'id': 'in_transit',           'text': '간선상차' },
        'SM인수': { 'id': 'in_transit',            'text': 'SM인수' },

        '배달출발': { 'id': 'out_for_delivery',     'text': '배송출발' },
        '배달완료': { 'id': 'delivered',            'text': '배송완료' },

    }.get(status, { 'id': 'information_received', 'text': '상품준비중' })

# 건영택배 
def getKunyoungTrack(trackId):
    tracker = {}
    try:
        url = f'https://www.kunyoung.com/goods/goods_02.php?mulno={trackId}'
        response = requests.get(url)
        html = response.content.decode('euc-kr', 'replace')
        soup = BeautifulSoup(html, 'html.parser')
        tables = soup.select('table[width="717"]')        
        list = tables[3].select('tr:nth-child(2n+4)')

        progresses = []
        
        for item in list:
            td = item.select("td:nth-child(2n+1)")
            if td:
                rows = {}
                if '배송완료' in td[1].text:
                    rows['status'] = parseKunyoungStatus("배송완료")
                else:
                    rows['status'] = parseKunyoungStatus("이동중")
                rows['time'] = getDateTimeString(datetime.datetime.strptime(f'{td[0].text}', '%Y-%m-%d %H:%M:%S'))
                rows['location'] = td[1].text
                progresses.append(rows)

        tracker['progresses'] = progresses

        if progresses:
            status = next((item for item in progresses if item['status']['id'] == 'delivered'), None)
            if not status:
                status = progresses[-1]
        else:
            if '미등록운송장' in html:
                status = {'status' : parseCJlogisticsStatus('송장오류'), 'time': '', 'location': ''} 
            else:
                status = {'status' : parseCJlogisticsStatus(''), 'time': '', 'location': ''}

        print(f"건영택배: {trackId} - {status['status']['text']}")
        tracker['status'] = status

    except Exception as e:
        print('Error getKunyoungTrack: ', e)
    return tracker

def parseKunyoungStatus(status): 
    return {
        '이동중': { 'id': 'in_transit',           'text': '이동중' },
        '배송완료': { 'id': 'delivered',            'text': '배송완료' },
    }.get(status, { 'id': 'information_received', 'text': '상품준비중' })

# 택배 배송 조회 실행
def runDeliveryTracks():
    try:
        # Firebase 연결 
        # cred = credentials.Certificate('./serviceAccountKey.json')
        # cred = credentials.Certificate('/Users/hantong/HantongLab/P_HantonDeliveryTrack/ht_delivery_track_crawler/serviceAccountKey.json')
        cred = credentials.Certificate('/home/intosharp/project/serviceAccountKey.json')
        firebase_admin.initialize_app(cred)
        db = firestore.client()

        # 배송조회 로딩 상태
        infoDocRef = db.collection('Hantong').document('Info')
        if infoDocRef.get().to_dict()['isUpdate'] == True:
            raise Exception('업데이트중입니다...') 

        try:
            infoDocRef.set({u'isUpdate' : True})

            batch = db.batch()

            # 조회 데이터 가져오기
            trackData = db.collection(u'TrackData').stream()
            for track in trackData:
                dict = track.to_dict()

                # 배송 상태 조회 
                data = {}
                if dict['carrierId'] == 'kr.hanjin':
                    # 한진택배
                    data = getHanJinTrack(dict['trackId'])
                elif dict['carrierId'] == 'kr.cjlogistics':
                    # CJ대한통운
                    data = getCJlogisticsTrack(dict['trackId'])
                elif dict['carrierId'] == 'kr.kunyoung':
                    # 건영택배
                    data = getKunyoungTrack(dict['trackId'])
                                
                if not data: 
                    # SSONG TODO: 배송 상태를 못 가져왔을때 어떻게 처리 할까 ? 
                    continue

                statusId        = data['status']['status']['id']
                statusText      = data['status']['status']['text']
                statusLocation  = data['status']['location']
                statusTime      = data['status']['time']

                # try:
                #     dateTime = statusTime.strftime('%m-%d %p %I:%M')
                #     statusTime = dateTime.replace('AM', '오전').replace('PM', '오후')
                # except Exception as ec:
                #     print('Worring: ', ec)

                # 기존 상태와 동일하면 업데이트 하지 않는다
                # if dict['statusId'] == statusId: 
                #    continue

                print(f"{dict['fileName']}: {dict['trackId']} - {statusId}")

                # ViewData Update
                viewData = db.collection('ViewData').document(dict['fileName'])
                field = {}
                field[dict['key'] + '.' + 'statusId']       = statusId
                field[dict['key'] + '.' + 'statusText']     = statusText
                field[dict['key'] + '.' + 'statusLocation'] = statusLocation
                field[dict['key'] + '.' + 'statusTime']     = statusTime
                batch.update(viewData, field)

                # TrackData Update or Delete
                # trackDocRef = db.collection('TrackData').document(dict['fileName']+'_'+dict['trackId'])
                trackDocRef = db.collection('TrackData').document(track.id)
                if statusId == 'delivered':
                    batch.delete(trackDocRef)
                else:
                    batch.update(trackDocRef, {u'statusId': statusId})

            # 조회 상태(카운트) 저장 
            time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            trackErrorCount = 0
            receivedCount = 0
            atPickupCount = 0
            inTransitCount = 0
            deliveryCount = 0
            deliveredCount = 0
            
            docDataIds = []
            docData = db.collection(u'ViewData').stream()            
            for doc in docData:
                docDataIds.append(doc.id)    

                dict = doc.to_dict()
                for key in dict['keys']:
                    statusId = dict[key]['statusId']
                    if statusId == 'track_error':
                        trackErrorCount += 1
                    elif statusId == 'information_received':
                        receivedCount += 1
                    elif statusId == 'at_pickup':
                        atPickupCount += 1
                    elif statusId == 'in_transit':
                        inTransitCount += 1
                    elif statusId == 'out_for_delivery':
                        deliveryCount += 1
                    elif statusId == 'delivered':
                        deliveredCount += 1
           
            startId = ""
            lastId  = ""

            try:
                startId = "%s-%s-%s" % (docDataIds[0][:4],docDataIds[0][4:6],docDataIds[0][6:8])
                lastId  = "%s-%s-%s" % (docDataIds[-1][:4],docDataIds[-1][4:6],docDataIds[-1][6:8])
            except Exception as eb:
                print('Error: ', eb)

            stateDocRef = db.collection('Hantong').document('State')
            batch.update(stateDocRef, 
                {
                    u'time':            time,
                    u'trackErrorCount': trackErrorCount,
                    u'receivedCount':   receivedCount,
                    u'atPickupCount':   atPickupCount,
                    u'inTransitCount':  inTransitCount,
                    u'deliveryCount':   deliveryCount,
                    u'deliveredCount':  deliveredCount,
                    u'startId':         startId,
                    u'lastId':          lastId,
                }
            )

            batch.commit()

        except Exception as ea:
            print('Error: ', ea)
        finally:
            infoDocRef.set({u'isUpdate' : False})

    except Exception as e:
        print('Error Firebase_admin Initialize_app: ', e)

print('=========================== Start');

time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
print('조회시간 : ', time)

runDeliveryTracks()

# print(getKunyoungTrack('1134426448'))

# SSONG TODO: 사유가 있으면 사유를 저장 - 화면에 어떻게 표시할지는 고민중. 

print('=========================== End');


