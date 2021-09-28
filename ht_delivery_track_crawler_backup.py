
# pip install firebase-admin

from datetime import datetime

import firebase_admin
from firebase_admin import credentials, firestore
import requests
from common import *

# 택배 배송 데이터 백업
def runBackup():
    try:
        # Firebase 연결 
        #cred = credentials.Certificate('./serviceAccountKey.json')
        # cred = credentials.Certificate('/Users/hantong/HantongLab/P_HantonDeliveryTrack/ht_delivery_track_crawler/serviceAccountKey.json')
        cred = credentials.Certificate('/home/intosharp/project/serviceAccountKey.json')
        firebase_admin.initialize_app(cred)
        db = firestore.client()

        try:
            batch = db.batch()

            today = datetime.datetime.now()

            # 백업 데이터 가져오기            
            viewData = db.collection(u'ViewData').stream()
            for view in viewData:
                date_diff = today - datetime.datetime.strptime(f'{view.id}', '%Y%m%d')

                # 15일이 지난 데이터는 백업 한다. 
                if date_diff.days < 15:
                    continue

                print(f"{view.id} - {date_diff.days}")

                # TrackData 가 있는지 체크 해서 있으면 백업하지 않습니다.
                trackDocRefs = db.collection('TrackData').where(u'fileName', u'==', f'{view.id}').limit(1).stream()

                stream_empty = False
                for trackDocRef in trackDocRefs:
                    stream_empty = True

                if stream_empty:
                    print(f"TrackData에 데이터가 있습니다. {view.id}")
                    continue

                # BackupData로 데이터 이동 
                backupDocRef = db.collection(u'BackupData').document(f'{view.id}')
                batch.set(backupDocRef, view.to_dict(), merge=True)

                # ViewData로 데이터 삭제
                viewDocRef = db.collection('ViewData').document(f'{view.id}')
                batch.delete(viewDocRef)

            batch.commit()

        except Exception as ea:
            print('Error: ', ea)

    except Exception as e:
        print('Error Firebase_admin Initialize_app: ', e)

print('=========================== Start');
time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
print('백업시간 : ', time)

runBackup()

print('=========================== End');
