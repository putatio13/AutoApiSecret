# -*- coding: UTF-8 -*-
import requests as req
import json,sys,time
#先注册azure应用,确保应用有以下权限:
#files:	Files.Read.All、Files.ReadWrite.All、Sites.Read.All、Sites.ReadWrite.All
#user:	User.Read.All、User.ReadWrite.All、Directory.Read.All、Directory.ReadWrite.All
#mail:  Mail.Read、Mail.ReadWrite、MailboxSettings.Read、MailboxSettings.ReadWrite
#注册后一定要再点代表xxx授予管理员同意,否则outlook api无法调用






path=sys.path[0]+r'/1.txt'
num1 = 0

def gettoken(refresh_token):
    headers={'Content-Type':'application/x-www-form-urlencoded'
            }
    data={'grant_type': 'refresh_token',
          'refresh_token': 0.AAAAaoJe83w4gUmNaWLH5nfHKxQ3BP0MjsJEj6Wtfer59-lwAO4.AgABAAAAAAB2UyzwtQEKR7-rWbgdcBZIAQDs_wIA9P9H7zxR-z4llW1-VXEx9tEp0lQgngy5qYfZAv2w7YHw1EVDJHrGvx2U5vdnAUg6IL37fKNQnzRR4N0Be6hATREJyB9dnEfW1piLr4A_RAsI_SZYaQOEYCi-D9r967ddesi04e_zxFMOrBKV71Eq8mqHG2Tvx-d7Sw2zbhZFHMu-lGPrp3h4MRFOI3NokKeRVzFEqBpE3yFMHAgpL93IoIbNK2zGMmqB4YXQVJ0eb4go9eeR17mo48KXzJeXDMQCGFo-G95STqN0RW8FYb-LokT7vQaWuKDo6TMTyU5lOXSkm7UbNdGVDUrMBAfXXxbiYh7Lwl7uqOgGCNhc4Z0KYFg8PnIvAVTafkpUhE5Mu4avaiYbBGGJdHyEtyg2lCbnLXxfQVuUdEPeIMU7-ZF5IjkEve9DPt4skYNUm1rZiqOpVhB-ph-TS4VU4kjZ4KTTxfxdMabKNykWrBDZhMLGs3bZh_5RBLBPv4k4mDoES_NOFA13_Jk8VqBibOnjxtEp3_ySBSOALDxTY_nBMxusB8IdDmO6cvllWQx5c9PdUd1iER9Vw8DRlysofAdXe9pf2DOvdSy54jdqv4_REnw5jEUWot2ifxYIAFOnjIIpXdMk97Eg6B8je3qZdmg0IT-jXyyV-t6YXBX6qetwtjXfZQycxPddUP_m0aF-Hk-B6aVvCvfkW22MVTDdwKNNPTgsnyJCMw634N0lEpEfNepdp6UhiECH4aqtY0lq5K7xPYmSIupaxzh577A3DZx0bqFYmFAB0Q-WpfzYjRsfSVIc9rBVYnIRb8HOmfKSrB3ouxhHhkMjncy7MpCqWL20ZTdbLwFNyjHy4XnOwCdt8LHw9Xa6Vvgzdft7PhVApoZmAJ_n4_n_YVPhKoBHRdEAvfw-bHiUosjDE9mJ4iAQAj1NuDFaynpGzRPOAzqtJnROE8hiqarzpK7s,
          'client_id':fd043714-8e0c-44c2-8fa5-ad7deaf9f7e9,
          'client_secret':QY8_Cbo-.ydZ-.t2u3720n.M4ULfo~~HGg,
          'redirect_uri':'http://localhost:53682/'
         }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    with open(path, 'w+') as f:
        f.write(refresh_token)
    return access_token
def main():
    fo = open(path, "r+")
    refresh_token = fo.read()
    fo.close()
    global num1
    localtime = time.asctime( time.localtime(time.time()) )
    access_token=gettoken(refresh_token)
    headers={
    'Authorization':access_token,
    'Content-Type':'application/json'
    }
    try:
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive',headers=headers).status_code == 200:
            num1+=1
            print("2调用成功"+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/drive/root',headers=headers).status_code == 200:
            num1+=1
            print('3调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/users ',headers=headers).status_code == 200:
            num1+=1
            print('4调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/messages',headers=headers).status_code == 200:
            num1+=1
            print('5调用成功'+str(num1)+'次')    
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',headers=headers).status_code == 200:
            num1+=1
            print('6调用成功'+str(num1)+'次')    
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta',headers=headers).status_code == 200:
            num1+=1
            print('7调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root/children',headers=headers).status_code == 200:
            num1+=1
            print('8调用成功'+str(num1)+'次')
        if req.get(r'https://api.powerbi.com/v1.0/myorg/apps',headers=headers).status_code == 200:
            num1+=1
            print('8调用成功'+str(num1)+'次') 
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders',headers=headers).status_code == 200:
            num1+=1
            print('9调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories',headers=headers).status_code == 200:
            num1+=1
            print('10调用成功'+str(num1)+'次')
            print('此次运行结束时间为 :', localtime)
    except:
        print("pass")
        pass
for _ in range(3):
    main()
