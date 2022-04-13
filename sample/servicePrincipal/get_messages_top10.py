# 必要な API アクセス許可
# Microsoft Graph/アプリケーションの許可/Mail.Read

import os
import msal
import requests

config = {
    'user_id': ''
}

authority = f"https://login.microsoftonline.com/{os.environ['TENANT_ID']}"
scopes = ['https://graph.microsoft.com/.default']
endpoint = f"https://graph.microsoft.com/v1.0/users/{config['user_id']}/messages?$select=sender,subject"

def connect_aad():
    cred = msal.ConfidentialClientApplication(
        client_id=os.environ['CLIENT_ID'],
        client_credential=os.environ['CLIENT_SECRET'],
        authority=authority)

    return cred

def get_access_token(cred):
    res = None
    res = cred.acquire_token_silent(scopes,account=None)

    if not res:
        print("キャッシュ上に有効なアクセストークンがありません。Azure AD から最新のアクセストークンを取得します。")
        res = cred.acquire_token_for_client(scopes=scopes)

    # if "access_token" in res:
    #     print(f"有効なアクセストークンを取得しました。token={res['access_token']}")
    return res

def get_mail(token):
    res = requests.get(
        endpoint,
        headers={'Authorization': 'Bearer ' + token['access_token']})

    if res.ok:
        print('最新 10 件分のメール取得に成功しました。')
        data = res.json()
        for email in data['value']:
            print(f"件名: {email['subject']}, 差出人: {email['sender']['emailAddress']['name']} <{email['sender']['emailAddress']['address']}>")
            # email = {
            #     '@odata.etag': '***',
            #     'id': '***',
            #     'subject': '***',
            #     'sender': {
            #         'emailAddress': {
            #             'name': '***',
            #             'address': '***'
            #         }
            #     }
            # }

    else:
        print(res.json())

def main():
    cred = connect_aad()
    token = get_access_token(cred)

    if "access_token" in token:
        get_mail(token)
    else:
        print(token.get("error"))
        print(token.get("error_description"))
        print(token.get("correlation_id"))

if __name__ == '__main__':
    main()