# 必要な API アクセス許可
# Microsoft Graph/アプリケーションの許可/Mail.Send

import os
import msal
import requests

config = {
    'user_from': '',
    'user_to': '',
    'user_cc': '',
    'user_bcc': ''
}

email_msg = {
    'Message': {
        'Subject': 'Python でのメール送信テスト',
        'Body': {
            'ContentType': 'Text',
            'Content': 'これはテストメールです。'
        },
        # 'BccRecipients': [{
        #     'EmailAddress': {'Address': config['user_bcc']}
        # }],
        # 'CcRecipients': [{
        #     'EmailAddress': {'Address': config['user_cc']}
        # }],
        'ToRecipients': [{
            'EmailAddress': {'Address': config['user_to']}
        }]
    },
    'SaveToSentItems': 'true'
}


authority = f"https://login.microsoftonline.com/{os.environ['TENANT_ID']}"
scopes = ['https://graph.microsoft.com/.default']
endpoint = f"https://graph.microsoft.com/v1.0/users/{config['user_from']}/sendMail"

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

def send_mail(token):
    res = requests.post(
        endpoint,
        headers={'Authorization': 'Bearer ' + token['access_token']},
        json=email_msg)

    if res.ok:
        print("メール送信に成功しました。")
    else:
        print(res.json())

def main():
    cred = connect_aad()
    token = get_access_token(cred)

    if "access_token" in token:
        send_mail(token)
    else:
        print(token.get("error"))
        print(token.get("error_description"))
        print(token.get("correlation_id"))

if __name__ == '__main__':
    main()