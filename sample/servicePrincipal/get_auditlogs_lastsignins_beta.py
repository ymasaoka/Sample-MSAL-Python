# 必要な API アクセス許可
# Microsoft Graph/アプリケーションの許可/AuditLog.Read.All
# Microsoft Graph/アプリケーションの許可/Directory.Read.All

import os
import msal
import requests

authority = f"https://login.microsoftonline.com/{os.environ['TENANT_ID']}"
scopes = ['https://graph.microsoft.com/.default']
endpoint_users = 'https://graph.microsoft.com/beta/users'

def connect_aad():
    # クライアントシークレットを使用するパターン
    # cred = msal.ConfidentialClientApplication(
    #     client_id=os.environ['CLIENT_ID'],
    #     client_credential=os.environ['CLIENT_SECRET'],
    #     authority=authority)

    # SSL 証明書を使用するパターン
    # thumbprint (拇印) の値は、サービスプリンシパルに SSL 証明書を登録した際に表示される値を使用します
    secret_key_path = os.environ['CLIENT_CERTIFICATION_PATH']
    cred = msal.ConfidentialClientApplication(
        client_id=os.environ['CLIENT_ID'],
        client_credential={"thumbprint": os.environ['CLIENT_CERTIFICATION_THUMBPRINT'],"private_key": open(secret_key_path).read()},
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

def get_users_all_and_lastsignins(token):
    q = '?$select=id,userPrincipalName,userType,signInActivity'
    res = requests.get(
        f"{endpoint_users}{q}",
        headers={'Authorization': 'Bearer ' + token['access_token']})

    if res.ok:
        data = res.json()
        print(f"{len(data['value'])} 件のユーザー情報と最終サインイン日時の取得に成功しました。")
        print(data)

        for user in data['value']:
            print(f"UPN: {user['userPrincipalName']}, ユーザー種別: {user['userType']}, 最終ログイン日時(UTC): {user['signInActivity']['lastSignInDateTime']}")
            # user = {
            #     'userPrincipalName': '***', 
            #     'userType': 'Member', # Member or Guest 
            #     'id': '***', 
            #     'signInActivity': {
            #         'lastSignInDateTime': '2021-06-27T16:03:52Z', 
            #         'lastSignInRequestId': '18d3b341-727f-418b-995f-3be4bf11bd00', 
            #         'lastNonInteractiveSignInDateTime': '0001-01-01T00:00:00Z', 
            #         'lastNonInteractiveSignInRequestId': ''
            #     }
            # }
    else:
        print(res.json())

def main():
    cred = connect_aad()
    token = get_access_token(cred)

    if "access_token" in token:
        get_users_all_and_lastsignins(token)
    else:
        print(token.get("error"))
        print(token.get("error_description"))
        print(token.get("correlation_id"))

if __name__ == '__main__':
    main()