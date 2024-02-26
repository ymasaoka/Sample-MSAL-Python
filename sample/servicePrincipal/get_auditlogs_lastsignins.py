# 必要な API アクセス許可
# Microsoft Graph/アプリケーションの許可/AuditLog.Read.All
# Microsoft Graph/アプリケーションの許可/Directory.Read.All

import os
import msal
import requests

authority = f"https://login.microsoftonline.com/{os.environ['TENANT_ID']}"
scopes = ['https://graph.microsoft.com/.default']
endpoint_users = 'https://graph.microsoft.com/v1.0/users'
endpoint_auditlogs = 'https://graph.microsoft.com/v1.0/auditLogs/signIns'

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
        client_credential={"thumbprint": os.environ['THUMBPRINT'],"private_key": open(secret_key_path).read()},
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

def get_users_all(token):
    p = '?$select=id,userPrincipalName'
    res = requests.get(
        f"{endpoint_users}{p}",
        headers={'Authorization': 'Bearer ' + token['access_token']})

    if res.ok:
        data = res.json()
        print(f"{len(data['value'])} 件のユーザー情報取得に成功しました。")
        return data
    else:
        print(res.json())

def get_auditlog_lastsingin(token, users):
    for user in users['value']:
        q = f"?$filter=userId eq '{user['id']}'&$orderby=createdDateTime desc&$top=1"
        res = requests.get(
            f"{endpoint_auditlogs}{q}",
            headers={'Authorization': 'Bearer ' + token['access_token']})
        data = res.json()

        if res.ok & len(data['value']) > 0:
            data = res.json()
            print(f"UPN: {data['value'][0]['userPrincipalName']}, 最終ログイン日時(UTC): {data['value'][0]['createdDateTime']}")
        elif res.ok & len(data['value']) == 0:
            print(f"UPN: {user['userPrincipalName']}, 最終ログイン日時(UTC): 0")
        else:
            print(data)

def main():
    cred = connect_aad()
    token = get_access_token(cred)

    if "access_token" in token:
        users = get_users_all(token)
        get_auditlog_lastsingin(token, users)
    else:
        print(token.get("error"))
        print(token.get("error_description"))
        print(token.get("correlation_id"))

if __name__ == '__main__':
    main()