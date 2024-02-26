# 必要な API アクセス許可
# Microsoft Graph/アプリケーションの許可/User.ReadWrite.All

import os
import msal
import requests

authority = f"https://login.microsoftonline.com/{os.environ['TENANT_ID']}"
scopes = ['https://graph.microsoft.com/.default']

is_physical_delete = False

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

def get_disabled_guests(token):
    endpoint = 'https://graph.microsoft.com/v1.0/users'
    q = '$select=id,displayName,userPrincipalName,mail,userType,accountEnabled&$filter=userType eq \'Guest\' and accountEnabled eq false'
    res = requests.get(
        f"{endpoint}{q}",
        headers={'Authorization': 'Bearer ' + token['access_token']})

    if res.ok:
        data = res.json()
        print(f"{len(data['value'])} 件の無効化されているゲストユーザー情報の取得に成功しました。")
        return data
    else:
        print(res.json())

def delete_disabled_guests(token, users):
    for user in users['value']:
        # user = {
        #     'id': '***',
        #     'displayName': '***',
        #     'userPrincipalName': '***#EXT#@***',
        #     'mail': "***@***",
        #     'userType': 'Guest'
        #     'accountEnabled': false
        # }
        endpoint = f"https://graph.microsoft.com/v1.0/users/{user['id']}"
        res = requests.delete(
            endpoint,
            headers={'Authorization': 'Bearer ' + token['access_token']})
        data = res.json()

        if res.ok:
            print(f"ゲストユーザー [{user['mail']}] の削除に成功しました。")
            if is_physical_delete:
                delete_guests_from_deleteditem(token, user)
        else:
            print(data)

def delete_guests_from_deleteditem(token, user):
    endpoint = f"https://graph.microsoft.com/v1.0/directory/deletedItems/{user['id']}"
    res = requests.delete(
        endpoint,
        headers={'Authorization': 'Bearer ' + token['access_token']})
    data = res.json()

    if res.ok:
        print(f"ゲストユーザー [{user['mail']}] を削除済みユーザーから削除しました。")
    else:
        print(data)

def main():
    cred = connect_aad()
    token = get_access_token(cred)

    if "access_token" in token:
        guests = get_disabled_guests(token)
        get_disabled_guests(token, guests)
    else:
        print(token.get("error"))
        print(token.get("error_description"))
        print(token.get("correlation_id"))

if __name__ == '__main__':
    main()