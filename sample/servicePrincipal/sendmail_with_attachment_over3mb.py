# 必要な API アクセス許可
# Microsoft Graph/アプリケーションの許可/Mail.ReadWrite
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

email_msg_draft = {
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
}

authority = f"https://login.microsoftonline.com/{os.environ['TENANT_ID']}"
scopes = ['https://graph.microsoft.com/.default']
secretkey_path = f""

attachment_file_path = "" # 添付ファイルのフルパス
attachment_file_isBinary = True # バイナリファイルか否か
bytes_send_size = 327680 # アップロードセッションで 1 PUT あたりに送信するバイト容量 (最大バイト数は 60MiB、分割する際は 320 KiB の倍数を指定)

def connect_aad():
    # クライアントシークレットを使用するパターン
    # cred = msal.ConfidentialClientApplication(
    #     client_id=os.environ['CLIENT_ID'],
    #     client_credential=os.environ['CLIENT_SECRET'],
    #     authority=authority)

    # SSL 証明書を使用するパターン
    # thumbprint (拇印) の値は、サービスプリンシパルに SSL 証明書を登録した際に表示される値を使用します
    cred = msal.ConfidentialClientApplication(
        client_id=os.environ['CLIENT_ID'],
        client_credential={"thumbprint": os.environ['THUMBPRINT'],"private_key": open(secretkey_path).read()},
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

def create_msg_draft(token):
    endpoint = f"https://graph.microsoft.com/v1.0/users/{config['user_from']}/messages"
    res = requests.post(
        endpoint,
        headers={'Authorization': 'Bearer ' + token['access_token']},
        json=email_msg_draft)

    if res.ok:
        print(f"メール下書き作成に成功しました。Message ID: {res.json()['id']}")
        return res.json()['id']

    else:
        print(res.json())

def get_file_info(full_filepath):
    return {
        'AttachmentItem': {
            'attachmentType': 'file',
            'name': os.path.basename(full_filepath),
            'size': os.path.getsize(full_filepath)
        }
    }

def upload_attachment(token, mid, file_info):
    endpoint = f"https://graph.microsoft.com/v1.0/users/{config['user_from']}/messages/{mid}/attachments/createUploadSession"
    res_session = requests.post(
        endpoint,
        headers={'Authorization': 'Bearer ' + token['access_token']},
        json=file_info)

    if res_session.ok:
        print("アップロードセッションの作成に成功しました。")
        # res_session = {
        #     '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#microsoft.graph.uploadSession', 
        #     'uploadUrl': "https://outlook.office.com/api/gv1.0/users('objectID')/messages('messageID')/AttachmentSessions('sessionID')?authtoken={tokenValue}", 
        #     'expirationDateTime': 'yyyy-mm-ddTHH:mm:ss.ffffffZ', 
        #     'nextExpectedRanges': ['0-']
        # }
    else:
        print(res_session.json())

    print("ファイルデータのアップロードを開始します。")
    file_size = os.path.getsize(attachment_file_path)
    print(f"アップロードするファイルのデータ容量: {file_size} bytes")

    next_bytes_start_point = 0
    next_bytes_end_point = (bytes_send_size - 1)
    next_bytes_range = f"{res_session.json()['nextExpectedRanges'][0]}{next_bytes_end_point}"

    send_data = None
    if attachment_file_isBinary:
        send_data = open(attachment_file_path,'rb').read()
    else:
        send_data = open(attachment_file_path,'r').read()

    while True:
        send_data_part = send_data[next_bytes_start_point:next_bytes_end_point + 1]

        if not len(send_data_part):
            break
        elif next_bytes_end_point + 1 == file_size:
            bytes_send_size = (file_size - next_bytes_start_point)

        res_upload = requests.put(
            res_session.json()['uploadUrl'],
            data=send_data_part,
            headers={'Content-Length': f"{bytes_send_size}",'Content-Range': f"bytes {next_bytes_range}/{file_size}"})

        if res_upload.status_code == 200:
            # res_upload = {
            #     '@odata.context': "https://outlook.office.com/api/gv1.0/$metadata#users('objectID')/messages('messageID')/attachmentSessions/$entity", 
            #     'expirationDateTime': 'yyyy-mm-ddTHH:mm:ss.ffffffZ', 
            #     'nextExpectedRanges': ['327680']
            # }
            print(f"uploading next {bytes_send_size} bytes, bytes_range={next_bytes_range}/{file_size - 1}: success")
            next_bytes_start_point = int(res_upload.json()['nextExpectedRanges'][0])
            next_bytes_end_point = (next_bytes_start_point + bytes_send_size)
            if next_bytes_end_point > file_size:
                next_bytes_end_point = file_size - 1
            else:
                next_bytes_end_point -= 1
            next_bytes_range = f"{next_bytes_start_point}-{next_bytes_end_point}"
        elif res_upload.status_code == 201:
            print(f"uploading next {bytes_send_size} bytes, bytes_range={next_bytes_range}/{file_size - 1}: success")
            requests.delete(res_session.json()['uploadUrl'])
            print("ファイルのアップロードセッションを終了しました。")
            print(f"ファイルのアップロードが完了しました。Message ID: {mid}")
            break
        else:
            print(f"uploading next {bytes_send_size} bytes, bytes_range={next_bytes_range}/{file_size}: failed")
            print(res_upload.json())
            break

def send_mail(token, mid):
    print(f"メール送信処理を開始します。Message ID: {mid}")
    endpoint = f"https://graph.microsoft.com/v1.0/users/{config['user_from']}/messages/{mid}/send"
    res = requests.post(
        endpoint,
        headers={'Authorization': 'Bearer ' + token['access_token']})

    if res.ok:
        print("メール送信に成功しました。")
    else:
        print(res.json())

def error(token):
    print(token.get("error"))
    print(token.get("error_description"))
    print(token.get("correlation_id"))

def main():
    cred = connect_aad()
    token = get_access_token(cred)

    mid = create_msg_draft(token) if "access_token" in token else error(token)
    file_info = get_file_info(attachment_file_path)
    upload_attachment(token, mid, file_info)
    send_mail(token, mid)

if __name__ == '__main__':
    main()