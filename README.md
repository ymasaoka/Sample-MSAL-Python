# Sample-MSAL-Python
本リポジトリは、MSAL for Python のサンプルコードを掲載しています。  
本リポジトリにある各種コードは、Docker Compose を使用して簡単に実行できるようになっています。  

# 実行環境準備

## Docker コンテナーの起動
本リポジトリをクローンし、ルートディレクトリに `web-variables.env` ファイルを作成します。  
`web-variables.env` ファイルには、以下の情報を記載し、起動する Python コンテナーの環境変数として利用できるようにします。  

```:web-variables.env
CLIENT_ID={サービスプリンシパル ID}
CLIENT_SECRET={サービスプリンシパルのクライアントシークレット値}
TENANT_ID={テナント ID}
THUMBPRINT={サービスプリンシパル認証時に SSL 証明書を使用する場合の thumbprint (拇印) 値}
```

`web-variables.env` ファイルを用意したら、Docker コンテナーを起動します。  

```bash
docker-compose up -d --build
```

## 自己署名証明書の作成

サービスプリンシパルを利用する際の認証方式には、クライアントシークレットではなく証明書を利用することも可能です。  
秘密鍵ありの SSL 証明書を サービスプリンシパルに登録し、MSAL for Python での接続時には、thumbprint と秘密鍵ファイルを使用します。

証明書に、自己署名証明書を使用する場合は、以下を参考に自己署名証明書を作成してください。

:::note warn
この例では、openssl を使用して作成しています。  
本リポジトリに記載のコードは、下記の手順で払い出した自己署名証明書を使用して動作確認をしたものになります。
:::

```bash
openssl genrsa 2024 > ca-key.pem
openssl req -new -key ca-key.pem > server.csr
openssl x509 -req -days {証明書が有効期限切れになるまでの日数} -signkey ca-key.pem < server.csr > server.crt
```

`ca-key.pem` が秘密鍵ファイル、`server.crt` が自己署名証明書 (SSL 証明書) です。  

## サンプルコードの実行

`sample` フォルダにある Python コードは、以下の形で実行が可能です。  

```bash
docker-compose run --rm app python3 {実行したい Python ファイル}
```
