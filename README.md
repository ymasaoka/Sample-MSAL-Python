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
```

`web-variables.env` ファイルを用意したら、Docker コンテナーを起動します。  

```bash
docker-compose up -d --build
```

## サンプルコードの実行

`sample` フォルダにある Python コードは、以下の形で実行が可能です。  

```bash
docker-compose run --rm app python3 {実行したい Python ファイル}
```
