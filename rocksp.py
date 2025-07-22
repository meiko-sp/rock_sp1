from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
import pyodbc

app = FastAPI()

# CORS設定（必要に応じて制限）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # GASからのアクセスを許可
    allow_methods=["*"],
    allow_headers=["*"],
)

# Accessに書き込む関数
def write_to_access(val):
    conn = pyodbc.connect(
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=\\Meikoh-svn\f全部署共有\0業務システム共有\製造\線ばね1班\ロックSPラベルmdb;'
    )
    cursor = conn.cursor()
    cursor.execute("INSERT INTO テーブル1 (フィールド1) VALUES (?)", val)
    conn.commit()
    conn.close()

# POSTエンドポイント
@app.post("/insert")
async def insert_data(request: Request):
    data = await request.json()
    print("Received JSON:", data)  # デバッグ用ログ
    write_to_access(data["val"])
    return {"status": "success"}


# uvicorn rocksp:app --host 192.168.1.23 --port 8000