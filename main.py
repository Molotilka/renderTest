# main.py
from fastapi import FastAPI

app = FastAPI()  # <-- вот этот объект `app`

@app.get("/")
def hello():
    return {"message": "Hello World"}
