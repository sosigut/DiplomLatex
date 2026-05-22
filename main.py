import asyncio
import sys

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

from app.db.base import Base
from app.db.session import engine
from app.routers import admin, auth, checker, statistics, title_page, rio

if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
app = FastAPI(title="Diplom Checker API")

Base.metadata.create_all(bind=engine)

app.mount("/static", StaticFiles(directory="app/static"), name="static")

app.include_router(auth.router)
app.include_router(checker.router)
app.include_router(admin.router)
app.include_router(statistics.router)
app.include_router(title_page.router)
app.include_router(rio.router)


@app.get("/")
def root():
    return FileResponse("app/templates/index.html")