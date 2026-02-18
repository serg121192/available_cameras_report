import os
import json
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine, Engine, text
from sqlalchemy.engine import URL


load_dotenv()

BASE_DIR = Path(__file__).resolve().parent.parent
CONFIG_DIR = BASE_DIR / "config"


def create_connection_engine() -> Engine:
    db_url = URL.create(
        drivername="postgresql+psycopg",
        username=os.getenv("POSTGRES_USER"),
        password=os.getenv("POSTGRES_PASSWORD"),
        host=os.getenv("POSTGRES_HOST"),
        port=os.getenv("POSTGRES_PORT"),
        database=os.getenv("POSTGRES_DB_NAME"),
    )

    return create_engine(db_url)


def camera_info_grabber() -> pd.DataFrame:
    map_files = CONFIG_DIR.glob("*_map.json")
    engine = create_connection_engine()

    complete_report = []

    with engine.connect() as connection:
        for file in map_files:
            with open(file, "r", encoding="utf-8") as f:
                dep_map = json.load(f)

            request = text(
                (CONFIG_DIR / "camera_info_request.txt").read_text(
                    encoding="utf-8"
                )
            )

            departments_list = dep_map.keys()

            district_info = pd.read_sql(
                request,
                connection,
                params={
                    "departments": [
                        int(department) for department in departments_list
                    ]
                },
            )

            district_info["department"] = (
                district_info["department"].astype(str).map(dep_map)
            )

            district_info["district"] = file.stem.replace("_map", "")
            complete_report.append(district_info)

    complete_report = pd.concat(complete_report, ignore_index=True)

    return complete_report
