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
    map_files = list(CONFIG_DIR.glob("*_map.json"))
    if not map_files:
        return pd.DataFrame()

    # Read all map files and collect department ids to run a single DB query
    dep_maps = []
    all_departments = set()
    for file in map_files:
        with open(file, "r", encoding="utf-8") as f:
            dep_map = json.load(f)
        dep_maps.append((file, dep_map))
        for k in dep_map.keys():
            try:
                all_departments.add(int(k))
            except Exception:
                continue

    request = text(
        (CONFIG_DIR / "camera_info_request.txt").read_text(encoding="utf-8")
    )

    engine = create_connection_engine()
    complete_report = []

    if not all_departments:
        return pd.DataFrame()

    with engine.connect() as connection:
        # Single read for all departments across map files
        district_info_all = pd.read_sql(
            request, connection, params={"departments": list(all_departments)}
        )

        # Split results per map file and apply local mappings
        for file, dep_map in dep_maps:
            try:
                departments_set = {int(k) for k in dep_map.keys()}
            except Exception:
                departments_set = set()

            if not departments_set:
                continue

            district_info = district_info_all[
                district_info_all["department"].isin(departments_set)
            ].copy()

            if district_info.empty:
                continue

            district_info["department"] = (
                district_info["department"].astype(str).map(dep_map)
            )
            district_info["district"] = file.stem.replace("_map", "")
            complete_report.append(district_info)

    if not complete_report:
        return pd.DataFrame()

    complete_report = pd.concat(complete_report, ignore_index=True)

    return complete_report
