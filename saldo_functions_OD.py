import os
import pandas as pd

"""
Saldo-Berechnung pro Station/Trip für die On-Demand-Daten.

Ablauf:
    1. load_data           -> Excel einlesen
    2. clean_station_names -> falsche/verunstaltete Stationsnamen korrigieren
    3. add_sequence        -> Stationsreihenfolge (SQ) pro Trip nummerieren
    4. compute_saldo       -> laufenden Saldo pro Trip berechnen
    5. build_final_table   -> Spalten auswählen/umbenennen
    6. compute_pkm         -> Personenkilometer berechnen
    7. export_data         -> als Excel speichern
"""

INPUT_PATH = "input/ondemand_07_2026.xlsx"
OUTPUT_PATH = "data/final_07_V2.xlsx"

# Mapping falsche Stationennamen
STATION_CORRECTIONS = {
    " HÃ¶lstein Station": "Hölstein Station",
    "HÃ¶lstein Station": "Hölstein Station",
    "Hoelstein Station": "Hölstein Station",
    "Lampenberg BÃ¼ndtenweg": "Lampenberg Bündtenweg",
    "MÃ¼lizelg": "Mülizelg",
}

FINAL_COLUMNS = [
    "date_trip",
    "vehicle_id",
    "line_number",
    "starts_at",
    "trip_no",
    "station_id",
    "station_name",
    "SQ",
    "distance",
    "time",
    "boarding",
    "alighting",
    "saldo_station",
]

RENAME_MAP = {
    "date_trip": "Datum",
    "vehicle_id": "Fahrzeug_id",
    "line_number": "LineNo",
    "distance": "Distanz",
    "starts_at": "Trip_Aktivierungszeit",
    "trip_no": "tripID",
    "station_id": "StopNo",
    "station_name": "Station",
    "time": "IST_Abfahrtzeit",
    "boarding": "Einsteiger",
    "alighting": "Aussteiger",
    "saldo_station": "Saldo_Fahrt",
}


def load_data(path: str = INPUT_PATH) -> pd.DataFrame:
    """Excel-Rohdaten einlesen."""
    return pd.read_excel(path)


def clean_station_names(df: pd.DataFrame, col: str = "station_name") -> pd.DataFrame:
    """Falsch kodierte Stationsnamen korrigieren, trimmen und Bindestriche normalisieren."""
    df = df.copy()
    df[col] = df[col].replace(STATION_CORRECTIONS)
    df[col] = df[col].str.replace(r"^\s+|\s+$", "", regex=True)
    df[col] = df[col].str.replace(r"\s*-\s*", "-", regex=True)
    return df


def add_sequence(df: pd.DataFrame) -> pd.DataFrame:
    """Stationsreihenfolge (SQ) pro Trip hinzufügen."""
    df = df.copy()
    df["SQ"] = df.groupby("trip_no").cumcount() + 1
    return df


def compute_saldo(df: pd.DataFrame) -> pd.DataFrame:
    """Laufenden Saldo (Ein- minus Aussteiger) pro Trip berechnen."""
    df = df.sort_values(["trip_no", "SQ"]).copy()
    df["saldo_station"] = (
        df.groupby("trip_no")["boarding"].cumsum()
        - df.groupby("trip_no")["alighting"].cumsum()
    )
    return df


def build_final_table(df: pd.DataFrame) -> pd.DataFrame:
    """Relevante Spalten auswählen und ins Zielschema umbenennen."""
    return df[FINAL_COLUMNS].copy().rename(columns=RENAME_MAP)


def compute_pkm(df_final: pd.DataFrame) -> pd.DataFrame:
    """
    Personenkilometer (PKM) je Zeile berechnen:
    Saldo der VORHERIGEN Station (geshiftet) * Distanz der aktuellen Station / 1000.
    """
    df_final = df_final.copy()
    df_final["PKM"] = (
        df_final.groupby("tripID")["Saldo_Fahrt"].shift(1)
        * df_final["Distanz"]
        / 1000
    )
    return df_final


def export_data(df: pd.DataFrame, path: str = OUTPUT_PATH) -> None:
    """Ergebnis als Excel-Datei speichern."""
    os.makedirs(os.path.dirname(str(path)) or ".", exist_ok=True)
    df.to_excel(path, index=False)