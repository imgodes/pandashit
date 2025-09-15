#!/usr/bin/env python3

"""
pandashit.py
Converte um arquivo .json (diversos formatos: array, jsonlines, múltiplos objetos)
para .xlsx usando pandas, tentando detectar o formato automaticamente.
Uso: python pandashit.py entrada.json saida.xlsx
"""

import argparse
import json
from pathlib import Path

import pandas as pd
from pandas import json_normalize
from pandas.api.types import is_datetime64_any_dtype, DatetimeTZDtype

def read_json_smart(path: Path):
    try:
        df = pd.read_json(path)
        return df
    except ValueError:
        pass

    try:
        df = pd.read_json(path, lines=True)
        return df
    except ValueError:
        pass

    objs = []
    with path.open("r", encoding="utf-8") as f:
        text = f.read().strip()
        try:
            data = json.loads(text)
            if isinstance(data, list):
                return json_normalize(data)
            if isinstance(data, dict):
                for v in data.values():
                    if isinstance(v, list):
                        return json_normalize(v)
                return json_normalize([data])
        except json.JSONDecodeError:
            pass

    objs = []
    with path.open("r", encoding="utf-8") as f:
        for i, line in enumerate(f, start=1):
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
                objs.append(obj)
            except json.JSONDecodeError:
                continue

    if objs:
        return json_normalize(objs)

    raise ValueError("Não foi possível interpretar o JSON. Formatos suportados: array JSON, JSON Lines (NDJSON), ou múltiplos objetos por linha.")

def remove_timezones_from_df(df: pd.DataFrame) -> pd.DataFrame:
    """Converte séries tz-aware para tz-naive (UTC) e tenta detectar strings de datetime."""
    df = df.copy()


    for col in df.columns:
        try:
            if isinstance(df[col].dtype, DatetimeTZDtype):
                df[col] = df[col].dt.tz_convert("UTC").dt.tz_localize(None)
        except Exception:
            pass

    for col in df.select_dtypes(include=["object"]).columns:
        series = df[col]
        sample = series.dropna().astype(str)
        if sample.empty:
            continue

        hint = ("T" in sample.iloc[0]) or ("+" in sample.iloc[0]) or ("-" in sample.iloc[0] and ":" in sample.iloc[0])
        if not hint:
            hint = any(("T" in s) or ("+" in s) or (":" in s and "-" in s) for s in sample.iloc[:10].astype(str))

        if not hint:
            continue


        try:
            parsed = pd.to_datetime(series, utc=True, errors="coerce")
            if parsed.notna().sum() > 0:
                parsed = parsed.dt.tz_convert("UTC").dt.tz_localize(None)
                df[col] = parsed
        except Exception:
            continue

    return df


def json_to_xlsx(json_file: str, xlsx_file: str):
    path = Path(json_file)
    if not path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {json_file}")

    df = read_json_smart(path)


    if not isinstance(df, pd.DataFrame):
        df = json_normalize(df)


    df = remove_timezones_from_df(df)


    df.to_excel(xlsx_file, index=False, engine="openpyxl")
    return xlsx_file

def main():
    parser = argparse.ArgumentParser(description="Converte JSON (variados formatos) para XLSX usando pandas.")
    parser.add_argument("input", help="caminho para arquivo .json")
    parser.add_argument("output", nargs="?", help="arquivo de saída .xlsx (default: saida.xlsx)", default="saida.xlsx")
    args = parser.parse_args()

    try:
        out = json_to_xlsx(args.input, args.output)
        print(f"OK — arquivo salvo em: {out}")
    except Exception as e:
        print("Erro:", e)
        raise SystemExit(1)

if __name__ == "__main__":
    main()

