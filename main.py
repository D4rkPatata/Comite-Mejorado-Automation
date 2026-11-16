import os
import sys

from config import INT_SEMANA, SCHEMA_DATA_BASE
from extract import extract_data
from load import update
from transform import group_data


# DATA.xlsx's Power Query reads this file.
DATA_CSV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ouput", "DATA.csv")


def main(semana=INT_SEMANA):
    df = extract_data()
    df = group_data(df, SCHEMA_DATA_BASE)
    df.to_csv(DATA_CSV_PATH, index=False, encoding="utf-8-sig")
    print(f"DATA.csv: {len(df):,} rows -> {DATA_CSV_PATH}")

    update(semana)


if __name__ == "__main__":
    # python main.py [semana]
    main(int(sys.argv[1]) if len(sys.argv) > 1 else INT_SEMANA)
