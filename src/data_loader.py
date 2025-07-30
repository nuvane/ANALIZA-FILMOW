import pandas as pd

def load_data(filepath):
    df = pd.read_csv(filepath)

    df['date_added'] = pd.to_datetime(df['date_added'], errors='coerce').dt.date

    df['country'] = df['country'].fillna('Unknown')
    df['rating'] = df['rating'].fillna('Unknown')

    dt_temp = pd.to_datetime(df['date_added'], errors='coerce')
    df['year_added'] = dt_temp.dt.year
    df['month_added'] = dt_temp.dt.month

    return df
