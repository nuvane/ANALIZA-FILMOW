from collections import Counter
import pandas as pd

def get_basic_stats(df):
    types = df['type'].value_counts()
    countries = df['country'].value_counts().head(10)
    years = df['release_year'].value_counts().sort_index()

    genres = Counter()
    for entry in df['listed_in'].dropna():
        genres.update([g.strip() for g in entry.split(',')])
    top_genres = pd.Series(genres).sort_values(ascending=False).head(10)

    top_directors = df['director'].value_counts().dropna().head(10)

    return types, countries, years, top_genres, top_directors
