import os
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px

from config import CHARTS_DIR, INTERACTIVE_DIR

def ensure_folders():
    os.makedirs(CHARTS_DIR, exist_ok=True)
    os.makedirs(INTERACTIVE_DIR, exist_ok=True)

def save_chart_static(data, kind, title, filename):
    plt.figure(figsize=(10, 6))
    sns.set(style="whitegrid")

    if kind == "bar":
        data.plot(kind="bar", color="skyblue")
    elif kind == "barh":
        data.plot(kind="barh", color="skyblue")
    elif kind == "line":
        data.plot(kind="line", marker="o")

    plt.title(title)
    plt.tight_layout()

    path = os.path.join(CHARTS_DIR, f"{filename}.png")
    plt.savefig(path)
    plt.close()
    return path

def save_chart_interactive(data, kind, title, filename):
    path = os.path.join(INTERACTIVE_DIR, f"{filename}.html")

    if kind == "bar":
        fig = px.bar(x=data.index, y=data.values, labels={'x': title, 'y': 'Liczba'}, title=title)
    elif kind == "barh":
        fig = px.bar(x=data.values, y=data.index, orientation='h', labels={'x': 'Liczba', 'y': title}, title=title)
    elif kind == "line":
        fig = px.line(x=data.index, y=data.values, labels={'x': title, 'y': 'Liczba'}, title=title)

    fig.write_html(path)
    return path

def generate_charts(stats):
    types, countries, years, genres, directors = stats

    static = {
        "typy": save_chart_static(types, "bar", "Typy produkcji", "typy_produkcji"),
        "kraje": save_chart_static(countries, "bar", "Top kraje", "kraje"),
        "rocznik": save_chart_static(years, "line", "Produkcje wg lat", "rocznik"),
        "gatunki": save_chart_static(genres, "barh", "Top gatunki", "gatunki"),
        "rezyserzy": save_chart_static(directors, "bar", "Top reżyserzy", "rezyserzy"),
    }

    # Interaktywne
    save_chart_interactive(types, "bar", "Typy produkcji", "typy_produkcji")
    save_chart_interactive(countries, "bar", "Top kraje", "kraje")
    save_chart_interactive(years, "line", "Produkcje wg lat", "rocznik")
    save_chart_interactive(genres, "barh", "Top gatunki", "gatunki")
    save_chart_interactive(directors, "bar", "Top reżyserzy", "rezyserzy")

    return static
