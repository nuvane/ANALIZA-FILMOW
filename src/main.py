from config import INPUT_FILE
from data_loader import load_data
from analysis import get_basic_stats
from charts import ensure_folders, generate_charts
from excel_writer import write_to_excel
from filtering import filter_and_analyze

def main():
    ensure_folders()
    df = load_data(INPUT_FILE)
    stats = get_basic_stats(df)
    chart_paths = generate_charts(stats)
    write_to_excel(df, chart_paths)


    filter_and_analyze(df, 'country', 'United States')
    filter_and_analyze(df, 'type', 'Movie')

if __name__ == "__main__":
    main()
