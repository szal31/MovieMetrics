# MovieMetrics: Box Office Performance and Profitability

DSC651 Data Visualization group project. MovieMetrics is a Power BI dashboard that shows how movies perform at the box office and how much profit they make, using 3,907 films released between 1929 and 2016.

## What's here

| File | What it is |
|---|---|
| `MovieMetrics.pbix` | The Power BI report. Open it in Power BI Desktop; the cleaned data is already inside, so no refresh is needed. |
| `Chapter 3 Methodology (MovieMetrics) - FULL.docx` | Chapter 3 of the report (methodology, data preparation, EDA, features). |
| `movies_data.csv` | The raw dataset (3,974 rows, 16 columns). |
| `figures/` | The Chapter 3 figures exported from the report. |
| `MovieMetrics.pbip`, `MovieMetrics.Report/`, `MovieMetrics.SemanticModel/` | The same report as a Power BI project (text files, easier to compare between versions). |
| `build_pbip.py` | Python script that generates the Power BI project files. |

## Report pages

1. **About Us**: what the dashboard is for and how to use it
2. **Main Overview**: KPI cards, box office and median ROI by genre, trends by decade
3. **Financial and Award Analysis**: loss rate by budget level, budget against box office, IMDb and award comparisons
4. **Talent Analysis**: top 10 directors and lead actors by box office
5. **Film Details**: a table of every film
6. **EDA (Chapter 3)**: the descriptive charts used in Chapter 3 (not part of the dashboard navigation)

In Power BI Desktop, hold **Ctrl** and click the navigation buttons to move between pages.

## Opening the Power BI project version

The `.pbip` version loads the CSV from `C:\DSC651\MovieMetrics\movies_data.csv`. To refresh it on another computer, open **Transform data**, select the `movies_data` query and change the file path in the `Source` step to where you saved the CSV. The `.pbix` file does not need this.

## Dataset

Oliva, D. (2024). *Movies (IMDb, Earnings and more)* [Data set]. Kaggle. https://www.kaggle.com/datasets/delfinaoliva/movies

The data was compiled by the uploader from Wikipedia, IMDb and The Movie Database (TMDB). All credit for the data goes to the original author.
