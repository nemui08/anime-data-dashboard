import requests
import pandas as pd
import time
import os

# =========================================================
# Anime Data Pipeline Project
# Source: Jikan API
# Tools: Python, Pandas, Excel
# Purpose: Prepare data for Excel Dashboard
# =========================================================

URL = "https://api.jikan.moe/v4/top/anime"
FILE_NAME = "anime_dashboard.xlsx"

anime_list = []

# ---------------- FETCH DATA ----------------
for page in range(1, 6):
    print(f"Page {page}...")

    try:
        response = requests.get(URL, params={"page": page})
        response.raise_for_status()
        data = response.json()
    except Exception as e:
        print("Error:", e)
        continue

    for anime in data.get("data", []):

        # -------- BASIC INFO --------
        title = anime.get("title_english") or anime.get("title")
        score = anime.get("score")
        episodes = anime.get("episodes")
        popularity = anime.get("popularity")
        status = anime.get("status")
        rank = anime.get("rank")

        # -------- YEAR --------
        # Use direct year if available, otherwise fallback to aried.from.year
        year = anime.get("year") or anime.get("aired", {}).get("prop", {}).get("from", {}).get("year")

        # -------- DATE CLEANING --------
        # Extract start and end dates and convert to YYYY-MM-DD format
        start_date = anime.get("aired", {}).get("from")
        end_date = anime.get("aired", {}).get("to")

        start_date = start_date.split("T")[0] if start_date else None
        end_date = end_date.split("T")[0] if end_date else None

        # -------- GENRES (LIST FORMAT)--------
        # Store genres as list to enable explode later
        genres = [g.get("name") for g in anime.get("genres", [])]

        # Append cleaned data into list
        anime_list.append({
            "title": title,
            "score": score,
            "episodes": episodes,
            "rank": rank,
            "popularity": popularity,
            "year": year,
            "status": status,
            "start_date": start_date,
            "end_date": end_date,
            "genres": genres
        })

    # Prevent hitting API rate limit
    time.sleep(1)


# ---------------- CREATE DATAFRAME ----------------
df = pd.DataFrame(anime_list)

# -------- DATA CLEANING --------
# Convert rank to numeric and remove invalid values
df["rank"] = pd.to_numeric(df["rank"], errors="coerce")
df = df[df["rank"].notna()]

# -------- EXPLODE GENRES --------
# Convert list of genres into multiple rows (1 genres per row)
df = df.explode("genres")

# -------- DATE CONVERSION --------
df["start_date"] = pd.to_datetime(df["start_date"], errors="coerce")
df["end_date"] = pd.to_datetime(df["end_date"], errors="coerce")

# -------- DURATION CALCULATION --------
# Calculate duration in days between start and end date
df["duration_days"] = (df["end_date"] - df["start_date"]).dt.days

# Set duration to None if anime is not finished or missing end date
df.loc[
    (df["status"] != "Finished Airing") | (df["end_date"].isna()),
    "duration_days"
] = None

# -------- FINAL DATE FORMAT --------
# Convert datetime to date format for Excel readability
df["start_date"] = df["start_date"].dt.date
df["end_date"] = df["end_date"].dt.date

# Debug output
print("\nTotal rows:", len(df))
print(df.head())


# ---------------- TOP ANALYSIS ----------------
# Filter only valid duration data
df_valid = df[df["duration_days"].notna()]

# Top 10 by rank (best ranked anime)
top_rank = (
    df_valid.sort_values("rank")
    .drop_duplicates("title") # Prevent duplicates due to exploded genres 
    .head(10)
)

# Top 10 by duration (longest anime)
top_duration = (
    df_valid.sort_values("duration_days", ascending=False)
    .drop_duplicates("title")
    .head(10)
)


# ---------------- EXPORT TO EXCEL ----------------
# Check if file already exists
if os.path.exists(FILE_NAME):
    print("\nFile exists → Updating raw_data...")
    mode = "a"
    if_sheet = "replace"
else:
    print("\nFile not found → Creating new file...")
    mode = "w"
    if_sheet = None

# Use ExcelWriter to safely write data
with pd.ExcelWriter(
    FILE_NAME,
    engine="openpyxl",
    mode=mode,
    if_sheet_exists=if_sheet
) as writer:

    # Main dataset (used for Pivot & Dashboard)
    df.to_excel(writer, sheet_name="raw_data", index=False)

    # Optional sheets for analysis/debugging
    top_rank.to_excel(writer, sheet_name="top_rank", index=False)
    top_duration.to_excel(writer, sheet_name="top_duration", index=False)

print("\nExport completed successfully.")