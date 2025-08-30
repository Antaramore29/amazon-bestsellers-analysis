import pandas as pd   # Step 1: import pandas

# Step 2: load the CSV into a DataFrame
df = pd.read_csv('bestsellers.csv')

# Step 3: explore the data
#print(df.head())      # first 5 rows
#print(df.shape)       # shape of the spreadsheet- number of rows, columns
#print(df.columns)     # list of column names
#print(df.describe())  # summary stats for numeric data/ for each column

# Quick peek at the data
#print("\n=== HEAD (first 5 rows) ===")
#print(df.head())

#print("\n=== SHAPE (rows, cols) ===")
#print(df.shape)

#print("\n=== COLUMNS ===")
#print(df.columns.tolist())

#print("\n=== DATA TYPES ===")
#print(df.dtypes)

#print("\n=== NULLS PER COLUMN ===")
#print(df.isna().sum())

#print("\n=== BASIC STATS (numeric cols) ===")
#print(df.describe())

# A couple of sanity checks
#print("\n=== UNIQUE GENRES ===")
#print(df["Genre"].unique())

#print("\n=== SAMPLE ROW (iloc) ===")
#print(df.iloc[0])

#python3 main.py

# STEP 4 - clean data
df.drop_duplicates(inplace=True) #removed duplicates
#renamed the columns
df.rename(
    columns={
        "Name": "Title",
        "Year": "Publication Year",
        "User Rating": "Rating"
    },
    inplace=True
)
df["Price"] = df["Price"].astype(float)

# check results
print("\n=== CLEANED DATA SAMPLE ===")
print(df.head())

print("\n=== COLUMNS AFTER RENAME ===")
print(df.columns)

# STEP 5 - ANALYSIS ===

# 1) author popularity
print("\n=== TOP AUTHORS BY BOOK COUNT ===")
author_counts = df["Author"].value_counts()
print(author_counts.head(10))

# 2) average rating by genre
print("\n=== AVERAGE RATING BY GENRE ===")
avg_rating_by_genre = df.groupby("Genre")["Rating"].mean()
print(avg_rating_by_genre)

# 3) most reviewed books
print("\n=== MOST REVIEWED BOOKS ===")
print(df[["Publication Year", "Title", "Author", "Reviews"]]
      .sort_values(by="Reviews", ascending=False)
      .head(10))

# 4)most expensive books
print("\n=== MOST EXPENSIVE BOOKS ===")
print(df[["Publication Year", "Title", "Author", "Price"]]
      .sort_values(by="Price", ascending=False)
      .head(10))

# 5) average price by genre
print("\n=== AVERAGE PRICE BY GENRE ===")
avg_price_by_genre = df.groupby("Genre")["Price"].mean()
print(avg_price_by_genre)

# STEP 6 — EXPORT ALL RESULTS TO ONE EXCEL FILE
# Requires: pip3 install openpyxl

with pd.ExcelWriter("analysis_results.xlsx", engine="openpyxl") as writer:
    # 1) Top authors
    author_counts.head(10).to_frame("Count").to_excel(writer, sheet_name="Top Authors")

    # 2) Average rating by genre
    avg_rating_by_genre.to_frame("Avg Rating").to_excel(writer, sheet_name="Avg Rating by Genre")

    # 3) Most reviewed books (same code as print, just with .to_excel at the end)
    (
        df[["Publication Year", "Title", "Author", "Reviews"]]
          .sort_values(by="Reviews", ascending=False)
          .head(10)
          .to_excel(writer, sheet_name="Most Reviewed Books", index=False)
    )

    # 4) Most expensive books (same code as print)
    (
        df[["Publication Year", "Title", "Author", "Price"]]
          .sort_values(by="Price", ascending=False)
          .head(10)
          .to_excel(writer, sheet_name="Most Expensive Books", index=False)
    )

    # 5) Average price by genre
    avg_price_by_genre.to_frame("Avg Price").to_excel(writer, sheet_name="Avg Price by Genre")
