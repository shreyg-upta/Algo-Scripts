import pandas as pd
import requests
from datetime import datetime, timedelta

# Load the Excel file
file_path = "crewdb.xlsx"
df = pd.read_excel(file_path)

# Print column names for debugging
print("Columns in the Excel file:", df.columns)

# Column names for handles (make sure these match your Excel file)
handle_columns = ['H1', 'H2', 'H3', 'H4', 'H5']

# Read contest IDs from the first row of the Excel file (first 5 columns)
contests =[2019, 2020, 2021, 2022, 2025]

# Get the timestamp for one month ago
one_month_ago = int((datetime.now() - timedelta(days=30)).timestamp())

# Function to get user info (current rating, peak rating, etc.)
def get_user_info(handle):
    url = f"https://codeforces.com/api/user.info?handles={handle}"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        if data["status"] == "OK" and len(data["result"]) > 0:
            user_info = data["result"][0]
            current_rating = user_info.get("rating", 0)
            peak_rating = user_info.get("maxRating", 0)
            return current_rating, peak_rating
    return 0, 0

# Function to get the number of submissions with verdict "OK" in the last month
def get_ok_submissions_last_month(handle):
    url = f"https://codeforces.com/api/user.status?handle={handle}&from=1"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        if data["status"] == "OK":
            submissions = data["result"]
            ok_count = 0
            for submission in submissions:
                submission_time = submission["creationTimeSeconds"]
                verdict = submission.get("verdict", "")
                if submission_time >= one_month_ago: 
                    if verdict == "OK":
                        ok_count += 1
                else:
                    break
            return ok_count
    return 0

# Function to get number of contests participated in the last 5 contests
def get_contests_participated(handle, contests):
    url = f"https://codeforces.com/api/user.rating?handle={handle}"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        if data["status"] == "OK":
            contest_ids = [rating_change["contestId"] for rating_change in data["result"]]
            count = sum(1 for contest_id in contests if contest_id in contest_ids)
            return count
    return 0

# Process each row to calculate the desired values
for i, row in df.iterrows():
    # Get handles from H1, H2, H3, H4, H5 columns
    handles = [row[h] for h in handle_columns if pd.notna(row[h])]
    
    # Initialize variables
    peak_rating = 0
    total_ok_submissions = 0
    contests_participated = 0
    
    # Loop through each handle to find peak rating, OK submissions, and contests participated
    for handle in handles:
        # Get current and peak rating for each handle
        _, peak = get_user_info(handle)
        peak_rating = max(peak_rating, peak)  # Keep track of the highest peak rating
        
        # Add OK submissions in the last month
        total_ok_submissions += get_ok_submissions_last_month(handle)
        
        # Add contest participation count for last 5 contests
        contests_participated += get_contests_participated(handle, contests)

    # Write the values back into the DataFrame
    df.at[i, 'Peak'] = peak_rating
    df.at[i, 'LastMonth'] = total_ok_submissions
    df.at[i, 'Last5Contests'] = contests_participated

# Save the updated DataFrame back to Excel
df.to_excel(file_path, index=False)

print("Updated Excel file saved successfully!")
