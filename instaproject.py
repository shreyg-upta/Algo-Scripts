import pandas as pd
import requests
from datetime import datetime, timedelta

# Load the Excel file and the 'A' column which contains the handles
file_path = "database.xlsx"
df = pd.read_excel(file_path)

# Get the timestamp for one week ago
one_week_ago = int((datetime.now() - timedelta(days=7)).timestamp())

# Function to get the sum of ratings for successful submissions in the past week
def get_sum_of_ratings_last_week(handle):
    url = f"https://codeforces.com/api/user.status?handle={handle}&from=1"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        if data["status"] == "OK":
            submissions = data["result"]
            total_rating = 0
            problem_ratings = []

            print(f"Handle: {handle}")

            for submission in submissions:
                # Check if submission is within the last week and if verdict is 'OK'
                submission_time = submission["creationTimeSeconds"]
                verdict = submission["verdict"]
                if submission_time >= one_week_ago and verdict == "OK":
                    # Get the problem rating, use 800 if not available
                    problem = submission["problem"]
                    rating = problem.get("rating", 800)
                    problem_ratings.append(rating)
                    total_rating += rating
            
            # Output all problem ratings solved by this handle
            print(f"Problem ratings solved: {problem_ratings}")
            print(f"Total rating sum for the last week: {total_rating}\n")
            return total_rating
        else:
            print(f"Handle: {handle} - No submissions")
            return "No submissions"
    else:
        print(f"Handle: {handle} - Error fetching data")
        return "Error"

# Apply the function to all handles in the dataframe and add the results to column 'B'
df['TotalRatingLastWeek'] = df['Handle'].apply(get_sum_of_ratings_last_week)

# Save the updated dataframe back to Excel
df.to_excel(file_path, index=False)

print("Updated Excel file saved successfully!")
