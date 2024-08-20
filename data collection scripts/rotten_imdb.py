import pandas as pd
import requests

# Function to get ratings from OMDb API
def get_ratings(title, api_key):
    base_url = 'http://www.omdbapi.com/'
    params = {
        't': title,
        'apikey': api_key
    }
    response = requests.get(base_url, params=params)
    data = response.json()
    
    if response.status_code == 200 and data['Response'] == 'True':
        imdb_rating = data.get('imdbRating', 'N/A')
        rt_rating = 'N/A'
        
        # Extract Rotten Tomatoes rating if available
        for rating in data.get('Ratings', []):
            if rating['Source'] == 'Rotten Tomatoes':
                rt_rating = rating['Value']
                break
        
        return imdb_rating, rt_rating
    else:
        return 'N/A', 'N/A'

# Your OMDb API key
api_key = '514b5453'

# Load the Excel file with the existing titles
input_file = 'moistmeter.xlsx'  # Replace with your actual file name
df = pd.read_excel(input_file)

# Initialize lists to store ratings
imdb_ratings = []
rt_ratings = []

# Iterate over each title and fetch the ratings
for title in df['Title']:  # Assuming 'Title' is the column name
    imdb, rt = get_ratings(title, api_key)
    imdb_ratings.append(imdb)
    rt_ratings.append(rt)
    print(f"Fetched ratings for: {title}")

# Add the ratings to the DataFrame
df['IMDB Rating'] = imdb_ratings
df['Rotten Tomatoes Rating'] = rt_ratings

# Write the updated data to a new Excel file
output_file = 'moistmeter_with_ratings.xlsx'  # Adjust as needed
df.to_excel(output_file, index=False)

print(f"Data with ratings has been successfully written to '{output_file}'")
