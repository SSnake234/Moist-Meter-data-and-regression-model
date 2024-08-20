import pandas as pd
import requests

# Function to get the genre from OMDb API
def get_genre(title, api_key):
    base_url = 'http://www.omdbapi.com/'
    params = {
        't': title,
        'apikey': api_key
    }
    try:
        response = requests.get(base_url, params=params)
        response.raise_for_status()  # Raises an HTTPError for bad responses
        data = response.json()
        
        if data.get('Response') == 'True':
            return data.get('Genre', 'N/A')
        else:
            print(f"OMDb API error: {data.get('Error')} for title '{title}'")
            return 'N/A'
    
    except requests.exceptions.RequestException as e:
        print(f"Request failed for title '{title}': {e}")
        return 'N/A'
    except ValueError:
        print(f"Invalid JSON response for title '{title}'")
        return 'N/A'

# Your OMDb API key
api_key = '514b5453'

# Load the Excel file with the existing titles
input_file = "C:/Users/ASUS/Downloads/moistmeter.xlsx"  # Replace with your actual file name
df = pd.read_excel(input_file)

# Initialize list to store genres
genres = []

# Iterate over each title and fetch the genre
for title in df['Title']:  # Assuming 'Title' is the column name
    genre = get_genre(title, api_key)
    genres.append(genre)
    print(f"Fetched genre for: {title}")

# Add the genres to the DataFrame
df['Genre'] = genres

# Write the updated data to a new Excel file
output_file = 'moistmeter_with_genres.xlsx'  # Adjust as needed
df.to_excel(output_file, index=False)

print(f"Data with genres has been successfully written to '{output_file}'")
