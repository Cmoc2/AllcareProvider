# -*- coding: utf-8 -*-
"""
Created on Thu Feb 13 14:40:26 2025

@author: ChristianOrtiz
"""

import pandas as pd
from geopy.distance import geodesic

# Example coordinates for demonstration purposes
coordinates = {
    '90260': (33.8872, -118.3526),
    '91104': (34.1654, -118.1316),
    '93063': (34.2694, -118.7815),
    '91505': (34.1663, -118.3453),
    '90023': (34.0273, -118.2029),
    '90221': (33.8958, -118.2176),
    '91762': (34.0633, -117.6509),
    '91791': (34.0686, -117.8975),
    '91706': (34.0853, -117.9609),
    '90043': (33.9888, -118.3316)
}

# Calculate distances between zip codes
distances = []
zip_codes = list(coordinates.keys())
for i in range(len(zip_codes)):
    for j in range(i + 1, len(zip_codes)):
        zip1 = zip_codes[i]
        zip2 = zip_codes[j]
        coord1 = coordinates[zip1]
        coord2 = coordinates[zip2]
        distance = geodesic(coord1, coord2).miles
        distances.append([zip1, zip2, distance])

# Create a DataFrame for the distances
distances_df = pd.DataFrame(distances, columns=['Zip Code 1', 'Zip Code 2', 'Distance (miles)'])

# Save the distances to a new Excel file
output_file_path = './zip_code_proximity.xlsx'
distances_df.to_excel(output_file_path, index=False)

print(f"Zip code proximities have been calculated and saved to {output_file_path}.")