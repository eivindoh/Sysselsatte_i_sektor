
# Importing necessary libraries
import pandas as pd
import numpy as np

# Reading in the detailed SSB data from the Excel file
detailed_ssb_data = pd.read_excel('13470_20230919-124013.xlsx', skiprows=4, header=None)
detailed_ssb_data.columns = ['Kommune', 'NACE', 'Sysselsatte']

# Fill down the 'Kommune' column
detailed_ssb_data['Kommune'].fillna(method='ffill', inplace=True)

# Separate the municipality number and name
#detailed_ssb_data[['KommuneNummer', 'KommuneNavn']] = detailed_ssb_data['Kommune'].str.split(' ', 1, expand=True)

detailed_ssb_data[['KommuneNummer', 'KommuneNavn']] = detailed_ssb_data['Kommune'].apply(lambda x: pd.Series(str(x).split(' ', 1)))

# Remove the 'K-' prefix from the 'KommuneNummer' values
detailed_ssb_data['KommuneNummer'] = detailed_ssb_data['KommuneNummer'].str.replace('K-', "")

# Filter out rows that don't represent municipalities (like '21-22 Industri')
detailed_ssb_data = detailed_ssb_data[detailed_ssb_data['KommuneNummer'].str.isnumeric() == True]

# Convert 'KommuneNummer' to integer
detailed_ssb_data['KommuneNummer'] = detailed_ssb_data['KommuneNummer'].astype(int)

# Drop the original 'Kommune' column
detailed_ssb_data.drop(columns=['Kommune'], inplace=True)

# Import the customer data
customer_data = pd.read_csv('f6a371cb-6df5-fb33-1a92-a4eac8f746df_brukere_normal.csv', delimiter=';', encoding='utf-8')

# Group the customer data by 'primarysector' and 'job_municipality', and count the unique users
unique_users_per_sector = customer_data.groupby(['primarysector', 'job_municipality'])['id'].nunique().reset_index()
unique_users_per_sector.columns = ['Sector', 'KommuneNummer', 'UniqueUsers']

# Renaming sectors in the customer data to match with SSB data
sector_mapping = {
    'Skole': 'Skole',
    'Barnehage': 'Barnehage',
    'Helse og omsorg': 'Helse og omsorg',
    'Barnevern': 'Barnevern',
    'Annet': 'Annet'
}

unique_users_per_sector['Sector'] = unique_users_per_sector['Sector'].map(sector_mapping)

# Filter out the rows in detailed_ssb_data that match the NACE codes specified for each sector
sector_nace_mapping = {
    'Skole': ['85.100', '85.201', '85.202', '85.521', '85.594', '85.601', '88.913'],
    'Barnehage': ['85.100', '88.911'],
    'Helse og omsorg': ['86.901', '86.902', '87.202', '86.903', '87.101', '87.102', '87.201', '87.203', '87.301', '87.302', '87.303', '87.304', '87.305', '87.909', '87.901', '88.101', '88.102', '88.103'],
    'Barnevern': ['88.991']
}

# Prepare a DataFrame to hold the final output
final_df = pd.DataFrame()

# Loop through each sector and filter the data
for sector, nace_codes in sector_nace_mapping.items():
    filtered_data = detailed_ssb_data[detailed_ssb_data['NACE'].str.split(' ', expand=True)[0].isin(nace_codes)]
    summed_data = filtered_data.groupby('KommuneNummer')['Sysselsatte'].sum().reset_index()
    summed_data['Sector'] = sector
    final_df = pd.concat([final_df, summed_data], ignore_index=True)

# Merge the final_df with unique_users_per_sector to calculate the ratio
merged_df = pd.merge(final_df, unique_users_per_sector, how='inner', on=['KommuneNummer', 'Sector'])
merged_df['Ratio'] = merged_df['UniqueUsers'] / merged_df['Sysselsatte']

# Merge the detailed SSB data to get the municipality names
merged_df_with_names = pd.merge(merged_df, detailed_ssb_data[['KommuneNummer', 'KommuneNavn']].drop_duplicates(), how='left', on='KommuneNummer')

# Write the data to an Excel file with different sheets for each sector
with pd.ExcelWriter('merged_data_detailed_sectors.xlsx') as writer:
    for sector in sector_nace_mapping.keys():
        temp_df = merged_df_with_names[merged_df_with_names['Sector'] == sector]
        temp_df = temp_df.sort_values(by='Ratio', ascending=False)
        temp_df.to_excel(writer, sheet_name=sector, index=False)
