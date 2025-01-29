# -*- coding: utf-8 -*-
"""
Created on Tue Nov 21 11:34:55 2023

@author: MahdaviAl
"""

import pandas as pd
import matplotlib.pyplot as plt
from itertools import product


# Read df and preprocess essentials
df = pd.read_excel('compliance_air_database.xlsx')

# Rank based on sample and manufacturer frequency (creating dictionaries based on sample type and manufacturers)
dict_ = df['Sample Type'].value_counts().to_dict()
dict_ = dict(sorted(dict_.items(), key=lambda x: x[1], reverse = True))
samples = [k for k in dict_.keys()]
dict_rank = dict(zip(samples , range(1, len(dict_) + 1)))
w_rank = df['W. Manuf.'].value_counts().to_dict()
w_rank = dict(zip([k for k in w_rank.keys()] , range(1, len(w_rank) + 1)))

# Aggregate over sample type and window manufacturer, encoding with dictionaries, and sorting
df_agg = df.groupby(['Sample Type', 'W. Manuf.']).size().reset_index(name='Count')
df_agg['Rank'] = df_agg['Sample Type'].replace(dict_rank)
df_agg['Rank2'] = df_agg['W. Manuf.'].replace(w_rank)
df_agg.sort_values(['Rank', 'Rank2'], ascending = [True, True], inplace = True)


# Make tuples of sample type and manufacturers for new df generation and merging combined groups in the main df_agg
combinations = list(product(df_agg['Sample Type'].unique(), df_agg['W. Manuf.'].unique()))
df_new = pd.DataFrame(combinations, columns = ['Sample Type', 'W. Manuf.'])
df_new['Rank'] = df_new['Sample Type'].replace(dict_rank)
df_new['Rank2'] = df_new['W. Manuf.'].replace(w_rank)
df_new = df_new.merge(df_agg[['Rank', 'Rank2', 'Count']], on = ['Rank', 'Rank2'], how ='outer')
df_new = df_new.fillna(0)


# Create stacked bar groups based on sample type and manufacturer
for i in df_new['Rank2'].unique():
    locals()['list_%s' %i] = list(df_new[df_new['Rank2'] == i]['Count'])

categories = ['Sliding Door', 'Combination Window', 'Swing Door',
               'Awning Window', 'Fixed Window', 'Curtain Wall', 'Casement Window',
               'Entry Door', 'Operable Window', 'Spandrel Panel']

# Plot the stacked bar chart
fig, ax = plt.subplots()
c = ['g', 'r', 'b', 'k', 'purple', 'y', 'orange', 'pink', 'gray'] # colors for various window manufacturers

# Initialize the bottom list to keep track of the cumulative heights
bottom_values = [0] * len(categories)

# Loop through each data list and create a stacked bar
handles = []
for i, data_list in enumerate([list_1, list_2, list_3, list_4, list_5, list_6, list_7, list_8, list_9]):
    locals()[f'bar{i+1}'] = ax.bar(categories, data_list, color=c[i], bottom=bottom_values, label=categories[i + 1])
    
    # Update the bottom values by adding the current list values
    bottom_values = [b + d for b, d in zip(bottom_values, data_list)]
    
    # Update handles for legends later
    handles.append(locals()[f'bar{i+1}'])

# Set axes, legend, and title properties
ax.set_xticklabels(categories, rotation = 90)
ax.set_ylabel('Number of Tests')
ax.legend(handles)
ax.set_title('Air Tightness Compliance Test Per Sample Type and Window Manufacturer', fontsize = 14)

# Save and display
plt.savefig('Air_compliance_SB.jpg',  bbox_inches = 'tight', pad_inches = 0.1, dpi = 1200)
plt.show()
