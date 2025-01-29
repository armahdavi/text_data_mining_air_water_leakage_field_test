# -*- coding: utf-8 -*-
"""
Created on Fri Nov 24 12:03:01 2023

@author: MahdaviAl
"""

import matplotlib.pyplot as plt
import pandas as pd

# ETL
df = pd.read_excel('compliance_air_database.xlsx')
a = df[(df['W. Manuf.'] == 'A') & 
          (df['Sample Type'] == 'Sliding Door')]['Leakage Flux']
b = df[(df['W. Manuf.'] == 'B') & 
          (df['Sample Type'] == 'Sliding Door')]['Leakage Flux']
c = df[(df['W. Manuf.'] == 'C') & 
          (df['Sample Type'] == 'Sliding Door')]['Leakage Flux']
data = [a, b, c]

# Plot the horizontal box plots
fig, ax = plt.subplots()
ax.boxplot(data, vert = False, labels = ['A', 'B', 'C'])

# Set the axes properties
ax.set_xlabel('Air Leakage Flux (L/(s.m$^{2}$)')
plt.xscale('log')
plt.xlim(0.003, 10)
ax.set_xticks([0.1, 1, 10])

# Set threshold lines
ax.axvline(x = 0.5, color = 'r', linestyle = '--')
ax.axvline(x = 1.5, color = 'g', linestyle = '--')
plt.text(2, 1.5, 'A2', fontsize = 10, ha = 'center', va = 'center', color = 'g')
plt.text(0.35, 1.5, 'A3', fontsize = 10, ha = 'center', va = 'center', color = 'r')

ax.set_title('Air Leakage Flux Distributions')

# Save and display
plt.savefig('box_plot_air_leakage_vs_manuf.jpg',  bbox_inches = 'tight', pad_inches = 0.1, dpi = 1200)
plt.show()
