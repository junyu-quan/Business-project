import pandas as pd
import matplotlib.pyplot as plt

url = 'https://raw.githubusercontent.com/junyu-quan/Business-project/main/Visualization.xlsx'
df = pd.read_excel(url, sheet_name=None, usecols="A:F", skiprows=1, nrows=10)
data = df['Sheet1']
categories = data['Visitors from...']
values = data['Total']
regions = ['Paris', 'Rest of France', 'Rest of Europe', 'Overseas']


# Plotting the Pie Chart of Paris Olympics Spectators
data_no_total = data[data['Visitors from...'] != 'Total']
categories_no_total = data_no_total['Visitors from...']
values_no_total = data_no_total['Total']
colors_pie = ['#004F6C', '#0094D3', '#8BC6E7', '#1F7A8C', '#B4A9A0', '#B0C8C2', '#D6E5E1']
plt.figure(figsize=(14, 7))
plt.pie(values_no_total, labels=categories_no_total, startangle=140, colors=colors_pie, wedgeprops=dict(width=0.3))
plt.title('Paris Olympics Spectator Carbon Emissions')
plt.legend(title='Categories', bbox_to_anchor=(1.2, 1), loc='upper left', fontsize=10, title_fontsize='13')
plt.subplots_adjust(right=0.75)
plt.show()

# Plotting the Composite Bar Chart of Paris Olympics Spectators
data_melted = data_no_total.melt(id_vars=['Visitors from...'], value_vars=regions, var_name='Region', value_name='Spectators')
data_pivot = data_melted.pivot_table(index='Region', columns='Visitors from...', values='Spectators')
colors_bar = ['#004F6C', '#0094D3', '#8BC6E7', '#1F7A8C',  '#B4A9A0', '#B0C8C2', '#D6E5E1']
plt.subplot(1, 2, 2)
ax = data_pivot.plot(kind='bar', stacked=True, figsize=(14, 7), color=colors_bar)
plt.title('Paris Olympics Spectator Carbon Emissions by Region and Category')
plt.ylabel('Carbon Emissions (kgCo2e)')
plt.xticks(rotation=45, ha='right')
plt.legend(title='Categories', bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=10, title_fontsize='13')
plt.tight_layout()
plt.show()

# Plotting the Pie Chart of Paris Olympics Operation
df1 = pd.read_excel(url, sheet_name='Sheet1', usecols="A:C", skiprows=14, nrows=11)
df1 = df1[df1['Type'] != 'Total']
colors_pie = ['#004F6C', '#0094D3', '#8BC6E7', '#1F7A8C', '#B4A9A0', '#B0C8C2', '#D6E5E1']
plt.figure(figsize=(10, 7))
plt.pie(df1['Percentage'], labels=df1['Type'], autopct='%1.1f%%', colors=colors_pie, wedgeprops=dict(width=0.3))
plt.legend(df1['Type'], title='Types', loc='center left', bbox_to_anchor=(-0.1, 0.8))
plt.title('Carbon Emissions Distribution of the Paris Olympics Operation')
plt.axis('equal')
plt.show()

# Plotting the Pie Chart of Paris Olympics Venues
df_venues = pd.read_excel(url, sheet_name='Sheet1', usecols="A:C", skiprows=27, nrows=4)
df_venues = df_venues[df_venues['Type'] != 'Total']
colors_pie_venues = ['#004F6C', '#0094D3', '#8BC6E7']
plt.figure(figsize=(12, 8))
plt.pie(df_venues['Percentage'], labels=df_venues['Type'], autopct='%1.1f%%', colors=colors_pie_venues, wedgeprops=dict(width=0.3))
plt.legend(df_venues['Type'], title='Types', loc='center left', bbox_to_anchor=(0.8, 0.5))
plt.title('Carbon Emissions Distribution of the Paris Olympics Venues')
plt.axis('equal')
plt.show()

# Plotting the Bar Chart of Carbon Emissions of the previous and Paris 2024 Olympics
df_emissions = pd.read_excel(url, sheet_name='Sheet1', usecols="A:E", skiprows=34, nrows=6)
df_emissions.columns = ['Category', 'London 2012', 'Rio 2016', 'Paris 2024', 'Total']
df_emissions = df_emissions[df_emissions['Category'] != 'Total']
plt.figure(figsize=(14, 8))
bar_width = 0.2
index = range(len(df_emissions))
bar1 = plt.bar(index, df_emissions['London 2012'], bar_width, label='London 2012', color='#004F6C')
bar2 = plt.bar([i + bar_width for i in index], df_emissions['Rio 2016'], bar_width, label='Rio 2016', color='#0094D3')
bar3 = plt.bar([i + 2 * bar_width for i in index], df_emissions['Paris 2024'], bar_width, label='Paris 2024', color='#8BC6E7')
plt.ylabel('Carbon Emissions (kg CO2e)')
plt.title('Carbon Emissions of the previous and Paris 2024 Olympics')
plt.xticks([i + bar_width for i in index], df_emissions['Category'])
plt.legend()
plt.tight_layout()
plt.show()

# Plotting the Pie Chart of Carbon Emissions Distribution by Category for London, Rio & Paris Olympics
df_emissions = pd.read_excel(url, sheet_name='Sheet1', usecols="A:E", skiprows=34, nrows=6)
df_emissions.columns = ['Category', 'London 2012', 'Rio 2016', 'Paris 2024', 'Total']
df_emissions = df_emissions[df_emissions['Category'] != 'Total']
df_total = df_emissions[['Category', 'Total']]
df_total.set_index('Category', inplace=True)
colors_pie_total = ['#004F6C', '#0094D3', '#8BC6E7', '#1F7A8C']
plt.figure(figsize=(12, 8))
plt.pie(df_total['Total'], labels=df_total.index, autopct='%1.1f%%', colors=colors_pie_total, wedgeprops=dict(width=0.3))
plt.legend(df_total.index, title='Categories', loc='center left', bbox_to_anchor=(0.8, 0.8))
plt.title('Carbon Emissions Distribution by Category for London, Rio & Paris Olympics')
plt.axis('equal')
plt.show()

# Plotting the Pie Chart of Paris 2024 Carbon Emissions
paris_data = df_emissions[['Category', 'Paris 2024']].set_index('Category')
colors_pie_paris = ['#004F6C', '#0094D3', '#8BC6E7', '#1F7A8C']
plt.figure(figsize=(12, 8))
plt.pie(paris_data['Paris 2024'], labels=paris_data.index, autopct='%1.1f%%', colors=colors_pie_paris, wedgeprops=dict(width=0.3))
plt.legend(paris_data.index, title='Categories', loc='center left', bbox_to_anchor=(0.8, 0.5))
plt.title('Carbon Emissions Distribution for Paris 2024')
plt.axis('equal')
plt.show()
