import os
import pandas as pd
import plotly.express as px
import dash
from dash import html, dcc
from dash.dash_table import DataTable

# Initialize Dash app
app = dash.Dash(__name__)

folder_path = r"C:\Users\hp\Desktop\First_official_project\supermarket" 
files = os.listdir(folder_path)

# Initialize combined_data as an empty DataFrame before the loop
combined_data = pd.DataFrame()

# Read data in chunks
#chunk_size = 1000  # Adjust the chunk size based on your memory constraints
for file in files:
    if file.endswith('.xls'):
        month = int(file[-6:-4])  # Extract the last two digits from the file name
        excel_file = pd.ExcelFile(os.path.join(folder_path, file), engine='openpyxl')
        for sheet_name in excel_file.sheet_names:
            # Read the sheet in chunks
            chunk = pd.read_excel(excel_file, sheet_name=sheet_name)
                # Create a new column with the 'Month' value
            chunk['Month'] = month
                #chunk.insert(loc=0, column='Month', value=month)
            combined_data = pd.concat([combined_data, chunk], ignore_index=True)

#Save the combined data to a new Excel file with .xlsx extension
combined_file_path = os.path.join(folder_path, "combined_data.xlsx")
combined_data.to_excel(combined_file_path, index=False)

# Assuming 'unit' is the name of the column you want on the y-axis
data = combined_data['unit'].value_counts().nlargest(10).reset_index()

# Create Plotly figures
fig1 = px.bar(combined_data.groupby('item')['amount'].sum().nlargest(10).reset_index(), x='item', y='amount',
              title='Top-Selling Items', labels={'item': 'Item', 'amount': 'Total Sales Amount'}, color='amount')

fig2 = px.histogram(combined_data, x='quantity', nbins=30, title='Quantity Distribution',
                    labels={'quantity': 'Quantity Sold', 'count': 'Frequency'}, color='quantity')
fig3 = px.line(combined_data, x='item', y='amount', color='item', title='Seasonal Sales Patterns',
               labels={'item': 'Item', 'amount': 'Total Sales Amount'})
fig4 = px.bar(combined_data['unit'].value_counts().reset_index(), x=combined_data['unit'].value_counts().index,
              y='unit', title='Unit Breakdown', labels={'index': 'Unit', 'unit': 'Number of Sales'}, color='unit')
fig5 = px.box(combined_data, x='item', y='quantity', title='Quantity Distribution for Items',
              labels={'item': 'Item', 'quantity': 'Quantity Sold'}, color='quantity')
fig6 = px.bar(combined_data.groupby('item')['amount'].mean().reset_index(), x='item', y='amount',
              title='Average Transaction Size', labels={'item': 'Item', 'amount': 'Average Amount per Transaction'},
              color='amount')
fig7 = px.line(combined_data, x='Month', y='quantity', color='item', title='Quantity Sold Over Time',
               labels={'Month': 'Month', 'quantity': 'Quantity Sold', 'item': 'Item'})
fig8 = px.bar(combined_data.groupby('item')['amount'].sum().nlargest(10).reset_index(),
              x='item', y='amount', title='Most Profitable Items',
              labels={'item': 'Item', 'amount': 'Total Sales Amount'}, color='amount')
fig9 = px.line(combined_data, x='Month', y='price', color='item', title='Unit Price Trends',
                labels={'Month': 'Month', 'price': 'Unit Price', 'item': 'Item'})
fig10 = px.bar(combined_data.groupby('item')['quantity'].std().nlargest(10).reset_index(), x='item', y='quantity',
               title='Items with Consistent Sales', labels={'item': 'Item', 'quantity': 'Quantity Standard Deviation'},
               color='quantity')
fig11 = px.bar(combined_data.groupby('itemno')['amount'].sum().nlargest(10).reset_index(),
               x='itemno', y='amount', title='High-Value Transactions',
               labels={'itemno': 'Item Number', 'amount': 'Total Amount'}, color='amount')
fig12 = px.scatter(combined_data, x='quantity', y='amount', title='Quantity-Amount Relationship',
                   labels={'quantity': 'Quantity Sold', 'amount': 'Total Amount'}, color='amount')
fig13 = px.pie(combined_data.groupby('item')['amount'].sum().nlargest(10).reset_index(), names='item', values='amount',
               title='Revenue Concentration', color='item')
fig14 = px.box(combined_data, x='item', y='price', title='Price Variation',
               labels={'item': 'Item', 'price': 'Price'}, color='price')
fig15 = px.bar(combined_data.groupby('item')['amount'].sum().nsmallest(10).reset_index(), x='item', y='amount',
               title='Lowest Selling Items', labels={'item': 'Item', 'amount': 'Total Sales Amount'}, color='amount')
fig16 = px.bar(combined_data['unit'].value_counts().nlargest(10))
app.layout = html.Div([
    html.H1("Supermarket Data Visualization"),
dcc.Graph(figure=fig1, id='top-selling-items'),
dcc.Graph(figure=fig2, id='quantity-distribution'),
dcc.Graph(figure=fig3, id='seasonal-sales-patterns'),
dcc.Graph(figure=fig4, id='unit-breakdown'),
dcc.Graph(figure=fig5, id='quantity-distribution-for-items'),
dcc.Graph(figure=fig6, id='average-transaction-size'),
dcc.Graph(figure=fig7, id='quantity-sold-over-time'),
dcc.Graph(figure=fig8, id='Most Profitable Items'),
dcc.Graph(figure=fig9, id='unit-price-trends'),
dcc.Graph(figure=fig10, id='items-with-consistent-sales'),
dcc.Graph(figure=fig11, id='high-value-transactions'),
dcc.Graph(figure=fig12, id='quantity-amount-relationship'),
dcc.Graph(figure=fig13, id='revenue-concentration'),
dcc.Graph(figure=fig14, id='price-variation'),
dcc.Graph(figure=fig15, id='lowest-selling-items'),
dcc.Graph(figure=fig16, id='popular-units-placeholder'),
   DataTable(
        id='combined-data-table',
        columns=[{"name": col, "id": col} for col in combined_data.columns],
        data=combined_data.to_dict('records'),
    ),
])
if __name__ == '__main__':
    app.run_server(debug=True)