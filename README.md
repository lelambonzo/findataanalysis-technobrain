# Financial Data Analysis Tool

## Pseudocode

### Initialization of the document file
1. Load the Excel file using a library like EPPlus.
2. Get the worksheet that contains the financial data.
3. Define lists to store asset data, returns, and volatilities.

### Extract data from document file
1. Iterate over the rows of the worksheet starting from the second row:
   - Create a list to store the data for the current asset.
   - Iterate over the columns of the worksheet starting from the second column:
     - Read the value from the current cell and add it to the asset's list.
   - Add the asset's list to the asset data list.
2. Iterate over each asset in the asset data list:
   - Create a list to store the returns for the current asset.
   - Iterate over the data points for the current asset starting from the second data point:
     - Calculate the return as (current value - previous value) / previous value.
     - Add the return to the returns list for the current asset.
### Calculations
1. Calculate the volatility for each asset:
   - Iterate over each asset's returns list:
     - Calculate the standard deviation of the returns.
     - Add the standard deviation to the volatilities list.
2. Calculate the correlation matrix between assets:
   - Create a 2D list to store the correlation matrix.
   - Iterate over each pair of assets:
     - Calculate the correlation between the returns of the two assets.
     - Add the correlation value to the correlation matrix at the corresponding position.

### Display the final data
1. Display the returns, volatilities, and correlation matrix.

### Exit the program.
