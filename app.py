import pandas as pd

# Load the data
products_bins_df = pd.read_excel('./ProductsBins (3).xlsx')
incoming_inventory_df = pd.read_excel('./TRKU4446897 - packingslip (1).xlsx', skiprows=1)

# Rename columns for easier reference
incoming_inventory_df.columns = ['QUANTITY', 'SKU']

# Define a function to allocate inventory to bins
def allocate_inventory(products_bins_df, incoming_inventory_df):
    # Create a dictionary to hold the bin allocations
    bin_allocations = []

    for _, row in incoming_inventory_df.iterrows():
        sku = row['SKU']
        incoming_qty = row['QUANTITY']
        remaining_qty = incoming_qty
        
        # Find existing bins for the SKU
        existing_bins = products_bins_df[products_bins_df['ProductID'] == sku]
        
        # Allocate to existing bins if there is space
        for i, bin_row in existing_bins.iterrows():
            if remaining_qty <= 0:
                break
            
            bin_capacity = bin_row['UnitCapacity']
            qty_available = bin_row['QtyAvailable']
            if qty_available < bin_capacity:
                space_available = bin_capacity - qty_available
                qty_to_allocate = min(space_available, remaining_qty)
                products_bins_df.at[i, 'QtyAvailable'] += qty_to_allocate
                remaining_qty -= qty_to_allocate
                bin_allocations.append((sku, qty_to_allocate, bin_row['BinName']))
        
        # Allocate to new bins if needed
        if remaining_qty > 0:
            empty_bins = products_bins_df[(products_bins_df['QtyAvailable'] == 0) & (products_bins_df['ProductID'] != sku)]
            for i, bin_row in empty_bins.iterrows():
                if remaining_qty <= 0:
                    break
                
                bin_capacity = bin_row['UnitCapacity']
                qty_to_allocate = min(bin_capacity, remaining_qty)
                products_bins_df.at[i, 'QtyAvailable'] = qty_to_allocate
                products_bins_df.at[i, 'ProductID'] = sku
                remaining_qty -= qty_to_allocate
                bin_allocations.append((sku, qty_to_allocate, bin_row['BinName']))

    # Create the output DataFrame
    output_df = incoming_inventory_df.copy()
    output_df['BinLocation'] = [None] * len(output_df)
    
    for i, row in output_df.iterrows():
        sku = row['SKU']
        allocations = [allocation[2] for allocation in bin_allocations if allocation[0] == sku]
        output_df.at[i, 'BinLocation'] = ', '.join(allocations)
    
    return output_df, products_bins_df

# Run the allocation function
output_df, updated_products_bins_df = allocate_inventory(products_bins_df, incoming_inventory_df)

# Save the result to a new Excel file
output_file_path = './output00xlsx'
with pd.ExcelWriter(output_file_path) as writer:
    output_df.to_excel(writer, sheet_name='PackingSlipWithBins', index=False)
    updated_products_bins_df.to_excel(writer, sheet_name='UpdatedProductsBins', index=False)

# import ace_tools as tools; tools.display_dataframe_to_user(name="Packing Slip with Bins", dataframe=output_df)
