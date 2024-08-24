import pandas as pd

# Load the data from the Excel files
products_bins_df = pd.read_excel('./ProductsBins.xlsx')
packing_slip_df = pd.read_excel('./packingslip.xlsx')

# Rename the columns to match expected names
products_bins_df.rename(columns={
    'ProductID': 'ProductID',
    'QtyAvailable': 'QtyAvailable',
    'UnitCapacity': 'BinCapacity',
    'BinName': 'BinName'
}, inplace=True)

# Rename the columns in packing slip
packing_slip_df.columns = ['ProductID', 'Qty']

# Initialize the output dataframe
output_df = packing_slip_df.copy()

# Function to allocate inventory
def allocate_inventory():
    # Create a dictionary to hold the bin data
    bin_dict = {}
    
    # Populate the bin dictionary with the product bin data
    for index, row in products_bins_df.iterrows():
        product_code = row['ProductID']
        bin_name = row['BinName']
        qty_available = row['QtyAvailable']
        bin_capacity = row['BinCapacity']
        
        if product_code not in bin_dict:
            bin_dict[product_code] = []
        
        bin_dict[product_code].append({
            'BinName': bin_name,
            'QtyAvailable': qty_available,
            'BinCapacity': bin_capacity
        })

    # Process each item in the packing slip
    for index, row in packing_slip_df.iterrows():
        product_code = row['ProductID']
        qty_incoming = row['Qty']
        allocated_qty = 0
        allocation_details = []

        # Check if the product code exists in the warehouse
        if product_code in bin_dict:
            for bin_info in bin_dict[product_code]:
                if allocated_qty >= qty_incoming:
                    break
                
                qty_available = bin_info['QtyAvailable']
                bin_capacity = bin_info['BinCapacity']
                bin_name = bin_info['BinName']
                
                if qty_available < bin_capacity:
                    if (bin_capacity - qty_available) >= (qty_incoming - allocated_qty):
                        bin_info['QtyAvailable'] += (qty_incoming - allocated_qty)
                        allocation_details.append(bin_name)
                        allocated_qty = qty_incoming
                    else:
                        allocation_details.append(f"{bin_name} (Partially Allocated)")
                        allocated_qty += (bin_capacity - qty_available)
                        bin_info['QtyAvailable'] = bin_capacity

            if allocated_qty < qty_incoming:
                for bin_info in bin_dict[product_code]:
                    if bin_info['QtyAvailable'] == 0 and bin_info['BinName'] not in ['PickingBin-Warehouse-274', 'Returns-Damaged']:
                        bin_info['QtyAvailable'] = qty_incoming - allocated_qty
                        allocation_details.append(f"{bin_info['BinName']} (New Bin)")
                        break
        else:
            for index, bin_info in products_bins_df.iterrows():
                if bin_info['QtyAvailable'] == 0 and bin_info['BinName'] not in ['PickingBin-Warehouse-274', 'Returns-Damaged']:
                    bin_info['QtyAvailable'] = qty_incoming
                    allocation_details.append(f"{bin_info['BinName']} (New SKU)")
                    break
        
        output_df.loc[index, 'Bin Allocation'] = ', '.join(allocation_details)

# Run the inventory allocation function
allocate_inventory()

# Save the output dataframe to an Excel file
output_df.to_excel('./output1232.xlsx', index=False)
