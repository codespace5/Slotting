import pandas as pd
import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QHBoxLayout, QWidget, QLabel, QFileDialog, QMessageBox, QLineEdit
)
from PyQt5.QtCore import QThread, pyqtSignal

class WorkerThread(QThread):
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, products_file, incoming_file, bins_file):
        super().__init__()
        self.products_file = products_file
        self.incoming_file = incoming_file
        self.bins_file = bins_file

    def run(self):
        try:
            bins_df = allocate_inventory(self.products_file, self.incoming_file, self.bins_file)
            bins_df.to_excel(self.bins_file, index=False)
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))

# def extract_bin_level(bin_name: str) -> int:
#     parts = bin_name.split('-')
#     if len(parts) >= 3:
#         return int(parts[1])
#     elif len(parts) == 2:
#         dash_pos = bin_name.find('-')
#         return int(bin_name[dash_pos + 1:])
#     return 0
def extract_bin_level(bin_name: str) -> int:
    parts = bin_name.split('-')
    if len(parts) == 2:
        alphanumeric_part = parts[0]
        numeric_part = ''.join(filter(str.isdigit, alphanumeric_part))
        return int(numeric_part)
    elif len(parts) >= 3:
        alphanumeric_part = parts[1]
        numeric_part = ''.join(filter(str.isdigit, alphanumeric_part))
        return int(numeric_part)
    return 0

def extract_sku_level(bin_name: str) -> int:
    parts = bin_name.split('-')
    if len(parts) == 2:
        dash_pos = bin_name.find('-')
        bin_number = int(bin_name[dash_pos + 1:])

        if bin_number > 0:
            if bin_number in [4, 3, 2, 28]:
                return 3
            elif bin_number in [6]:
                return 2
            elif bin_number in [5]:
                return 3
            elif bin_number in [2, 0, 9, 8]:
                return 1
        return None  # Default value if bin_number is not in any case

    parts = bin_name.split('-')
    if len(parts) >= 3:
        return int(parts[1])
    elif len(parts) == 2:
        dash_pos = bin_name.find('-')
        return int(bin_name[dash_pos + 1:])
    return 0

def extract_sku_num(bin_name: str) -> int:
    parts = bin_name.split('-')
    if len(parts) == 2:
        dash_pos = bin_name.find('-')
        bin_number = int(bin_name[dash_pos + 1:])
        return bin_number


def allocate_inventory(products_file, incoming_file, bins_file):
    products_df = pd.read_excel(products_file)
    incoming_df = pd.read_excel(incoming_file)
    bins_df = pd.read_excel(bins_file)
    
    excluded_bins = {"PutAwayBin-274", "ReceivingBin-Warehouse-274", "Returns-Damaged", "PickingBin-Warehouse-274", "Returns-Damaged-Light", "F1-13", "F1-14", "Returns-Putaway"}
    
    output_df = pd.DataFrame(columns=['Quantity', 'SKU', 'Bin', 'Notes'])

    for i, row in incoming_df.iterrows():
        incoming_sku = row['SKU']
        incoming_qty = row['QUANTITY']
        remaining_qty = incoming_qty
        incoming_num = extract_sku_num(incoming_sku)
        print("sku", incoming_sku, ": ", incoming_num)
        new_row = {'Quantity': incoming_qty, 'SKU': incoming_sku, 'Bin': '', 'Notes': ''}
        
        for _, prod_row in products_df.iterrows():
            if prod_row['ProductID'] == incoming_sku:
                available_qty = prod_row['QtyAvailable']
                bin_capacity = prod_row['UnitCapacity']
                bin_name = prod_row['BinName']
                if incoming_num in [12, 10, 9, 8] and "Aisle" in bin_name:
                    continue
                if bin_name not in excluded_bins and (available_qty + remaining_qty) <= bin_capacity and available_qty > 0:
                    new_row['Bin'] = bin_name
                    remaining_qty = 0
                    break
                if bin_name not in excluded_bins and bin_capacity == 20 and (available_qty + remaining_qty) <= (bin_capacity + 5 ) and available_qty > 0:
                    new_row['Bin'] = bin_name
                    new_row['Notes'] = 'Check capacity'
                    remaining_qty = 0
                    break
                if bin_name not in excluded_bins and bin_capacity == 50 and (available_qty + remaining_qty) <= (bin_capacity + 15) and available_qty > 0:
                    new_row['Bin'] = bin_name
                    new_row['Notes'] = 'Check capacity'
                    remaining_qty = 0
                    break
                elif bin_name not in excluded_bins and (available_qty + remaining_qty) > bin_capacity and available_qty > 0:
                    new_row['Bin'] += bin_name + ', '
                    remaining_qty -= (bin_capacity - available_qty)
        
        if remaining_qty > 0:
            unique_capacities = bins_df['UnitCapacity'].unique()
            capacities_array = sorted(unique_capacities, reverse=True)
            
            for _, bin_row in bins_df.iterrows():
                bin_name = bin_row['BinName']
                bin_qty = bin_row['TotalQty']
                bin_capacity = bin_row['UnitCapacity']
                # if "Aisle" in bin_name:
                #     print("bin_name", bin_name)
                for bin_unit_capacity in capacities_array:
                    if bin_qty == 0 and remaining_qty > 0 and bin_capacity == bin_unit_capacity:
                        if incoming_num in [12, 10, 9, 8] and "Aisle" in bin_name:
                            continue
                        #small bins
                        if remaining_qty <= bin_unit_capacity:
                            bin_level = extract_bin_level(bin_name)
                            
                            incoming_bin_level = extract_sku_level(incoming_sku)
                            # print("bin_name", bin_name, "bin_level",  bin_level)
                            # print("sku_name", incoming_sku, "sku_level",  incoming_bin_level)
                            if bin_level == incoming_bin_level:
                                new_row['Bin'] += bin_name + ', '
                                bins_df.loc[bins_df['BinName'] == bin_name, 'TotalQty'] = 888888
                                remaining_qty = 0
                                break
                
                        #over bins
                        if bin_unit_capacity == 20 and (remaining_qty - bin_unit_capacity) > 5:
                            bin_level = extract_bin_level(bin_name)
                            incoming_bin_level = extract_sku_level(incoming_sku)
                            # print("bin_name", bin_name, "bin_level",  bin_level)
                            # print("sku_name", incoming_sku, "sku_level",  incoming_bin_level)
                            if bin_level == incoming_bin_level:
                                new_row['Bin'] += bin_name + ', '
                                new_row['Notes'] = 'Check capacity'
                                bins_df.loc[bins_df['BinName'] == bin_name, 'TotalQty'] = 999999
                                remaining_qty = 0
                                break
                        if bin_unit_capacity == 50 and (remaining_qty - bin_unit_capacity) > 15:
                            bin_level = extract_bin_level(bin_name)
                            incoming_bin_level = extract_sku_level(incoming_sku)
                            # print("bin_name", bin_name, "bin_level",  bin_level)
                            # print("sku_name", incoming_sku, "sku_level",  incoming_bin_level)
                            if bin_level == incoming_bin_level:
                                new_row['Bin'] += bin_name + ', '
                                new_row['Notes'] = 'Check capacity'
                                bins_df.loc[bins_df['BinName'] == bin_name, 'TotalQty'] = 999999
                                remaining_qty = 0
                                break
                        # Multi Bins
                        if remaining_qty <= bin_unit_capacity * 2:
                            bin_level = extract_bin_level(bin_name)
                            incoming_bin_level = extract_sku_level(incoming_sku)
                            # print("bin_name", bin_name, "bin_level",  bin_level)
                            # print("sku_name", incoming_sku, "sku_level",  incoming_bin_level)
                            if bin_level == incoming_bin_level:
                                new_row['Bin'] += bin_name + ', '
                                bins_df.loc[bins_df['BinName'] == bin_name, 'TotalQty'] = 77777
                                remaining_qty = remaining_qty - bin_unit_capacity
                                break

        # new_row['Notes'] = 'Check capacity' if remaining_qty > 0 else ''
        output_df = pd.concat([output_df, pd.DataFrame([new_row])], ignore_index=True)
    
    output_df.to_excel('output.xlsx', index=False)
    return bins_df

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Inventory Allocation')
        self.setGeometry(100, 100, 700, 300)
        
        self.layout = QVBoxLayout()
        
        # Create widgets for Products File
        self.products_label = QLabel('Products File:', self)
        self.products_path = QLineEdit(self)
        self.products_path.setReadOnly(True)
        self.products_button = QPushButton('Select', self)
        self.products_button.clicked.connect(self.select_products_file)
        
        # Create layout for Products File
        self.products_layout = QHBoxLayout()
        self.products_layout.addWidget(self.products_label)
        self.products_layout.addWidget(self.products_path)
        self.products_layout.addWidget(self.products_button)
        self.layout.addLayout(self.products_layout)
        
        # Create widgets for Incoming File
        self.incoming_label = QLabel('Incoming File:', self)
        self.incoming_path = QLineEdit(self)
        self.incoming_path.setReadOnly(True)
        self.incoming_button = QPushButton('Select', self)
        self.incoming_button.clicked.connect(self.select_incoming_file)
        
        # Create layout for Incoming File
        self.incoming_layout = QHBoxLayout()
        self.incoming_layout.addWidget(self.incoming_label)
        self.incoming_layout.addWidget(self.incoming_path)
        self.incoming_layout.addWidget(self.incoming_button)
        self.layout.addLayout(self.incoming_layout)
        
        # Create widgets for Bins File
        self.bins_label = QLabel('Bins File:', self)
        self.bins_path = QLineEdit(self)
        self.bins_path.setReadOnly(True)
        self.bins_button = QPushButton('Select', self)
        self.bins_button.clicked.connect(self.select_bins_file)
        
        # Create layout for Bins File
        self.bins_layout = QHBoxLayout()
        self.bins_layout.addWidget(self.bins_label)
        self.bins_layout.addWidget(self.bins_path)
        self.bins_layout.addWidget(self.bins_button)
        self.layout.addLayout(self.bins_layout)
        
        # Run Button
        self.run_button = QPushButton('Run Allocation', self)
        self.run_button.clicked.connect(self.run_allocation)
        self.run_button.setEnabled(False)
        self.layout.addWidget(self.run_button)
        
        # Status Label
        self.status_label = QLabel('', self)
        self.layout.addWidget(self.status_label)
        
        container = QWidget()
        container.setLayout(self.layout)
        self.setCentralWidget(container)
        
        self.products_file = ''
        self.incoming_file = ''
        self.bins_file = ''
        
        self.show()

    def select_products_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, 'Select Products File', '', 'Excel Files (*.xlsx)')
        if file_name:
            self.products_file = file_name
            self.products_path.setText(file_name)
            self.check_files_ready()

    def select_incoming_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, 'Select Incoming File', '', 'Excel Files (*.xlsx)')
        if file_name:
            self.incoming_file = file_name
            self.incoming_path.setText(file_name)
            self.check_files_ready()

    def select_bins_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, 'Select Bins File', '', 'Excel Files (*.xlsx)')
        if file_name:
            self.bins_file = file_name
            self.bins_path.setText(file_name)
            self.check_files_ready()

    def check_files_ready(self):
        if self.products_file and self.incoming_file and self.bins_file:
            self.run_button.setEnabled(True)
        else:
            self.run_button.setEnabled(False)

    def run_allocation(self):
        if not (self.products_file and self.incoming_file and self.bins_file):
            return

        self.status_label.setText('Running allocation...')
        self.run_button.setEnabled(False)

        self.worker = WorkerThread(self.products_file, self.incoming_file, self.bins_file)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_finished(self):
        self.status_label.setText('Allocation complete. Output saved to output.xlsx.')
        self.run_button.setEnabled(True)

    def on_error(self, error):
        QMessageBox.critical(self, 'Error', f'An error occurred: {error}')
        self.run_button.setEnabled(True)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec_())




#
#  import pandas as pd
# import sys
# from PyQt5.QtWidgets import (
#     QApplication, QMainWindow, QPushButton, QVBoxLayout, QHBoxLayout, QWidget, QLabel, QFileDialog, QMessageBox, QLineEdit
# )
# from PyQt5.QtCore import QThread, pyqtSignal

# class WorkerThread(QThread):
#     finished = pyqtSignal()

#     def __init__(self, products_file, incoming_file, bins_file):
#         super().__init__()
#         self.products_file = products_file
#         self.incoming_file = incoming_file
#         self.bins_file = bins_file

#     def run(self):
#         try:
#             allocate_inventory(self.products_file, self.incoming_file, self.bins_file)
#             self.finished.emit()
#         except Exception as e:
#             self.finished.emit()
#             QMessageBox.critical(None, "Error", str(e))

# def extract_bin_level(bin_name: str) -> int:
#     parts = bin_name.split('-')
#     if len(parts) >= 3:
#         return int(parts[1])
#     elif len(parts) == 2:
#         dash_pos = bin_name.find('-')
#         return int(bin_name[dash_pos + 1:])
#     return 0

# def allocate_inventory(products_file, incoming_file, bins_file):
#     products_df = pd.read_excel(products_file)
#     incoming_df = pd.read_excel(incoming_file)
#     bins_df = pd.read_excel(bins_file)
    
#     excluded_bins = {"PutAwayBin-274", "ReceivingBin-Warehouse-274", "Returns-Damaged", "PickingBin-Warehouse-274", "Returns-Damaged-Light"}
    
#     output_df = pd.DataFrame(columns=['Quantity', 'SKU', 'Bin', 'Notes'])
    
#     for i, row in incoming_df.iterrows():
#         incoming_sku = row['SKU']
#         incoming_qty = row['QUANTITY']
#         remaining_qty = incoming_qty
        
#         new_row = {'Quantity': incoming_qty, 'SKU': incoming_sku, 'Bin': '', 'Notes': ''}
        
#         for _, prod_row in products_df.iterrows():
#             if prod_row['ProductID'] == incoming_sku:
#                 available_qty = prod_row['QtyAvailable']
#                 bin_capacity = prod_row['UnitCapacity']
#                 bin_name = prod_row['BinName']
                
#                 if bin_name not in excluded_bins and (available_qty + remaining_qty) <= bin_capacity and available_qty > 0:
#                     new_row['Bin'] = bin_name
#                     remaining_qty = 0
#                     break
#                 elif bin_name not in excluded_bins and (available_qty + remaining_qty) > bin_capacity and available_qty > 0:
#                     new_row['Bin'] += bin_name + ', '
#                     remaining_qty -= (bin_capacity - available_qty)
        
#         if remaining_qty > 0:
#             unique_capacities = bins_df['UnitCapacity'].unique()
#             capacities_array = sorted(unique_capacities, reverse=True)
            
#             for _, bin_row in bins_df.iterrows():
#                 bin_name = bin_row['BinName']
#                 bin_qty = bin_row['TotalQty']
#                 bin_capacity = bin_row['UnitCapacity']

#                 for bin_unit_capacity in capacities_array:
#                     if bin_qty == 0 and remaining_qty > 0 and bin_capacity == bin_unit_capacity:
#                         #small bins
#                         if remaining_qty <= bin_unit_capacity:
#                             bin_level = extract_bin_level(bin_name)
#                             incoming_bin_level = extract_bin_level(incoming_sku)
#                             if bin_level == incoming_bin_level:
#                                 new_row['Bin'] += bin_name + ', '
#                                 bins_df.loc[bins_df['BinName'] == bin_name, 'UnitCapacity'] = 999999
#                                 remaining_qty = 0
#                                 break
                
#                         #over bins
#                         if bin_unit_capacity == 20 and (remaining_qty - bin_unit_capacity) > 5:
#                             bin_level = extract_bin_level(bin_name)
#                             incoming_bin_level = extract_bin_level(incoming_sku)
#                             if bin_level == incoming_bin_level:
#                                 new_row['Bin'] += bin_name + ', '
#                                 new_row['Notes'] = 'Check capacity'
#                                 bins_df.loc[bins_df['BinName'] == bin_name, 'UnitCapacity'] = 999999
#                                 remaining_qty = 0
#                                 break
#                         if bin_unit_capacity == 50 and (remaining_qty - bin_unit_capacity) > 15:
#                             bin_level = extract_bin_level(bin_name)
#                             incoming_bin_level = extract_bin_level(incoming_sku)
#                             if bin_level == incoming_bin_level:
#                                 new_row['Bin'] += bin_name + ', '
#                                 new_row['Notes'] = 'Check capacity'
#                                 bins_df.loc[bins_df['BinName'] == bin_name, 'UnitCapacity'] = 999999
#                                 remaining_qty = 0
#                                 break
#                         # Multi Bins
#                         if remaining_qty <= bin_unit_capacity * 2:
#                             bin_level = extract_bin_level(bin_name)
#                             incoming_bin_level = extract_bin_level(incoming_sku)
#                             if bin_level == incoming_bin_level:
#                                 new_row['Bin'] += bin_name + ', '
#                                 bins_df.loc[bins_df['BinName'] == bin_name, 'UnitCapacity'] = 999999
#                                 remaining_qty = remaining_qty - bin_unit_capacity
#                                 break

#         new_row['Notes'] = 'Check capacity' if remaining_qty > 0 else ''
#         output_df = pd.concat([output_df, pd.DataFrame([new_row])], ignore_index=True)
    
#     output_df.to_excel('output.xlsx', index=False)
#     bins_df.to_excel(bins_file, index=False)
# class MainWindow(QMainWindow):
#     def __init__(self):
#         super().__init__()
#         self.setWindowTitle('Inventory Allocation')
#         self.setGeometry(100, 100, 700, 300)
        
#         self.layout = QVBoxLayout()
        
#         # Create widgets for Products File
#         self.products_label = QLabel('Products File:', self)
#         self.products_path = QLineEdit(self)
#         self.products_path.setReadOnly(True)
#         self.products_button = QPushButton('Select', self)
#         self.products_button.clicked.connect(self.select_products_file)
        
#         # Create layout for Products File
#         self.products_layout = QHBoxLayout()
#         self.products_layout.addWidget(self.products_label)
#         self.products_layout.addWidget(self.products_path)
#         self.products_layout.addWidget(self.products_button)
#         self.layout.addLayout(self.products_layout)
        
#         # Create widgets for Incoming File
#         self.incoming_label = QLabel('Incoming File:', self)
#         self.incoming_path = QLineEdit(self)
#         self.incoming_path.setReadOnly(True)
#         self.incoming_button = QPushButton('Select', self)
#         self.incoming_button.clicked.connect(self.select_incoming_file)
        
#         # Create layout for Incoming File
#         self.incoming_layout = QHBoxLayout()
#         self.incoming_layout.addWidget(self.incoming_label)
#         self.incoming_layout.addWidget(self.incoming_path)
#         self.incoming_layout.addWidget(self.incoming_button)
#         self.layout.addLayout(self.incoming_layout)
        
#         # Create widgets for Bins File
#         self.bins_label = QLabel('Bins File:', self)
#         self.bins_path = QLineEdit(self)
#         self.bins_path.setReadOnly(True)
#         self.bins_button = QPushButton('Select', self)
#         self.bins_button.clicked.connect(self.select_bins_file)
        
#         # Create layout for Bins File
#         self.bins_layout = QHBoxLayout()
#         self.bins_layout.addWidget(self.bins_label)
#         self.bins_layout.addWidget(self.bins_path)
#         self.bins_layout.addWidget(self.bins_button)
#         self.layout.addLayout(self.bins_layout)
        
#         # Run Button
#         self.run_button = QPushButton('Run Allocation', self)
#         self.run_button.clicked.connect(self.run_allocation)
#         self.run_button.setEnabled(False)
#         self.layout.addWidget(self.run_button)
        
#         # Status Label
#         self.status_label = QLabel('', self)
#         self.layout.addWidget(self.status_label)
        
#         container = QWidget()
#         container.setLayout(self.layout)
#         self.setCentralWidget(container)
        
#         self.products_file = ''
#         self.incoming_file = ''
#         self.bins_file = ''
        
#         self.show()

#     def select_products_file(self):
#         file_name, _ = QFileDialog.getOpenFileName(self, 'Select Products File', '', 'Excel Files (*.xlsx)')
#         if file_name:
#             self.products_file = file_name
#             self.products_path.setText(file_name)
#             self.check_files_ready()

#     def select_incoming_file(self):
#         file_name, _ = QFileDialog.getOpenFileName(self, 'Select Incoming File', '', 'Excel Files (*.xlsx)')
#         if file_name:
#             self.incoming_file = file_name
#             self.incoming_path.setText(file_name)
#             self.check_files_ready()

#     def select_bins_file(self):
#         file_name, _ = QFileDialog.getOpenFileName(self, 'Select Bins File', '', 'Excel Files (*.xlsx)')
#         if file_name:
#             self.bins_file = file_name
#             self.bins_path.setText(file_name)
#             self.check_files_ready()

#     def check_files_ready(self):
#         if self.products_file and self.incoming_file and self.bins_file:
#             self.run_button.setEnabled(True)
#         else:
#             self.run_button.setEnabled(False)

#     def run_allocation(self):
#         if not (self.products_file and self.incoming_file and self.bins_file):
#             QMessageBox.warning(self, 'Warning', 'Please select all files before running.')
#             return
        
#         self.run_button.setEnabled(False)
#         self.status_label.setText('Processing...')
#         self.thread = WorkerThread(self.products_file, self.incoming_file, self.bins_file)
#         self.thread.finished.connect(self.on_finished)
#         self.thread.start()

#     def on_finished(self):
#         self.run_button.setEnabled(True)
#         self.status_label.setText('Allocation completed')

# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     window = MainWindow()
#     sys.exit(app.exec_())













# import pandas as pd
# import sys
# from PyQt5.QtWidgets import (
#     QApplication, QMainWindow, QPushButton, QVBoxLayout, QHBoxLayout, QWidget, QLabel, QFileDialog, QMessageBox, QLineEdit
# )
# from PyQt5.QtCore import QThread, pyqtSignal

# class WorkerThread(QThread):
#     finished = pyqtSignal()

#     def __init__(self, products_file, incoming_file, bins_file):
#         super().__init__()
#         self.products_file = products_file
#         self.incoming_file = incoming_file
#         self.bins_file = bins_file

#     def run(self):
#         try:
#             allocate_inventory(self.products_file, self.incoming_file, self.bins_file)
#             self.finished.emit()
#         except Exception as e:
#             self.finished.emit()
#             QMessageBox.critical(None, "Error", str(e))

# def extract_bin_level(bin_name: str) -> int:
#     parts = bin_name.split('-')
#     if len(parts) >= 3:
#         return int(parts[1])
#     elif len(parts) == 2:
#         dash_pos = bin_name.find('-')
#         return int(bin_name[dash_pos + 1:])
#     return 0

# def allocate_inventory(products_file, incoming_file, bins_file):
#     # products_df = pd.read_excel(products_file, sheet_name='ProductsBins')
#     # incoming_df = pd.read_excel(incoming_file, sheet_name='PackingSlip')
#     # bins_df = pd.read_excel(bins_file, sheet_name='Bins')
#     products_df = pd.read_excel(products_file)
#     incoming_df = pd.read_excel(incoming_file )
#     bins_df = pd.read_excel(bins_file)
    
#     excluded_bins = {"PutAwayBin-274", "ReceivingBin-Warehouse-274", "Returns-Damaged", "PickingBin-Warehouse-274", "Returns-Damaged-Light"}
    
#     output_df = pd.DataFrame(columns=['Quantity', 'SKU', 'Bin', 'Notes'])
    
#     for i, row in incoming_df.iterrows():
#         incoming_sku = row['SKU']
#         incoming_qty = row['QUANTITY']
#         remaining_qty = incoming_qty
        
#         new_row = {'Quantity': incoming_qty, 'SKU': incoming_sku, 'Bin': '', 'Notes': ''}
        
#         for _, prod_row in products_df.iterrows():
#             if prod_row['ProductID'] == incoming_sku:
#                 available_qty = prod_row['QtyAvailable']
#                 bin_capacity = prod_row['UnitCapacity']
#                 bin_name = prod_row['BinName']
                
#                 if bin_name not in excluded_bins and (available_qty + remaining_qty) <= bin_capacity and available_qty > 0:
#                     new_row['Bin'] = bin_name
#                     remaining_qty = 0
#                     break
#                 elif bin_name not in excluded_bins and (available_qty + remaining_qty) > bin_capacity and available_qty > 0:
#                     new_row['Bin'] += bin_name + ', '
#                     remaining_qty -= (bin_capacity - available_qty)
        
#         if remaining_qty > 0:
#             unique_capacities = bins_df['UnitCapacity'].unique()
#             capacities_array = sorted(unique_capacities, reverse=True)
            
#             for _, bin_row in bins_df.iterrows():
#                 bin_name = bin_row['BinName']
#                 bin_qty = bin_row['TotalQty']
#                 bin_capacity = bin_row['UnitCapacity']
                
#                 if bin_qty == 0 and bin_capacity in capacities_array and remaining_qty > 0:
#                     bin_level = extract_bin_level(bin_name)
#                     incoming_bin_level = extract_bin_level(incoming_sku)
                    
#                     if bin_level == incoming_bin_level:
#                         new_row['Bin'] += bin_name + ', '
#                         bins_df.loc[bins_df['BinName'] == bin_name, 'UnitCapacity'] = remaining_qty
#                         remaining_qty = 0
#                         break
        
#         new_row['Notes'] = 'Check capacity' if remaining_qty > 0 else ''
#         output_df = output_df.append(new_row, ignore_index=True)
    
#     output_df.to_excel('output.xlsx', index=False)

# class MainWindow(QMainWindow):
#     def __init__(self):
#         super().__init__()
#         self.setWindowTitle('Inventory Allocation')
#         self.setGeometry(100, 100, 700, 300)
        
#         self.layout = QVBoxLayout()
        
#         # Create widgets for Products File
#         self.products_label = QLabel('Products File:', self)
#         self.products_path = QLineEdit(self)
#         self.products_path.setReadOnly(True)
#         self.products_button = QPushButton('Select', self)
#         self.products_button.clicked.connect(self.select_products_file)
        
#         # Create layout for Products File
#         self.products_layout = QHBoxLayout()
#         self.products_layout.addWidget(self.products_label)
#         self.products_layout.addWidget(self.products_path)
#         self.products_layout.addWidget(self.products_button)
#         self.layout.addLayout(self.products_layout)
        
#         # Create widgets for Incoming File
#         self.incoming_label = QLabel('Incoming File:', self)
#         self.incoming_path = QLineEdit(self)
#         self.incoming_path.setReadOnly(True)
#         self.incoming_button = QPushButton('Select', self)
#         self.incoming_button.clicked.connect(self.select_incoming_file)
        
#         # Create layout for Incoming File
#         self.incoming_layout = QHBoxLayout()
#         self.incoming_layout.addWidget(self.incoming_label)
#         self.incoming_layout.addWidget(self.incoming_path)
#         self.incoming_layout.addWidget(self.incoming_button)
#         self.layout.addLayout(self.incoming_layout)
        
#         # Create widgets for Bins File
#         self.bins_label = QLabel('Bins File:', self)
#         self.bins_path = QLineEdit(self)
#         self.bins_path.setReadOnly(True)
#         self.bins_button = QPushButton('Select', self)
#         self.bins_button.clicked.connect(self.select_bins_file)
        
#         # Create layout for Bins File
#         self.bins_layout = QHBoxLayout()
#         self.bins_layout.addWidget(self.bins_label)
#         self.bins_layout.addWidget(self.bins_path)
#         self.bins_layout.addWidget(self.bins_button)
#         self.layout.addLayout(self.bins_layout)
        
#         # Run Button
#         self.run_button = QPushButton('Run Allocation', self)
#         self.run_button.clicked.connect(self.run_allocation)
#         self.run_button.setEnabled(False)
#         self.layout.addWidget(self.run_button)
        
#         # Status Label
#         self.status_label = QLabel('', self)
#         self.layout.addWidget(self.status_label)
        
#         container = QWidget()
#         container.setLayout(self.layout)
#         self.setCentralWidget(container)
        
#         self.products_file = ''
#         self.incoming_file = ''
#         self.bins_file = ''
        
#         self.show()

#     def select_products_file(self):
#         file_name, _ = QFileDialog.getOpenFileName(self, 'Select Products File', '', 'Excel Files (*.xlsx)')
#         if file_name:
#             self.products_file = file_name
#             self.products_path.setText(file_name)
#             self.check_files_ready()

#     def select_incoming_file(self):
#         file_name, _ = QFileDialog.getOpenFileName(self, 'Select Incoming File', '', 'Excel Files (*.xlsx)')
#         if file_name:
#             self.incoming_file = file_name
#             self.incoming_path.setText(file_name)
#             self.check_files_ready()

#     def select_bins_file(self):
#         file_name, _ = QFileDialog.getOpenFileName(self, 'Select Bins File', '', 'Excel Files (*.xlsx)')
#         if file_name:
#             self.bins_file = file_name
#             self.bins_path.setText(file_name)
#             self.check_files_ready()

#     def check_files_ready(self):
#         if self.products_file and self.incoming_file and self.bins_file:
#             self.run_button.setEnabled(True)
#         else:
#             self.run_button.setEnabled(False)

#     def run_allocation(self):
#         if not (self.products_file and self.incoming_file and self.bins_file):
#             QMessageBox.warning(self, 'Warning', 'Please select all files before running.')
#             return
        
#         # self.run_button.setEnabled(False)
#         # self.status_label.setText('Processing...')
#         # self.thread = WorkerThread(self.products_file, self.incoming_file, self.bins_file)
#         # self.thread.finished.connect(self.on_finished)
#         # self.thread.start()
#         print("product", self.products_file)
#         print("probins", self.incoming_file)
#         print("product", self.bins_file)
#         allocate_inventory(self.products_file, self.incoming_file, self.bins_file)
#         print('complete')

#     def on_finished(self):
#         self.run_button.setEnabled(True)
#         self.status_label.setText('Allocation completed')
# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     window = MainWindow()
#     sys.exit(app.exec_())
