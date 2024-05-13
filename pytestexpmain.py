import subprocess

required_libraries = ["PyQt5", "openpyxl", "fuzzywuzzy", "win10toast"]

try:
    for lib in required_libraries:
        subprocess.check_call([sys.executable, "-m", "pip", "install","--upgrade", lib])
    print("All required libraries installed successfully.")
except Exception as e:
    print(f"An error occurred while installing required libraries: {str(e)}")
import sys
import os
from PyQt5.QtWidgets import QApplication, QAbstractScrollArea,QMainWindow, QAction, QToolBar, QMessageBox, QStatusBar, QTableWidget, QTableWidgetItem, QFileDialog, QInputDialog, QWidget, QVBoxLayout, QPushButton, QLabel, QHBoxLayout, QLineEdit,QListWidgetItem,QListWidget,QDialog,QAbstractItemView
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QTimer, QDate
from win10toast import ToastNotifier
import sqlite3
from openpyxl import load_workbook
from inspect import getsourcefile
from fuzzywuzzy import process as fw_process
from datetime import datetime
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QLineEdit
import module_locator

class NearExpiryDialog(QDialog):
    def __init__(self, near_expiry_products):
        super().__init__()
        print("init near expiry")
        self.setWindowTitle("Near Expiry Products")
        self.resize(800, 600)  # Set the size of the dialog window

        layout = QVBoxLayout()
        self.table_widget = QTableWidget()
        layout.addWidget(self.table_widget)
        self.setLayout(layout)

        self.populate_table(near_expiry_products)

    def populate_table(self, products):
        print("populate table")
        self.table_widget.setColumnCount(6)  # Set the number of columns

        # Set the header labels for each column
        self.table_widget.setHorizontalHeaderLabels(["ID", "Barcode", "Product Name", "Quantity","Vendor", "Expiry Date"])

        # Set the number of rows based on the number of products
        self.table_widget.setRowCount(len(products))

        for row, product in enumerate(products):
            # Populate each cell in the table
            for col, item in enumerate(product):
                # Create QTableWidgetItem for each item
                table_item = QTableWidgetItem(str(item))
                # Set alignment for each cell
                if col == 0:  # ID column
                    table_item.setTextAlignment(Qt.AlignCenter)  # Center alignment
                else:
                    table_item.setTextAlignment(Qt.AlignLeft)  # Left alignment
                self.table_widget.setItem(row, col, table_item)

        # Resize the columns to fit the contents
        self.table_widget.resizeColumnsToContents()

class MainWindow(QMainWindow):
    def __init__(self, pth):
        print("init")
        super().__init__()
        my_path = pth
        
        dir_path = my_path
        db_path = dir_path + "\\P-E-N-S-main\\new_database1.db"
        print(db_path)
        if not os.path.exists(db_path):
            try:
                self.conn = sqlite3.connect(db_path)
                self.cur = self.conn.cursor()
                self.cur.execute('''CREATE TABLE IF NOT EXISTS products (
                                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                                    barcode TEXT,
                                    product_name TEXT,
                                    quantity INTEGER,
                                    vendor TEXT,
                                    expiry_date DATE
                                    )''')
                self.conn.commit()
            except Exception as e:
                print(f"Error in SQLite: {e}")
                print(db_path)
                sys.exit(1)
        else:
            try:
                self.conn = sqlite3.connect(db_path)
                self.cur = self.conn.cursor()
            except Exception as e:
                print(f"Error in SQLite: {e}")
                print(db_path)
                sys.exit(1)

        try:
            self.toaster = ToastNotifier()
        except Exception as e:
            print(f"Error in win10toast: {e}")
            sys.exit(1)
        
        self.near_expiry_products = []

        self.setWindowTitle("Product Expiry Notification System")
        self.setWindowIcon(QIcon(dir_path+"\\P-E-N-S-main\\notification-icon-bell-alarm2.ico"))
        self.setMinimumSize(1024, 768)

        toolbar = QToolBar()
        self.addToolBar(toolbar)

        import_data_action = QAction(QIcon(dir_path+"\\P-E-N-S-main\\add-product.png"), "Import Data", self)
        import_data_action.triggered.connect(self.import_data)
        toolbar.addAction(import_data_action)

        search_action = QAction(QIcon(dir_path+"\\P-E-N-S-main\\s1.png"), "Search", self)
        search_action.triggered.connect(self.search)
        toolbar.addAction(search_action)

        refresh_action = QAction(QIcon(dir_path+"\\P-E-N-S-main\\r3.png"), "Refresh", self)
        refresh_action.triggered.connect(self.refresh_data)  # Change this to self.refresh_data
        toolbar.addAction(refresh_action)
        
        near_expiry_action = QAction(QIcon(dir_path+"\\P-E-N-S-main\\download.png"), "Near Expiry Products", self)
        near_expiry_action.triggered.connect(self.show_near_expiry_list)
        toolbar.addAction(near_expiry_action)

        delete_action = QAction(QIcon(dir_path+"\\P-E-N-S-main\\d1.png"), "Delete", self)
        delete_action.triggered.connect(self.delete_product)
        toolbar.addAction(delete_action)

        statusbar = QStatusBar()
        self.setStatusBar(statusbar)

        self.tableWidget = QTableWidget()
        self.setCentralWidget(self.tableWidget)
        self.tableWidget.setSizeAdjustPolicy(QAbstractScrollArea.AdjustToContents)
        self.tableWidget.setAlternatingRowColors(True)
        self.tableWidget.setColumnCount(6)
        self.tableWidget.setHorizontalHeaderLabels(("ID", "Barcode", "Product Name", "Quantity", "Vendor", "Expiry Date"))
        print("loading data")
        self.load_data()
        self.tableWidget.resizeColumnsToContents()

       # Setup timer for notifications
        self.notification_timer = QTimer(self)
        self.notification_timer.timeout.connect(self.check_expiry)

       # Initial notification after 2 seconds
        QTimer.singleShot(2000, self.check_expiry)

       # Start the timer to check every 30 minutes
        self.notification_timer.start(30 * 60 * 1000)  # 30 minutes in milliseconds

       # Set inline CSS styles for QMainWindow
        self.setStyleSheet("""
    QMainWindow {
        background-color: #f0f0f0; /* Light grey */
    }
    QToolBar {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #6A0DAD, stop:0.5 #A74DF6, stop:0.9 #FAC8FD, stop:1 #A74DF6) !important;
    color: #fff; /* Text color */
    }

    QTableWidget {
        background-color: #fff; /* White */
        border-radius: 10px;
    }

    QTableWidget QHeaderView::section {
        background-color: #3c486a; /* Dark blue-gray */
    color: #fff; /* Text color */
    padding: 8px;
    border-radius: 0;
    }

    QTableWidget::item {
        padding: 10px;
    }

    QAction {
        background-color: #4CAF50; /* Green */
    border: none;
    color: white;
    padding: 12px 24px;
    text-align: center;
    text-decoration: none;
    display: inline-block;
    font-size: 16px;
    margin: 4px 2px;
    transition: transform 0.3s ease; /* Add transition for smooth transform */
    border-radius: 12px;
    }

    QAction:hover {
        background-color: #45a049; /* Dark Green */
    transform: scale(1.2); /* Scale up on hover */
    }

    QAction:active {
    transform: scale(0.9); /* Scale down when clicked */
    }                       

    QLabel {
        color: #333;
    }
""")

    def import_data(self):
        try:
            print("import data")
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Select File", "", "Excel Files (*.xlsx *.xls)")

            if file_path:
                workbook = load_workbook(filename=file_path)
                sheet = workbook.active

                # Define expected column names
                expected_columns = {'Barcode', 'Product Name', 'Quantity', 'Vendor', 'Expiry Date'}

                # Extract actual column names from the header row
                header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
                actual_columns = set(header_row)

                # Match actual column names to expected column names
                column_mapping = {}
                for expected_column in expected_columns:
                    matched_column, _ = fw_process.extractOne(expected_column, actual_columns)
                    if matched_column:
                        column_mapping[expected_column] = header_row.index(matched_column) + 1
                    else:
                        QMessageBox.warning(
                            self, "Warning", f"Could not find a match for '{expected_column}' column.")

                # Check if all expected columns have been matched
                if len(column_mapping) < len(expected_columns):
                    QMessageBox.critical(
                        self, "Error", "Could not find matches for all required columns "
                                        "in the Excel file.")
                    return

                # Proceed with data extraction and import using column_mapping
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    barcode = row[column_mapping.get('Barcode') - 1]
                    product_name = row[column_mapping.get('Product Name') - 1]
                    quantity = row[column_mapping.get('Quantity') - 1]
                    vendor = row[column_mapping.get('Vendor') - 1]
                    expiry_date = row[column_mapping.get('Expiry Date') - 1]  # Assuming expiry_date is fetched as a date object
                    self.cur.execute(
                        "INSERT INTO products (barcode, product_name, quantity, vendor, expiry_date) "
                        "VALUES (?, ?, ?, ?, ?)",
                        (barcode, product_name, quantity, vendor, expiry_date))

                self.conn.commit()
                self.load_data()
                QMessageBox.information(self, "Success", "Data imported successfully.")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")

    def search(self) -> None:
        """Real-time search for products."""
        print("search")
        search_text, ok_pressed = QInputDialog.getText(self, "Search", "Enter text to search:")
        if ok_pressed:
            search_text = search_text.strip().lower()  # Convert search text to lowercase and remove leading/trailing spaces
            if search_text:
                found = False
                for row in range(self.tableWidget.rowCount()):
                    for column in range(self.tableWidget.columnCount()):
                        item = self.tableWidget.item(row, column)
                        if item and search_text in item.text().lower():
                            self.tableWidget.selectRow(row)
                            self.tableWidget.scrollToItem(item, QAbstractItemView.PositionAtTop)
                            found = True
                            break  # Stop searching in this row after finding the first match
                    if found:
                        break  # Stop searching after finding the first match
                if not found:
                    QMessageBox.information(self, "Search Result", "No matching product found.")
            else:
                QMessageBox.information(self, "Search Result", "Please enter a search term.")
    
    def check_expiry(self):
        """Check products with near expiry dates and store them."""
        try:
            print("check expiry")
            current_date = QDate.currentDate()
            expiry_threshold = current_date.addDays(7)

            self.cur.execute("SELECT * FROM products WHERE expiry_date <= ?", (expiry_threshold.toString("yyyy-MM-dd"),))
            self.near_expiry_products = self.cur.fetchall()

            if self.near_expiry_products:
                QMessageBox.information(self, "Near Expiry Alert", "You have some products near expiry.")

        except Exception as e:
            print(f"Error checking expiry: {e}")

    def delete_product(self) -> None:
        """Delete a product from the table and database."""
        print("delete product")
        selected_items = self.tableWidget.selectedItems()
        if selected_items:
            id_to_delete = selected_items[0].text()
            try:
                self.cur.execute("DELETE FROM products WHERE id = ?", (id_to_delete,))
                self.conn.commit()
                self.load_data()
                QMessageBox.information(self, "Success", "Product deleted successfully.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
        else:
            QMessageBox.warning(self, "Warning", "Please select a product to delete.")

    def load_data(self) -> None:
     """Load non-empty data from the database into the table widget."""
     try:
        print("load data")
        self.cur.execute("SELECT * FROM products WHERE barcode IS NOT NULL AND product_name IS NOT NULL AND quantity IS NOT NULL AND vendor IS NOT NULL AND expiry_date IS NOT NULL")
        data = self.cur.fetchall()
        self.tableWidget.setRowCount(0)
        for row_number, row_data in enumerate(data):
            self.tableWidget.insertRow(row_number)
            for column_number, data in enumerate(row_data):
                item = QTableWidgetItem(str(data))
                self.tableWidget.setItem(row_number, column_number, item)
     except Exception as e:
        QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
        print("error: " + e)

    def refresh_data(self) -> None:
        """Clear the table and reload data from the database."""
        print("refresh data")
        confirmation = QMessageBox.question(self, "Confirmation", "Are you sure you want to refresh and clear all data?",
                                             QMessageBox.Yes | QMessageBox.No)
        if confirmation == QMessageBox.Yes:
            try:
                self.cur.execute("DELETE FROM products")
                self.conn.commit()
                self.load_data()
                QMessageBox.information(self, "Success", "Database cleared successfully.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
    
    def show_near_expiry_list(self):
       try:
        print("show near expiry")
        current_date = QDate.currentDate()
        expiry_threshold = current_date.addDays(7)

        self.cur.execute("SELECT * FROM products WHERE expiry_date <= ?", (expiry_threshold.toString("yyyy-MM-dd"),))
        data = self.cur.fetchall()

        # Create a new dialog to display near expiry products
        dialog = NearExpiryDialog(data)
        dialog.exec_()
        
       except Exception as e:
        print(f"Error fetching near expiry products: {e}")

if __name__ == "__main__":
    app = QApplication([])
    file_path = module_locator.module_path()
    main_window = MainWindow(file_path)
    main_window.show()
    try:
        sys.exit(app.exec())
    except Exception as e:
        print(f"An error occurred: {e}")
