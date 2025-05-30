import sys
import pandas as pd
import pymysql
from sqlalchemy import create_engine, types as sqlalchemy_types, text, inspect as sqlalchemy_inspect
import hashlib
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget,
    QPushButton, QLabel, QLineEdit, QTextEdit, QFileDialog, QGroupBox,
    QDateEdit, QTableView, QComboBox, QGridLayout
)
from PyQt5.QtCore import QTimer, Qt, QDateTime, QDate
from PyQt5.QtGui import QFont, QColor, QPalette, QStandardItemModel, QStandardItem
import traceback
from sqlalchemy import exc as sqlalchemy_exc

class ExcelToMySQLApp(QMainWindow):
    """
    Main application class for the Excel to MySQL Import and Query Tool.

    Provides a PyQt5 GUI for users to:
    1. Select an Excel file and import its contents into a MySQL database.
       This includes features like data fingerprinting to avoid duplicate imports
       and automatic creation of 'id' and 'import_time' columns.
    2. Configure database connection parameters for both import and query operations.
    3. Query data from a specified MySQL table based on a date range.
    4. Preview queried data in a table view.
    5. Export queried data to CSV or XLSX format.
    Features a themed UI ("WinXP transparent宝石蓝色") and console logging.
    """
    def __init__(self):
        """
        Initializes the application window, styles, UI components, timer,
        and status variables.
        """
        super().__init__()
        self.setWindowTitle("Excel导入MySQL工具 (宝石天蓝半透明 - UTF8兼容)")
        self.setGeometry(100, 100, 820, 650) # x-pos, y-pos, width, height

        self.set_lightblue_style() # Apply custom QSS styling

        # Setup central widget and main layout
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)
        # self.main_layout is the top-level QVBoxLayout for the central widget.
        # It will contain the main horizontal layout (app_main_hbox) and the status bar.
        self.main_layout = QVBoxLayout(self.central_widget)

        self.init_ui() # Initialize all UI elements

        # Timer for updating the status bar clock
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)  # Update time every 1 second

        # Application state variables
        self.imported_rows_count = 0  # Counter for successfully imported rows in the current session
        self.current_status_text = "等待操作"  # User-facing status message
        self.current_queried_df = None  # Stores the DataFrame from the latest database query

        self.update_status_display() # Update UI elements showing status
        self.log_message("应用程序启动成功 (Application started successfully).")

    def set_lightblue_style(self):
        """
        Sets the application's visual theme using QSS (Qt Style Sheets).
        This method defines the "WinXP transparent宝石蓝色" theme, styling various
        widgets like QMainWindow, QGroupBox, QPushButton, QLineEdit, etc.,
        using gradients, RGBA for transparency, and specific fonts.
        """
        self.setWindowOpacity(0.95) # Slight overall window transparency
        self.setFont(QFont("微软雅黑", 10)) # Default application font

        # Main QSS string
        style_sheet = """
            /* Overall window styling */
            QMainWindow {
                background-color: rgba(230, 240, 255, 0.85); /* Light blue, semi-transparent */
            }
            /* GroupBox styling for distinct sections */
            QGroupBox {
                background-color: rgba(200, 220, 255, 0.7); /* Slightly darker blue for group boxes */
                border: 1px solid rgba(100, 150, 255, 0.9);
                border-radius: 8px;
                margin-top: 10px; /* Space for title */
                padding: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 10px;
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                  stop:0 rgba(150, 200, 255, 1), stop:1 rgba(120, 180, 255, 1));
                color: white;
                border-radius: 4px;
                font-weight: bold;
            }
            /* Standard Button styling with gradient and states */
            QPushButton {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                  stop:0 rgba(100, 180, 255, 1), stop:1 rgba(50, 130, 235, 1));
                color: white;
                border: 1px solid rgba(30, 100, 200, 1); /* Darker border for definition */
                border-radius: 5px;
                padding: 8px 15px; /* Ample padding for clickable area */
                font-size: 10pt;
            }
            QPushButton:hover { /* Style for button when mouse hovers */
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                  stop:0 rgba(120, 200, 255, 1), stop:1 rgba(70, 150, 255, 1));
            }
            QPushButton:pressed { /* Style for button when clicked */
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                  stop:0 rgba(50, 130, 235, 1), stop:1 rgba(30, 100, 200, 1));
            }
            QPushButton:disabled { /* Style for disabled button */
                background-color: rgba(180, 180, 180, 0.7);
                color: rgba(100, 100, 100, 0.8);
                border-color: rgba(150, 150, 150, 0.7);
            }
            /* LineEdit styling for text inputs */
            QLineEdit {
                background-color: rgba(255, 255, 255, 0.9); /* Whiteish, semi-transparent */
                border: 1px solid rgba(100, 150, 255, 0.9);
                border-radius: 4px;
                padding: 6px; /* Comfortable text padding */
                font-size: 10pt;
            }
            QLineEdit:read-only { /* Style for read-only LineEdits */
                background-color: rgba(230, 230, 230, 0.8);
                color: #555; /* Greyed out text */
            }
            /* Label styling */
            QLabel {
                color: #333; /* Dark grey text for good readability */
                padding: 2px;
                font-size: 10pt;
            }
            /* Specific styling for the console output QTextEdit */
            QTextEdit#ConsoleOutput {
                background-color: rgba(255, 255, 255, 0.9);
                border: 1px solid rgba(100, 150, 255, 0.9); /* Thematic border */
                border-radius: 4px;
                padding: 5px; /* Padding for text content */
                color: #333; /* Text color */
                font-family: "Consolas", "Courier New", monospace; /* Monospaced font for console */
            }
            /* QDateEdit styling */
            QDateEdit {
                background-color: rgba(255, 255, 255, 0.9);
                border: 1px solid rgba(100, 150, 255, 0.9);
                border-radius: 4px;
                padding: 5px; 
                font-size: 10pt;
            }
            QDateEdit::drop-down { /* Styling for the dropdown button of QDateEdit */
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px; /* Width of the dropdown button */
                border-left-width: 1px;
                border-left-color: rgba(100, 150, 255, 0.9);
                border-left-style: solid;
                border-top-right-radius: 3px; /* Rounded corners for the button */
                border-bottom-right-radius: 3px;
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                  stop:0 rgba(150, 200, 255, 1), stop:1 rgba(120, 180, 255, 1));
            }
            QDateEdit::down-arrow { /* Arrow icon for QDateEdit dropdown */
                image: url(down_arrow.png); /* Placeholder: requires an actual image file or remove for default arrow */
            }
            /* QComboBox styling */
            QComboBox {
                background-color: rgba(255, 255, 255, 0.9);
                border: 1px solid rgba(100, 150, 255, 0.9);
                border-radius: 4px;
                padding: 5px; /* Padding for text */
                font-size: 10pt;
                min-width: 6em; /* Minimum width to ensure readability */
            }
            QComboBox::drop-down { /* Styling for the dropdown button of QComboBox */
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left-width: 1px;
                border-left-color: rgba(100, 150, 255, 0.9);
                border-left-style: solid;
                border-top-right-radius: 3px;
                border-bottom-right-radius: 3px;
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                  stop:0 rgba(150, 200, 255, 1), stop:1 rgba(120, 180, 255, 1));
            }
            QComboBox::down-arrow { /* Arrow icon for QComboBox dropdown */
                 image: url(down_arrow.png); /* Placeholder: requires an actual image file or remove for default arrow */
            }
            /* QTableView styling for data preview */
            QTableView {
                background-color: rgba(255, 255, 255, 0.9);
                border: 1px solid rgba(100, 150, 255, 0.9);
                border-radius: 4px;
                gridline-color: rgba(150, 200, 255, 0.8); /* Thematic grid lines */
                font-size: 9pt; /* Slightly smaller font for table data */
            }
            /* QHeaderView styling for QTableView headers */
            QHeaderView::section {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                  stop:0 rgba(150, 200, 255, 1), stop:1 rgba(120, 180, 255, 1));
                color: white; /* Header text color */
                padding: 4px; /* Header padding */
                border: 1px solid rgba(100, 150, 255, 0.9); /* Border for header sections */
                font-weight: bold;
            }
        """
        self.setStyleSheet(style_sheet)

    def init_ui(self):
        """
        Initializes and arranges the main UI components.

        This method sets up the primary two-column layout:
        - Left Panel: Contains controls for Excel import and database query/export settings.
        - Right Panel: Contains the data preview table and the console output.
        The status bar is placed at the bottom of the window.
        """
        # Overall application layout (vertical: app_main_hbox on top, status_bar_label at bottom)
        # self.main_layout is already a QVBoxLayout set in __init__

        app_main_hbox = QHBoxLayout() # Main horizontal layout for the two panels

        # --- Left Panel: Controls and Info ---
        left_panel_widget = QWidget()
        left_panel_vbox = QVBoxLayout(left_panel_widget)
        
        # Create and add functional groups to the left panel
        self.control_group = self.create_control_group() # Excel selection and import button
        left_panel_vbox.addWidget(self.control_group)

        self.db_info_group = self.create_db_info_group() # DB credentials for import
        left_panel_vbox.addWidget(self.db_info_group)

        self.table_info_group = self.create_table_info_group() # Table name and import status
        left_panel_vbox.addWidget(self.table_info_group)
        
        self.query_export_settings_group = self.create_query_export_settings_group() # Query/export controls
        left_panel_vbox.addWidget(self.query_export_settings_group)

        left_panel_vbox.addStretch(1) # Pushes all group boxes to the top of the left panel

        # --- Right Panel: Data Preview and Console ---
        right_panel_widget = QWidget()
        right_panel_vbox = QVBoxLayout(right_panel_widget)

        # Group for Data Query Preview Table
        query_preview_group = QGroupBox("数据查询结果预览 (Query Results Preview)") 
        query_preview_layout = QVBoxLayout(query_preview_group)
        self.query_data_preview_table = QTableView() # Table to display queried data
        query_preview_layout.addWidget(self.query_data_preview_table)
        # Give table view more stretch factor compared to console
        right_panel_vbox.addWidget(query_preview_group, 2) 

        # Group for Console Output
        console_group_box = QGroupBox("控制台输出 (Console Output)")
        self.console_widget = self.create_console_output_widget() # QTextEdit for logs
        console_layout = QVBoxLayout(console_group_box)
        console_layout.addWidget(self.console_widget)
        right_panel_vbox.addWidget(console_group_box, 1) # Console gets less stretch factor

        # Add left and right panels to the main horizontal layout
        # Change stretch factors to 2 for left_panel_widget and 3 for right_panel_widget
        app_main_hbox.addWidget(left_panel_widget) # Add widget first
        app_main_hbox.setStretchFactor(left_panel_widget, 2) # Then set stretch factor
        app_main_hbox.addWidget(right_panel_widget) # Add widget first
        app_main_hbox.setStretchFactor(right_panel_widget, 3) # Then set stretch factor


        # Add the two-column layout (app_main_hbox) to the main vertical layout of the central widget
        self.main_layout.addLayout(app_main_hbox)
        
        # Status bar at the bottom
        self.status_bar_label = QLabel("状态栏初始化... (Status Bar Initialized...)") 
        self.main_layout.addWidget(self.status_bar_label)

        self.update_time() # Initialize time display

    def create_query_export_settings_group(self):
        """
        Creates the QGroupBox for "数据库查询与导出" (Database Query & Export) settings.
        
        This group contains input fields for database connection (host, user, password, DB name),
        table name, date range for querying, and controls for initiating the query,
        selecting export path/format, and downloading data.
        
        Returns:
            QGroupBox: The configured group box with all UI elements.
        """
        group_box = QGroupBox("数据库查询与导出 (Database Query & Export)")
        layout = QGridLayout() # Using QGridLayout for a structured label-field layout
        # Configure column stretch: column 1 and 3 (widgets) can expand more than 0 and 2 (labels)
        layout.setColumnStretch(1, 1)
        layout.setColumnStretch(3, 1)

        # Row 0: DB Host and DB Name
        layout.addWidget(QLabel("主机 (Host):"), 0, 0)
        self.query_db_host_edit = QLineEdit("localhost")
        layout.addWidget(self.query_db_host_edit, 0, 1)

        layout.addWidget(QLabel("数据库 (Database):"), 0, 2)
        self.query_db_name_edit = QLineEdit("michentestdb2")
        layout.addWidget(self.query_db_name_edit, 0, 3)

        # Row 1: DB User and Table Name
        layout.addWidget(QLabel("用户 (User):"), 1, 0)
        self.query_db_user_edit = QLineEdit("root")
        layout.addWidget(self.query_db_user_edit, 1, 1)

        layout.addWidget(QLabel("表名 (Table):"), 1, 2)
        self.query_table_name_edit = QLineEdit("report_data")
        layout.addWidget(self.query_table_name_edit, 1, 3)
        
        # Row 2: DB Password and Start Date
        layout.addWidget(QLabel("密码 (Password):"), 2, 0)
        self.query_db_password_edit = QLineEdit()
        self.query_db_password_edit.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.query_db_password_edit, 2, 1)

        layout.addWidget(QLabel("开始日期 (Start Date):"), 2, 2)
        self.query_start_date_edit = QDateEdit(QDate.currentDate())
        self.query_start_date_edit.setCalendarPopup(True)
        self.query_start_date_edit.setDisplayFormat("yyyy-MM-dd")
        layout.addWidget(self.query_start_date_edit, 2, 3)

        # Row 3: (empty) and End Date
        layout.addWidget(QLabel("结束日期 (End Date):"), 3, 2)
        self.query_end_date_edit = QDateEdit(QDate.currentDate())
        self.query_end_date_edit.setCalendarPopup(True)
        self.query_end_date_edit.setDisplayFormat("yyyy-MM-dd")
        layout.addWidget(self.query_end_date_edit, 3, 3)
        
        # Row 4: Query Button (spans all 4 columns)
        self.btn_query_data = QPushButton("查询数据 (Query Data)")
        self.btn_query_data.clicked.connect(self.query_data_from_db)
        layout.addWidget(self.btn_query_data, 4, 0, 1, 4) 

        # Row 5: Export File Path (spans 3 columns) and Browse button
        layout.addWidget(QLabel("导出路径 (Export Path):"), 5, 0)
        self.query_export_path_edit = QLineEdit()
        self.query_export_path_edit.setPlaceholderText("选择或输入导出文件路径... (Select or input export file path...)")
        layout.addWidget(self.query_export_path_edit, 5, 1, 1, 2) # Spans columns 1 and 2
        self.btn_browse_export_path = QPushButton("浏览... (Browse...)")
        self.btn_browse_export_path.clicked.connect(self.select_export_file_path)
        layout.addWidget(self.btn_browse_export_path, 5, 3) # Column 3

        # Row 6: Export Format (spans 3 columns for the ComboBox)
        layout.addWidget(QLabel("导出格式 (Export Format):"), 6, 0)
        self.query_export_format_combo = QComboBox()
        self.query_export_format_combo.addItems(["CSV", "XLSX"])
        layout.addWidget(self.query_export_format_combo, 6, 1, 1, 3)

        # Row 7: Download Button (spans all 4 columns)
        self.btn_download_data = QPushButton("下载数据 (Download Data)")
        self.btn_download_data.clicked.connect(self.download_queried_data)
        self.btn_download_data.setEnabled(False) 
        layout.addWidget(self.btn_download_data, 7, 0, 1, 4)
        
        group_box.setLayout(layout)
        return group_box


    def create_control_group(self):
        """
        Creates the QGroupBox for "控制面板" (Control Panel).
        This group contains widgets for selecting an Excel file and initiating the import.
        """
        group_box = QGroupBox("控制面板 (Control Panel)")
        layout = QVBoxLayout() # Simple vertical layout

        # Button to trigger file selection dialog
        self.btn_select = QPushButton("选择Excel文件 (Select Excel File)")
        self.btn_select.clicked.connect(self.select_excel_file)
        layout.addWidget(self.btn_select)

        # Read-only LineEdit to display selected file path
        self.file_path_edit = QLineEdit() 
        self.file_path_edit.setPlaceholderText("未选择文件 (No file selected)")
        self.file_path_edit.setReadOnly(True)
        layout.addWidget(self.file_path_edit)

        # Button to start the import process
        self.btn_import = QPushButton("导入数据库 (Import to Database)")
        self.btn_import.clicked.connect(self.import_to_mysql)
        self.btn_import.setEnabled(False) # Initially disabled until a file is selected
        layout.addWidget(self.btn_import)

        group_box.setLayout(layout)
        return group_box

    def create_db_info_group(self):
        """
        Creates the QGroupBox for "数据库信息" (Database Information).
        This group contains input fields for database connection details (name, user, password)
        primarily used for the Excel import functionality.
        """
        group_box = QGroupBox("数据库信息 (Database Information - For Import)")
        # Using QGridLayout for a more compact label-field layout
        layout = QGridLayout()
        layout.setColumnStretch(1, 1) # Allow the QLineEdit column to expand

        # Database Name input
        layout.addWidget(QLabel("数据库名称 (DB Name):"), 0, 0)
        self.db_name_edit = QLineEdit("michentestdb2") # Default DB name for import
        layout.addWidget(self.db_name_edit, 0, 1)

        # Database User input
        layout.addWidget(QLabel("用户名 (User):"), 1, 0)
        self.db_user_edit = QLineEdit("root") # Default user for import
        layout.addWidget(self.db_user_edit, 1, 1)

        # Database Password input
        layout.addWidget(QLabel("密码 (Password):"), 2, 0)
        self.db_password_edit = QLineEdit("123") # Default password for import
        self.db_password_edit.setEchoMode(QLineEdit.Password) # Mask password
        layout.addWidget(self.db_password_edit, 2, 1)
        
        group_box.setLayout(layout)
        return group_box

    def create_table_info_group(self):
        """
        Creates the QGroupBox for "表信息" (Table Information).
        This group displays information related to the import process, such as
        the target table name, number of rows imported, and current operation status.
        Uses QGridLayout for better alignment and compactness.
        """
        group_box = QGroupBox("表信息 (Table Information - For Import)")
        # Using QGridLayout for a more compact label-field layout
        layout = QGridLayout()
        layout.setColumnStretch(1, 1) # Allow the QLineEdit column to expand

        # Target Table Name input (for import)
        layout.addWidget(QLabel("表名称 (Table Name):"), 0, 0)
        self.table_name_edit = QLineEdit("report_data") # Default table name for import
        layout.addWidget(self.table_name_edit, 0, 1)

        # Display for Number of Rows Imported (read-only)
        layout.addWidget(QLabel("导入行数 (Rows Imported):"), 1, 0)
        self.rows_imported_edit = QLineEdit("0") 
        self.rows_imported_edit.setReadOnly(True)
        layout.addWidget(self.rows_imported_edit, 1, 1)
        
        # Display for Current Operation Status (read-only)
        layout.addWidget(QLabel("状态 (Status):"), 2, 0)
        self.operation_status_edit = QLineEdit("等待操作 (Waiting for operation)") 
        self.operation_status_edit.setReadOnly(True)
        layout.addWidget(self.operation_status_edit, 2, 1)

        group_box.setLayout(layout)
        return group_box

    def create_console_output_widget(self):
        """
        Creates the QTextEdit widget used for console logging.
        It's set to read-only and given an object name for specific QSS styling.
        """
        self.console_output = QTextEdit()
        self.console_output.setObjectName("ConsoleOutput") # For QSS styling
        self.console_output.setReadOnly(True) # User cannot type into console
        return self.console_output

    def select_excel_file(self):
        """
        Opens a QFileDialog to allow the user to select an Excel file (*.xlsx, *.xls).
        Updates the file path display and enables the import button if a file is selected.
        """
        options = QFileDialog.Options()
        # options |= QFileDialog.DontUseNativeDialog # Uncomment to use Qt's dialog over native
        file_path, _ = QFileDialog.getOpenFileName(self, 
                                                   "选择Excel文件 (Select Excel File)", 
                                                   "", # Default directory
                                                   "Excel Files (*.xlsx *.xls);;All Files (*)", 
                                                   options=options)
        if file_path:
            self.file_path_edit.setText(file_path)
            self.btn_import.setEnabled(True) # Enable import button
            self.log_message(f"已选择文件 (File selected): {file_path}")
            self.current_status_text = "文件已选择，待导入 (File selected, ready to import)"
            self.update_status_display()
        else:
            self.log_message("未选择文件 (No file selected).")
            self.btn_import.setEnabled(False) # Keep import button disabled
            self.current_status_text = "操作取消 (Operation cancelled)"
            self.update_status_display()

    def generate_data_fingerprint(self, dataframe):
        """
        Generates an MD5 fingerprint for a given DataFrame.
        The fingerprint is based on a string representation of the DataFrame's head (first 5 rows)
        and its column names. This helps in identifying if an identical dataset (at least the beginning)
        has been imported before.

        Args:
            dataframe (pd.DataFrame): The DataFrame to fingerprint.

        Returns:
            str: The hex digest of the MD5 hash.
        """
        # Using a sample of the dataframe to generate fingerprint
        # This includes the first 5 rows and all column names
        sample_data = dataframe.head().to_string() + "".join(dataframe.columns)
        return hashlib.md5(sample_data.encode('utf-8')).hexdigest()

    def import_to_mysql(self):
        """
        Handles the entire Excel to MySQL import process.
        
        Steps:
        1. Retrieves Excel file path and database/table details from UI.
        2. Validates inputs.
        3. Reads data from Excel using pandas.
        4. Connects to MySQL using pymysql for DDL operations (CREATE DATABASE/TABLE IF NOT EXISTS).
        5. Generates a data fingerprint for the Excel data.
        6. Connects to MySQL using SQLAlchemy for data operations.
        7. Checks if the table exists and if the data fingerprint is already present to avoid duplicates.
        8. Appends the DataFrame to the MySQL table using SQLAlchemy's `to_sql`.
           - Includes the data fingerprint as a new column.
           - Specifies data types for columns, especially for strings and potential large text.
        9. Ensures 'id' (auto-increment primary key) and 'import_time' (timestamp) columns exist,
           creating them if necessary using ALTER TABLE.
        10. Logs success or errors to the console and updates status display.
        11. Handles various exceptions (file errors, database errors, etc.).
        12. Ensures database connections are closed/disposed in a finally block.
        """
        excel_file_path = self.file_path_edit.text()
        if not excel_file_path:
            self.log_message("错误: 未选择Excel文件 (Error: No Excel file selected).")
            return

        # Disable import button during operation and update status
        self.btn_import.setEnabled(False)
        self.current_status_text = "正在导入... (Importing...)"
        self.update_status_display()
        self.log_message(f"开始从 {excel_file_path} 导入数据 (Starting data import from {excel_file_path})...")
        QApplication.processEvents() # Keep UI responsive

        # Retrieve database and table details from UI (for import section)
        db_name = self.db_name_edit.text().strip()
        db_user = self.db_user_edit.text().strip()
        db_password = self.db_password_edit.text() # Password not stripped
        table_name = self.table_name_edit.text().strip()
        db_host = 'localhost' # Hardcoded as per original requirement, can be made a field

        # Input validation for DB details
        if not all([db_name, db_user, table_name]): # Password can be empty for some MySQL setups
            self.log_message("错误: 数据库名称、用户名和表名称不能为空。(Error: DB Name, User, and Table Name cannot be empty).")
            self.current_status_text = "导入失败 (Import failed)"
            self.update_status_display()
            self.btn_import.setEnabled(True) # Re-enable import button
            return
        
        pymysql_conn = None # For DDL operations
        sqlalchemy_engine = None # For data operations (to_sql)

        try:
            # --- 1. Read Excel file ---
            self.log_message("正在读取Excel文件 (Reading Excel file)...")
            QApplication.processEvents()
            df = pd.read_excel(excel_file_path)
            if df.empty:
                self.log_message("警告: Excel文件为空，没有数据可导入。(Warning: Excel file is empty).")
                self.current_status_text = "文件为空 (File empty)"
                self.update_status_display()
                self.btn_import.setEnabled(True)
                return
            self.log_message(f"成功读取 {df.shape[0]} 行, {df.shape[1]} 列数据。(Successfully read {df.shape[0]} rows, {df.shape[1]} columns).")

            # --- 2. Database/Table Creation (using pymysql for DDL) ---
            self.log_message("正在连接数据库并准备表 (Connecting to DB and preparing table)...")
            QApplication.processEvents()
            # Connect without specifying db_name first to create it if it doesn't exist
            pymysql_conn = pymysql.connect(host=db_host, user=db_user, password=db_password, charset='utf8')
            with pymysql_conn.cursor() as cursor:
                cursor.execute(f"CREATE DATABASE IF NOT EXISTS `{db_name}` CHARACTER SET utf8 COLLATE utf8_general_ci")
                self.log_message(f"数据库 '{db_name}' 已确保存在。(Database '{db_name}' ensured).")
                cursor.execute(f"USE `{db_name}`") # Switch to the target database
            pymysql_conn.commit() 
            
            # --- 3. Data Fingerprint & Duplication Check ---
            fingerprint = self.generate_data_fingerprint(df)
            self.log_message(f"数据指纹 (Data fingerprint): {fingerprint}")

            # Create SQLAlchemy engine for data insertion and inspection
            engine_url = f"mysql+pymysql://{db_user}:{db_password}@{db_host}/{db_name}?charset=utf8"
            sqlalchemy_engine = create_engine(engine_url)
            
            table_exists = sqlalchemy_inspect(sqlalchemy_engine).has_table(table_name)
            if table_exists:
                self.log_message(f"表 '{table_name}' 已存在。检查重复数据... (Table '{table_name}' exists. Checking for duplicates...)")
                try:
                    # Check if 'data_fingerprint' column exists in the table
                    columns_in_db = [col['name'] for col in sqlalchemy_inspect(sqlalchemy_engine).get_columns(table_name)]
                    if 'data_fingerprint' in columns_in_db:
                        # Check if this specific fingerprint already exists
                        with sqlalchemy_engine.connect() as connection:
                            result = connection.execute(text(f"SELECT 1 FROM `{table_name}` WHERE `data_fingerprint` = :fp LIMIT 1"), {'fp': fingerprint})
                            if result.scalar_one_or_none(): # If a row with this fingerprint is found
                                self.log_message(f"数据指纹 {fingerprint} 已存在于表 '{table_name}'。跳过导入以避免重复。(Fingerprint {fingerprint} already exists. Skipping import.)")
                                self.current_status_text = "数据重复，跳过 (Duplicate data, skipped)"
                                self.update_status_display()
                                self.btn_import.setEnabled(True)
                                return # Stop import process
                    else:
                        self.log_message("警告: 表已存在但无 'data_fingerprint' 列。无法精确去重，将追加数据。(Warning: Table exists but no 'data_fingerprint' column. Appending data.)")
                except Exception as e_dup_check:
                    self.log_message(f"检查重复数据时出错 (Error checking duplicates): {e_dup_check}。将尝试追加数据。(Appending data.)")


            # --- 4. Import Data (using SQLAlchemy's to_sql) ---
            self.log_message("正在导入数据到MySQL (Importing data to MySQL)...")
            QApplication.processEvents()
            df['data_fingerprint'] = fingerprint # Add fingerprint column to DataFrame
            
            # Define SQLAlchemy types for DataFrame columns to ensure correct table creation/insertion
            dtypedict = {'data_fingerprint': sqlalchemy_types.VARCHAR(32)} # MD5 hash is 32 chars
            for col in df.columns:
                if col != 'data_fingerprint': # Skip already defined fingerprint column
                    if pd.api.types.is_datetime64_any_dtype(df[col]):
                        dtypedict[col] = sqlalchemy_types.DateTime()
                    elif pd.api.types.is_integer_dtype(df[col]):
                        if df[col].isnull().any(): # If integers have NaNs, pandas loads as float
                             dtypedict[col] = sqlalchemy_types.Float() 
                        else: # Pure integers
                             dtypedict[col] = sqlalchemy_types.BigInteger() # Use BigInteger for safety
                    elif pd.api.types.is_float_dtype(df[col]):
                        dtypedict[col] = sqlalchemy_types.Float()
                    elif pd.api.types.is_string_dtype(df[col]):
                        # Estimate max length for VARCHAR or use TEXT for very long strings
                        max_len = df[col].str.len().max()
                        if pd.isna(max_len) or max_len > 1000 : # Heuristic for using TEXT (e.g., > 1000 chars)
                            dtypedict[col] = sqlalchemy_types.TEXT()
                        else:
                            # Ensure length is at least 1 if max_len is 0 (e.g. all empty strings)
                            dtypedict[col] = sqlalchemy_types.VARCHAR(int(max_len) if max_len > 0 else 1) 
            
            # Perform the import using df.to_sql
            df.to_sql(name=table_name, 
                      con=sqlalchemy_engine, 
                      if_exists='append',  # Append data if table exists
                      index=False,         # Do not write DataFrame index as a column
                      chunksize=1000,      # Import in chunks for large DataFrames
                      dtype=dtypedict)     # Specify column types
            
            self.imported_rows_count = df.shape[0]
            self.log_message(f"成功导入 {self.imported_rows_count} 行数据到表 '{table_name}'。(Successfully imported {self.imported_rows_count} rows).")

            # --- 5. Add 'id' and 'import_time' columns if they don't exist ---
            # These are common utility columns.
            with sqlalchemy_engine.connect() as connection:
                trans = connection.begin() # Start a transaction for DDL changes
                try:
                    current_columns_meta = sqlalchemy_inspect(sqlalchemy_engine).get_columns(table_name, schema=db_name)
                    current_columns_names = [col['name'] for col in current_columns_meta]
                    
                    if 'id' not in current_columns_names:
                        self.log_message(f"正在为表 '{table_name}' 添加 'id' 列 (Adding 'id' column to '{table_name}')...")
                        # Add 'id' as auto-incrementing primary key, placed first
                        connection.execute(text(f"ALTER TABLE `{db_name}`.`{table_name}` ADD COLUMN `id` INT AUTO_INCREMENT PRIMARY KEY FIRST;"))
                    
                    if 'import_time' not in current_columns_names:
                        self.log_message(f"正在为表 '{table_name}' 添加 'import_time' 列 (Adding 'import_time' column to '{table_name}')...")
                        # Add 'import_time' with default current timestamp
                        connection.execute(text(f"ALTER TABLE `{db_name}`.`{table_name}` ADD COLUMN `import_time` TIMESTAMP DEFAULT CURRENT_TIMESTAMP;"))
                    
                    trans.commit() # Commit DDL changes
                    self.log_message("必要的 'id' 和 'import_time' 列已确保存在。(Ensured 'id' and 'import_time' columns exist).")
                except Exception as alter_e:
                    trans.rollback() # Rollback on error
                    self.log_message(f"添加辅助列时出错 (Error adding auxiliary columns): {alter_e}")

            self.current_status_text = "导入成功 (Import successful)"

        except pd.errors.EmptyDataError: # Specific pandas error for empty Excel
            self.log_message("错误: Excel文件为空或格式不受支持。(Error: Excel file empty or unsupported format).")
            self.current_status_text = "导入失败: 文件空 (Import failed: File empty)"
        except pymysql.Error as e_pymysql: # Errors from pymysql (DDL part)
            self.log_message(f"数据库连接或操作错误 (pymysql): {e_pymysql} (DB connection/operation error (pymysql)).")
            self.current_status_text = "导入失败: DB错误 (Import failed: DB error)"
        except sqlalchemy_exc.SQLAlchemyError as e_sqlalchemy: # Errors from SQLAlchemy (Data part)
            self.log_message(f"数据库操作错误 (SQLAlchemy): {e_sqlalchemy} (DB operation error (SQLAlchemy)).")
            self.current_status_text = "导入失败: DB错误 (Import failed: DB error)"
            self.log_message(traceback.format_exc()) # Log full traceback for SQLAlchemy errors
        except Exception as e_general: # Catch-all for other errors
            self.log_message(f"导入过程中发生未知错误: {e_general} (Unknown error during import).")
            self.current_status_text = "导入失败: 未知错误 (Import failed: Unknown error)"
            self.log_message(traceback.format_exc()) # Log full traceback
        finally:
            # Ensure connections are closed/disposed
            if pymysql_conn and pymysql_conn.open:
                pymysql_conn.close()
                self.log_message("Pymysql连接已关闭。(Pymysql connection closed).")
            if sqlalchemy_engine:
                sqlalchemy_engine.dispose()
                self.log_message("SQLAlchemy引擎已释放。(SQLAlchemy engine disposed).")
            
            self.btn_import.setEnabled(True) # Re-enable import button
            self.update_status_display() # Update status on UI
            self.log_message("导入操作完成。(Import operation finished).")
            QApplication.processEvents()

    def log_message(self, message):
        """
        Appends a timestamped message to the console output QTextEdit.
        Ensures the console automatically scrolls to the latest message.
        
        Args:
            message (str): The message to log.
        """
        timestamp = QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss")
        self.console_output.append(f"[{timestamp}] {message}")
        self.console_output.ensureCursorVisible() # Auto-scroll to the bottom

    def update_time(self):
        """
        Updates the status bar label with the current system time.
        Called periodically by self.timer.
        """
        current_time = QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss")
        self.status_bar_label.setText(f"当前时间 (Current Time): {current_time}")

    def update_status_display(self): 
        """
        Updates the QLineEdit fields that display the number of imported rows
        and the current operation status text.
        """
        self.rows_imported_edit.setText(str(self.imported_rows_count))
        self.operation_status_edit.setText(self.current_status_text)

    def query_data_from_db(self):
        """
        Queries data from the specified MySQL database and table based on user inputs.
        
        Steps:
        1. Retrieves database connection details, table name, and date range from UI.
        2. Validates inputs (required fields, date order).
        3. Constructs an SQL query to select data within the date range.
           (Assumes a date column named 'Date' in the target table).
        4. Connects to the database using SQLAlchemy.
        5. Fetches data into a pandas DataFrame using `pd.read_sql_query()`.
        6. If data is found:
           - Logs DataFrame structure and head for diagnostics.
           - Populates `self.query_data_preview_table` (QTableView) with the data.
           - Enables the "Download Data" button.
           - Stores the DataFrame in `self.current_queried_df`.
        7. If no data or an error occurs, clears the table, disables download, and resets `self.current_queried_df`.
        8. Handles various exceptions (DB connection, SQL errors, etc.).
        9. Ensures the SQLAlchemy engine is disposed of.
        """
        self.log_message("Initiating data query from database...")
        QApplication.processEvents() # Keep UI responsive

        # Retrieve query parameters from UI fields
        host = self.query_db_host_edit.text().strip()
        user = self.query_db_user_edit.text().strip()
        password = self.query_db_password_edit.text() # Password not stripped
        db_name = self.query_db_name_edit.text().strip()
        table_name = self.query_table_name_edit.text().strip() 

        # Input validation for query parameters
        if not all([host, user, db_name, table_name]): # Password can be empty
            self.log_message("错误: 主机, 用户名, 数据库名称和表名称为必填项。(Error: Host, User, DB Name, and Table Name are required).")
            self.query_data_preview_table.setModel(None) # Clear table
            self.btn_download_data.setEnabled(False)
            self.current_queried_df = None
            return

        start_date_q = self.query_start_date_edit.date() # QDate object
        end_date_q = self.query_end_date_edit.date()   # QDate object

        if start_date_q > end_date_q:
            self.log_message("错误: 开始日期不能晚于结束日期。(Error: Start date cannot be after end date).")
            self.query_data_preview_table.setModel(None)
            self.btn_download_data.setEnabled(False)
            self.current_queried_df = None
            return

        start_date_str = start_date_q.toString("yyyy-MM-dd")
        end_date_str = end_date_q.toString("yyyy-MM-dd") 

        self.log_message(f"Querying table `{table_name}` in database `{db_name}` from {start_date_str} to {end_date_str}.")
        
        # Construct SQL query. Assumes a date column named 'Date'.
        # This might need to be configurable if the date column name varies.
        # The DATE() function in SQL extracts the date part from a datetime column.
        sql_query = f"SELECT * FROM `{table_name}` WHERE DATE(`Date`) >= '{start_date_str}' AND DATE(`Date`) <= '{end_date_str}'"
        self.log_message(f"Executing SQL: {sql_query}")

        engine = None # Initialize engine to None for finally block
        try:
            # --- Database Connection (SQLAlchemy) ---
            self.log_message(f"Connecting to database '{db_name}' on host '{host}' with user '{user}'...")
            QApplication.processEvents()
            engine_url = f"mysql+pymysql://{user}:{password}@{host}/{db_name}?charset=utf8"
            engine = create_engine(engine_url) # Create SQLAlchemy engine
            
            self.log_message("Database connection successful via SQLAlchemy engine.")
            self.log_message("Fetching data...")
            QApplication.processEvents()

            # --- Data Fetching (pandas) ---
            df = pd.read_sql_query(sql_query, engine)

            # --- Process Fetched DataFrame ---
            if df.empty:
                self.log_message("未查询到数据 (No data found for the given criteria).")
                self.query_data_preview_table.setModel(None) # Clear previous results
                self.btn_download_data.setEnabled(False)
                self.current_queried_df = None
            else:
                self.log_message(f"成功获取 {len(df)} 行数据。正在填充表格...(Successfully fetched {len(df)} rows. Populating table...)")
                # Log some info about the DataFrame for diagnostics
                self.log_message(f"DataFrame columns: {list(df.columns)}")
                self.log_message(f"First few rows of DataFrame (up to 5):\n{df.head().to_string()}")
                
                # Populate QTableView with the DataFrame data
                model = QStandardItemModel(df.shape[0], df.shape[1]) # Rows, Columns
                model.setHorizontalHeaderLabels(list(df.columns)) # Set column headers

                for i in range(df.shape[0]): # Iterate rows
                    for j in range(df.shape[1]): # Iterate columns
                        item_value = str(df.iloc[i, j]) # Get cell value as string
                        item = QStandardItem(item_value)
                        item.setEditable(False) # Make cells in table view non-editable
                        model.setItem(i, j, item) # Set item in model
                
                self.query_data_preview_table.setModel(model) # Set model to table view
                self.query_data_preview_table.resizeColumnsToContents() # Adjust column widths
                self.btn_download_data.setEnabled(True) # Enable download button
                self.current_queried_df = df # Store DataFrame for potential download
                self.log_message("数据已成功加载到预览表格。(Data successfully loaded into preview table).")

        except sqlalchemy_exc.OperationalError as e_op_sql: # Specific error for DB connection issues
            self.log_message(f"数据库连接失败 (SQLAlchemy OperationalError): {e_op_sql} (DB Connection Failed).")
            self.query_data_preview_table.setModel(None)
            self.btn_download_data.setEnabled(False)
            self.current_queried_df = None
        except pymysql.Error as e_pymysql_query: # Errors from underlying pymysql driver during query
            self.log_message(f"数据库错误 (pymysql): {e_pymysql_query} (DB Error (pymysql)).")
            self.query_data_preview_table.setModel(None)
            self.btn_download_data.setEnabled(False)
            self.current_queried_df = None
        except pd.errors.DatabaseError as e_pd_db: # Errors from pandas during read_sql_query
             self.log_message(f"Pandas数据库读取错误: {e_pd_db}. 可能查询或表名有问题。(Pandas DB Read Error. Possible issue with query or table name).")
             self.query_data_preview_table.setModel(None)
             self.btn_download_data.setEnabled(False)
             self.current_queried_df = None
        except Exception as e_generic_query: # Catch-all for other errors
            self.log_message(f"查询数据时发生未知错误: {e_generic_query} (Unknown error during data query).")
            self.log_message(traceback.format_exc()) # Log full traceback
            self.query_data_preview_table.setModel(None)
            self.btn_download_data.setEnabled(False)
            self.current_queried_df = None
        finally:
            # --- Resource Cleanup ---
            if engine:
                engine.dispose() # Release database connection pool
                self.log_message("SQLAlchemy engine disposed.")
            QApplication.processEvents() # Final UI update

    def select_export_file_path(self):
        """
        Opens a QFileDialog to allow the user to select a file path and name
        for exporting the queried data. The file type (CSV/XLSX) is determined
        by the current selection in the export format ComboBox.
        """
        export_format = self.query_export_format_combo.currentText() # "CSV" or "XLSX"
        default_filename = "exported_data" # Base for suggested filename
        filter_string = "" # File dialog filter

        if export_format == "CSV":
            default_filename += ".csv"
            filter_string = "CSV files (*.csv);;All Files (*)"
        elif export_format == "XLSX":
            default_filename += ".xlsx"
            filter_string = "Excel files (*.xlsx);;All Files (*)"
        else: # Fallback, though current UI restricts to CSV/XLSX
            filter_string = "All Files (*)"

        options = QFileDialog.Options()
        # options |= QFileDialog.DontUseNativeDialog # Uncomment for Qt's own dialog
        file_name, _ = QFileDialog.getSaveFileName(self, 
                                                   "保存文件 (Save File)", 
                                                   default_filename, # Suggested filename
                                                   filter_string, 
                                                   options=options)
        
        if file_name: # If a file name was chosen (dialog not cancelled)
            # QFileDialog might not enforce the extension from the filter on all platforms/settings.
            # Ensure the filename has the correct extension corresponding to the selected format.
            if export_format == "CSV" and not file_name.lower().endswith(".csv"):
                file_name += ".csv"
            elif export_format == "XLSX" and not file_name.lower().endswith(".xlsx"):
                file_name += ".xlsx"
            
            self.query_export_path_edit.setText(file_name) # Update the QLineEdit
            self.log_message(f"导出路径已选择 (Export path selected): {file_name}")

    def download_queried_data(self):
        """
        Downloads the currently stored queried data (self.current_queried_df)
        to a file in the format specified by the user (CSV or XLSX).
        The file path is taken from `self.query_export_path_edit`.
        """
        self.log_message("Initiating data download...")
        QApplication.processEvents() # Keep UI responsive

        # Validate that there is data to export
        if self.current_queried_df is None or self.current_queried_df.empty:
            self.log_message("没有可供导出的数据 (No data available for export).")
            return

        # Validate that an export path has been specified
        file_path = self.query_export_path_edit.text().strip()
        if not file_path:
            self.log_message("请先选择或输入导出文件路径 (Please select or enter an export file path).")
            # Optionally, could trigger self.select_export_file_path() here if desired.
            return

        export_format = self.query_export_format_combo.currentText() # "CSV" or "XLSX"
        self.log_message(f"Attempting to export data as {export_format} to: {file_path}")

        try:
            # --- Perform Export using pandas ---
            if export_format == "CSV":
                # Use utf-8-sig for CSV to ensure BOM for Excel compatibility with UTF-8 chars
                self.current_queried_df.to_csv(file_path, index=False, encoding='utf-8-sig')
            elif export_format == "XLSX":
                # Requires 'openpyxl' engine for .xlsx format.
                # User might need to install it: pip install openpyxl
                self.current_queried_df.to_excel(file_path, index=False, engine='openpyxl') 
            else: # Should not be reached with current ComboBox setup
                self.log_message(f"不支持的导出格式: {export_format} (Unsupported export format).")
                return
            
            self.log_message(f"数据已成功导出到: {file_path} (Data successfully exported to {file_path}).")

        except Exception as e_export: # Catch any error during file writing
            self.log_message(f"导出数据时发生错误: {e_export} (Error during data export).")
            self.log_message(traceback.format_exc()) # Log full traceback
        finally:
            QApplication.processEvents() # Final UI update


if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # It's good practice to set application-wide font here if desired,
    # though it's also set in set_lightblue_style for the main window.
    # default_font = QFont("微软雅黑", 10)
    # app.setFont(default_font)

    window = ExcelToMySQLApp()
    window.show()
    sys.exit(app.exec_())
