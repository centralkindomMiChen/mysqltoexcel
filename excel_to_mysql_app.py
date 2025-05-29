import sys
import pandas as pd
import pymysql
from sqlalchemy import create_engine, text
import sqlalchemy.types
import hashlib
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout,
                             QWidget, QPushButton, QLabel, QLineEdit, QTextEdit,
                             QFileDialog, QGroupBox, QDateEdit, QTableView,
                             QComboBox, QGridLayout)
from PyQt5.QtCore import QTimer, Qt, QDateTime, QDate
from PyQt5.QtGui import QFont, QColor, QPalette, QStandardItemModel, QStandardItem

class ExcelToMySQLApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel导入MySQL工具 (宝石天蓝半透明 - UTF8兼容)")
        self.setGeometry(300, 300, 820, 650)
        self.set_lightblue_style()

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        # self.main_layout will be reassigned in init_ui to the new QHBoxLayout
        # self.main_layout = QVBoxLayout(self.central_widget) 
        # self.main_layout.setSpacing(10)

        self.init_ui() # Call to the refactored UI setup

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)

        self.imported_rows = 0
        self.status = "等待操作"
        self.update_status()
        self.current_queried_df = None # DataFrame to store data from the last successful query

    def set_lightblue_style(self):
        # Sets the application's visual style using QSS (Qt Style Sheets).
        # This specific style aims for a "WinXP transparent宝石蓝色" (WinXP transparent sapphire blue) theme.
        # It includes gradients, rounded corners, and specific color choices for various widgets.
        # 半透明天蓝、宝石蓝渐变风格
        self.setStyleSheet("""
            QMainWindow {
                /* Main window background: vertical gradient from light blue to a slightly darker blue, semi-transparent */
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                stop:0 rgba(180,222,255,190),
                                                stop:1 rgba(90,180,255,170));
                border: 1px solid rgba(120, 180, 240, 180);
                border-radius: 8px;
            }
            QGroupBox {
                background-color: rgba(210, 235, 255, 160); /* Light blue, semi-transparent background for group boxes */
                border: 1px solid rgba(100,180,240,200); /* Slightly darker blue border */
                border-radius: 6px;
                margin-top: 12px; /* Margin to prevent title overlap or provide spacing */
                padding: 12px; /* Inner padding */
                padding-top: 25px; /* Extra top padding to make space for the title */
            }
            QGroupBox::title {
                subcontrol-origin: margin; /* Position relative to the margin area */
                subcontrol-position: top left; /* Place at the top left */
                left: 12px; /* Offset from the left edge */
                padding: 3px 6px;
                background-color: rgba(120, 180, 240, 180); /* Themed blue background for the title */
                border-radius: 4px;
                color: #2B4C77; /* Dark blue text color for contrast */
                font-weight: bold;
            }
            QPushButton {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                stop:0 rgba(120,180,240,220),
                                                stop:1 rgba(90,160,220,220));
                border: 1px solid rgba(100,180,240,200);
                border-radius: 5px;
                padding: 7px 14px;
                min-width: 110px;
                color: #255A8A;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                stop:0 rgba(140,200,255,230),
                                                stop:1 rgba(100,180,240,230));
                border: 1.5px solid rgba(90,140,200,220);
                color: #1366bb;
            }
            QPushButton:pressed {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                stop:0 rgba(90,160,220,230),
                                                stop:1 rgba(80,130,180,230));
                border: 1px solid rgba(60,110,170,220);
            }
            QPushButton:disabled {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                stop:0 rgba(180,222,255,100),
                                                stop:1 rgba(120,180,240,100));
                border: 1px solid rgba(120,180,240,100);
                color: rgba(25, 90, 160, 100);
            }
            QLineEdit {
                background-color: rgba(255, 255, 255, 210);
                border: 1px solid rgba(120,180,240,180);
                border-radius: 4px;
                padding: 6px;
                color: #1865A0;
            }
            QLineEdit[readOnly="true"] {
                background-color: rgba(200, 230, 255, 150);
                color: #2B4C77;
                border: 1px solid rgba(120,180,240,140);
                border-radius: 4px;
                padding: 6px;
            }
            QDateEdit, QComboBox {
                background-color: rgba(255, 255, 255, 210); /* Matches QLineEdit */
                border: 1px solid rgba(120,180,240,180); /* Matches QLineEdit */
                border-radius: 4px; /* Matches QLineEdit */
                padding: 6px; /* Harmonized with QLineEdit */
                color: #1865A0; /* Matches QLineEdit */
                selection-background-color: rgba(120, 180, 240, 180); /* Background color for selected text inside QDateEdit/QComboBox */
                selection-color: #FFFFFF; /* Text color for selected text */
            }
            /* Styling for the dropdown button of QDateEdit */
            QDateEdit::drop-down {
                subcontrol-origin: padding; /* Position relative to the padding edge */
                subcontrol-position: top right; /* Place at the top right */
                width: 20px; /* Width of the dropdown button */
                border-left-width: 1px; /* Border separating button from the text field part */
                border-left-color: rgba(120,180,240,180);
                border-left-style: solid;
                border-top-right-radius: 3px; /* Rounded corners for the button part */
                border-bottom-right-radius: 3px;
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1, /* Gradient background for the button */
                                                stop:0 rgba(180,222,255,190),
                                                stop:1 rgba(120,180,240,190));
            }
            /* Styling for the dropdown arrow of QDateEdit */
            QDateEdit::down-arrow {
                /* image: url(down_arrow.png); -- Removed to allow default Qt arrow or custom QSS arrow */
                width: 12px; 
                height: 12px;
                /* Example for a simple QSS-drawn triangle arrow:
                   border-left: 4px solid transparent;
                   border-right: 4px solid transparent;
                   border-top: 4px solid #1865A0; 
                   margin: auto; 
                */
            }
            /* Styling for the dropdown button of QComboBox, similar to QDateEdit */
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left-width: 1px;
                border-left-color: rgba(120,180,240,180);
                border-left-style: solid;
                border-top-right-radius: 3px;
                border-bottom-right-radius: 3px;
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                stop:0 rgba(180,222,255,190),
                                                stop:1 rgba(120,180,240,190));
            }
            /* Styling for the dropdown arrow of QComboBox */
            QComboBox::down-arrow {
                /* image: url(down_arrow.png); -- Removed for default arrow */
                width: 12px; 
                height: 12px;
            }
            /* Styling for the dropdown list (popup) of QComboBox */
            QComboBox QAbstractItemView { 
                border: 1px solid rgba(120,180,240,180); /* Border for the popup list */
                background-color: rgba(230, 245, 255, 250); /* Background of the list, slightly more opaque */
                color: #1865A0; /* Text color for items */
                selection-background-color: rgba(120, 180, 240, 200); /* Background for selected item */
                selection-color: #FFFFFF; /* Text color for selected item */
                padding: 3px; /* Padding for items within the list */
                outline: 0px; /* Removes focus outline from items if not desired */
            }
            /* General styling for QTableView */
            QTableView {
                background-color: rgba(255, 255, 255, 210); /* Matches QLineEdit background */
                border: 1px solid rgba(120,180,240,180); /* Matches QLineEdit border */
                border-radius: 4px; /* Matches QLineEdit border-radius */
                gridline-color: rgba(180,222,255,190); /* Grid lines with a theme color */
                color: #1865A0; /* Default text color for items */
                alternate-background-color: rgba(230, 245, 255, 180); /* Alternating row color for readability */
            }
            /* Styling for individual items/cells in QTableView */
            QTableView::item {
                padding: 5px;
                border-bottom: 1px solid rgba(180,222,255,170); /* Subtle separator line between rows */
            }
            /* Styling for selected items in QTableView */
            QTableView::item:selected {
                background-color: rgba(120, 180, 240, 180); /* Themed blue background for selected items */
                color: #FFFFFF; /* White text for selected items for contrast */
            }
            /* Styling for the cell that currently has focus in QTableView */
             QTableView::item:focus { 
                outline: 1px solid rgba(90,160,220,220); /* Outline to indicate focus */
                outline-offset: -1px; /* Draw outline slightly inside the cell boundaries */
            }
            /* Styling for the header sections of QTableView (both horizontal and vertical headers) */
            QHeaderView::section {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                stop:0 rgba(180,222,255,190),
                                                stop:1 rgba(120,180,240,190));
                border: 1px solid rgba(120,180,240,180);
                padding: 5px;
                color: #2B4C77;
                font-weight: bold;
            }
            QLabel {
                color: #2B4C77;
                background-color: transparent;
                padding: 3px;
            }
            QTextEdit#ConsoleOutput {
                background-color: rgba(190, 230, 255, 140);
                border: 1px solid rgba(100,180,240,150);
                color: #255A8A;
                font-family: 'Consolas', 'Courier New', monospace;
                border-radius: 4px;
            }
        """)
        self.setWindowOpacity(0.95)
        font = QFont("微软雅黑", 10)
        self.setFont(font)

    def init_ui(self):
        """
        Initializes the main UI layout.
        The UI is structured into a two-column layout:
        - Left Panel: Contains all control and input group boxes.
        - Right Panel: Displays the data preview table and console output.
        An outer QVBoxLayout ensures the status bar remains at the bottom of the central widget.
        """
        
        # This QHBoxLayout forms the main two-column structure (left and right panels).
        app_main_hbox_layout = QHBoxLayout() 
        app_main_hbox_layout.setSpacing(10)

        # --- Left Panel: Controls and Inputs ---
        # This widget and its QVBoxLayout will hold all user input group boxes.
        left_panel_widget = QWidget()
        left_panel_vbox = QVBoxLayout(left_panel_widget) # Layout for the left panel
        left_panel_vbox.setSpacing(10)
        left_panel_vbox.setContentsMargins(0,0,0,0) 

        # Create and add 'Control Panel' group box to the left panel.
        self.control_group = self.create_control_group()
        left_panel_vbox.addWidget(self.control_group)

        # Create and add 'Database Info' group box to the left panel.
        self.db_info_group = self.create_db_info_group()
        left_panel_vbox.addWidget(self.db_info_group)

        # Create and add 'Table Info' group box to the left panel.
        self.table_info_group = self.create_table_info_group()
        left_panel_vbox.addWidget(self.table_info_group)
        
        # Create and add 'Database Query and Export' settings group box to the left panel.
        self.query_export_settings_group = self.create_query_export_group() 
        left_panel_vbox.addWidget(self.query_export_settings_group)
        
        left_panel_vbox.addStretch(1) # Pushes all group boxes in the left panel upwards.

        # Add the left panel widget to the main horizontal layout.
        # The stretch factor of 1 for the left panel vs 3 for the right makes the right panel wider.
        app_main_hbox_layout.addWidget(left_panel_widget, 1) 

        # --- Right Panel: Data Display and Console ---
        # This widget and its QVBoxLayout will hold the data preview table and console output.
        right_panel_widget = QWidget()
        right_panel_vbox = QVBoxLayout(right_panel_widget) # Layout for the right panel
        right_panel_vbox.setSpacing(10)
        right_panel_vbox.setContentsMargins(0,0,0,0)

        # Add the Query Data Preview Table to the right panel.
        # self.query_data_preview_table is instantiated within create_query_export_group().
        # Stretch factor of 3 for the table view makes it take more vertical space than the console.
        right_panel_vbox.addWidget(self.query_data_preview_table, 3) 

        # Add the Console Output (QTextEdit) wrapped in a QGroupBox to the right panel.
        self.console_output = self.create_console_output_widget() 
        console_group_for_right_panel = QGroupBox("控制台输出") # Title in Chinese
        console_layout_for_right_panel = QVBoxLayout()
        console_layout_for_right_panel.addWidget(self.console_output)
        console_group_for_right_panel.setLayout(console_layout_for_right_panel)
        # Stretch factor of 2 for the console group.
        right_panel_vbox.addWidget(console_group_for_right_panel, 2) 
        
        # Add the right panel widget to the main horizontal layout.
        # Stretch factor of 3 makes the right panel wider than the left panel (which has a factor of 1).
        app_main_hbox_layout.addWidget(right_panel_widget, 3)

        # --- Overall Layout and Status Bar ---
        # The outer_vbox_layout ensures that the status_bar is positioned correctly
        # at the bottom of the central widget, underneath the two-column (app_main_hbox_layout) structure.
        outer_vbox_layout = QVBoxLayout(self.central_widget) # Set this as the layout for the central widget
        outer_vbox_layout.addLayout(app_main_hbox_layout) # Add the two-column layout first
        
        self.status_bar = QLabel() 
        self.status_bar.setAlignment(Qt.AlignCenter)
        self.status_bar.setStyleSheet(
            "padding: 4px; color: #255A8A; "
            "background-color: rgba(180,222,255,190); "
            "border-radius: 3px; font-weight:bold;"
        )
        outer_vbox_layout.addWidget(self.status_bar) # Add status bar at the bottom
        self.main_layout = outer_vbox_layout # The outermost layout is now the main_layout for central widget.
        self.update_time()


    def create_control_group(self):
        """Creates and returns the 'Control Panel' QGroupBox."""
        group = QGroupBox("控制面板") # Title in Chinese
        layout = QHBoxLayout()
        layout.setSpacing(10)

        self.btn_select = QPushButton("选择Excel文件")
        self.btn_select.clicked.connect(self.select_excel_file)
        layout.addWidget(self.btn_select)

        self.file_path = QLineEdit()
        self.file_path.setPlaceholderText("未选择文件")
        self.file_path.setReadOnly(True)
        layout.addWidget(self.file_path, 1)

        self.btn_import = QPushButton("导入数据库")
        self.btn_import.clicked.connect(self.import_to_mysql)
        self.btn_import.setEnabled(False)
        layout.addWidget(self.btn_import)

        group.setLayout(layout)
        # self.main_layout.addWidget(group) # Removed: init_ui will add it
        return group

    def create_db_info_group(self):
        """Creates and returns the 'Database Info' QGroupBox."""
        group = QGroupBox("数据库信息") # Title in Chinese
        layout = QHBoxLayout()
        layout.setSpacing(8)

        lbl_db_name = QLabel("数据库名称:")
        lbl_db_name.setMinimumWidth(80)
        layout.addWidget(lbl_db_name)
        self.db_name = QLineEdit("michentestdb2")
        layout.addWidget(self.db_name, 1)

        lbl_db_user = QLabel("用户名:")
        lbl_db_user.setMinimumWidth(60)
        layout.addWidget(lbl_db_user)
        self.db_user = QLineEdit("root")
        layout.addWidget(self.db_user, 1)

        lbl_db_password = QLabel("密码:")
        lbl_db_password.setMinimumWidth(50)
        layout.addWidget(lbl_db_password)
        self.db_password = QLineEdit("123")
        self.db_password.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.db_password, 1)

        group.setLayout(layout)
        # self.main_layout.addWidget(group) # Removed: init_ui will add it
        return group

    def create_table_info_group(self):
        """Creates and returns the 'Table Info' QGroupBox."""
        group = QGroupBox("表信息") # Title in Chinese
        layout = QHBoxLayout()
        layout.setSpacing(8)

        lbl_table_name = QLabel("表名称:")
        lbl_table_name.setMinimumWidth(65)
        layout.addWidget(lbl_table_name)
        self.table_name = QLineEdit("report_data")
        layout.addWidget(self.table_name, 1)

        lbl_rows_imported = QLabel("导入行数:")
        lbl_rows_imported.setMinimumWidth(70)
        layout.addWidget(lbl_rows_imported)
        self.rows_imported = QLineEdit("0")
        self.rows_imported.setReadOnly(True)
        self.rows_imported.setMaximumWidth(100)
        layout.addWidget(self.rows_imported, 0)

        lbl_op_status = QLabel("状态:")
        lbl_op_status.setMinimumWidth(45)
        layout.addWidget(lbl_op_status)
        self.operation_status = QLineEdit("等待操作")
        self.operation_status.setReadOnly(True)
        layout.addWidget(self.operation_status, 1)

        group.setLayout(layout)
        # self.main_layout.addWidget(group) # Removed: init_ui will add it
        return group

    def create_console_output_widget(self):
        """Creates and returns the QTextEdit widget for console output."""
        if not hasattr(self, 'console_output') or self.console_output is None:
             # Ensure console_output is created if not already (e.g. if create_console_group was removed entirely)
            self.console_output = QTextEdit()
            self.console_output.setObjectName("ConsoleOutput")
            self.console_output.setReadOnly(True)
        return self.console_output


    def create_query_export_group(self):
        """
        Creates the 'Database Query and Export' QGroupBox (settings part only).
        The QTableView for data preview is handled separately in init_ui.
        """
        group = QGroupBox("数据库查询与导出") # Group box title in Chinese as per UI
        
        # Main layout for this groupbox (vertical) - for settings only now
        query_export_settings_layout = QVBoxLayout()
        query_export_settings_layout.setSpacing(10) # Spacing between child layouts/widgets

        # --- Database Connection Info ---
        # Layout for database connection parameters
        db_info_layout = QGridLayout() 
        db_info_layout.setSpacing(8)

        # Input fields for database connection
        self.query_db_host_edit = QLineEdit('localhost') # Default host
        self.query_db_user_edit = QLineEdit() 
        self.query_db_password_edit = QLineEdit()
        self.query_db_password_edit.setEchoMode(QLineEdit.Password) # Mask password input
        self.query_db_name_edit = QLineEdit() 
        self.query_table_name_edit = QLineEdit() 
        
        # Adding labels and input fields to the grid layout
        db_info_layout.addWidget(QLabel("主机:"), 0, 0)
        db_info_layout.addWidget(self.query_db_host_edit, 0, 1)
        db_info_layout.addWidget(QLabel("用户:"), 0, 2)
        db_info_layout.addWidget(self.query_db_user_edit, 0, 3)
        
        db_info_layout.addWidget(QLabel("密码:"), 1, 0)
        db_info_layout.addWidget(self.query_db_password_edit, 1, 1)
        db_info_layout.addWidget(QLabel("数据库名:"), 1, 2)
        db_info_layout.addWidget(self.query_db_name_edit, 1, 3)

        db_info_layout.addWidget(QLabel("表名:"), 2, 0)
        db_info_layout.addWidget(self.query_table_name_edit, 2, 1, 1, 3) # Table name input spans 3 columns

        query_export_settings_layout.addLayout(db_info_layout) # Add DB info grid to the settings layout

        # --- Date Selection & Query Button ---
        # Layout for date range selection and the query button
        date_query_layout = QHBoxLayout()
        date_query_layout.setSpacing(8)

        date_query_layout.addWidget(QLabel("开始日期:"))
        self.query_start_date_edit = QDateEdit(QDate.currentDate()) # Default to current date
        self.query_start_date_edit.setCalendarPopup(True) # Use a popup calendar
        date_query_layout.addWidget(self.query_start_date_edit)
        
        date_query_layout.addWidget(QLabel("结束日期:"))
        self.query_end_date_edit = QDateEdit(QDate.currentDate()) # Default to current date
        self.query_end_date_edit.setCalendarPopup(True)
        date_query_layout.addWidget(self.query_end_date_edit)
        
        date_query_layout.addStretch(1) # Pushes the query button to the right
        
        self.btn_query_data = QPushButton("查询数据") 
        self.btn_query_data.clicked.connect(self.query_data_from_db) # Connect button to query method
        date_query_layout.addWidget(self.btn_query_data)
        query_export_settings_layout.addLayout(date_query_layout) # Add date/query layout to settings

        # --- Data Display Table (creation only, not added to this group's layout) ---
        if not hasattr(self, 'query_data_preview_table') or self.query_data_preview_table is None:
            self.query_data_preview_table = QTableView() 
            # Basic properties like setAlternatingRowColors could be set here if desired,
            # but it's primarily for display in the right panel.

        # --- Export Section ---
        # Layout for export path selection
        export_controls_layout = QHBoxLayout()
        export_controls_layout.setSpacing(8)

        export_controls_layout.addWidget(QLabel("导出路径:"))
        self.query_export_path_edit = QLineEdit() 
        self.query_export_path_edit.setPlaceholderText("选择或输入导出文件路径...")
        export_controls_layout.addWidget(self.query_export_path_edit, 1) # Path edit takes available space
        
        self.btn_browse_export_path = QPushButton("浏览...") 
        self.btn_browse_export_path.clicked.connect(self.select_export_file_path) # Connect to path selection dialog
        export_controls_layout.addWidget(self.btn_browse_export_path)
        query_export_settings_layout.addLayout(export_controls_layout)

        # Layout for export format selection and download button
        export_options_layout = QHBoxLayout()
        export_options_layout.setSpacing(8)
        export_options_layout.addWidget(QLabel("导出格式:"))
        self.query_export_format_combo = QComboBox() 
        self.query_export_format_combo.addItems(["CSV", "XLSX"]) # Available export formats
        export_options_layout.addWidget(self.query_export_format_combo)
        
        export_options_layout.addStretch(1) # Pushes download button to the right

        self.btn_download_data = QPushButton("下载数据") 
        self.btn_download_data.setEnabled(False) # Initially disabled, enabled after successful query
        self.btn_download_data.clicked.connect(self.download_queried_data) # Connect to download method
        export_options_layout.addWidget(self.btn_download_data)
        query_export_settings_layout.addLayout(export_options_layout)
        
        group.setLayout(query_export_settings_layout) # Set the settings layout for the group box
        return group

    def select_export_file_path(self):
        """
        Opens a file dialog to allow the user to select a path and filename for exporting data.
        The suggested filename and file type filter are based on the currently selected
        export format in the QComboBox.
        """
        self.log_message("Opening dialog to select export file path...")
        current_format = self.query_export_format_combo.currentText().lower()
        default_filename = f"exported_data.{current_format}" # Suggest e.g., "exported_data.csv"
        
        # Define file filters based on the selected format for the dialog
        if current_format == "csv":
            filter_str = "CSV Files (*.csv);;All Files (*)"
        elif current_format == "xlsx":
            filter_str = "Excel Files (*.xlsx);;All Files (*)"
        else:
            filter_str = "All Files (*)" # Fallback, should not happen with current ComboBox items

        options = QFileDialog.Options()
        # options |= QFileDialog.DontUseNativeDialog # Useful for testing or if native dialogs cause issues
        
        # Open the "Save File" dialog
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "选择导出文件路径", # Dialog title
            default_filename,  # Suggested filename
            filter_str,        # File type filters
            options=options
        )

        if file_path:
            # Automatically append the correct extension if the user didn't type it
            # or if they selected a filter but typed a name without an extension.
            if current_format == "csv" and not file_path.lower().endswith(".csv"):
                file_path += ".csv"
            elif current_format == "xlsx" and not file_path.lower().endswith(".xlsx"):
                file_path += ".xlsx"
                
            self.query_export_path_edit.setText(file_path) # Update the QLineEdit with the chosen path
            self.log_message(f"Export file path selected: {file_path}")
        else:
            self.log_message("Export file path selection cancelled.")

    def download_queried_data(self):
        """
        Saves the data currently stored in `self.current_queried_df` to a file.
        The file path and format (CSV/XLSX) are taken from the UI elements.
        """
        self.log_message("Initiating data download...")

        # Check if there's any data to export
        if self.current_queried_df is None or self.current_queried_df.empty:
            self.log_message("没有可供导出的数据。请先成功查询数据。") # "No data available for export. Please query data first."
            return

        # Get the file path from the QLineEdit
        file_path = self.query_export_path_edit.text().strip()
        if not file_path:
            self.log_message("错误: 请先选择或输入导出文件路径。") # "Error: Please select or enter an export file path."
            return

        # Get the selected export format
        export_format = self.query_export_format_combo.currentText()
        self.log_message(f"Attempting to export data as {export_format} to: {file_path}")

        try:
            # Save the DataFrame to the specified format
            if export_format == "CSV":
                self.current_queried_df.to_csv(file_path, index=False, encoding='utf-8-sig') # utf-8-sig for CSV with BOM
            elif export_format == "XLSX":
                self.current_queried_df.to_excel(file_path, index=False)
            else:
                # This case should ideally not be reached if ComboBox items are fixed
                self.log_message(f"错误: 不支持的导出格式 {export_format}。") # "Error: Unsupported export format."
                return
            
            self.log_message(f"数据已成功导出到: {file_path}") # "Data successfully exported to: {file_path}"

        except Exception as e:
            self.log_message(f"导出数据时发生错误: {e}") # "Error occurred during data export:"
            import traceback
            self.log_message(traceback.format_exc()) # Log full traceback for debugging

    def query_data_from_db(self):
        """
        Queries data from the specified MySQL database based on user inputs.
        Retrieves connection details, table name, and date range from UI elements.
        Fetches data using pandas and pymysql, then populates a QTableView.
        Handles errors during the process and logs messages to the console.
        """
        self.log_message("Initiating data query from database...")

        # Retrieve database connection details and query parameters from QLineEdit and QDateEdit widgets
        host = self.query_db_host_edit.text().strip()
        user = self.query_db_user_edit.text().strip()
        password = self.query_db_password_edit.text() # Password is taken as is, without stripping whitespace
        db_name = self.query_db_name_edit.text().strip()
        table_name = self.query_table_name_edit.text().strip()
        
        start_date_q = self.query_start_date_edit.date() # QDate object for start date
        end_date_q = self.query_end_date_edit.date()     # QDate object for end date

        # --- Input Validation ---
        # Check if essential fields are filled
        if not all([host, user, db_name, table_name]):
            self.log_message("错误: 主机, 用户, 数据库名, 和 表名 不能为空。") # "Error: Host, User, DB Name, and Table Name cannot be empty."
            self.query_data_preview_table.setModel(None) # Clear table view
            self.btn_download_data.setEnabled(False)     # Disable download button
            self.current_queried_df = None               # Clear stored DataFrame
            return

        # Check if the start date is after the end date
        if start_date_q > end_date_q:
            self.log_message("错误: 开始日期不能晚于结束日期。") # "Error: Start date cannot be later than end date."
            self.query_data_preview_table.setModel(None)
            self.btn_download_data.setEnabled(False)
            self.current_queried_df = None
            return

        # --- Date Formatting for SQL Query ---
        start_date_str = start_date_q.toString("yyyy-MM-dd")
        # End date string, also in 'yyyy-MM-dd' format for DATE() comparison.
        end_date_str = end_date_q.toString("yyyy-MM-dd") 
        
        self.log_message(f"Querying table `{table_name}` from {start_date_str} to {end_date_str}.")

        # --- SQL Query Construction ---
        # Uses DATE() function on the 'Date' column to compare only the date part.
        # This makes the comparison robust against varying time components in the 'Date' field.
        sql_query = f"SELECT * FROM `{table_name}` WHERE DATE(`Date`) >= '{start_date_str}' AND DATE(`Date`) <= '{end_date_str}'"
        self.log_message(f"Executing SQL: {sql_query}")

        conn = None # Initialize connection variable
        try:
            # --- Database Connection ---
            self.log_message(f"Connecting to database '{db_name}' on host '{host}' with user '{user}'...")
            conn_config = {
                'host': host,
                'user': user,
                'password': password,
                'database': db_name,
                'charset': 'utf8', # Changed from utf8mb4 to utf8 for compatibility
                'cursorclass': pymysql.cursors.DictCursor # Fetch results as dictionaries (optional, good for pandas)
            }
            conn = pymysql.connect(**conn_config)
            self.log_message("Database connection successful.")

            # --- Data Fetching ---
            self.log_message("Fetching data...")
            # Use pandas to execute the SQL query and load results into a DataFrame
            df = pd.read_sql_query(sql_query, conn)
            
            if df.empty:
                self.log_message("未查询到数据。") # "No data found."
                self.query_data_preview_table.setModel(None) # Clear table
                self.btn_download_data.setEnabled(False)     # Disable download
                self.current_queried_df = None               # Clear DataFrame
            else:
                self.log_message(f"成功获取 {len(df)} 行数据。正在填充表格...") # "Successfully fetched {len(df)} rows. Populating table..."
                
                # --- Populate QTableView ---
                # Create an empty QStandardItemModel. Rows will be appended.
                model = QStandardItemModel() 
                model.setHorizontalHeaderLabels(list(df.columns)) # Set column headers correctly.
                
                # Iterate over the DataFrame rows and columns to populate the model
                # using df.iloc for precise cell access.
                for i in range(len(df)):  # Iterate through row indices
                    row_items = []
                    for j in range(len(df.columns)):  # Iterate through column indices
                        item_value = df.iloc[i, j]
                        # Convert item_value to string, handle None or other types if necessary
                        item_text = str(item_value) if item_value is not None else ""
                        standard_item = QStandardItem(item_text)
                        standard_item.setEditable(False) # Make cells non-editable
                        row_items.append(standard_item)
                    model.appendRow(row_items) # Append the list of items as a new row
                
                self.query_data_preview_table.setModel(model) # Set the model to the QTableView
                self.query_data_preview_table.resizeColumnsToContents() # Adjust column widths to fit content
                self.btn_download_data.setEnabled(True)     # Enable the download button
                self.current_queried_df = df                # Store the fetched DataFrame
                self.log_message("数据已成功加载到预览表格。") # "Data successfully loaded into preview table."

        except pymysql.Error as e: # Catch specific PyMySQL errors (e.g., connection, query execution)
            self.log_message(f"数据库错误: {e}") # "Database error:"
            self.query_data_preview_table.setModel(None)
            self.btn_download_data.setEnabled(False)
            self.current_queried_df = None
        except pd.errors.DatabaseError as e: # Catch pandas-specific database errors
             self.log_message(f"Pandas数据库读取错误: {e}") # "Pandas database read error:"
             self.query_data_preview_table.setModel(None)
             self.btn_download_data.setEnabled(False)
             self.current_queried_df = None
        except Exception as e: # Catch any other unexpected errors
            self.log_message(f"查询数据时发生未知错误: {e}") # "An unknown error occurred while querying data:"
            import traceback
            self.log_message(traceback.format_exc()) # Log the full traceback for debugging
            self.query_data_preview_table.setModel(None)
            self.btn_download_data.setEnabled(False)
            self.current_queried_df = None
        finally:
            # --- Resource Cleanup ---
            if conn:
                conn.close() # Ensure the database connection is closed
                self.log_message("Database connection closed.")
            QApplication.processEvents() # Process any pending UI events


    def select_excel_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "",
            "Excel Files (*.xlsx *.xls);;All Files (*)",
            options=options
        )
        if file_name:
            self.file_path.setText(file_name)
            self.btn_import.setEnabled(True)
            self.log_message(f"已选择文件: {file_name}")
        else:
            self.btn_import.setEnabled(False)

    def import_to_mysql(self):
        excel_file = self.file_path.text()
        if not excel_file:
            self.log_message("错误: 请先选择Excel文件")
            return

        self.btn_import.setEnabled(False)
        self.status = "处理中..."
        self.update_status()
        QApplication.processEvents()

        conn = None
        cursor = None
        engine = None

        try:
            self.log_message("正在读取Excel文件...")
            df = pd.read_excel(excel_file)
            if df.empty:
                self.log_message("警告: Excel文件为空，跳过导入")
                self.status = "完成 (无数据)"
                return
            self.log_message(f"成功读取Excel文件，共 {len(df)} 行数据")

            config = {
                'user': self.db_user.text(),
                'password': self.db_password.text(),
                'host': 'localhost',
                'charset': 'utf8',
                'use_unicode': True
            }
            self.log_message("正在连接MySQL数据库...")
            conn = pymysql.connect(**config)
            cursor = conn.cursor()
            db_name = self.db_name.text()
            self.log_message(f"正在创建/使用数据库: {db_name}")
            cursor.execute(f"CREATE DATABASE IF NOT EXISTS `{db_name}` CHARACTER SET utf8 COLLATE utf8_general_ci")
            cursor.execute(f"USE `{db_name}`")
            table_name = self.table_name.text()

            def generate_data_fingerprint(dataframe):
                sample_data = dataframe.head(min(5, len(dataframe))).to_string(index=False)
                return hashlib.md5(sample_data.encode('utf-8')).hexdigest()

            current_fingerprint = generate_data_fingerprint(df)
            self.log_message("已生成数据指纹用于重复检查")
            cursor.execute(f"SHOW TABLES LIKE '{table_name}'")
            table_exists = cursor.fetchone() is not None

            if table_exists:
                self.log_message(f"表 {table_name} 已存在，检查重复数据...")
                try:
                    cursor.execute(f"SHOW COLUMNS FROM `{table_name}` LIKE 'data_fingerprint'")
                    has_fingerprint_column = cursor.fetchone() is not None
                    if has_fingerprint_column:
                        cursor.execute(f"SELECT 1 FROM `{table_name}` WHERE `data_fingerprint` = %s LIMIT 1", (current_fingerprint,))
                        if cursor.fetchone() is not None:
                            self.log_message("检测到相似数据批次 (基于指纹)，跳过导入")
                            self.status = "完成 (跳过重复数据)"
                            return
                    else:
                        self.log_message(f"警告: 表 {table_name} 不存在 'data_fingerprint' 列。")
                except pymysql.Error as e:
                    self.log_message(f"检查表结构时出错: {e}. 继续尝试导入...")

            engine_url = f"mysql+pymysql://{config['user']}:{config['password']}@{config['host']}/{db_name}?charset=utf8"
            engine = create_engine(engine_url)
            df['data_fingerprint'] = current_fingerprint
            custom_dtype = {'data_fingerprint': sqlalchemy.types.VARCHAR(32)}

            self.log_message("正在导入数据到MySQL...")
            df.to_sql(
                name=table_name,
                con=engine,
                if_exists='append',
                index=False,
                chunksize=1000,
                dtype=custom_dtype if not table_exists else None
            )

            if not table_exists:
                with engine.connect() as connection:
                    trans = connection.begin()
                    try:
                        inspector = sqlalchemy.inspect(engine)
                        columns_in_table = [col['name'] for col in inspector.get_columns(table_name)]
                        if 'id' not in columns_in_table:
                            connection.execute(text(f"ALTER TABLE `{table_name}` ADD COLUMN `id` INT AUTO_INCREMENT PRIMARY KEY FIRST;"))
                            self.log_message(f"Added AUTO_INCREMENT PRIMARY KEY 'id' to new table '{table_name}'.")
                        if 'import_time' not in columns_in_table:
                            connection.execute(text(f"ALTER TABLE `{table_name}` ADD COLUMN `import_time` TIMESTAMP DEFAULT CURRENT_TIMESTAMP;"))
                            self.log_message(f"Added 'import_time' column to new table '{table_name}'.")
                        trans.commit()
                    except Exception as alter_e:
                        trans.rollback()
                        self.log_message(f"添加列时出错 (可能已存在或其它问题): {alter_e}")

            self.imported_rows = len(df)
            self.log_message(f"成功导入 {self.imported_rows} 行数据到表 '{table_name}'")
            self.log_message("正在从数据库导出当前批次数据到CSV和Excel...")
            export_df = pd.read_sql(f"SELECT * FROM `{table_name}` WHERE `data_fingerprint` = '{current_fingerprint}'", engine)

            if not export_df.empty:
                csv_output = 'output_current_import.csv'
                export_df.to_csv(csv_output, index=False, encoding='utf-8-sig')
                self.log_message(f"当前导入批次数据已导出到: {csv_output}")
                excel_output = 'output_current_import.xlsx'
                export_df.to_excel(excel_output, index=False)
                self.log_message(f"当前导入批次数据已导出到: {excel_output}")
            else:
                self.log_message("没有数据导出（当前批次未找到或为空）。")
            self.status = "完成 (成功)"

        except Exception as e_outer:
            self.log_message(f"发生错误: {str(e_outer)}")
            import traceback
            self.log_message(traceback.format_exc())
            self.status = "完成 (失败)"
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()
            if engine:
                engine.dispose()

            self.btn_import.setEnabled(True)
            self.update_status()
            QApplication.processEvents()

    def log_message(self, message):
        timestamp = QDateTime.currentDateTime().toString("hh:mm:ss")
        self.console_output.append(f"[{timestamp}] {message}")
        self.console_output.ensureCursorVisible()

    def update_time(self):
        current_time = QDateTime.currentDateTime().toString("yyyy-MM-dd hh:mm:ss")
        self.status_bar.setText(f"系统时间: {current_time}")

    def update_status(self):
        self.rows_imported.setText(str(self.imported_rows))
        self.operation_status.setText(self.status)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    font = QFont("微软雅黑", 10)
    app.setFont(font)
    window = ExcelToMySQLApp()
    window.show()
    sys.exit(app.exec_())
