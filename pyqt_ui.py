"""Modern PyQt5-based UI for attendance tracker."""

import sys
import os
from datetime import datetime
from pathlib import Path
import logging

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QGridLayout, QLabel, QPushButton, QTextEdit, QComboBox, QLineEdit,
    QFrame, QScrollArea, QGroupBox, QRadioButton, QButtonGroup,
    QFileDialog, QMessageBox,
    QProgressBar, QTabWidget, QSplitter, QStatusBar, QDialog,
    QListWidget, QListWidgetItem
)
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal, QPropertyAnimation, QEasingCurve
from PyQt5.QtGui import (
    QFont, QIcon, QPalette, QColor, QLinearGradient, QPainter, 
    QBrush, QPen, QPixmap, QFontDatabase
)

logger = logging.getLogger(__name__)

class ModernButton(QPushButton):
    """Custom modern button with hover effects."""
    
    def __init__(self, text, color="#007bff", hover_color="#0056b3", parent=None):
        super().__init__(text, parent)
        self.color = color
        self.hover_color = hover_color
        self.setup_style()
    
    def setup_style(self):
        """Setup modern button styling."""
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.color};
                color: white;
                border: none;
                border-radius: 8px;
                padding: 12px 24px;
                font-size: 14px;
                font-weight: bold;
                font-family: 'Segoe UI', Arial, sans-serif;
            }}
            QPushButton:hover {{
                background-color: {self.hover_color};
            }}
            QPushButton:pressed {{
                background-color: {self.hover_color};
            }}
            QPushButton:disabled {{
                background-color: #6c757d;
                color: #adb5bd;
            }}
        """)
        self.setMinimumHeight(45)

class ModernCard(QFrame):
    """Modern card widget with shadow effect."""
    
    def __init__(self, title="", parent=None):
        super().__init__(parent)
        self.setFrameStyle(QFrame.NoFrame)
        self.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 12px;
                border: 1px solid #e9ecef;
                margin: 5px;
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        if title:
            title_label = QLabel(title)
            title_label.setStyleSheet("""
                QLabel {
                    font-size: 18px;
                    font-weight: bold;
                    color: #2c3e50;
                    margin-bottom: 10px;
                    border: none;
                }
            """)
            layout.addWidget(title_label)

class ModernRadioButton(QRadioButton):
    """Custom modern radio button."""
    
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setStyleSheet("""
            QRadioButton {
                font-size: 14px;
                color: #495057;
                spacing: 10px;
                padding: 8px;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
                border-radius: 9px;
                border: 2px solid #007bff;
                background-color: white;
            }
            QRadioButton::indicator:checked {
                background-color: #007bff;
                border: 2px solid #007bff;
            }
            QRadioButton::indicator:hover {
                border: 2px solid #0056b3;
            }
        """)

class AttendancePyQtGUI(QMainWindow):
    """Modern PyQt-based attendance tracker GUI."""
    
    def __init__(self, config_manager, event_monitor, pdf_handler, word_handler=None, web_server=None):
        super().__init__()
        self.config = config_manager
        self.pdf_handler = pdf_handler
        self.word_handler = word_handler
        self.web_server = web_server
        self.last_checkin_time = None
        # Setup UI
        self.setup_ui()
        self.setup_timers()
        if self.web_server:
            self.start_web_server()
        self.apply_modern_theme()
        
    def setup_ui(self):
        """Setup the main user interface."""
        self.setWindowTitle("🕒 Khaled's Attendance Tracker - Modern Edition")
        self.setGeometry(100, 100, 1200, 800)
        self.setMinimumSize(1000, 700)
        
        # Central widget with splitter
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)
        
        # Create splitter for resizable sections
        splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(splitter)
        
        # Left panel with scroll area
        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setMaximumWidth(500)
        left_scroll.setMinimumWidth(400)
        
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setSpacing(15)
        
        # Create sections
        self.create_config_section(left_layout)
        self.create_status_section(left_layout)
        self.create_controls_section(left_layout)
        
        left_scroll.setWidget(left_widget)
        splitter.addWidget(left_scroll)
        
        # Right panel for logs
        self.create_log_section(splitter)
        
        # Set splitter proportions
        splitter.setSizes([400, 600])
        
        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")
        
    def create_config_section(self, layout):
        """Create configuration section."""
        config_card = ModernCard("📁 Configuration")
        config_layout = QVBoxLayout()
        
        # Document type selection
        doc_type_group = QGroupBox("Document Type")
        doc_type_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #dee2e6;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
        """)
        
        doc_type_layout = QHBoxLayout(doc_type_group)
        
        self.doc_type_group = QButtonGroup()
        self.word_radio = ModernRadioButton("📝 Word Document")
        
        self.doc_type_group.addButton(self.word_radio, 1)
        
        # Connect radio button changes
        self.word_radio.toggled.connect(self.on_doc_type_change)
        
        doc_type_layout.addWidget(self.word_radio)
        
        config_layout.addWidget(doc_type_group)
        
        # File path selection
        path_layout = QHBoxLayout()
        self.path_input = QLineEdit()
        self.path_input.setPlaceholderText("Select document file...")
        self.path_input.setStyleSheet("""
            QLineEdit {
                padding: 12px;
                border: 2px solid #dee2e6;
                border-radius: 8px;
                font-size: 14px;
                background-color: white;
            }
            QLineEdit:focus {
                border-color: #007bff;
            }
        """)
        
        browse_btn = ModernButton("📂 Browse", "#28a745", "#1e7e34")
        browse_btn.clicked.connect(self.browse_file)
        
        path_layout.addWidget(self.path_input, 3)
        path_layout.addWidget(browse_btn, 1)
        config_layout.addLayout(path_layout)
        
        # Set initial selection and path AFTER creating path_input
        self.word_radio.setChecked(True)
        doc_path = self.config.get('document_path', '')
        if doc_path and Path(doc_path).exists():
            self.path_input.setText(doc_path)
        else:
            self.path_input.setText("")
            self.show_info("Document Required", "Please upload your Word document. The previous file was not found.")
        
        # Month selection
        month_layout = QHBoxLayout()
        month_label = QLabel("📅 Month:")
        month_label.setStyleSheet("font-weight: bold; color: #495057;")
        
        self.month_combo = QComboBox()
        self.month_combo.setStyleSheet("""
            QComboBox {
                padding: 12px;
                border: 2px solid #dee2e6;
                border-radius: 8px;
                font-size: 14px;
                background-color: white;
            }
            QComboBox:focus {
                border-color: #007bff;
            }
            QComboBox::drop-down {
                border: none;
                width: 30px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #007bff;
            }
        """)
        
        # Populate months
        self.populate_months()
        
        # Connect month change
        self.month_combo.currentTextChanged.connect(self.on_month_change)
        
        month_layout.addWidget(month_label)
        month_layout.addWidget(self.month_combo, 1)
        config_layout.addLayout(month_layout)
        
        config_card.layout().addLayout(config_layout)
        layout.addWidget(config_card)
        
    def create_status_section(self, layout):
        """Create status section."""
        status_card = ModernCard("📊 Status")
        status_layout = QVBoxLayout()
        
        # Current time display
        self.time_label = QLabel()
        self.time_label.setStyleSheet("""
            QLabel {
                font-size: 24px;
                font-weight: bold;
                color: #007bff;
                text-align: center;
                padding: 15px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #f8f9fa, stop:1 #e9ecef);
                border-radius: 8px;
                border: none;
            }
        """)
        self.time_label.setAlignment(Qt.AlignCenter)
        status_layout.addWidget(self.time_label)
        
        # Monitoring status
        self.status_label = QLabel("⏹️ Monitoring: Stopped")
        self.status_label.setStyleSheet("""
            QLabel {
                font-size: 16px;
                font-weight: bold;
                padding: 10px;
                background-color: #f8d7da;
                color: #721c24;
                border-radius: 8px;
                border: none;
            }
        """)
        self.status_label.setAlignment(Qt.AlignCenter)
        status_layout.addWidget(self.status_label)
        
        status_card.layout().addLayout(status_layout)
        layout.addWidget(status_card)
        
    def create_controls_section(self, layout):
        """Create controls section."""
        controls_card = ModernCard("🎮 Quick Actions")
        controls_layout = QVBoxLayout()
        
        # Monitoring controls
        # Monitoring controls removed (fully manual mode)
        
        # Manual actions
        manual_layout = QHBoxLayout()
        checkin_btn = ModernButton("🟢 Manual Check-in", "#17a2b8", "#138496")
        checkout_btn = ModernButton("🔴 Manual Check-out", "#ffc107", "#e0a800")
        
        checkin_btn.clicked.connect(self.manual_checkin)
        checkout_btn.clicked.connect(self.manual_checkout)
        
        manual_layout.addWidget(checkin_btn)
        manual_layout.addWidget(checkout_btn)
        controls_layout.addLayout(manual_layout)
        
        # Download button
        download_btn = ModernButton("💾 Download Filled Document", "#6f42c1", "#5a2d91")
        download_btn.clicked.connect(self.download_document)
        controls_layout.addWidget(download_btn)
        
        # Mobile access info (if web server is enabled)
        if self.web_server:
            mobile_info = QLabel()
            mobile_info.setStyleSheet("""
                QLabel {
                    background-color: #e3f2fd;
                    border: 1px solid #2196f3;
                    border-radius: 6px;
                    padding: 8px;
                    color: #0d47a1;
                    font-size: 12px;
                    margin: 5px 0;
                }
            """)
            mobile_info.setWordWrap(True)
            mobile_info.setText("📱 Mobile access enabled! Check console for iPhone setup instructions.")
            controls_layout.addWidget(mobile_info)
        
        controls_card.layout().addLayout(controls_layout)
        layout.addWidget(controls_card)
        
        # Add stretch to push everything up
        layout.addStretch()
        
    def create_log_section(self, parent):
        """Create log section."""
        log_card = ModernCard("📝 Activity Log")
        
        # Clear button
        clear_btn = ModernButton("🗑️ Clear Log", "#6c757d", "#5a6268")
        clear_btn.clicked.connect(self.clear_log)
        log_card.layout().addWidget(clear_btn)
        
        # Log text area
        self.log_text = QTextEdit()
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #f8f9fa;
                border: 2px solid #dee2e6;
                border-radius: 8px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 12px;
                padding: 10px;
                color: #495057;
            }
        """)
        self.log_text.setReadOnly(True)
        log_card.layout().addWidget(self.log_text)
        
        parent.addWidget(log_card)
        
    def apply_modern_theme(self):
        """Apply modern theme to the application."""
        self.setStyleSheet("""
            QMainWindow {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #f8f9fa, stop:1 #e9ecef);
            }
            QScrollArea {
                background: transparent;
                border: none;
            }
            QSplitter::handle {
                background-color: #dee2e6;
                width: 2px;
            }
            QSplitter::handle:hover {
                background-color: #007bff;
            }
        """)
        
    # Tray functionality removed
            
    def setup_timers(self):
        """Setup timers for updates."""
        # Timer for clock update
        self.clock_timer = QTimer()
        self.clock_timer.timeout.connect(self.update_time)
        self.clock_timer.start(1000)  # Update every second
        
        # Initial time update
        self.update_time()
        
    def update_time(self):
        """Update the time display."""
        current_time = datetime.now().strftime("%H:%M:%S")
        current_date = datetime.now().strftime("%A, %B %d, %Y")
        self.time_label.setText(f"{current_time}\n{current_date}")
        
    def populate_months(self):
        """Populate month dropdown."""
        import calendar
        from datetime import datetime
        
        current_year = datetime.now().year
        months = []
        
        for year in [current_year - 1, current_year, current_year + 1]:
            for month in range(1, 13):
                month_name = calendar.month_name[month]
                months.append(f"{month_name} {year}")
        
        self.month_combo.addItems(months)
        
        # Set current month
        current_month = datetime.now().strftime("%B %Y")
        saved_month = self.config.get('selected_month', current_month)
        
        index = self.month_combo.findText(saved_month)
        if index >= 0:
            self.month_combo.setCurrentIndex(index)
        
    def browse_file(self):
        """Browse for document file."""
        file_type = "Word Documents (*.docx)"
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Document", "", file_type
        )
        
        if file_path:
            self.path_input.setText(file_path)
            # Save to config
            self.config.set('document_path', file_path)
            self.config.set('document_type', 'word')
            self.log_message(f"📁 Document selected: {Path(file_path).name}")
        
    def start_monitoring(self):
        """Start event monitoring."""
        try:
            self.event_monitor.start_monitoring()
            self.start_btn.setEnabled(False)
            self.stop_btn.setEnabled(True)
            self.status_label.setText("▶️ Monitoring: Active")
            self.status_label.setStyleSheet("""
                QLabel {
                    font-size: 16px;
                    font-weight: bold;
                    padding: 10px;
                    background-color: #d4edda;
                    color: #155724;
                    border-radius: 8px;
                    border: none;
                }
            """)
            self.status_bar.showMessage("Monitoring started")
            self.log_message("🟢 Event monitoring started")
        except Exception as e:
            self.show_error(f"Failed to start monitoring: {str(e)}")
            
    def stop_monitoring(self):
        """Stop event monitoring."""
        try:
            self.event_monitor.stop_monitoring()
            self.start_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            self.status_label.setText("⏹️ Monitoring: Stopped")
            self.status_label.setStyleSheet("""
                QLabel {
                    font-size: 16px;
                    font-weight: bold;
                    padding: 10px;
                    background-color: #f8d7da;
                    color: #721c24;
                    border-radius: 8px;
                    border: none;
                }
            """)
            self.status_bar.showMessage("Monitoring stopped")
            self.log_message("🔴 Event monitoring stopped")
        except Exception as e:
            self.show_error(f"Failed to stop monitoring: {str(e)}")
            
    def manual_checkin(self):
        """Manual check-in."""
        try:
            doc_type = 'word' if self.word_radio.isChecked() else 'pdf'
            current_time = datetime.now()
            if doc_type == 'word' and self.word_handler:
                # Use word handler for manual check-in
                result = self.word_handler.fill_attendance_sheet(time_in=current_time)
                if result:
                    self.last_checkin_time = current_time
                    self.log_message("🟢 Manual check-in recorded successfully")
                    self.status_bar.showMessage("Manual check-in recorded")
                    self.show_info("Success", f"Check-in recorded at {current_time.strftime('%H:%M:%S')}")
                else:
                    self.log_message("❌ Failed to record check-in")
                    self.show_error("Failed to record check-in. Please check document path.")
            elif doc_type == 'pdf' and self.pdf_handler:
                # Use PDF handler for manual check-in
                result = self.pdf_handler.fill_attendance_sheet(time_in=current_time)
                if result:
                    self.last_checkin_time = current_time
                    self.log_message("🟢 Manual check-in recorded successfully")
                    self.status_bar.showMessage("Manual check-in recorded")
                    self.show_info("Success", f"Check-in recorded at {current_time.strftime('%H:%M:%S')}")
                else:
                    self.log_message("❌ Failed to record check-in")
                    self.show_error("Failed to record check-in. Please check document path.")
            else:
                self.show_error("No document handler available. Please select a document type and file path.")
                
        except Exception as e:
            self.log_message(f"❌ Error during manual check-in: {e}")
            self.show_error(f"Failed to record check-in: {e}")
        
    def manual_checkout(self):
        """Manual check-out."""
        try:
            doc_type = 'word' if self.word_radio.isChecked() else 'pdf'
            current_time = datetime.now()
            import traceback
            # Debug: log the type and value of last_checkin_time and current_time
            self.log_message(f"[DEBUG] last_checkin_time: {repr(getattr(self, 'last_checkin_time', None))} (type: {type(getattr(self, 'last_checkin_time', None))})")
            self.log_message(f"[DEBUG] current_time: {repr(current_time)} (type: {type(current_time)})")
            doc_type = 'word' if self.word_radio.isChecked() else 'pdf'
            if doc_type == 'word' and self.word_handler:
                time_in = getattr(self, 'last_checkin_time', None)
                if time_in is None:
                    time_in = current_time
                self.log_message(f"[DEBUG] Using time_in for word: {repr(time_in)} (type: {type(time_in)})")
                result = self.word_handler.fill_attendance_sheet(time_in=time_in, time_out=current_time)
                if result:
                    self.log_message("🔴 Manual check-out recorded successfully")
                    self.status_bar.showMessage("Manual check-out recorded")
                    self.show_info("Success", f"Check-out recorded at {current_time.strftime('%H:%M:%S')}")
                else:
                    self.log_message("❌ Failed to record check-out")
                    self.show_error("Failed to record check-out. Please check document path.")
            elif doc_type == 'pdf' and self.pdf_handler:
                time_in = getattr(self, 'last_checkin_time', None)
                if time_in is None:
                    time_in = current_time
                self.log_message(f"[DEBUG] Using time_in for pdf: {repr(time_in)} (type: {type(time_in)})")
                result = self.pdf_handler.fill_attendance_sheet(time_in=time_in, time_out=current_time)
                if result:
                    self.log_message("🔴 Manual check-out recorded successfully")
                    self.status_bar.showMessage("Manual check-out recorded")
                    self.show_info("Success", f"Check-out recorded at {current_time.strftime('%H:%M:%S')}")
                else:
                    self.log_message("❌ Failed to record check-out")
                    self.show_error("Failed to record check-out. Please check document path.")
            else:
                self.show_error("No document handler available. Please select a document type and file path.")
        except Exception as e:
            self.log_message(f"❌ Error during manual check-out: {e}\n{traceback.format_exc()}")
            self.show_error(f"Failed to record check-out: {e}")
        
    def download_document(self):
        """Download/open the most recent filled documents."""
        try:
            from pathlib import Path
            from datetime import datetime
            import os
            import shutil
            
            output_dir = Path(self.config.get('output_directory', 'filled_docs'))
            
            if not output_dir.exists():
                self.show_warning("No Documents", "No filled documents found. Please record some attendance first.")
                self.log_message("No filled documents directory found")
                return
            
            # Get all document files (PDF and Word) sorted by modification time (newest first)
            doc_files = list(output_dir.glob("*.pdf")) + list(output_dir.glob("*.docx")) + list(output_dir.glob("*.doc"))
            if not doc_files:
                self.show_warning("No Documents", "No filled documents found. Please record some attendance first.")
                self.log_message("No filled documents found in output directory")
                return
            
            doc_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            
            # Show document selection dialog
            self.show_document_selection_dialog(doc_files)
            
        except Exception as e:
            self.log_message(f"Error accessing filled documents: {e}")
            self.show_error(f"Failed to access filled documents: {e}")
    
    def show_document_selection_dialog(self, doc_files):
        """Show dialog to select and download document."""
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QListWidget, QListWidgetItem
        from datetime import datetime
        import os
        import shutil
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Select Document to Download")
        dialog.setGeometry(0, 0, 600, 400)
        dialog.move(self.x() + 50, self.y() + 50)
        
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # Title
        title_label = QLabel("Select a filled document file:")
        title_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #2c3e50;")
        layout.addWidget(title_label)
        
        # Document list
        doc_list = QListWidget()
        doc_list.setStyleSheet("""
            QListWidget {
                border: 2px solid #dee2e6;
                border-radius: 8px;
                background-color: white;
                font-size: 14px;
                padding: 5px;
            }
            QListWidget::item {
                padding: 10px;
                border-bottom: 1px solid #e9ecef;
            }
            QListWidget::item:selected {
                background-color: #007bff;
                color: white;
            }
            QListWidget::item:hover {
                background-color: #f8f9fa;
            }
        """)
        
        # Populate list
        for doc_file in doc_files:
            mod_time = datetime.fromtimestamp(doc_file.stat().st_mtime)
            file_type = "PDF" if doc_file.suffix.lower() == '.pdf' else "Word"
            display_text = f"{doc_file.name}\n📅 {mod_time.strftime('%Y-%m-%d %H:%M:%S')} • 📄 {file_type}"
            
            item = QListWidgetItem(display_text)
            item.setData(Qt.UserRole, doc_file)  # Store file path
            doc_list.addItem(item)
        
        # Select first item
        if doc_files:
            doc_list.setCurrentRow(0)
        
        layout.addWidget(doc_list)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        open_btn = ModernButton("📂 Open Document", "#28a745", "#1e7e34")
        save_btn = ModernButton("💾 Save As...", "#17a2b8", "#138496")
        folder_btn = ModernButton("📁 Open Folder", "#ffc107", "#e0a800")
        cancel_btn = ModernButton("❌ Cancel", "#6c757d", "#5a6268")
        
        def open_document():
            current_item = doc_list.currentItem()
            if current_item:
                doc_file = current_item.data(Qt.UserRole)
                self.open_document_file(doc_file)
                dialog.accept()
        
        def save_document():
            current_item = doc_list.currentItem()
            if current_item:
                doc_file = current_item.data(Qt.UserRole)
                self.save_document_as(doc_file)
                dialog.accept()
        
        def open_folder():
            output_dir = Path(self.config.get('output_directory', 'filled_docs'))
            self.open_folder(output_dir)
            dialog.accept()
        
        open_btn.clicked.connect(open_document)
        save_btn.clicked.connect(save_document)
        folder_btn.clicked.connect(open_folder)
        cancel_btn.clicked.connect(dialog.reject)
        
        button_layout.addWidget(open_btn)
        button_layout.addWidget(save_btn)
        button_layout.addWidget(folder_btn)
        button_layout.addStretch()
        button_layout.addWidget(cancel_btn)
        
        layout.addLayout(button_layout)
        
        dialog.exec_()
    
    def open_document_file(self, doc_file):
        """Open document file with default application."""
        try:
            import os
            os.startfile(str(doc_file))  # Windows-specific
            self.log_message(f"📂 Opened document: {doc_file.name}")
            self.status_bar.showMessage(f"Opened: {doc_file.name}")
        except Exception as e:
            self.log_message(f"❌ Error opening document: {e}")
            self.show_error(f"Failed to open document: {e}")
    
    def save_document_as(self, doc_file):
        """Save document to a user-selected location."""
        try:
            import shutil
            
            # Determine file types based on original file
            if doc_file.suffix.lower() == '.pdf':
                file_filter = "PDF files (*.pdf);;All files (*.*)"
                default_suffix = ".pdf"
            else:
                file_filter = "Word documents (*.docx);;Word 97-2003 (*.doc);;All files (*.*)"
                default_suffix = doc_file.suffix
            
            save_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Document As",
                str(doc_file.name),
                file_filter
            )
            
            if save_path:
                # Ensure proper extension
                if not save_path.endswith(default_suffix):
                    save_path += default_suffix
                
                shutil.copy2(str(doc_file), save_path)
                self.log_message(f"💾 Document saved to: {save_path}")
                self.status_bar.showMessage(f"Saved to: {Path(save_path).name}")
                self.show_info("Success", f"Document saved successfully to:\n{save_path}")
                
        except Exception as e:
            self.log_message(f"❌ Error saving document: {e}")
            self.show_error(f"Failed to save document: {e}")
    
    def open_folder(self, folder_path):
        """Open folder in file explorer."""
        try:
            import os
            os.startfile(str(folder_path))  # Windows-specific
            self.log_message(f"📁 Opened folder: {folder_path}")
            self.status_bar.showMessage(f"Opened folder: {folder_path.name}")
        except Exception as e:
            self.log_message(f"❌ Error opening folder: {e}")
            self.show_error(f"Failed to open folder: {e}")
    
    def on_doc_type_change(self):
        """Handle document type change."""
        if self.word_radio.isChecked():
            doc_type = 'word'
            path = self.config.get('document_path', '')
        else:
            doc_type = 'pdf'
            path = self.config.get('pdf_path', '')
        
        self.config.set('document_type', doc_type)
        
        # Update path input if it exists (avoid initialization order issues)
        if hasattr(self, 'path_input'):
            self.path_input.setText(path)
        
        self.log_message(f"📄 Document type changed to: {doc_type.upper()}")
    
    def on_month_change(self):
        """Handle month selection change."""
        selected_month = self.month_combo.currentText()
        self.config.set('selected_month', selected_month)
        self.log_message(f"📅 Month changed to: {selected_month}")

    def show_info(self, title, message):
        """Show info message."""
        QMessageBox.information(self, title, message)
    
    def show_warning(self, title, message):
        """Show warning message."""
        QMessageBox.warning(self, title, message)
        
    def clear_log(self):
        """Clear the log."""
        self.log_text.clear()
        self.log_message("📝 Log cleared")
        
    def log_message(self, message):
        """Add message to log."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        
        # Only log to UI if log_text exists (avoid initialization order issues)
        if hasattr(self, 'log_text'):
            self.log_text.append(formatted_message)
            
            # Auto-scroll to bottom
            scrollbar = self.log_text.verticalScrollBar()
            scrollbar.setValue(scrollbar.maximum())
        
        # Always log to console/file with safe encoding
        try:
            # Remove emojis for logging compatibility
            safe_message = message.encode('ascii', 'ignore').decode('ascii').strip()
            if safe_message:  # Only log if there's text left after removing emojis
                logger.info(safe_message)
            else:
                logger.info("Action completed")
        except Exception:
            logger.info("Action completed")
        
    def show_error(self, message):
        """Show error message."""
        QMessageBox.critical(self, "Error", message)
        self.log_message(f"❌ Error: {message}")
        
    def on_login(self, event_data=None):
        """Handle login event."""
        self.log_message("🔑 User logged in")
        
    def on_logout(self, event_data=None):
        """Handle logout event."""
        self.log_message("🚪 User logged out")
        
    # Tray functionality removed
    
    def start_web_server(self):
        """Start the web server for mobile access."""
        if self.web_server:
            try:
                self.web_server.start_server()
                self.log_message("📱 Web server started for mobile access")
            except Exception as e:
                self.log_message(f"❌ Failed to start web server: {e}")
    
    def stop_web_server(self):
        """Stop the web server."""
        if self.web_server:
            try:
                self.web_server.stop_server()
                self.log_message("📱 Web server stopped")
            except Exception as e:
                self.log_message(f"❌ Failed to stop web server: {e}")
                
    def closeEvent(self, event):
        """Handle close event."""
        # Stop web server when closing
        if self.web_server:
            self.stop_web_server()
        event.accept()

def create_pyqt_app(config_manager, event_monitor, pdf_handler, word_handler=None, web_server=None):
    """Create and return PyQt application."""
    window = AttendancePyQtGUI(config_manager, event_monitor, pdf_handler, word_handler, web_server)
    return window
