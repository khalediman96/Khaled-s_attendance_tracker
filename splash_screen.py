"""Splash screen for attendance tracker application."""

import sys
from PyQt5.QtWidgets import QApplication, QSplashScreen, QLabel, QVBoxLayout, QWidget, QProgressBar
from PyQt5.QtCore import Qt, QTimer, pyqtSignal
from PyQt5.QtGui import QPixmap, QPainter, QFont, QColor, QLinearGradient, QBrush
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

class ModernSplashScreen(QSplashScreen):
    """Modern splash screen with custom branding."""
    
    progress_updated = pyqtSignal(int, str)
    
    def __init__(self):
        # Create custom splash pixmap
        pixmap = self.create_splash_pixmap()
        super().__init__(pixmap, Qt.WindowStaysOnTopHint)
        
        self.setWindowFlags(self.windowFlags() | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        # Setup progress tracking
        self.progress = 0
        self.status_text = "Initializing..."
        
        # Connect progress signal
        self.progress_updated.connect(self.update_progress)
        
        logger.info("Splash screen initialized")
    
    def create_splash_pixmap(self):
        """Create custom splash screen pixmap."""
        width, height = 800, 600
        pixmap = QPixmap(width, height)
        pixmap.fill(Qt.transparent)
        
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Background gradient
        gradient = QLinearGradient(0, 0, 0, height)
        gradient.setColorAt(0, QColor(102, 126, 234))  # Blue
        gradient.setColorAt(1, QColor(118, 75, 162))   # Purple
        
        brush = QBrush(gradient)
        painter.fillRect(0, 0, width, height, brush)
        
        # Semi-transparent overlay
        overlay_gradient = QLinearGradient(0, 0, 0, height)
        overlay_gradient.setColorAt(0, QColor(255, 255, 255, 30))
        overlay_gradient.setColorAt(1, QColor(255, 255, 255, 5))
        
        overlay_brush = QBrush(overlay_gradient)
        painter.fillRect(0, 0, width, height, overlay_brush)
        
        # Load and draw logo if exists
        logo_path = Path("assets/logo.png")
        if logo_path.exists():
            logo_pixmap = QPixmap(str(logo_path))
            logo_size = 120
            logo_scaled = logo_pixmap.scaled(logo_size, logo_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_x = (width - logo_size) // 2
            logo_y = height // 2 - 150
            painter.drawPixmap(logo_x, logo_y, logo_scaled)
        
        # Title text
        painter.setPen(QColor(255, 255, 255))
        title_font = QFont("Segoe UI", 32, QFont.Bold)
        painter.setFont(title_font)
        
        title_text = "Khaled's Attendance Tracker"
        title_rect = painter.fontMetrics().boundingRect(title_text)
        title_x = (width - title_rect.width()) // 2
        title_y = height // 2 - 50
        painter.drawText(title_x, title_y, title_text)
        
        # Subtitle
        subtitle_font = QFont("Segoe UI", 16)
        painter.setFont(subtitle_font)
        painter.setPen(QColor(255, 255, 255, 200))
        
        subtitle_text = "Professional Time Management Solution"
        subtitle_rect = painter.fontMetrics().boundingRect(subtitle_text)
        subtitle_x = (width - subtitle_rect.width()) // 2
        subtitle_y = title_y + 50
        painter.drawText(subtitle_x, subtitle_y, subtitle_text)
        
        # Version info
        version_font = QFont("Segoe UI", 12)
        painter.setFont(version_font)
        painter.setPen(QColor(255, 255, 255, 180))
        
        version_text = "Version 2.0 - Modern PyQt Edition with Mobile PWA"
        version_rect = painter.fontMetrics().boundingRect(version_text)
        version_x = (width - version_rect.width()) // 2
        version_y = subtitle_y + 40
        painter.drawText(version_x, version_y, version_text)
        
        painter.end()
        return pixmap
    
    def update_progress(self, value, message):
        """Update progress and status message."""
        self.progress = value
        self.status_text = message
        
        # Show message on splash screen
        self.showMessage(
            f"{message}\n\nProgress: {value}%",
            Qt.AlignBottom | Qt.AlignCenter,
            QColor(255, 255, 255)
        )
        
        # Force update
        QApplication.processEvents()
        
        logger.info(f"Splash progress: {value}% - {message}")
    
    def set_progress(self, value, message=""):
        """Set progress value and optional message."""
        self.progress_updated.emit(value, message)
    
    def finish_loading(self):
        """Complete the loading process."""
        self.set_progress(100, "Loading complete!")
        QTimer.singleShot(500, self.close)  # Auto-close after 500ms

class SplashScreenManager:
    """Manager for splash screen operations."""
    
    def __init__(self):
        self.splash = None
        self.steps = [
            ("Initializing configuration...", 10),
            ("Loading document handlers...", 25),
            ("Setting up event monitoring...", 40),
            ("Initializing web server...", 60),
            ("Creating user interface...", 80),
            ("Finalizing setup...", 95),
            ("Ready to start!", 100)
        ]
        self.current_step = 0
        
    def show_splash(self):
        """Show the splash screen."""
        self.splash = ModernSplashScreen()
        self.splash.show()
        QApplication.processEvents()
        return self.splash
    
    def next_step(self, custom_message=None):
        """Move to the next loading step."""
        try:
            if self.splash and self.current_step < len(self.steps):
                message, progress = self.steps[self.current_step]
                if custom_message:
                    message = custom_message
                
                self.splash.set_progress(progress, message)
                self.current_step += 1
                
                # Small delay for visual effect
                QTimer.singleShot(200, lambda: None)
        except RuntimeError:
            # Splash screen was deleted, ignore
            pass
    
    def update_step(self, message, progress=None):
        """Update current step with custom message and progress."""
        try:
            if self.splash:
                if progress is None and self.current_step < len(self.steps):
                    _, progress = self.steps[self.current_step]
                elif progress is None:
                    progress = 100
                
                self.splash.set_progress(progress, message)
        except RuntimeError:
            # Splash screen was deleted, ignore
            pass
    
    def finish(self):
        """Finish the splash screen."""
        if self.splash:
            self.splash.finish_loading()
            self.splash = None

def create_splash_screen():
    """Create and return splash screen manager."""
    return SplashScreenManager()
