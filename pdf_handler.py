"""PDF handling module for attendance sheet processing."""

import os
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional, Tuple
import logging

import PyPDF2
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import NameObject, TextStringObject
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import pdfplumber

logger = logging.getLogger(__name__)

class PDFHandler:
    """Handles PDF reading, filling, and generation."""
    
    def __init__(self, config_manager):
        """Initialize PDF handler.
        
        Args:
            config_manager: Configuration manager instance
        """
        self.config = config_manager
        self.last_filled_pdf = None
        
    def fill_attendance_sheet(self, time_in: datetime, time_out: datetime = None) -> Optional[str]:
        """Fill attendance sheet with timestamps.
        
        Args:
            time_in: Login timestamp
            time_out: Logout timestamp (optional)
            
        Returns:
            Path to filled PDF or None on error
        """
        pdf_path = self.config.get('pdf_path')
        
        if not pdf_path or not os.path.exists(pdf_path):
            logger.error(f"PDF not found: {pdf_path}")
            return None
        
        try:
            # Prepare field values
            field_values = self._prepare_field_values(time_in, time_out)
            
            # Try filling as form first
            output_path = self._fill_pdf_form(pdf_path, field_values)
            
            if not output_path and self.config.get('pdf_fallback.enabled', True):
                # Fallback to overlay method
                logger.info("Using fallback PDF overlay method")
                output_path = self._overlay_pdf_text(pdf_path, field_values)
            
            if output_path:
                self.last_filled_pdf = output_path
                logger.info(f"PDF filled successfully: {output_path}")
            
            return output_path
            
        except Exception as e:
            logger.error(f"Error filling PDF: {e}")
            return None
    
    def _prepare_field_values(self, time_in: datetime, time_out: datetime = None) -> Dict[str, str]:
        """Prepare field values for PDF.
        
        Args:
            time_in: Login timestamp
            time_out: Logout timestamp (optional)
            
        Returns:
            Dictionary of field names and values
        """
        date_format = self.config.get('date_format', '%Y-%m-%d')
        time_format = self.config.get('time_format', '%H:%M:%S')
        field_names = self.config.get('field_names', {})
        
        values = {}
        
        # Date field
        if 'date' in field_names:
            values[field_names['date']] = time_in.strftime(date_format)
        
        # Time in field
        if 'time_in' in field_names:
            values[field_names['time_in']] = time_in.strftime(time_format)
        
        # Time out field
        if time_out and 'time_out' in field_names:
            values[field_names['time_out']] = time_out.strftime(time_format)
        
        # Selected month field
        selected_month = self.config.get('selected_month')
        if selected_month:
            # Try common month field names
            month_field_names = ['month', 'Month', 'MONTH', 'month_year', 'Month/Year', 'period', 'Period']
            for month_field in month_field_names:
                values[month_field] = selected_month
        
        # Employee name field (keeping for backward compatibility)
        employee_name = self.config.get('employee_name')
        if employee_name and 'employee_name' in field_names:
            values[field_names['employee_name']] = employee_name
        
        return values
    
    def _fill_pdf_form(self, pdf_path: str, field_values: Dict[str, str]) -> Optional[str]:
        """Fill PDF form fields.
        
        Args:
            pdf_path: Path to source PDF
            field_values: Dictionary of field names and values
            
        Returns:
            Path to filled PDF or None on error
        """
        try:
            # Read the PDF
            reader = PdfReader(pdf_path)
            writer = PdfWriter()
            
            # Check if PDF has form fields
            if '/AcroForm' not in reader.trailer['/Root']:
                logger.warning("PDF has no form fields")
                return None
            
            # Get form fields
            fields = reader.get_form_text_fields()
            
            if not fields:
                logger.warning("No fillable fields found in PDF")
                return None
            
            logger.info(f"Found form fields: {list(fields.keys())}")
            
            # Fill the form
            for page in reader.pages:
                writer.add_page(page)
            
            # Update form field values
            writer.update_page_form_field_values(
                writer.pages[0], 
                field_values
            )
            
            # Generate output filename
            output_path = self._generate_output_path(pdf_path)
            
            # Write the filled PDF
            with open(output_path, 'wb') as output_file:
                writer.write(output_file)
            
            return output_path
            
        except Exception as e:
            logger.error(f"Error filling PDF form: {e}")
            return None
    
    def _overlay_pdf_text(self, pdf_path: str, field_values: Dict[str, str]) -> Optional[str]:
        """Overlay text on PDF using reportlab.
        
        Args:
            pdf_path: Path to source PDF
            field_values: Dictionary of field names and values
            
        Returns:
            Path to filled PDF or None on error
        """
        try:
            from PyPDF2 import PdfMerger
            from io import BytesIO
            
            # Read original PDF
            reader = PdfReader(pdf_path)
            
            # Get coordinates from config
            coordinates = self.config.get('pdf_fallback.coordinates', {})
            
            # Create overlay PDF
            packet = BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            
            # Add text at specified coordinates
            for field_name, value in field_values.items():
                if field_name in coordinates:
                    x, y = coordinates[field_name]
                    can.drawString(x, y, value)
                    logger.debug(f"Added '{value}' at ({x}, {y})")
            
            can.save()
            
            # Merge PDFs
            packet.seek(0)
            overlay = PdfReader(packet)
            writer = PdfWriter()
            
            # Add the overlay to first page
            page = reader.pages[0]
            if len(overlay.pages) > 0:
                page.merge_page(overlay.pages[0])
            writer.add_page(page)
            
            # Add remaining pages
            for i in range(1, len(reader.pages)):
                writer.add_page(reader.pages[i])
            
            # Generate output filename
            output_path = self._generate_output_path(pdf_path)
            
            # Write the result
            with open(output_path, 'wb') as output_file:
                writer.write(output_file)
            
            return output_path
            
        except Exception as e:
            logger.error(f"Error overlaying PDF text: {e}")
            return None
    
    def _generate_output_path(self, source_path: str) -> str:
        """Generate output path for filled PDF.
        
        Args:
            source_path: Path to source PDF
            
        Returns:
            Output path for filled PDF
        """
        output_dir = Path(self.config.get('output_directory', 'filled_pdfs'))
        # Only create output_dir if/when a file is actually written
        if not output_dir.exists():
            output_dir.mkdir(parents=True, exist_ok=True)
        
        # Generate filename with timestamp
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        source_name = Path(source_path).stem
        output_name = f"{source_name}_filled_{timestamp}.pdf"
        
        return str(output_dir / output_name)
    
    def generate_report_pdf(self, start_date: datetime, end_date: datetime) -> Optional[str]:
        """Generate attendance report PDF.
        
        Args:
            start_date: Report start date
            end_date: Report end date
            
        Returns:
            Path to generated report or None on error
        """
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import letter, A4
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
            from reportlab.lib.styles import getSampleStyleSheet
            from reportlab.lib.units import inch
            
            # Load event logs
            events = self._load_events_for_period(start_date, end_date)
            
            if not events:
                logger.warning("No events found for report period")
                return None
            
            # Generate report filename
            output_dir = Path(self.config.get('output_directory', 'filled_pdfs'))
            # Only create output_dir if/when a file is actually written
            
            report_name = f"attendance_report_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.pdf"
            output_path = str(output_dir / report_name)
            
            # Create PDF document
            doc = SimpleDocTemplate(output_path, pagesize=letter)
            elements = []
            
            # Add title
            styles = getSampleStyleSheet()
            title = Paragraph("Attendance Report", styles['Title'])
            elements.append(title)
            
            # Add date range
            date_range = Paragraph(
                f"Period: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}",
                styles['Normal']
            )
            elements.append(date_range)
            
            # Create table data
            data = [['Date', 'Time In', 'Time Out', 'Duration']]
            
            for event in events:
                if event['type'] == 'login':
                    # Find corresponding logout
                    logout = self._find_logout_for_login(event, events)
                    
                    row = [
                        event['date'],
                        event['time'],
                        logout['time'] if logout else 'N/A',
                        self._calculate_duration(event, logout) if logout else 'N/A'
                    ]
                    data.append(row)
            
            # Create table
            table = Table(data)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 14),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            
            elements.append(table)
            
            # Build PDF
            doc.build(elements)
            
            logger.info(f"Report generated: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"Error generating report: {e}")
            return None
    
    def _load_events_for_period(self, start_date: datetime, end_date: datetime) -> list:
        """Load events for specified period.
        
        Args:
            start_date: Period start date
            end_date: Period end date
            
        Returns:
            List of events
        """
        events = []
        log_dir = Path(self.config.get('log_directory', 'logs'))
        
        # Iterate through date range
        current_date = start_date
        while current_date <= end_date:
            log_file = log_dir / f"events_{current_date.strftime('%Y%m%d')}.json"
            
            if log_file.exists():
                try:
                    import json
                    with open(log_file, 'r') as f:
                        day_events = json.load(f)
                        
                        for event in day_events:
                            event_time = datetime.fromisoformat(event['timestamp'])
                            events.append({
                                'type': event['type'],
                                'date': event_time.strftime('%Y-%m-%d'),
                                'time': event_time.strftime('%H:%M:%S'),
                                'timestamp': event_time
                            })
                except Exception as e:
                    logger.error(f"Error loading events from {log_file}: {e}")
            
            current_date = current_date.replace(day=current_date.day + 1)
        
        return sorted(events, key=lambda x: x['timestamp'])
    
    def _find_logout_for_login(self, login_event: dict, all_events: list) -> Optional[dict]:
        """Find corresponding logout for a login event.
        
        Args:
            login_event: Login event
            all_events: All events
            
        Returns:
            Logout event or None
        """
        login_time = login_event['timestamp']
        
        for event in all_events:
            if event['type'] == 'logout' and event['timestamp'] > login_time:
                # Check if this is the same day
                if event['date'] == login_event['date']:
                    return event
                break
        
        return None
    
    def _calculate_duration(self, login_event: dict, logout_event: dict) -> str:
        """Calculate duration between login and logout.
        
        Args:
            login_event: Login event
            logout_event: Logout event
            
        Returns:
            Duration string
        """
        duration = logout_event['timestamp'] - login_event['timestamp']
        hours = int(duration.total_seconds() // 3600)
        minutes = int((duration.total_seconds() % 3600) // 60)
        
        return f"{hours:02d}:{minutes:02d}"