"""Web server for mobile control of attendance tracker."""

import os
import json
import logging
from datetime import datetime
from pathlib import Path
from threading import Thread
import socket

from flask import Flask, render_template, request, jsonify
from flask_socketio import SocketIO, emit

from config_manager import ConfigManager
from word_handler import WordHandler
from pdf_handler import PDFHandler

logger = logging.getLogger(__name__)

class AttendanceWebServer:
    """Web server for mobile control of attendance tracker."""
    
    def __init__(self, config_manager, word_handler, pdf_handler):
        """Initialize web server.
        
        Args:
            config_manager: Configuration manager instance
            word_handler: Word handler instance
            pdf_handler: PDF handler instance
        """
        self.config = config_manager
        self.word_handler = word_handler
        self.pdf_handler = pdf_handler
        # Create Flask app
        self.app = Flask(__name__, 
                        template_folder='web_templates',
                        static_folder='web_templates/static')
        self.app.config['SECRET_KEY'] = 'attendance_tracker_secret_2025'
        # Create SocketIO instance with threading mode for PyInstaller compatibility
        self.socketio = SocketIO(self.app, cors_allowed_origins="*", async_mode='threading')
        # Setup routes
        self._setup_routes()
        self._setup_socketio_events()
        # Server state
        self.server_thread = None
        self.is_running = False
        
    def _setup_routes(self):
        @self.app.route('/api/download/<filename>')
        def download_document(filename):
            """Serve a filled document for download."""
            from flask import send_from_directory, abort
            import os
            output_dir = self.config.get('output_directory', 'filled_docs')
            file_path = os.path.join(output_dir, filename)
            if not os.path.isfile(file_path):
                abort(404)
            return send_from_directory(output_dir, filename, as_attachment=True)
        """Setup Flask routes."""
        
        @self.app.route('/')
        def index():
            """Main PWA page."""
            return render_template('index.html')
        
        @self.app.route('/api/status')
        def get_status():
            """Get current system status."""
            try:
                doc_type = self.config.get('document_type', 'word')
                doc_path = self.config.get('document_path' if doc_type == 'word' else 'pdf_path', '')
                
                status = {
                    'timestamp': datetime.now().isoformat(),
                    'document_type': doc_type,
                    'document_path': Path(doc_path).name if doc_path else 'Not selected',
                    'month': self.config.get('selected_month', 'August 2025'),
                    'server_status': 'running'
                }
                return jsonify(status)
            except Exception as e:
                logger.error(f"Error getting status: {e}")
                return jsonify({'error': str(e)}), 500
        
        @self.app.route('/api/checkin', methods=['POST'])
        def manual_checkin():
            """Handle manual check-in from mobile."""
            try:
                current_time = datetime.now()
                doc_type = self.config.get('document_type', 'word')
                
                if doc_type == 'word' and self.word_handler:
                    result = self.word_handler.fill_attendance_sheet(time_in=current_time)
                elif doc_type == 'pdf' and self.pdf_handler:
                    result = self.pdf_handler.fill_attendance_sheet(time_in=current_time)
                else:
                    return jsonify({'error': 'No document handler available'}), 400
                
                if result:
                    message = f"Check-in recorded at {current_time.strftime('%H:%M:%S')}"
                    logger.info(f"Mobile check-in: {message}")
                    
                    # Emit to all connected clients
                    self.socketio.emit('attendance_update', {
                        'type': 'checkin',
                        'time': current_time.strftime('%H:%M:%S'),
                        'message': message
                    })
                    
                    return jsonify({
                        'success': True,
                        'message': message,
                        'time': current_time.strftime('%H:%M:%S'),
                        'document': Path(result).name if result else None
                    })
                else:
                    return jsonify({'error': 'Failed to record check-in'}), 500
                    
            except Exception as e:
                logger.error(f"Error during mobile check-in: {e}")
                return jsonify({'error': str(e)}), 500
        
        @self.app.route('/api/checkout', methods=['POST'])
        def manual_checkout():
            """Handle manual check-out from mobile."""
            try:
                current_time = datetime.now()
                doc_type = self.config.get('document_type', 'word')
                
                if doc_type == 'word' and self.word_handler:
                    # Try to retrieve last check-in time if available, else use current_time
                    time_in = getattr(self, 'last_checkin_time', None)
                    if time_in is None:
                        time_in = current_time
                    result = self.word_handler.fill_attendance_sheet(time_in=time_in, time_out=current_time)
                elif doc_type == 'pdf' and self.pdf_handler:
                    # Try to retrieve last check-in time if available, else use current_time
                    time_in = getattr(self, 'last_checkin_time', None)
                    if time_in is None:
                        time_in = current_time
                    result = self.pdf_handler.fill_attendance_sheet(time_in=time_in, time_out=current_time)
                else:
                    return jsonify({'error': 'No document handler available'}), 400
                
                if result:
                    message = f"Check-out recorded at {current_time.strftime('%H:%M:%S')}"
                    logger.info(f"Mobile check-out: {message}")
                    
                    # Emit to all connected clients
                    self.socketio.emit('attendance_update', {
                        'type': 'checkout',
                        'time': current_time.strftime('%H:%M:%S'),
                        'message': message
                    })
                    
                    return jsonify({
                        'success': True,
                        'message': message,
                        'time': current_time.strftime('%H:%M:%S'),
                        'document': Path(result).name if result else None
                    })
                else:
                    return jsonify({'error': 'Failed to record check-out'}), 500
                    
            except Exception as e:
                logger.error(f"Error during mobile check-out: {e}")
                return jsonify({'error': str(e)}), 500
        
        @self.app.route('/api/recent-documents')
        def get_recent_documents():
            """Get list of recent filled documents."""
            try:
                output_dir = Path(self.config.get('output_directory', 'filled_docs'))
                
                if not output_dir.exists():
                    return jsonify({'documents': []})
                
                # Get recent documents
                doc_files = list(output_dir.glob("*.pdf")) + list(output_dir.glob("*.docx"))
                doc_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
                
                documents = []
                for doc_file in doc_files[:5]:  # Latest 5 documents
                    mod_time = datetime.fromtimestamp(doc_file.stat().st_mtime)
                    documents.append({
                        'name': doc_file.name,
                        'modified': mod_time.strftime('%Y-%m-%d %H:%M:%S'),
                        'type': 'PDF' if doc_file.suffix.lower() == '.pdf' else 'Word',
                        'size': f"{doc_file.stat().st_size / 1024:.1f} KB"
                    })
                
                return jsonify({'documents': documents})
                
            except Exception as e:
                logger.error(f"Error getting recent documents: {e}")
                return jsonify({'error': str(e)}), 500
        
        @self.app.route('/manifest.json')
        def manifest():
            """PWA manifest file."""
            manifest_data = {
                "name": "Khaled's Attendance Tracker",
                "short_name": "Attendance",
                "description": "Mobile control for attendance tracking",
                "start_url": "/",
                "display": "standalone",
                "background_color": "#ffffff",
                "theme_color": "#007bff",
                "orientation": "portrait",
                "icons": [
                    {
                        "src": "/static/icon-192.png",
                        "sizes": "192x192",
                        "type": "image/png",
                        "purpose": "any maskable"
                    },
                    {
                        "src": "/static/icon-512.png", 
                        "sizes": "512x512",
                        "type": "image/png",
                        "purpose": "any maskable"
                    }
                ]
            }
            return jsonify(manifest_data)
    
    def _setup_socketio_events(self):
        """Setup SocketIO events for real-time communication."""
        
        @self.socketio.on('connect')
        def handle_connect():
            """Handle client connection."""
            logger.info("Mobile client connected")
            emit('status', {'message': 'Connected to attendance tracker'})
        
        @self.socketio.on('disconnect')
        def handle_disconnect():
            """Handle client disconnection."""
            logger.info("Mobile client disconnected")
        
        @self.socketio.on('ping')
        def handle_ping():
            """Handle ping from client."""
            emit('pong', {'timestamp': datetime.now().isoformat()})
    
    def get_local_ip(self):
        """Get the local IP address."""
        try:
            # Connect to a remote server to determine local IP
            sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            sock.connect(('8.8.8.8', 80))
            local_ip = sock.getsockname()[0]
            sock.close()
            return local_ip
        except Exception:
            return '127.0.0.1'
    
    def start_server(self, host='0.0.0.0', port=5000):
        """Start the web server in a separate thread."""
        if self.is_running:
            logger.warning("Web server is already running")
            return
        
        def run_server():
            logger.info(f"Starting web server on {host}:{port}")
            self.socketio.run(self.app, host=host, port=port, debug=False, allow_unsafe_werkzeug=True)
        
        self.server_thread = Thread(target=run_server, daemon=True)
        self.server_thread.start()
        self.is_running = True
        
        # Get and log access URLs
        local_ip = self.get_local_ip()
        print("\n" + "="*60)
        print("🚀 ATTENDANCE TRACKER WEB SERVER STARTED")
        print("="*60)
        print(f"📱 Local Access: http://localhost:{port}")
        print(f"🌐 Mobile Access: http://{local_ip}:{port}")
        print("="*60)
        print("📋 Instructions for iPhone PWA:")
        print("1. Connect your iPhone to the same WiFi network")
        print(f"2. Open Safari and go to: http://{local_ip}:{port}")
        print("3. Tap the Share button (📤) at the bottom")
        print("4. Select 'Add to Home Screen' from the menu")
        print("5. Tap 'Add' to install the PWA on your home screen")
        print("6. The app will now work like a native iPhone app!")
        print("="*60)
        print("🔧 To stop the server: Close the desktop application")
        print("="*60 + "\n")
        
        logger.info(f"Web server started!")
        logger.info(f"Local access: http://localhost:{port}")
        logger.info(f"Mobile access: http://{local_ip}:{port}")
        
        return local_ip, port
    
    def stop_server(self):
        """Stop the web server."""
        if not self.is_running:
            return
        
        self.is_running = False
        logger.info("Web server stopped")
    
    def get_server_url(self):
        """Get the server URL for mobile access."""
        if not self.is_running:
            return None
        
        local_ip = self.get_local_ip()
        return f"http://{local_ip}:5000"

def create_web_server(config_manager, word_handler, pdf_handler):
    """Create and return web server instance."""
    return AttendanceWebServer(config_manager, word_handler, pdf_handler)
