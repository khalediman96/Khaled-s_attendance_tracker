from config_manager import ConfigManager
from word_handler import WordHandler
from pdf_handler import PDFHandler
from web_server import create_web_server

# Initialize config and handlers
config = ConfigManager()
word_handler = WordHandler(config)
pdf_handler = PDFHandler(config)

# Create the web server and expose the Flask app
web_server = create_web_server(config, word_handler, pdf_handler)
app = web_server.app  # This exposes the Flask app for Gunicorn

if __name__ == "__main__":
    # Start the web server when running directly
    web_server.start_server(host="0.0.0.0", port=5000)
    
    # Keep the main thread alive
    import time
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nShutting down web server...")