from config_manager import ConfigManager
from word_handler import WordHandler
from pdf_handler import PDFHandler
from web_server import create_web_server

if __name__ == "__main__":
    # Initialize config and handlers
    config = ConfigManager()
    word_handler = WordHandler(config)
    pdf_handler = PDFHandler(config)

    # Create and start the web server
    web_server = create_web_server(config, word_handler, pdf_handler)
    web_server.start_server(host="0.0.0.0", port=5000)
    
    # Keep the main thread alive
    import time
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nShutting down web server...")
