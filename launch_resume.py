#!/usr/bin/env python3
"""
=============================================================================
  SANSKAR YADAV - 3D ANIMATED INTERACTIVE RESUME LAUNCHER
  Data Science Student | Python & Machine Learning | AWS Cloud Practitioner
=============================================================================
"""

import http.server
import socketserver
import webbrowser
import os
import sys
import time
import socket

DEFAULT_PORT = 8000

def find_available_port(start_port=8000, max_attempts=50):
    for port in range(start_port, start_port + max_attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("", port))
                return port
            except OSError:
                continue
    return start_port

def main():
    # Change working directory to the directory where this script is located
    base_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(base_dir)

    port = find_available_port(DEFAULT_PORT)
    url = f"http://localhost:{port}/index.html"

    banner = rf"""
=============================================================================
  ____                  _                __   __          _             
 / ___|  __ _ _ __  ___| | ____ _ _ __   \ \ / /_ _  __| | __ ___   __
 \___ \ / _` | '_ \/ __| |/ / _` | '__|   \ V / _` |/ _` |/ _` \ \ / /
  ___) | (_| | | | \__ \   < (_| | |       | | (_| | (_| | (_| |\ V / 
 |____/ \__,_|_| |_|___/_|\_\__,_|_|       |_|\__,_|\__,_|\__,_| \_/  

  DATA SCIENCE STUDENT | PYTHON & MACHINE LEARNING | AWS CLOUD
  AAFT College, Noida  |  MJPRU Bareilly  |  MongoDB Certified
=============================================================================
 [✓] Local 3D Web Server Started at:
     >> {url} <<

 [✓] Resume Files Ready in Folder:
     - 3D Animated Resume : index.html
     - Word Document      : Sanskar_Yadav_Resume.docx
     - PDF Document       : Sanskar_Yadav_Resume.pdf

 [!] Opening your default web browser automatically...
 [!] Press CTRL + C in this terminal window to stop the server anytime.
=============================================================================
"""
    print(banner)

    # Open web browser after a tiny pause to let server socket initialize
    def open_browser():
        time.sleep(0.8)
        webbrowser.open(url)

    import threading
    threading.Thread(target=open_browser, daemon=True).start()

    # Create HTTP Server handler
    Handler = http.server.SimpleHTTPRequestHandler
    Handler.extensions_map.update({
        '.js': 'application/javascript',
        '.css': 'text/css',
        '.html': 'text/html',
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        '.pdf': 'application/pdf',
    })

    try:
        with socketserver.TCPServer(("", port), Handler) as httpd:
            httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n[+] 3D Resume server stopped gracefully. Have a great day, Sanskar!\n")
        sys.exit(0)

if __name__ == "__main__":
    main()
