#!/usr/bin/env python3
"""
Start the PO Tracker application.
Run:  python3 run.py
Then open:  http://localhost:5000
"""
import os
import sys
import webbrowser
import threading
from app import app

HOST = "127.0.0.1"
PORT = 5000


def open_browser():
    webbrowser.open(f"http://{HOST}:{PORT}")


if __name__ == "__main__":
    print(f"\n  PO Tracker is starting...")
    print(f"  Open your browser at: http://{HOST}:{PORT}\n")

    # Open browser automatically after a short delay
    t = threading.Timer(1.2, open_browser)
    t.daemon = True
    t.start()

    app.run(host=HOST, port=PORT, debug=False)
