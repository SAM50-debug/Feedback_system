# ============================================
# run_production.py
# Waitress WSGI Server for Windows
# Optimized for 300 concurrent users
# ============================================

"""
Place this file in your project root directory
(same directory as manage.py)

Run with: python run_production.py
"""

from waitress import serve
import os
import sys

# Add project to path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

# Set Django settings module
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'feedback_system.settings')

# Import Django application
from feedback_system.wsgi import application

if __name__ == '__main__':
    print("=" * 60)
    print("Django Feedback System - Production Server")
    print("=" * 60)
    print("Server: Waitress WSGI")
    print("Optimized for: 300 concurrent users")
    print("Configuration:")
    print("  - Threads: 8")
    print("  - Channel Timeout: 60s")
    print("  - Connection Limit: 500")
    print("  - Host: 0.0.0.0 (All interfaces)")
    print("  - Port: 8000")
    print("=" * 60)
    print("\nStarting server...")
    print("Access at: http://202.164.44.102")
    print("Admin at: http://202.164.44.102/admin/")
    print("\nPress Ctrl+C to stop the server")
    print("=" * 60)
    
    try:
        serve(
            application,
            host='127.0.0.1',  # Listen on all network interfaces
            port=8000,       # Port number
            threads=8,       # Number of worker threads (handles concurrency)
            channel_timeout=60,  # Timeout for idle connections
            connection_limit=500,  # Maximum simultaneous connections
            backlog=128,     # Socket backlog
            send_bytes=18000,  # Send buffer size
            url_scheme='http',  # Use 'https' if you have SSL
            ident='FeedbackSystem/1.0',  # Server identification
        )
    except KeyboardInterrupt:
        print("\n\nServer stopped by user")
    except Exception as e:
        print(f"\n\nError starting server: {e}")
        sys.exit(1)