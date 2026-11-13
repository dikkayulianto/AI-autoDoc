import os
import sys
import time
import socket
import re
import subprocess
import threading
import webbrowser

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
CLOUDFLARED_BIN = os.path.join(CURRENT_DIR, "bin", "cloudflared.exe")
PORT = 5000

def get_local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('10.255.255.255', 1))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()
    return ip

def ensure_cloudflared():
    if os.path.exists(CLOUDFLARED_BIN):
        return True
    print(" [i] Binary cloudflared.exe belum ada, mengunduh otomatis...")
    os.makedirs(os.path.join(CURRENT_DIR, "bin"), exist_ok=True)
    url = "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe"
    try:
        import urllib.request
        urllib.request.urlretrieve(url, CLOUDFLARED_BIN)
        print(" [i] Binary cloudflared.exe siap digunakan!")
        return True
    except Exception as e:
        print(f" [!] Gagal mengunduh cloudflared: {e}")
        return False

def start_flask():
    python_exe = sys.executable
    app_py = os.path.join(CURRENT_DIR, "app.py")
    env = os.environ.copy()
    env["PORT"] = str(PORT)
    return subprocess.Popen([python_exe, app_py], cwd=CURRENT_DIR, env=env)

def main():
    print("\n" + "=" * 68)
    print("      🏛️  SISTEM GENERATOR BUKTI FISIK JABAR BERAKSI (OPSI 1)  ")
    print("          Local Python Server & Cloudflare Remote Tunnel")
    print("=" * 68)
    print(" [1/3] Menyalakan Web Server Flask (Python-docx native)...")
    
    flask_proc = start_flask()
    time.sleep(1.5)
    
    local_ip = get_local_ip()
    tunnel_url = None
    cloudflared_proc = None
    
    if ensure_cloudflared():
        print(" [2/3] Membuka Terowongan Online Cloudflare Tunnel (Akses HP/Luar)...")
        try:
            cloudflared_proc = subprocess.Popen(
                [CLOUDFLARED_BIN, "tunnel", "--url", f"http://127.0.0.1:{PORT}"],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                cwd=CURRENT_DIR
            )
            
            # Read lines to find the trycloudflare url
            start_time = time.time()
            while time.time() - start_time < 12:
                line = cloudflared_proc.stdout.readline()
                if not line:
                    break
                match = re.search(r'https://[a-zA-Z0-9-]+\.trycloudflare\.com', line)
                if match:
                    tunnel_url = match.group(0)
                    break
        except Exception as e:
            print(f" [!] Gagal membuka cloudflared: {e}")
    else:
        print(" [!] File bin/cloudflared.exe tidak ditemukan. Mode hanya jaringan lokal.")

    print(" [3/3] Membuka peramban (browser) otomatis...")
    try:
        webbrowser.open(f"http://localhost:{PORT}")
    except Exception:
        pass

    print("\n" + "#" * 68)
    print("  ✅ SERVER SIAP & AKTIF! SILAKAN BUKA SALAH SATU TAUTAN BERIKUT:")
    print("#" * 68)
    print(f"  💻 PC Ini (Lokal)           : http://localhost:{PORT}")
    print(f"  📶 HP / Laptop (Satu Wi-Fi) : http://{local_ip}:{PORT}")
    if tunnel_url:
        print(f"  🌐 Online / Luar Rumah      : {tunnel_url}")
        print("     (Tautan aman HTTPS di atas bisa dibuka langsung dari HP di mana saja!)")
    print("#" * 68)
    print("\n  👉 Tekan [CTRL + C] di jendela ini kapan saja untuk mematikan server.\n")

    try:
        while True:
            time.sleep(1)
            if flask_proc.poll() is not None:
                print("Server Flask berhenti.")
                break
    except KeyboardInterrupt:
        print("\nMematikan server dan terowongan...")
    finally:
        if flask_proc:
            flask_proc.terminate()
        if cloudflared_proc:
            cloudflared_proc.terminate()
        print("Selesai. Sampai jumpa!")

if __name__ == '__main__':
    main()
