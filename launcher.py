import os
import sys
import time
import socket
import re
import subprocess
import webbrowser

# Pastikan output konsol selalu UTF-8 dan tidak buffering
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace', line_buffering=True)
    except Exception:
        pass
if hasattr(sys.stderr, 'reconfigure'):
    try:
        sys.stderr.reconfigure(encoding='utf-8', errors='replace', line_buffering=True)
    except Exception:
        pass

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
CLOUDFLARED_BIN = os.path.join(CURRENT_DIR, "bin", "cloudflared.exe")
LOG_FILE = os.path.join(CURRENT_DIR, "tunnel.log")
FLASK_LOG = os.path.join(CURRENT_DIR, "flask.log")
PORT = 5000

def kill_existing_on_port(port):
    try:
        cmd = f'netstat -ano | findstr :{port}'
        output = subprocess.check_output(cmd, shell=True).decode(errors='ignore')
        for line in output.strip().splitlines():
            parts = line.split()
            if len(parts) >= 5 and 'LISTENING' in parts:
                pid = parts[-1]
                if pid != '0':
                    subprocess.run(f'taskkill /F /PID {pid}', shell=True, capture_output=True)
    except Exception:
        pass

def get_local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()
    return ip

def ensure_cloudflared():
    if os.path.exists(CLOUDFLARED_BIN):
        return True
    print(" [i] Binary cloudflared.exe belum ada, mengunduh otomatis...", flush=True)
    os.makedirs(os.path.join(CURRENT_DIR, "bin"), exist_ok=True)
    url = "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe"
    try:
        import urllib.request
        urllib.request.urlretrieve(url, CLOUDFLARED_BIN)
        print(" [i] Binary cloudflared.exe siap digunakan!", flush=True)
        return True
    except Exception as e:
        print(f" [!] Gagal mengunduh cloudflared: {e}", flush=True)
        return False

def start_flask():
    python_exe = sys.executable
    app_py = os.path.join(CURRENT_DIR, "app.py")
    env = os.environ.copy()
    env["PORT"] = str(PORT)
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    f_log = open(FLASK_LOG, "w", encoding="utf-8")
    return subprocess.Popen([python_exe, app_py], cwd=CURRENT_DIR, env=env, stdout=f_log, stderr=subprocess.STDOUT)

def main():
    print("\n" + "=" * 70, flush=True)
    print("   [+] SISTEM GENERATOR BUKTI FISIK & MODUL AJAR (AI-AUTODOC)  ", flush=True)
    print("         1-Click Launcher (Flask Engine + Cloudflare Tunnel)", flush=True)
    print("=" * 70, flush=True)
    
    # 0. Bersihkan port 5000 jika ada proses lama tertinggal
    kill_existing_on_port(PORT)
    
    # Bersihkan file log lama
    for f in [LOG_FILE, FLASK_LOG]:
        if os.path.exists(f):
            try:
                os.remove(f)
            except Exception:
                pass

    print(" [1/3] Menyalakan Web Server Flask (Backend Python)...", flush=True)
    flask_proc = start_flask()
    time.sleep(2.0)

    # Cek apakah Flask langsung crash saat startup
    if flask_proc.poll() is not None:
        print("\n [!] ERROR: Web server Flask gagal dijalankan!", flush=True)
        if os.path.exists(FLASK_LOG):
            with open(FLASK_LOG, "r", encoding="utf-8", errors="ignore") as f:
                print(f.read())
        input("\nTekan [ENTER] untuk keluar...")
        return
    
    local_ip = get_local_ip()
    tunnel_url = None
    cloudflared_proc = None
    
    if ensure_cloudflared():
        print(" [2/3] Menghubungkan Cloudflare Tunnel (Akses Luar / HP Tanpa Batas)...", flush=True)
        try:
            cloudflared_proc = subprocess.Popen(
                [CLOUDFLARED_BIN, "tunnel", "--url", f"http://127.0.0.1:{PORT}", "--logfile", LOG_FILE],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                cwd=CURRENT_DIR
            )
            
            # Polling file log tunnel untuk menangkap URL trycloudflare.com
            start_time = time.time()
            while time.time() - start_time < 20:
                if os.path.exists(LOG_FILE):
                    try:
                        with open(LOG_FILE, "r", encoding="utf-8", errors="ignore") as f:
                            content = f.read()
                            match = re.search(r'https://[a-zA-Z0-9-]+\.trycloudflare\.com', content)
                            if match:
                                tunnel_url = match.group(0)
                                break
                    except Exception:
                        pass
                time.sleep(0.5)
        except Exception as e:
            print(f" [!] Catatan Cloudflare: {e}", flush=True)
    else:
        print(" [!] File bin/cloudflared.exe tidak ditemukan. Berjalan dalam mode jaringan lokal.", flush=True)

    print(" [3/3] Membuka browser default otomatis...", flush=True)
    try:
        webbrowser.open(f"http://localhost:{PORT}")
    except Exception:
        pass

    print("\n" + "=" * 70, flush=True)
    print("  [OK] SERVER & SISTEM SIAP! BUKA SALAH SATU LINK DI BAWAH:", flush=True)
    print("=" * 70, flush=True)
    print(f"  * PC Ini (Localhost)       : http://localhost:{PORT}", flush=True)
    print(f"  * HP/Laptop (Satu Wi-Fi)   : http://{local_ip}:{PORT}", flush=True)
    if tunnel_url:
        print(f"  * Online / Luar Jaringan   : {tunnel_url}", flush=True)
        print("     (Link aman HTTPS di atas bisa dibuka langsung dari HP di mana saja!)", flush=True)
    else:
        print(f"  * Online / Luar Jaringan   : Sedang proses / Cek http://{local_ip}:{PORT}", flush=True)
    print("=" * 70, flush=True)
    print("  Tips: Halaman web otomatis dibuka di browser PC Anda.", flush=True)
    print("  Tekan [CTRL + C] di jendela ini kapan saja untuk mematikan server.", flush=True)
    print("=" * 70 + "\n", flush=True)

    try:
        while True:
            time.sleep(1)
            if flask_proc.poll() is not None:
                print("\n [!] Web server Flask terhenti.", flush=True)
                if os.path.exists(FLASK_LOG):
                    with open(FLASK_LOG, "r", encoding="utf-8", errors="ignore") as f:
                        print(f.read())
                input("\nTekan [ENTER] untuk keluar...")
                break
    except KeyboardInterrupt:
        print("\n [i] Mematikan server dan menghentikan tunnel...", flush=True)
    finally:
        if flask_proc:
            try:
                flask_proc.terminate()
            except Exception:
                pass
        if cloudflared_proc:
            try:
                cloudflared_proc.terminate()
            except Exception:
                pass
        try:
            subprocess.run("taskkill /F /IM cloudflared.exe", shell=True, capture_output=True)
        except Exception:
            pass
        print(" [OK] Semua proses dimatikan dengan bersih. Sampai jumpa!\n", flush=True)

if __name__ == '__main__':
    main()
