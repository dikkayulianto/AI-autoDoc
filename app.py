import os
import io
import sys
import re
import json
import socket
import zipfile
import requests
from flask import Flask, request, send_file, jsonify, send_from_directory

# Ensure engine is in python path
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

import engine.document_generator as dg

app = Flask(__name__, static_folder=current_dir, static_url_path='')

def get_local_ip():
    """Finds the local IP address of this machine on the local network (Wi-Fi/LAN)"""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        # doesn't even have to be reachable
        s.connect(('10.255.255.255', 1))
        IP = s.getsockname()[0]
    except Exception:
        IP = '127.0.0.1'
    finally:
        s.close()
    return IP

@app.route('/')
def index():
    return send_from_directory(current_dir, 'index.html')

@app.route('/<path:filename>')
def serve_static(filename):
    return send_from_directory(current_dir, filename)

@app.route('/api/health')
def health():
    return jsonify({
        "status": "online",
        "engine": "Python docx generator",
        "version": "2.0.0",
        "indicators_count": len(dg.INDICATORS_27)
    })

@app.route('/api/network-info')
def network_info():
    ip = get_local_ip()
    port = int(os.environ.get('PORT', 5000))
    return jsonify({
        "local_ip": ip,
        "port": port,
        "local_url": f"http://localhost:{port}",
        "network_url": f"http://{ip}:{port}"
    })

@app.route('/api/indicators')
def get_indicators():
    data = []
    for item in dg.INDICATORS_27:
        data.append({
            "num": item[0],
            "title": item[1],
            "evidence": item[2],
            "target": item[3]
        })
    return jsonify(data)

@app.route('/api/generate-main', methods=['POST'])
def generate_main():
    sch = request.get_json(silent=True) or {}
    bio = dg.generate_main_instrument_docx(sch)
    clean_name = (sch.get('name') or 'SDN_1_KESAMBI').replace(' ', '_')
    filename = f"Instrumen_Pemantauan_Jabar_Beraksi_27Poin_{clean_name}_2026.docx"
    return send_file(
        bio,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

@app.route('/api/generate-indicator/<int:poin>', methods=['POST'])
def generate_indicator(poin):
    sch = request.get_json(silent=True) or {}
    bio = dg.generate_indicator_docx(poin, sch)
    clean_name = (sch.get('name') or 'SDN_1_KESAMBI').replace(' ', '_')
    filename = f"Bukti_Fisik_Poin_{poin}_{clean_name}.docx"
    return send_file(
        bio,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

@app.route('/api/download-all-zip', methods=['POST'])
def download_all_zip():
    sch = request.get_json(silent=True) or {}
    z_buf = io.BytesIO()
    clean_name = (sch.get('name') or 'SDN_1_KESAMBI').replace(' ', '_')
    
    with zipfile.ZipFile(z_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        # 1. Main Instrument
        main_bio = dg.generate_main_instrument_docx(sch)
        zf.writestr(f"00_Instrumen_Pemantauan_Jabar_Beraksi_27Poin_{clean_name}.docx", main_bio.getvalue())
        
        # 2. All 27 Indicators
        for item in dg.INDICATORS_27:
            num = item[0]
            title_slug = item[1].replace(' ', '_')[:30]
            ind_bio = dg.generate_indicator_docx(num, sch)
            zf.writestr(f"Poin_{num:02d}_{title_slug}.docx", ind_bio.getvalue())
            
    z_buf.seek(0)
    zip_filename = f"Paket_Lengkap_Jabar_Beraksi_27Poin_{clean_name}.zip"
    return send_file(
        z_buf,
        as_attachment=True,
        download_name=zip_filename,
        mimetype='application/zip'
    )

def is_text_model(name):
    low = name.lower()
    for bad in ['tts', 'audio', 'embedding', 'imagen', 'aqa', 'realtime', 'whisper']:
        if bad in low:
            return False
    return True

@app.route('/api/test-key', methods=['POST'])
def test_key():
    data = request.get_json(silent=True) or {}
    api_key = (data.get('apiKey') or os.environ.get('GEMINI_API_KEY') or '').strip()
    if not api_key:
        return jsonify({"success": False, "error": "API Key tidak boleh kosong"}), 400
        
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
        res = requests.get(url, timeout=10)
        if not res.ok:
            err_data = res.json().get('error', {}) if res.content else {}
            msg = err_data.get('message', res.reason)
            return jsonify({"success": False, "error": f"Google API Error ({res.status_code}): {msg}"}), res.status_code
            
        models = res.json().get('models', [])
        valid_models = []
        for m in models:
            m_name = m.get('name', '').replace('models/', '')
            methods = m.get('supportedGenerationMethods', [])
            if 'generateContent' in methods and is_text_model(m_name):
                valid_models.append(m_name)
                
        return jsonify({
            "success": True,
            "models_count": len(valid_models),
            "models": valid_models
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/generate-ai-text', methods=['POST'])
def generate_ai_text():
    data = request.get_json(silent=True) or {}
    api_key = (data.get('apiKey') or os.environ.get('GEMINI_API_KEY') or '').strip()
    topic = (data.get('topic') or "Implementasi Program Pembiasaan Karakter").strip()
    evidence_types = data.get('evidenceTypes') or "Surat Tugas, Laporan Kegiatan, Jurnal Log 7 Kolom, Lembar Foto Dokumentasi, Lembar Pengesahan"
    custom_prompt = data.get('customPrompt') or ""
    chosen_model = data.get('model') or "auto"
    sch = data.get('sch') or {}
    
    school_name = sch.get('name') or 'SD NEGERI 1 KESAMBI'
    head_name = sch.get('head') or sch.get('ks') or 'Drs. H. Mulyono, M.Pd.'
    supervisor_name = sch.get('supervisor') or 'Ari Haerani, S.Pd.SD, M.Pd'
    year = str(sch.get('year') or '2026')
    
    from engine.smart_content_engine import build_system_prompt, get_smart_fallback, is_modul_ajar_topic
    
    is_modul = is_modul_ajar_topic(topic)
    system_prompt = build_system_prompt(topic, evidence_types, custom_prompt, school_name, head_name, supervisor_name)
    
    candidates = []
    if chosen_model and chosen_model != 'auto' and is_text_model(chosen_model):
        candidates.append(chosen_model)
    candidates.extend(['gemini-1.5-flash', 'gemini-2.0-flash', 'gemini-1.5-pro'])
    
    raw_ai_text = None
    source = "smart_engine"
    if api_key:
        headers = {'Content-Type': 'application/json'}
        payload = {
            "contents": [{
                "parts": [{"text": system_prompt}]
            }]
        }
        for m in candidates:
            for ver in ['v1beta', 'v1']:
                try:
                    url = f"https://generativelanguage.googleapis.com/{ver}/models/{m}:generateContent?key={api_key}"
                    resp = requests.post(url, headers=headers, json=payload, timeout=35)
                    if resp.ok:
                        data_res = resp.json()
                        candidates_list = data_res.get('candidates', [])
                        if candidates_list:
                            raw_ai_text = candidates_list[0].get('content', {}).get('parts', [{}])[0].get('text')
                            if raw_ai_text:
                                source = f"gemini ({m})"
                                break
                except Exception as ex:
                    print(f"Error calling {m} ({ver}): {ex}")
            if raw_ai_text:
                break

    if not raw_ai_text:
        raw_ai_text = get_smart_fallback(topic, evidence_types, school_name, head_name, supervisor_name, year)

    return jsonify({
        "success": True,
        "content": raw_ai_text,
        "is_modul": is_modul,
        "source": source
    })

@app.route('/api/generate-ai', methods=['POST'])
def generate_ai():
    data = request.get_json(silent=True) or {}
    api_key = (data.get('apiKey') or os.environ.get('GEMINI_API_KEY') or '').strip()
    topic = (data.get('topic') or "Implementasi Program Pembiasaan Karakter").strip()
    evidence_types = data.get('evidenceTypes') or "Surat Tugas, Laporan Kegiatan, Jurnal Log 7 Kolom, Lembar Foto Dokumentasi, Lembar Pengesahan"
    custom_prompt = data.get('customPrompt') or ""
    chosen_model = data.get('model') or "auto"
    sch = data.get('sch') or {}
    
    school_name = sch.get('name') or 'SD NEGERI 1 KESAMBI'
    head_name = sch.get('head') or sch.get('ks') or 'Drs. H. Mulyono, M.Pd.'
    supervisor_name = sch.get('supervisor') or 'Ari Haerani, S.Pd.SD, M.Pd'
    year = str(sch.get('year') or '2026')
    
    from engine.smart_content_engine import build_system_prompt, get_smart_fallback, is_modul_ajar_topic
    
    is_modul = is_modul_ajar_topic(topic)
    system_prompt = build_system_prompt(topic, evidence_types, custom_prompt, school_name, head_name, supervisor_name)
    
    # Priority candidate models (strictly text only)
    candidates = []
    if chosen_model and chosen_model != 'auto' and is_text_model(chosen_model):
        candidates.append(chosen_model)
    candidates.extend(['gemini-1.5-flash', 'gemini-2.0-flash', 'gemini-1.5-pro'])
    
    raw_ai_text = None
    if api_key:
        headers = {'Content-Type': 'application/json'}
        payload = {
            "contents": [{
                "parts": [{"text": system_prompt}]
            }]
        }
        
        for m in candidates:
            for ver in ['v1beta', 'v1']:
                try:
                    url = f"https://generativelanguage.googleapis.com/{ver}/models/{m}:generateContent?key={api_key}"
                    resp = requests.post(url, headers=headers, json=payload, timeout=35)
                    if resp.ok:
                        data_res = resp.json()
                        candidates_list = data_res.get('candidates', [])
                        if candidates_list:
                            raw_ai_text = candidates_list[0].get('content', {}).get('parts', [{}])[0].get('text')
                            if raw_ai_text:
                                break
                except Exception as ex:
                    print(f"Error calling {m} ({ver}): {ex}")
            if raw_ai_text:
                break

    # If API call failed or no API key, use the intelligent contextual generator
    if not raw_ai_text:
        raw_ai_text = get_smart_fallback(topic, evidence_types, school_name, head_name, supervisor_name, year)

    bio = dg.generate_custom_ai_docx(topic, evidence_types, raw_ai_text, sch)
    clean_topic = re.sub(r'[^a-zA-Z0-9_]', '_', topic)[:30]
    prefix = "Modul_Ajar" if is_modul else "Bukti_Fisik"
    filename = f"{prefix}_{clean_topic}.docx"
    return send_file(
        bio,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
    return send_file(
        bio,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    ip = get_local_ip()
    print("=" * 60)
    print("  🚀 AI-autoDoc Local Engine Server Siap Berjalan!")
    print(f"  📍 Akses Lokal PC:      http://localhost:{port}")
    print(f"  📱 Akses Jaringan/HP:   http://{ip}:{port}")
    print("=" * 60)
    app.run(host='0.0.0.0', port=port, debug=False)
