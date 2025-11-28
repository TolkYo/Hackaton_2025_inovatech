import cv2
import numpy as np
import threading
import time
import math
import queue
import sys
import os
from typing import Tuple
from flask import Flask, render_template, Response, jsonify

current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
template_dir = os.path.join(parent_dir, 'templates')

sys.path.append(parent_dir)

app = Flask(__name__, template_folder=template_dir)

try:
    import win32com.client as winclient
    WINDOWS_TTS_AVAILABLE = True
except ImportError:
    WINDOWS_TTS_AVAILABLE = False
    print("[WARN] pywin32 não instalado. Instale com: pip install pywin32")

CAMERA_INDEX = 0
FRAME_WIDTH = 960
FRAME_HEIGHT = 720
REAL_WIDTH_CM = 30.0 
SAFE_ZONE_PX = 40
TTS_INTERVAL_SECONDS = 2.0

COLOR_MIN = np.array([95, 100, 20])
COLOR_MAX = np.array([145, 255, 255]) 

MIN_AREA_THRESHOLD = 500
CIRCLE_CIRCULARITY_MIN = 0.5
CIRCLE_RADIUS_MIN = 8

ultimo_comando_tempo = 0
msg_atual_interface = "Sistema Iniciado."

class SinteseVozWorker:
    def __init__(self):
        self.fila_fala = queue.Queue()
        self.tts = None
        self._inicializar_tts()
        self.thread = None
        self._iniciar_thread()
    
    def _inicializar_tts(self) -> None:
        if not WINDOWS_TTS_AVAILABLE:
            print("[TTS] pywin32 não disponível, fala desabilitada")
            return
        
        try:
            self.tts = winclient.Dispatch("SAPI.SpVoice")
            self.tts.Rate = 4
            self.tts.Volume = 100
            
            voices = self.tts.GetVoices()
            for voice in voices:
                voice_name = voice.GetAttribute("Name")
                if 'portuguese' in voice_name.lower() or 'pt' in voice_name.lower():
                    self.tts.Voice = voice
                    print(f"[TTS] Voz em português selecionada: {voice_name}")
                    break
            
            print("[TTS] Motor SAPI (Windows TTS) inicializado com sucesso")
        except Exception as e:
            print(f"[TTS] Erro ao inicializar SAPI: {e}")
            self.tts = None
    
    def _iniciar_thread(self) -> None:
        self.thread = threading.Thread(target=self._worker_loop, daemon=True)
        self.thread.start()
        print("[TTS] Thread de fala iniciada")
    
    def _worker_loop(self) -> None:
        print("[TTS] Worker loop ativo")
        while True:
            try:
                texto = self.fila_fala.get(timeout=1)
                
                if self.tts and WINDOWS_TTS_AVAILABLE:
                    try:
                        print(f"[TTS] Falando: {texto}")
                        self.tts.Speak(texto, 0)
                        print(f"[TTS] Fala concluída")
                    except Exception as e:
                        print(f"[TTS] Erro ao reproduzir fala: {e}")
                else:
                    print(f"[TTS] SAPI não disponível para: {texto}")
                
                self.fila_fala.task_done()
            
            except queue.Empty:
                continue
            except Exception as e:
                print(f"[TTS] Erro na worker loop: {e}")
    
    def falar(self, texto: str) -> None:
        self.fila_fala.put(texto)
    
    def aguardar_conclusao(self) -> None:
        self.fila_fala.join()

tts_worker = SinteseVozWorker()

def falar_comando(texto: str) -> None:
    tts_worker.falar(texto)

def desenhar_seta(img: cv2.Mat, direcao: str, centro: Tuple[int, int], cor: Tuple[int, int, int]) -> None:
    cx, cy = centro
    tamanho = 100
    espessura = 5
    
    p_ini = (cx, cy)
    p_fim = (cx, cy)

    if '-' in direcao:
        parts = direcao.split('-')
        if len(parts) == 2:
            vertical, horizontal = parts[0], parts[1]
            dx_sign = 1 if horizontal == "Direita" else -1
            dy_sign = 1 if vertical == "Baixo" else -1
            p_ini = (cx + dx_sign * 50, cy + dy_sign * 50)
            p_fim = (cx + dx_sign * (50 + tamanho), cy + dy_sign * (50 + tamanho))
    else:
        if direcao == "Direita":
            p_ini = (cx + 50, cy)
            p_fim = (cx + 50 + tamanho, cy)
        elif direcao == "Esquerda":
            p_ini = (cx - 50, cy)
            p_fim = (cx - 50 - tamanho, cy)
        elif direcao == "Baixo":
            p_ini = (cx, cy + 50)
            p_fim = (cx, cy + 50 + tamanho)
        elif direcao == "Cima":
            p_ini = (cx, cy - 50)
            p_fim = (cx, cy - 50 - tamanho)
    
    if p_ini != p_fim:
        cv2.arrowedLine(img, p_ini, p_fim, cor, espessura, tipLength=0.3)

def desenhar_referencial(frame: cv2.Mat, centro_alvo: Tuple[int, int]) -> None:
    cv2.drawMarker(frame, centro_alvo, (200, 200, 200), cv2.MARKER_CROSS, 30, 2)
    cv2.circle(frame, centro_alvo, SAFE_ZONE_PX, (100, 100, 100), 1)

def desenhar_deteccao(frame: cv2.Mat, cx: int, cy: int, raio: int, distancia_pixels: float,
                      pixels_por_cm: float, centro_alvo: Tuple[int, int], direcao: str) -> None:
    distancia_cm = distancia_pixels / pixels_por_cm
    
    if distancia_pixels < SAFE_ZONE_PX:
        cor_status = (0, 255, 0)
        cv2.putText(frame, "ALVO TRAVADO", (cx - 80, cy - 60), cv2.FONT_HERSHEY_SIMPLEX, 0.8, cor_status, 2)
    elif distancia_pixels < 150:
        cor_status = (0, 255, 255)
    else:
        cor_status = (0, 0, 255)

    cv2.circle(frame, (cx, cy), raio, cor_status, 2)
    cv2.line(frame, (cx, cy), centro_alvo, cor_status, 2)
    cv2.putText(frame, f"{distancia_cm:.1f}cm", (cx - 30, cy - raio - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.7, cor_status, 2)

    if distancia_pixels >= SAFE_ZONE_PX and direcao:
        desenhar_seta(frame, direcao, centro_alvo, cor_status)

def detectar_circulo(frame: cv2.Mat) -> Tuple[bool, int, int, int, float]:
    hsv = cv2.cvtColor(frame, cv2.COLOR_BGR2HSV)
    mask = cv2.inRange(hsv, COLOR_MIN, COLOR_MAX)
    
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (5, 5))
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel, iterations=1)
    
    contornos, _ = cv2.findContours(mask.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    if len(contornos) == 0:
        return False, 0, 0, 0, 0
    
    c = max(contornos, key=cv2.contourArea)
    area = cv2.contourArea(c)
    
    if area < MIN_AREA_THRESHOLD:
        return False, 0, 0, 0, area
    
    (cx, cy), raio = cv2.minEnclosingCircle(c)
    cx, cy, raio = int(cx), int(cy), int(raio)
    
    perimetro = cv2.arcLength(c, True)
    circularidade = (4 * math.pi * area) / (perimetro * perimetro) if perimetro > 0 else 0
    
    if raio >= CIRCLE_RADIUS_MIN and circularidade >= CIRCLE_CIRCULARITY_MIN:
        return True, cx, cy, raio, area
    
    if area > MIN_AREA_THRESHOLD * 2:
        return True, cx, cy, raio, area
    
    return False, 0, 0, 0, area

def calcular_distancia_e_direcao(cx: int, cy: int, centro_alvo: Tuple[int, int]) -> Tuple[float, str]:
    dx = centro_alvo[0] - cx
    dy = centro_alvo[1] - cy
    distancia_pixels = math.sqrt(dx**2 + dy**2)
    
    direcao = ""
    if distancia_pixels >= SAFE_ZONE_PX:
        if abs(dx) >= SAFE_ZONE_PX and abs(dy) >= SAFE_ZONE_PX:
            horizontal = "Direita" if dx > 0 else "Esquerda"
            vertical = "Baixo" if dy > 0 else "Cima"
            direcao = f"{vertical}-{horizontal}"
        else:
            if abs(dx) > abs(dy):
                direcao = "Direita" if dx > 0 else "Esquerda"
            else:
                direcao = "Baixo" if dy > 0 else "Cima"
    return distancia_pixels, direcao

def gerar_frames():
    global ultimo_comando_tempo, msg_atual_interface
    cap = cv2.VideoCapture(CAMERA_INDEX)
    centro_alvo = (FRAME_WIDTH // 2, FRAME_HEIGHT // 2)
    pixels_por_cm = FRAME_WIDTH / REAL_WIDTH_CM

    while True:
        ret, frame = cap.read()
        if not ret: break

        frame = cv2.resize(frame, (FRAME_WIDTH, FRAME_HEIGHT))
        frame = cv2.flip(frame, 1)
        
        desenhar_referencial(frame, centro_alvo)
        encontrado, cx, cy, raio, area = detectar_circulo(frame)
        msg_comando = ""
        tempo_atual = time.time()

        if encontrado:
            distancia_pixels, direcao = calcular_distancia_e_direcao(cx, cy, centro_alvo)
            distancia_cm = distancia_pixels / pixels_por_cm
            desenhar_deteccao(frame, cx, cy, raio, distancia_pixels, pixels_por_cm, centro_alvo, direcao)

            if distancia_pixels < SAFE_ZONE_PX:
                msg_comando = "Posição correta."
            elif direcao:
                distancia_cm_int = int(round(distancia_cm))
                msg_comando = f"Mova {distancia_cm_int} centímetros para {direcao}"

            if msg_comando and (tempo_atual - ultimo_comando_tempo) > TTS_INTERVAL_SECONDS:
                falar_comando(msg_comando)
                ultimo_comando_tempo = tempo_atual
        else:
            if (tempo_atual - ultimo_comando_tempo) > 5:
                msg_atual_interface = "Procurando peça..."
                ultimo_comando_tempo = tempo_atual

        ret, buffer = cv2.imencode('.jpg', frame)
        if not ret: continue
        frame_bytes = buffer.tobytes()
        yield (b'--frame\r\n'
               b'Content-Type: image/jpeg\r\n\r\n' + frame_bytes + b'\r\n')
    
    cap.release()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/video_feed')
def video_feed():
    return Response(gerar_frames(), mimetype='multipart/x-mixed-replace; boundary=frame')

@app.route('/status')
def status():
    return jsonify({'mensagem': msg_atual_interface})

def main():
    print("--- SISTEMA DE DETECÇÃO E GUIA DE OBJETOS ---")
    print(f"Diretório de execução: {current_dir}")
    print(f"Diretório de Templates (Front): {template_dir}")
    print(f"Calibração: {FRAME_WIDTH / REAL_WIDTH_CM:.2f} pixels por cm")
    
    time.sleep(0.5)
    falar_comando("Sistema pronto.")
    time.sleep(1)
    
    print("[MAIN] Iniciando servidor Flask...")
    app.run(host='0.0.0.0', port=5000, debug=False, threaded=True)

if __name__ == '__main__':
    main()