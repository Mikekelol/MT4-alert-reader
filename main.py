import numpy as np
import soundfile as sf
from scipy.signal import correlate
import sounddevice as sd
import sys
import queue
import time
import win32gui
import win32con
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
from PIL import ImageGrab, Image, ImageEnhance
import psutil
import traceback
import pywintypes
import os
import datetime  # Add this import at the beginning of the file
# ### MODIFICATION ###: Required for window handling
try:
    import win32com.client
except ImportError:
    print("ERROR: Module 'pywin32' not found or incomplete.")
    print("Please install or reinstall it with: pip install pywin32")
    sys.exit(1)


# --- Configuration ---
SYSTEM_SAMPLE_RATE = 48000
REFERENCE_FILE = r"PATH\TO\YOUR\REFERENCE_AUDIO.wav"  # Modify with the correct path
DETECTION_THRESHOLD = 0.7
BLOCK_SIZE = 4096  # Increased for better stability
WINDOW_DURATION = 3  # Reduced for smaller buffer
LATENCY = 'high'    # Changed for better stability

# Window title to read from
ALERT_WINDOW_TITLE = "Allarme"

# OCR coordinates relative to "Alert" window
# Exact coordinates for text reading
ALERT_WINDOW_COORDS = (10, 45, 380, 65)  # Moved slightly higher

# Desired position for alert window
ALERT_WINDOW_X = 100
ALERT_WINDOW_Y = 100

# --- Setup and Global Variables ---
try:
    reference, _ = sf.read(REFERENCE_FILE, dtype='float32')
    if len(reference.shape) > 1:
        reference = reference.mean(axis=1) # Convert to mono if stereo
except Exception as e:
    print(f"ERROR: Unable to load audio reference file '{REFERENCE_FILE}'. Details: {e}")
    sys.exit(1)

window_size = SYSTEM_SAMPLE_RATE * WINDOW_DURATION
audio_buffer = np.zeros(window_size, dtype='float32')
action_queue = queue.Queue()
last_detection_time = 0
COOLDOWN_PERIOD = 30  # Aumentato a 30 secondi
SIGNAL_THRESHOLD = 3.0  # Tempo minimo tra segnali audio in secondi
last_signal_time = 0  # Per tracciare l'ultimo segnale audio
last_processed_signal = 0
last_trade = None  # Per tracciare l'ultimo trade eseguito

# --- Funzioni di Gestione Finestre e OCR ---

def find_alert_window():
    """Find the handle of the 'Alert' popup window."""
    try:
        # First try exact match
        hwnd = win32gui.FindWindow(None, ALERT_WINDOW_TITLE)
        if hwnd:
            print(f"Window '{ALERT_WINDOW_TITLE}' found.")
            return hwnd
            
        # If not found, search all windows
        def callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if 'allarme' in title.lower():
                    results.append(hwnd)
            return True
        
        results = []
        win32gui.EnumWindows(callback, results)
        
        if results:
            print(f"Alert window found with handle: {results[0]}")
            return results[0]
            
        print("No alert window found")
        return None
        
    except Exception as e:
        print(f"Error finding window: {e}")
        traceback.print_exc()
        return None

def bring_window_to_foreground(hwnd):
    """Bring specified window to foreground."""
    try:
        # More aggressive method to bring window to front
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')  # Send ALT to "unlock" system focus
        win32gui.SetForegroundWindow(hwnd)
        
        # A volte è necessario ripristinarla se è iconificata
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            
        time.sleep(0.5) # Attendi che la finestra sia effettivamente attiva
        return True
    except (pywintypes.error, Exception) as e:
        print(f"ERROR bringing window to foreground (handle: {hwnd}): {e}")
        print("-> Tip: Run script as Administrator.")
        return False

def position_alert_window(hwnd):
    """Force alert window to specific position."""
    try:
        rect = win32gui.GetWindowRect(hwnd)
        width = rect[2] - rect[0]
        height = rect[3] - rect[1]
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, ALERT_WINDOW_X, ALERT_WINDOW_Y, width, height, win32con.SWP_SHOWWINDOW)
        print(f"Finestra alert posizionata a ({ALERT_WINDOW_X}, {ALERT_WINDOW_Y}) e portata in primo piano.")
    except Exception as e:
        print(f"Errore nel posizionamento della finestra: {e}")

def read_text_from_alert_window(hwnd):
    """Read text from alert window using specified coordinates."""
    try:
        # Capture entire window
        window_rect = win32gui.GetWindowRect(hwnd)
        full_screenshot = ImageGrab.grab(bbox=window_rect)
        full_screenshot.save("debug_alert_window.png")
        
        # Usa le coordinate esatte per estrarre il testo
        roi = full_screenshot.crop(ALERT_WINDOW_COORDS)
        roi.save("debug_roi.png")
        print(f"ROI extracted from coordinates: {ALERT_WINDOW_COORDS}")
        
        # Pre-processing
        img = roi.convert('L')
        img = ImageEnhance.Contrast(img).enhance(2.5)
        img = ImageEnhance.Sharpness(img).enhance(2.0)
        img.save("debug_processed.png")
        
        # OCR
        text = pytesseract.image_to_string(img, config='--oem 3 --psm 7').strip()
        print(f"Text detected: [{text}]")
        
        # Save text to file with timestamp
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        output_file = os.path.join(os.path.dirname(__file__), "ocr_log.txt")
        with open(output_file, 'a', encoding='utf-8') as f:
            f.write(f"[{timestamp}] {text}\n")
        print(f"Text added to file: {output_file}")
        
        return text

    except Exception as e:
        print(f"Error during OCR on alert window: {e}")
        traceback.print_exc()
        return ""

# --- Funzioni Audio e Loop Principale ---
def audio_callback(indata, frames, time_info, status):
    if status:
        print(f"Audio callback status: {status}", file=sys.stderr)
        return
    
    global audio_buffer, last_processed_signal, last_signal_time
    current_time = datetime.datetime.now().timestamp()
    
    # Verifica del tempo minimo tra segnali audio
    if current_time - last_signal_time < SIGNAL_THRESHOLD:
        return
    
    try:    
        # Gestione più robusta del buffer audio
        audio_data = np.zeros_like(indata.flatten())
        audio_data[:] = indata.flatten()
        
        # Aggiornamento del buffer con controllo overflow
        buffer_len = len(audio_buffer)
        data_len = len(audio_data)
        if data_len > 0:
            audio_buffer = np.roll(audio_buffer, -min(data_len, buffer_len))
            audio_buffer[-data_len:] = audio_data[-min(data_len, buffer_len):]
        
        # Normalizzazione e correlazione con controlli
        norm_ref = np.linalg.norm(reference)
        norm_buf = np.linalg.norm(audio_buffer)
        
        if norm_ref > 1e-6 and norm_buf > 1e-6:
            correlation = correlate(audio_buffer, reference, mode='valid')
            max_corr = np.max(np.abs(correlation)) / (norm_ref * norm_buf)
            
            if max_corr > DETECTION_THRESHOLD and action_queue.empty():
                last_signal_time = current_time
                if current_time - last_processed_signal >= COOLDOWN_PERIOD:
                    print(f"\n--- AUDIO SIGNAL DETECTED (Correlation: {max_corr:.2f}) ---")
                    action_queue.put_nowait("READ_ALERT")
                    audio_buffer.fill(0)  # Reset buffer after detection
                else:
                    print(f"\nSignal ignored: cooldown active ({int(COOLDOWN_PERIOD - (current_time - last_processed_signal))}s)")
    
    except Exception as e:
        print(f"Error in audio callback: {e}")
        traceback.print_exc()

def process_alert():
    """Handle the entire alert process."""
    global last_processed_signal
    current_time = datetime.datetime.now().timestamp()
    

        
    print("Starting alert reading process...")
    time.sleep(1.5)  # Aumentato il tempo di attesa per l'apparizione della finestra

    hwnd_alert = find_alert_window()
    if not hwnd_alert:
        print("Alert window not found after signal. Returning to listening.")
        return

    print(f"Window '{ALERT_WINDOW_TITLE}' found. Trying to bring to front.")
    if not bring_window_to_foreground(hwnd_alert):
        print("Could not bring window to front. Process might fail.")

    position_alert_window(hwnd_alert)
    time.sleep(0.3)  # Attesa dopo il posizionamento
    
    text = read_text_from_alert_window(hwnd_alert)
    if not text:
        print("No text read from window. Returning to listening.")
        
    # Chiudi la finestra di allarme per prepararsi al prossimo segnale
    print("Closing alert window...")
    win32gui.PostMessage(hwnd_alert, win32con.WM_CLOSE, 0, 0)
    last_processed_signal = current_time

def main():
    """Main script loop."""
    if not any(p.name() == 'terminal.exe' for p in psutil.process_iter(['name'])):
        print("WARNING: MT4 (terminal.exe) does not appear to be running.")
    
    vb_cable_index = None
    try:
        devices = sd.query_devices()
        for i, device in enumerate(devices):
            if "CABLE" in device['name'] and device['max_input_channels'] > 0:
                vb_cable_index = i
                print(f"Dispositivo audio trovato: '{device['name']}' (indice {i})")
                break
    except Exception as e:
         print(f"Errore nella ricerca dei dispositivi audio: {e}")
         return

    if vb_cable_index is None:
        print("ERRORE: Dispositivo di output virtuale (es. 'VB-CABLE') non trovato.")
        return

    print("\n--- Script avviato. In ascolto per il segnale audio... (Premi Ctrl+C per uscire) ---")
    
    tick_count = 0
    stream = None
    is_listening = True
    
    try:
        while True:
            try:
                # Gestione dello stream audio
                if stream is None:
                    # Reset del buffer audio prima di creare un nuovo stream
                    audio_buffer.fill(0)
                    stream = sd.InputStream(
                        device=vb_cable_index,
                        channels=1,
                        samplerate=SYSTEM_SAMPLE_RATE,
                        blocksize=BLOCK_SIZE,
                        callback=audio_callback,
                        latency=LATENCY,
                        dtype=np.float32  # Specifica esplicita del tipo di dati
                    )
                    stream.start()
                    print("\nAscolto audio attivato")
                
                try:
                    action = action_queue.get_nowait()
                    if action == "READ_ALERT":
                        # Stoppa lo stream durante il processing
                        if stream is not None:
                            stream.stop()
                            stream.close()
                            stream = None
                        
                        process_alert()
                        time.sleep(2)  # Breve pausa dopo il processing
                        
                except queue.Empty:
                    time.sleep(0.1)
                    
                # Stampa un punto ogni 10 secondi
                tick_count += 1
                if tick_count % 10 == 0:
                    print(".", end="", flush=True)
                    
            except KeyboardInterrupt:
                print("\nProgramma terminato dall'utente.")
                break
            except Exception as e:
                print(f"\nERRORE nel loop principale: {e}")
                if stream is not None:
                    stream.stop()
                    stream.close()
                    stream = None
                time.sleep(1)

    finally:
        if stream is not None:
            stream.stop()
            stream.close()

if __name__ == "__main__":
    print("-----------------------------------------------------")
    print("  MT4 Alert Automation Script                       ")
    print("-----------------------------------------------------")
    print("Tip: Run this script as Administrator for proper window handling.")
    main()