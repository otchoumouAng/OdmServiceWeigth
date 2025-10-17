import os, sys, ctypes, traceback, socket
import win32serviceutil
import win32service
import win32event
import servicemanager
import logging
import logging.handlers
import serial
import requests
import time
import serial.tools.list_ports
from datetime import datetime, date
from collections import deque  

#####################################

# Solution robuste pour les DLLs dans les builds PyInstaller
def load_critical_dlls():
    """Charge manuellement les DLLs essentielles pour pywin32"""
    dlls_to_load = [
        "pythoncom{}.dll".format(sys.winver.replace('.', '')),
        "pywintypes{}.dll".format(sys.winver.replace('.', ''))
    ]
    
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    loaded = []
    
    for dll_name in dlls_to_load:
        try:
            dll_path = os.path.join(base_path, dll_name)
            if os.path.exists(dll_path):
                ctypes.WinDLL(dll_path)
                loaded.append(dll_name)
        except Exception as e:
            error_msg = f"Error loading {dll_name}: {str(e)}"
            if 'log_error' in globals():
                log_error(error_msg)
            else:
                print(error_msg)
    
    return loaded

# Essayer d'abord d'importer normalement
try:
    import win32api
    import pywintypes
except ImportError:
    # Chargement manuel si échec
    loaded_dlls = load_critical_dlls()
    
    # Réessayer après chargement manuel
    try:
        import win32api
        import pywintypes
        print(f"✅ DLLs chargées avec succès: {', '.join(loaded_dlls)}")
    except ImportError as e:
        error_msg = f"CRITICAL DLL LOAD ERROR: {str(e)}"
        if 'log_error' in globals():
            log_error(error_msg)
        else:
            print(error_msg)
        traceback.print_exc()
        sys.exit(1)

# Solution pour les DLLs pywin32 dans les builds PyInstaller
if getattr(sys, 'frozen', False):
    # Mode exécutable - charger les DLLs manuellement
    base_dir = sys._MEIPASS
    os.environ['PATH'] = base_dir + os.pathsep + os.environ['PATH']
    
    # Charger explicitement les DLLs critiques
    try:
        import pywintypes
        import pythoncom
    except ImportError:
        # Ajout manuel du chemin des DLLs
        dll_dir = os.path.join(base_dir)
        os.add_dll_directory(dll_dir)

# Chemin absolu pour les logs (adapté pour les services Windows)
LOG_DIR = os.path.join(os.getenv('ProgramData'), 'OdmService', 'logs')
if not os.path.exists(LOG_DIR):
    try:
        os.makedirs(LOG_DIR)
    except Exception as e:
        LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        if not os.path.exists(LOG_DIR):
            os.makedirs(LOG_DIR)

LOG_FILE = os.path.join(LOG_DIR, "OdmService.log")

# Configuration de l'API
API_URL = "https://capturepoidsapi.odmtec.com/api/poids"
COMPANY = "SITC, SAN-PEDRO"
DESKTOP = socket.gethostname()
#
FRAME_LENGTH = 11
SERVICE_NAME = "OdmService"
SERVICE_DISPLAY_NAME = "ODM - Balance Data Collector Service"

# Paramètres de stabilisation
STABILIZATION_COUNT = 3  # Nombre de lectures stables requises
# MODIFICATION: Augmentation de l'intervalle entre les envois
MIN_SEND_INTERVAL = 20   # Délai minimum entre 2 envois (secondes)

def configure_logging():
    """Configure la journalisation vers fichier et Event Viewer"""
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR)
    
    logger = logging.getLogger(SERVICE_NAME)
    logger.setLevel(logging.INFO)
    
    file_handler = logging.handlers.RotatingFileHandler(
        LOG_FILE, maxBytes=2*1024*1024, backupCount=5
    )
    file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(file_formatter)
    
    event_handler = logging.handlers.NTEventLogHandler(SERVICE_NAME)
    
    logger.addHandler(file_handler)
    logger.addHandler(event_handler)
    
    return logger

logger = configure_logging()

def find_scale_port():
    """Trouve automatiquement le port de la balance"""
    ports = serial.tools.list_ports.comports()
    logger.info(f"Ports disponibles: {[p.device for p in ports]}")
    
    for port in ports:
        try:
            logger.info(f"Test du port {port.device}")
            ser = serial.Serial(
                port=port.device,
                baudrate=9600,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=2
            )
            time.sleep(1)
            data = ser.read(ser.in_waiting or FRAME_LENGTH)
            if data and b'w' in data:
                logger.info(f"Balance détectée sur {port.device}")
                ser.reset_input_buffer()
                return ser
            ser.close()
        except Exception as e:
            logger.error(f"Erreur sur {port.device}: {type(e).__name__} - {e}")
    return None

def parse_weight_data(frame):
    """Parse une trame de données de poids"""
    try:
        frame_str = frame.decode('ascii')
        if not ((frame_str.startswith('ww') or frame_str.startswith('wn')) and frame_str.endswith('kg')):
            return None
        
        num_part = frame_str[2:9].replace(' ', '')
        
        if '-' in num_part:
            return -int(num_part.replace('-', '').strip())
        else:
            return int(num_part)
    except (UnicodeDecodeError, ValueError) as e:
        logger.error(f"Erreur de parsing: {e} pour la trame: {frame}")
        return None

def send_to_api(weight_kg):
    """Envoie le poids à l'API via une requête POST."""
    try:
        payload = {"poids": weight_kg, "company": COMPANY, "desktop": DESKTOP}
        response = requests.post(API_URL, json=payload, timeout=15)
        
        if response.status_code == 200:
            logger.info(f"Poids {weight_kg}kg envoyé avec succès (POST)")
            return True
        else:
            logger.error(f"Erreur POST API ({response.status_code}): {response.text} | Payload: {payload}")
            return False
    except requests.exceptions.RequestException as e:
        logger.error(f"ERREUR connexion POST API: {type(e).__name__} - {e}")
        return False

def get_latest_weight_from_api(desktop, company):
    """Récupère le dernier poids enregistré depuis l'API via une requête GET."""
    try:
        params = {"desktop": desktop, "company": company}
        response = requests.get(API_URL, params=params, timeout=15)

        if response.status_code == 200:
            data = response.json()
            if data and "valeur" in data:
                latest_weight = float(data["valeur"])
                logger.info(f"Dernier poids récupéré de l'API (GET): {latest_weight}kg")
                return latest_weight
            else:
                logger.warning("L'API a retourné une réponse valide mais sans données de poids. On considère 0kg.")
                return 0.0
        elif response.status_code == 404:
            logger.info("Aucun poids trouvé pour ce poste sur l'API (404). On considère 0kg.")
            return 0.0
        else:
            logger.error(f"Erreur GET API ({response.status_code}): {response.text}")
            return None
    except requests.exceptions.RequestException as e:
        logger.error(f"ERREUR connexion GET API: {type(e).__name__} - {e}")
        return None
    except (ValueError, KeyError) as e:
        logger.error(f"Erreur parsing réponse GET API: {e} | Réponse: {response.text}")
        return None

class OdmService(win32serviceutil.ServiceFramework):
    _svc_name_ = SERVICE_NAME
    _svc_display_name_ = SERVICE_DISPLAY_NAME
    _svc_description_ = "Capture de poids depuis une balance et les envoie à une API"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.is_alive = True
        self.ser = None

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.is_alive = False
        if self.ser and self.ser.is_open:
            self.ser.close()
            logger.info("Port série fermé")

    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, '')
        )
        logger.info(f"Démarrage du service {SERVICE_DISPLAY_NAME}")
        self.main()

    def main(self):
        while self.is_alive:
            try:
                self.ser = find_scale_port()
                
                if not self.ser:
                    logger.warning("Balance non détectée! Nouvelle tentative dans 10s")
                    time.sleep(10)
                    continue

                logger.info(f"Connexion établie sur {self.ser.port}")
                buffer = bytearray()
                
                # NOUVELLES VARIABLES D'ÉTAT
                recent_readings = deque(maxlen=STABILIZATION_COUNT)
                last_sent_time = 0
                last_sent_weight = None # Mémorise le dernier poids envoyé

                self.ser.timeout = 0.1
                
                while self.is_alive:
                    try:
                        chunk = self.ser.read(self.ser.in_waiting or 1)
                        if chunk:
                            buffer.extend(chunk)
                        
                        processed = True
                        while processed and len(buffer) >= FRAME_LENGTH:
                            processed = False
                            found_frame = False
                            
                            for i in range(len(buffer) - FRAME_LENGTH + 1):
                                if buffer[i] == ord('w'):
                                    frame_candidate = bytes(buffer[i:i+FRAME_LENGTH])
                                    
                                    if (frame_candidate.endswith(b'kg') and 
                                       (frame_candidate[1] in [ord('w'), ord('n')])):
                                        
                                        weight_kg = parse_weight_data(frame_candidate)
                                        if weight_kg is not None:
                                            recent_readings.append(weight_kg)
                                            
                                            # MODIFICATION: Refonte complète de la logique de décision
                                            if len(recent_readings) == STABILIZATION_COUNT:
                                                is_stable = (max(recent_readings) - min(recent_readings)) <= 1
                                                
                                                if is_stable:
                                                    stable_weight = recent_readings[-1]
                                                    
                                                    # 1. Ignorer les poids négatifs
                                                    if stable_weight < 0:
                                                        continue

                                                    # 2. Ignorer si identique au dernier poids envoyé
                                                    if stable_weight == last_sent_weight:
                                                        continue

                                                    # 3. Vérifier le délai minimum
                                                    time_since_last = time.time() - last_sent_time
                                                    if time_since_last < MIN_SEND_INTERVAL:
                                                        logger.debug(f"Valeur stable {stable_weight}kg, mais délai non écoulé ({MIN_SEND_INTERVAL - time_since_last:.1f}s restants).")
                                                        continue

                                                    # 4. Décision d'envoi
                                                    should_send = False
                                                    if stable_weight == 0:
                                                        logger.info("Poids stable à 0 détecté. Vérification de la valeur sur l'API...")
                                                        api_weight = get_latest_weight_from_api(DESKTOP, COMPANY)
                                                        if api_weight is not None and api_weight != 0:
                                                            should_send = True
                                                        else:
                                                            logger.info(f"L'API est déjà à 0 ou inaccessible -> {api_weight}")
                                                            last_sent_weight = 0 # Met à jour l'état local pour éviter des vérifs répétées
                                                    else: # Poids positif
                                                        should_send = True

                                                    if should_send:
                                                        if send_to_api(stable_weight):
                                                            last_sent_weight = stable_weight
                                                            last_sent_time = time.time()
                                                            recent_readings.clear()
                                        
                                        del buffer[:i+FRAME_LENGTH]
                                        processed = True
                                        found_frame = True
                                        break
                            
                            if not found_frame and len(buffer) > 100:
                                buffer.clear()
                                logger.debug("Buffer vidé (aucune trame valide trouvée).")
                        
                        if win32event.WaitForSingleObject(self.hWaitStop, 100) == win32event.WAIT_OBJECT_0:
                            self.is_alive = False
                            break
                            
                    except serial.SerialException as se:
                        logger.error(f"ERREUR PORT SÉRIE: {se}. Déconnexion.")
                        break # Sort de la boucle de lecture pour tenter une reconnexion
                    except Exception as e:
                        logger.exception(f"ERREUR LECTURE: {type(e).__name__} - {e}")
                        time.sleep(5)
                        break
                
                if self.ser and self.ser.is_open:
                    self.ser.close()
                recent_readings.clear()

            except Exception as e:
                logger.exception(f"ERREUR MAJEURE: {type(e).__name__} - {e}")
                time.sleep(10)

        logger.info("Arrêt du service")

if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(OdmService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(OdmService)
