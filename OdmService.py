# --- Early-stage error logging ---
# This helps debug startup issues before the main logger is configured.
try:
    import os, sys, ctypes, traceback, socket, threading
    import win32serviceutil
    import win32service
    import win32event
    import servicemanager
    import logging
    import logging.handlers
    import serial
    import time
    import serial.tools.list_ports
    import re 
    from collections import deque
    from flask import Flask, request, jsonify
    from flask_cors import CORS

    # --- DataStore (local database) ---
    import datastore

except Exception as e:
    log_dir_fallback = os.path.join(os.getenv('ProgramData', 'C:'), 'OdmService', 'logs')
    if not os.path.exists(log_dir_fallback):
        os.makedirs(log_dir_fallback, exist_ok=True)
    with open(os.path.join(log_dir_fallback, "startup_error.log"), "a") as f:
        f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - CRITICAL STARTUP ERROR\n")
        f.write(f"Error: {str(e)}\n")
        f.write(traceback.format_exc())
    sys.exit(1)


# --- Flask App ---
app = Flask(__name__)
CORS(app, origins=[re.compile(r".*odmtec.*"), re.compile(r".*otchoumouang\.github\.io.*")])

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
    base_dir = sys._MEIPASS
    os.environ['PATH'] = base_dir + os.pathsep + os.environ['PATH']
    try:
        import pywintypes
        import pythoncom
    except ImportError:
        dll_dir = os.path.join(base_dir)
        os.add_dll_directory(dll_dir)

# Chemin absolu pour les logs
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
COMPANY = "SITC, SAN-PEDRO"
DESKTOP = socket.gethostname()
FRAME_LENGTH = 11
SERVICE_NAME = "OdmService"
SERVICE_DISPLAY_NAME = "ODM - Balance Data Collector Service"

# Paramètres de stabilisation et d'envoi
STABILIZATION_COUNT = 3
MIN_SEND_INTERVAL = 2  # Délai minimum entre 2 envois (secondes) - MODIFIÉ
CLEANUP_INTERVAL = 600 # Intervalle de nettoyage en secondes (10 minutes)

def configure_logging():
    """Configure la journalisation vers fichier et Event Viewer"""
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR)
    
    logger = logging.getLogger(SERVICE_NAME)
    logger.setLevel(logging.INFO)
    
    # Garder une référence au handler pour la rotation manuelle
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

# --- API Endpoints ---
service_instance = None

@app.route('/api/lire_poids_reel', methods=['GET'])
def lire_poids_reel():
    """Nouvel endpoint pour lire le poids à la demande."""
    if service_instance and service_instance.ser and service_instance.ser.is_open:
        try:
            poids = service_instance.read_weight_on_demand()
            if poids is not None:
                return jsonify({"poids": poids, "unite": "kg"}), 200
            else:
                return jsonify({"error": "Impossible de lire un poids stable depuis la balance."}), 503
        except Exception as e:
            logger.error(f"Erreur lors de la lecture à la demande : {e}")
            return jsonify({"error": f"Erreur interne du service: {e}"}), 500
    else:
        return jsonify({"error": "Le service n'est pas connecté à la balance."}), 503

@app.route('/api/poids', methods=['POST'])
def post_poids():
    data = request.get_json()
    if not data or 'poids' not in data or data['poids'] < 0:
        return jsonify({"error": "Le modèle de données est invalide ou le poids est négatif."}), 400

    poids_valeur = data['poids']
    desktop = data.get('desktop', DESKTOP)
    company = data.get('company', COMPANY)

    try:
        datastore.add_poids(poids_valeur, desktop, company)
        return jsonify({"message": "Valeur ajoutée avec succès", "poids": poids_valeur}), 200
    except Exception as e:
        logger.error(f"API Error on POST: {e}")
        return jsonify({"error": "Une erreur interne est survenue."}), 500

@app.route('/api/poids', methods=['GET'])
def get_poids():
    desktop = request.args.get('desktop')
    company = request.args.get('company')

    try:
        dernier_poids = datastore.get_dernier_poids(desktop, company)
        if dernier_poids:
            return jsonify(dernier_poids)
        else:
            return jsonify({"message": "Aucun enregistrement trouvé pour les critères fournis."}), 404
    except Exception as e:
        logger.error(f"API Error on GET: {e}")
        return jsonify({"error": "Une erreur interne est survenue."}), 500

def run_flask_app(svc_instance):
    """Runs the Flask app in a separate thread."""
    global service_instance
    service_instance = svc_instance
    try:
        app.run(host='127.0.0.1', port=5000)
    except Exception as e:
        logger.error(f"Failed to start Flask server: {e}")

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

def save_weight_locally(weight_kg):
    """Saves the weight to the local database."""
    try:
        datastore.add_poids(weight_kg, DESKTOP, COMPANY)
        logger.info(f"Poids {weight_kg}kg enregistré localement.")
        return True
    except Exception as e:
        logger.error(f"Erreur d'enregistrement local: {e}")
        return False

def get_latest_weight_from_local_db():
    """Retrieves the last recorded weight from the local database."""
    try:
        data = datastore.get_dernier_poids(DESKTOP, COMPANY)
        if data and "valeur" in data:
            latest_weight = float(data["valeur"])
            logger.info(f"Dernier poids récupéré de la DB locale: {latest_weight}kg")
            return latest_weight
        else:
            logger.info("Aucun poids trouvé en local. On considère 0kg.")
            return 0.0
    except Exception as e:
        logger.error(f"Erreur de lecture de la DB locale: {e}")
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
        self.flask_thread = None
        self.cleanup_thread = None
        self.lock = threading.Lock()

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.is_alive = False
        if self.ser and self.ser.is_open:
            self.ser.close()
            logger.info("Port série fermé")
        logger.info("Service stop requested.")

    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, '')
        )
        logger.info(f"Démarrage du service {SERVICE_DISPLAY_NAME}")

        try:
            datastore.init_db()
            logger.info("Database initialized.")
        except Exception as e:
            logger.error(f"CRITICAL: Failed to initialize database: {e}")
            self.SvcStop()
            return

        self.flask_thread = threading.Thread(target=run_flask_app, args=(self,), daemon=True)
        self.flask_thread.start()
        logger.info("Flask server thread started.")
        
        # Démarrage du thread de nettoyage
        self.cleanup_thread = threading.Thread(target=self.run_cleanup_task, daemon=True)
        self.cleanup_thread.start()
        logger.info(f"Cleanup thread started. Will run every {CLEANUP_INTERVAL} seconds.")

        self.main()

    def run_cleanup_task(self):
        """Tâche de fond pour nettoyer la DB et les logs périodiquement."""
        while self.is_alive:
            stop_event_triggered = win32event.WaitForSingleObject(self.hWaitStop, CLEANUP_INTERVAL * 1000)
            if stop_event_triggered == win32event.WAIT_OBJECT_0:
                break

            if self.is_alive:
                try:
                    logger.info("--- Début du nettoyage périodique ---")
                    
                    deleted_count = datastore.cleanup_poids(keep=5)
                    logger.info(f"Nettoyage DB: {deleted_count} anciens enregistrements supprimés.")
                    
                    for handler in logger.handlers:
                        if isinstance(handler, logging.handlers.RotatingFileHandler):
                            handler.doRollover()
                            logger.info("Nettoyage Logs: Rotation effectuée.")
                            break
                    
                    logger.info("--- Fin du nettoyage périodique ---")
                except Exception as e:
                    logger.error(f"Erreur durant le nettoyage périodique: {e}")

    def read_weight_on_demand(self):
        """Lit une valeur de poids stable depuis la balance."""
        with self.lock:
            if not self.ser or not self.ser.is_open:
                logger.error("Tentative de lecture alors que le port série n'est pas ouvert.")
                return None

            self.ser.reset_input_buffer()
            buffer = bytearray()
            recent_readings = deque(maxlen=STABILIZATION_COUNT)
            start_time = time.time()

            logger.info("Début de la lecture de poids à la demande...")

            while time.time() - start_time < 5: # Timeout de 5 secondes
                try:
                    chunk = self.ser.read(self.ser.in_waiting or 1)
                    if chunk:
                        buffer.extend(chunk)

                    # On cherche une trame complète
                    while len(buffer) >= FRAME_LENGTH:
                        frame_start_index = buffer.find(b'w')
                        if frame_start_index == -1:
                            buffer.clear() # Pas de début de trame, on vide
                            break

                        # Si on a trouvé un début, on s'assure d'avoir une trame complète
                        if len(buffer) - frame_start_index >= FRAME_LENGTH:
                            frame = bytes(buffer[frame_start_index : frame_start_index + FRAME_LENGTH])
                            del buffer[:frame_start_index + FRAME_LENGTH]

                            if frame.endswith(b'kg') and (frame[1] in [ord('w'), ord('n')]):
                                weight_kg = parse_weight_data(frame)
                                if weight_kg is not None:
                                    recent_readings.append(weight_kg)

                                    if len(recent_readings) == STABILIZATION_COUNT:
                                        # Vérifier si les valeurs sont stables
                                        if (max(recent_readings) - min(recent_readings)) <= 1:
                                            stable_weight = recent_readings[-1]
                                            logger.info(f"Poids stable détecté: {stable_weight}kg")
                                            # Enregistrer immédiatement en local
                                            save_weight_locally(stable_weight)
                                            return stable_weight
                        else:
                            # Trame incomplète, on attend plus de données
                            break
                except Exception as e:
                    logger.error(f"Erreur pendant la lecture à la demande: {e}")
                    return None

                time.sleep(0.05) # Petite pause pour ne pas surcharger le CPU

            logger.warning("Timeout: Impossible de lire un poids stable en 5 secondes.")
            return None

    def main(self):
        """Boucle principale du service: maintient la connexion à la balance."""
        while self.is_alive:
            try:
                if self.ser and self.ser.is_open:
                    # Le port est déjà ouvert, on attend juste
                    if win32event.WaitForSingleObject(self.hWaitStop, 1000) == win32event.WAIT_OBJECT_0:
                        self.is_alive = False
                    continue

                # Si le port n'est pas ouvert, on tente de le trouver
                self.ser = find_scale_port()
                if self.ser:
                    logger.info(f"Balance connectée sur {self.ser.port}. En attente de requêtes.")
                else:
                    logger.warning("Balance non détectée. Nouvelle tentative dans 10s.")
                    if win32event.WaitForSingleObject(self.hWaitStop, 10000) == win32event.WAIT_OBJECT_0:
                        self.is_alive = False

            except serial.SerialException as se:
                logger.error(f"Erreur port série: {se}. Reconnexion...")
                if self.ser and self.ser.is_open:
                    self.ser.close()
                time.sleep(5)
            except Exception as e:
                logger.exception(f"Erreur majeure dans la boucle principale: {e}")
                time.sleep(10)

        logger.info("Arrêt du service.")

if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(OdmService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(OdmService)

