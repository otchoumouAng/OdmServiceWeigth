# Code gestion du service de capture de poids
import os
import sys
import threading
import time
import win32api
import win32con
import win32gui
import win32service
import win32serviceutil
import ctypes
import winreg
import serial
import serial.tools.list_ports
import requests
import traceback
import socket

# Configuration
SERVICE_NAME = "OdmService"  
APP_NAME = "OdmServiceTray"
COMPANY = "SITC, SAN-PEDRO"
DESKTOP = socket.gethostname()
API_URL = "https://capturepoidsapi.odmtec.com/api/poids"

# Constantes pour la capture
FRAME_LENGTH = 11
CAPTURE_TIMEOUT = 15  # secondes

# Chemin des logs
LOG_DIR = os.path.join(os.getenv('ProgramData'), 'OdmService', 'logs')
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, "OdmService.log")

# Création des icônes
def create_icons():
    icons = {}
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Chemin des icônes
    green_icon = os.path.join(base_dir, "vert.ico")
    red_icon = os.path.join(base_dir, "rouge.ico")
    
    # Charger les icônes avec gestion d'erreur
    try:
        icons['green'] = win32gui.LoadImage(
            0, green_icon, win32con.IMAGE_ICON, 
            0, 0, win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
        )
    except Exception as e:
        print(f"Erreur chargement icône verte: {e}")
        # Icône de secours (icône système)
        icons['green'] = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
    
    try:
        icons['red'] = win32gui.LoadImage(
            0, red_icon, win32con.IMAGE_ICON, 
            0, 0, win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
        )
    except Exception as e:
        print(f"Erreur chargement icône rouge: {e}")
        icons['red'] = win32gui.LoadIcon(0, win32con.IDI_ERROR)
    
    return icons

ICONS = create_icons()

def get_service_status():
    try:
        return win32serviceutil.QueryServiceStatus(SERVICE_NAME)[1]
    except Exception as e:
        print(f"Erreur statut service: {e}")
        return win32service.SERVICE_STOPPED

def is_user_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit(0)

def service_action(action):
    try:
        if not is_user_admin():
            win32api.MessageBox(0, "Privilèges administrateur requis", "Erreur", win32con.MB_ICONWARNING)
            run_as_admin()
            return False

        if action == "start":
            win32serviceutil.StartService(SERVICE_NAME)
        elif action == "stop":
            win32serviceutil.StopService(SERVICE_NAME)
        elif action == "restart":
            win32serviceutil.RestartService(SERVICE_NAME)
        return True
    except Exception as e:
        error_msg = f"Erreur action service: {str(e)}"
        print(error_msg)
        win32api.MessageBox(0, error_msg, "Action service", win32con.MB_ICONERROR)
        return False

def show_logs():
    if not os.path.exists(LOG_FILE):
        try:
            open(LOG_FILE, 'a').close()
        except Exception:
            pass
    try:
        os.startfile(LOG_FILE)
    except Exception as e:
        win32api.MessageBox(0, f"Impossible d'ouvrir les logs: {str(e)}", "Erreur", win32con.MB_ICONERROR)

def set_autostart(enabled=True):
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE)
        if enabled:
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, f'"{sys.executable}"')
        else:
            try:
                winreg.DeleteValue(key, APP_NAME)
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"Erreur autostart: {e}")
        return False

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
    except (UnicodeDecodeError, ValueError):
        return None

def capture_single_weight():
    """Capture un seul poids depuis la balance"""
    ser = None
    try:
        # Trouver le port de la balance
        ports = serial.tools.list_ports.comports()
        #print(f"Ports disponibles: {[p.device for p in ports]}")
        
        for port in ports:
            try:
                print(f"Essai port: {port.device}")
                ser = serial.Serial(
                    port=port.device,
                    baudrate=9600,
                    bytesize=serial.EIGHTBITS,
                    parity=serial.PARITY_NONE,
                    stopbits=serial.STOPBITS_ONE,
                    timeout=2
                )
                time.sleep(1)  # Délai de stabilisation
                
                # Lire les données
                start_time = time.time()
                buffer = bytearray()
                #print("Début de la capture...")
                
                while time.time() - start_time < CAPTURE_TIMEOUT:
                    to_read = ser.in_waiting
                    if to_read > 0:
                        chunk = ser.read(to_read)
                        buffer.extend(chunk)
                    
                    # Rechercher une trame valide
                    for i in range(len(buffer)):
                        if buffer[i] == ord('w') and len(buffer) - i >= FRAME_LENGTH:
                            frame_candidate = bytes(buffer[i:i+FRAME_LENGTH])
                            if frame_candidate.endswith(b'kg') and frame_candidate[1] in [ord('w'), ord('n')]:
                                weight = parse_weight_data(frame_candidate)
                                if weight is not None:
                                    #print(f"Poids capturé: {weight}kg")
                                    return weight
                    
                    time.sleep(0.1)
                
                ser.close()
            except Exception as e:
                print(f"Erreur port {port.device}: {str(e)}")
                if ser and ser.is_open:
                    ser.close()
                continue
        
        return None
    except Exception as e:
        print(f"Erreur capture: {str(e)}")
        traceback.print_exc()
        return None
    finally:
        if ser and ser.is_open:
            ser.close()

def send_to_api(weight_kg):
    """Envoie le poids à l'API"""
    try:
        response = requests.post(
            API_URL, 
            json={
                "poids": weight_kg,
                "company": COMPANY,
                "desktop": DESKTOP
            }, 
            timeout=5
        )
        print(f"Réponse API: {response.status_code} - {response.text[:50]}")
        return response.status_code == 200
    except Exception as e:
        print(f"Erreur API: {str(e)}")
        return False

class ScaleTrayApp:
    def __init__(self):
        self.hwnd = None
        self.nid = None
        self.last_status = None
        self.stop_event = threading.Event()
        self.icon_id = 1000  # ID unique pour l'icône
        self.capture_lock = threading.Lock()
        
        # Configurer le démarrage automatique
        set_autostart(True)
        
        # Enregistrer la classe de fenêtre
        wc = win32gui.WNDCLASS()
        wc.hInstance = win32gui.GetModuleHandle(None)
        wc.lpszClassName = "OdmServiceTrayClass"
        wc.lpfnWndProc = self.wnd_proc
        self.class_atom = win32gui.RegisterClass(wc)
        
    def create_window(self):
        # Créer la fenêtre de message
        self.hwnd = win32gui.CreateWindowEx(
            0,
            self.class_atom,
            "OdmService Tray",
            0,
            0, 0, 0, 0,
            0, 0,
            win32gui.GetModuleHandle(None),
            None
        )
        
        # Vérifier que la fenêtre a été créée
        if not self.hwnd:
            raise Exception("Échec de la création de la fenêtre")
            
        # Initialiser l'icône
        self.last_status = get_service_status()
        self.add_tray_icon()
        
        # Démarrer le thread de vérification
        threading.Thread(target=self.status_check_loop, daemon=True).start()
    
    def add_tray_icon(self):
        """Ajoute l'icône dans la zone de notification"""
        if not self.hwnd:
            return
            
        status = get_service_status()
        icon = ICONS['green' if status == win32service.SERVICE_RUNNING else 'red']
        status_text = "En cours d'exécution" if status == win32service.SERVICE_RUNNING else "Arrêté"
        tooltip = f"ODM Capture Poids Service - {status_text}"
        
        # Créer la structure NOTIFYICONDATA
        nid = (self.hwnd, self.icon_id,
               win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
               win32con.WM_USER + 20,
               icon,
               tooltip)
        
        # Ajouter l'icône
        try:
            win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
        except Exception as e:
            # Essayer avec une structure étendue pour les nouvelles versions de Windows
            try:
                nid = (self.hwnd, self.icon_id,
                       win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
                       win32con.WM_USER + 20,
                       icon,
                       tooltip,
                       0, 0, 0, None, None, win32gui.NIIF_NONE)
                win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
            except Exception as ex:
                error_msg = f"Erreur ajout icône: {str(ex)}"
                print(error_msg)
                win32api.MessageBox(0, error_msg, "Erreur", win32con.MB_ICONERROR)
    
    def update_tray_icon(self):
        """Met à jour l'icône dans la zone de notification"""
        if not self.hwnd:
            return
            
        status = get_service_status()
        icon = ICONS['green' if status == win32service.SERVICE_RUNNING else 'red']
        status_text = "En cours d'exécution" if status == win32service.SERVICE_RUNNING else "Arrêté"
        tooltip = f"ODM Capture Poids Service - {status_text}"
        
        # Créer la structure de mise à jour
        nid = (self.hwnd, self.icon_id,
               win32gui.NIF_ICON | win32gui.NIF_TIP,
               0,
               icon,
               tooltip)
        
        try:
            win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, nid)
        except Exception as e:
            # Si la modification échoue, réessayer avec une structure étendue
            try:
                nid = (self.hwnd, self.icon_id,
                       win32gui.NIF_ICON | win32gui.NIF_TIP,
                       0,
                       icon,
                       tooltip,
                       0, 0, 0, None, None, win32gui.NIIF_NONE)
                win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, nid)
            except Exception as ex:
                print(f"Erreur mise à jour icône: {ex}")
    
    def remove_tray_icon(self):
        """Supprime l'icône de la zone de notification"""
        if self.hwnd:
            try:
                win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, (self.hwnd, self.icon_id))
            except Exception as e:
                print(f"Erreur suppression icône: {e}")
    
    def wnd_proc(self, hwnd, msg, wparam, lparam):
        if msg == win32con.WM_DESTROY:
            self.cleanup()
            win32gui.PostQuitMessage(0)
            return 0
            
        elif msg == win32con.WM_USER + 20:
            if lparam == win32con.WM_RBUTTONUP:
                self.show_menu()
            elif lparam == win32con.WM_LBUTTONDBLCLK:
                show_logs()
            return 0
            
        elif msg == win32con.WM_COMMAND:
            cmd = win32gui.LOWORD(wparam)
            if cmd == 1001:  # Démarrer
                service_action("start")
            elif cmd == 1002:  # Arrêter
                service_action("stop")
            elif cmd == 1003:  # Redémarrer
                service_action("restart")
            elif cmd == 1004:  # Logs
                show_logs()
            elif cmd == 1005:  # Quitter
                self.cleanup()
                sys.exit(0)
            elif cmd == 1006:  # Capter le poids
                threading.Thread(target=self.capture_and_send_weight).start()
            return 0
            
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)
    
    def status_check_loop(self):
        """Vérifie périodiquement l'état du service"""
        while not self.stop_event.is_set():
            try:
                current_status = get_service_status()
                if current_status != self.last_status:
                    self.last_status = current_status
                    self.update_tray_icon()
            except Exception as e:
                print(f"Erreur vérification statut: {e}")
            time.sleep(5)
    
    def show_menu(self):
        menu = win32gui.CreatePopupMenu()
        status = get_service_status()
        
        # État actuel (élément non cliquable)
        status_str = "EN COURS" if status == win32service.SERVICE_RUNNING else "ARRÊTÉ"
        win32gui.AppendMenu(menu, win32con.MF_STRING | win32con.MF_DISABLED, 0, f"État: {status_str}")
        win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
        
        # Actions selon l'état
        if status != win32service.SERVICE_RUNNING:
            win32gui.AppendMenu(menu, win32con.MF_STRING, 1001, "Démarrer le service")
        else:
            win32gui.AppendMenu(menu, win32con.MF_STRING, 1002, "Arrêter le service")
            win32gui.AppendMenu(menu, win32con.MF_STRING, 1003, "Redémarrer le service")
        
        win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
        win32gui.AppendMenu(menu, win32con.MF_STRING, 1006, "Capter le poids")  
        win32gui.AppendMenu(menu, win32con.MF_STRING, 1004, "Voir les logs")
        win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
        win32gui.AppendMenu(menu, win32con.MF_STRING, 1005, "Quitter")
        
        # Afficher le menu au niveau du curseur
        pos = win32gui.GetCursorPos()
        win32gui.SetForegroundWindow(self.hwnd)
        win32gui.TrackPopupMenu(
            menu,
            win32con.TPM_LEFTALIGN | win32con.TPM_BOTTOMALIGN,
            pos[0],
            pos[1],
            0,
            self.hwnd,
            None
        )
        win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)
    
    def cleanup(self):
        """Nettoyage avant fermeture"""
        self.stop_event.set()
        self.remove_tray_icon()
        if self.hwnd:
            win32gui.DestroyWindow(self.hwnd)
        win32gui.UnregisterClass(self.class_atom, win32gui.GetModuleHandle(None))
    
    def capture_and_send_weight(self):
        """Gère le processus complet de capture et d'envoi"""
        # Vérifier le verrou pour éviter les captures simultanées
        if not self.capture_lock.acquire(blocking=False):
            win32api.MessageBox(0, "Une capture est déjà en cours", "Information", win32con.MB_ICONINFORMATION)
            return
        
        try:
            service_was_running = False
            current_status = get_service_status()
            
            # Si le service est en cours d'exécution, on l'arrête temporairement
            if current_status == win32service.SERVICE_RUNNING:
                service_was_running = True
                print("Arrêt temporaire du service...")
                if not service_action("stop"):
                    win32api.MessageBox(0, "Impossible d'arrêter le service", "Erreur", win32con.MB_ICONERROR)
                    return
                
                # Attendre l'arrêt complet du service (max 5 secondes)
                for _ in range(10):
                    if get_service_status() == win32service.SERVICE_STOPPED:
                        break
                    time.sleep(0.5)
                else:
                    win32api.MessageBox(0, "Le service n'a pas pu s'arrêter à temps", "Erreur", win32con.MB_ICONERROR)
                    return
            
            # Capture du poids
            # win32api.MessageBox(
            #     0, 
            #     "Placez l'objet sur la balance et attendez la stabilisation...", 
            #     "Capture en cours", 
            #     win32con.MB_ICONINFORMATION | win32con.MB_OK
            # )
            
            #print("Début de la capture du poids...")
            weight = capture_single_weight()
            
            if weight is None:
                win32api.MessageBox(
                    0, 
                    "Échec de la capture du poids\nVérifiez la connexion de la balance", 
                    "Erreur", 
                    win32con.MB_ICONERROR
                )
                return
            
            # Envoi à l'API
            #print(f"Envoi du poids {weight}kg à l'API...")
            if send_to_api(weight):
                win32api.MessageBox(
                    0, 
                    f"Poids capturé avec succès: {weight} kg", 
                    "Succès", 
                    win32con.MB_ICONINFORMATION
                )
            else:
                win32api.MessageBox(
                    0, 
                    f"Poids capturé: {weight} kg\nÉchec de l'envoi à l'API", 
                    "Avertissement", 
                    win32con.MB_ICONWARNING
                )
        except Exception as e:
            error_msg = f"Erreur critique: {str(e)}"
            print(error_msg)
            traceback.print_exc()
            win32api.MessageBox(0, error_msg, "Erreur", win32con.MB_ICONERROR)
        finally:
            # Redémarrer le service si nécessaire
            if service_was_running:
                print("Redémarrage du service...")
                service_action("start")
            
            self.capture_lock.release()
            print("Capture terminée")
    
    def run(self):
        """Lancer l'application"""
        self.create_window()
        win32gui.PumpMessages()

if __name__ == "__main__":
    # Cacher la console si compilé en exe
    if getattr(sys, 'frozen', False):
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    
    # Vérifier les privilèges admin
    if not is_user_admin():
        run_as_admin()
    
    app = ScaleTrayApp()
    try:
        app.run()
    except Exception as e:
        error_msg = f"Erreur initiale: {str(e)}"
        print(error_msg)
        traceback.print_exc()
        win32api.MessageBox(0, error_msg, "Erreur critique", win32con.MB_ICONERROR)