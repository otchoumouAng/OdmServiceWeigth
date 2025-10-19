import sqlite3
import os
from datetime import datetime

# --- Configuration ---
# Utilise ProgramData pour un stockage fiable, avec un fallback local
try:
    DB_DIR = os.path.join(os.getenv('ProgramData'), 'OdmService', 'database')
    if not os.path.exists(DB_DIR):
        os.makedirs(DB_DIR)
except Exception:
    DB_DIR = os.path.dirname(os.path.abspath(__file__))

DB_PATH = os.path.join(DB_DIR, 'poids.db')
print(f"Database path: {DB_PATH}")

# --- Fonctions de base de données ---

def get_db_connection():
    """Crée et retourne une connexion à la base de données SQLite."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row  # Permet d'accéder aux colonnes par nom
    return conn

def init_db():
    """Initialise la base de données et crée la table si elle n'existe pas."""
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS poids (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    valeur REAL NOT NULL,
                    desktop TEXT NOT NULL,
                    company TEXT NOT NULL,
                    date TEXT NOT NULL
                )
            """)
            conn.commit()
            print("Database initialized successfully.")
    except sqlite3.Error as e:
        print(f"Database initialization error: {e}")
        # Log this error appropriately in a real application
        raise

def add_poids(valeur, desktop, company):
    """
    Enregistre une nouvelle mesure de poids dans la base de données.
    Retourne l'ID de la nouvelle ligne ou None en cas d'erreur.
    """
    if valeur < 0:
        print("Error: Weight cannot be negative.")
        return None

    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            # Utilise le format ISO 8601 pour la date/heure
            current_date = datetime.utcnow().isoformat()
            cursor.execute(
                "INSERT INTO poids (valeur, desktop, company, date) VALUES (?, ?, ?, ?)",
                (valeur, desktop, company, current_date)
            )
            conn.commit()
            new_id = cursor.lastrowid
            print(f"Successfully added weight: {valeur} for {desktop}")
            return new_id
    except sqlite3.Error as e:
        print(f"Error adding weight to database: {e}")
        return None

def get_dernier_poids(desktop=None, company=None):
    """
    Récupère le dernier enregistrement de poids, avec filtres optionnels.
    Retourne un dictionnaire représentant la ligne, ou None si aucun résultat.
    """
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()

            # Construction de la requête de base
            query = "SELECT * FROM poids"

            # Ajout des filtres
            conditions = []
            params = []
            if desktop:
                conditions.append("desktop = ?")
                params.append(desktop)
            if company:
                conditions.append("company = ?")
                params.append(company)

            if conditions:
                query += " WHERE " + " AND ".join(conditions)

            # Tri pour obtenir le plus récent
            query += " ORDER BY date DESC LIMIT 1"

            cursor.execute(query, params)
            dernier_poids = cursor.fetchone()

            if dernier_poids:
                # Convertir l'objet Row en dictionnaire pour une utilisation facile
                return dict(dernier_poids)
            else:
                return None
    except sqlite3.Error as e:
        print(f"Error fetching last weight from database: {e}")
        return None

# --- Point d'entrée pour l'initialisation ---
if __name__ == '__main__':
    print("Initializing database...")
    init_db()

    # Exemple d'utilisation (peut être décommenté pour tester)
    # print("\n--- Testing datastore ---")
    # test_desktop = "TestPC"
    # test_company = "TestCorp"

    # # Ajout d'un poids
    # new_id = add_poids(12.5, test_desktop, test_company)
    # if new_id:
    #     print(f"Added new weight with ID: {new_id}")

    # # Récupération du dernier poids
    # dernier = get_dernier_poids(desktop=test_desktop, company=test_company)
    # if dernier:
    #     print(f"Retrieved last weight: {dernier['valeur']}kg on {dernier['date']}")
    # else:
    #     print("No weight found.")

    # # Test sans filtre
    # dernier_total = get_dernier_poids()
    # if dernier_total:
    #     print(f"Overall last weight: {dernier_total['valeur']}kg")
    # else:
    #     print("No weight found overall.")