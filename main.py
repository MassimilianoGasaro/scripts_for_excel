import openpyxl
import os
import configparser
import json
import requests
from typing import Optional, List, Dict, Any
from datetime import datetime

# Variabili globali
DEFAULT_FILE_PATH: str = "data.xlsx"
API_CONFIG: Dict[str, str] = {}
AUTH_TOKEN: Optional[str] = None

def main() -> None:
    workbook = open_excel_file(DEFAULT_FILE_PATH)
    if workbook is None:
        print("Impossibile aprire il file Excel")
        return
    
    sheet = workbook.active
    
    # Leggere i dati e creare il JSON
    data = read_excel_data(sheet)
    json_output = create_json_body(data)
    
    # Stampare il JSON
    print("\nJSON creato:")
    print(json.dumps(json_output, indent=2, ensure_ascii=False))
    
    # Salvare il JSON su file (opzionale)
    save_json_to_file(json_output, "output.json")
    
    # Eseguire il workflow API se la configurazione è disponibile
    if API_CONFIG and API_CONFIG['base_url']:
        execute_api_workflow(json_output)
    else:
        print("Configurazione API non disponibile. Saltando le chiamate HTTP.")
    
    workbook.close()

def read_excel_data(sheet) -> List[Dict[str, Any]]:
    """Legge i dati dal foglio Excel e li converte in una lista di dizionari"""
    data = []
    
    # Assumendo che la prima riga contenga le intestazioni
    headers = []
    visible_columns = []
    
    for col in range(1, 20):  # Aumentato il range per controllare più colonne
        # Verificare se la colonna è nascosta
        column_letter = sheet.cell(row=1, column=col).column_letter
        column_dimension = sheet.column_dimensions.get(column_letter)
        
        # Se la colonna è nascosta, saltarla
        if column_dimension and column_dimension.hidden:
            continue
            
        cell_value = sheet.cell(row=1, column=col).value
        if cell_value:
            headers.append(str(cell_value))
            visible_columns.append(col)
        elif not cell_value and col <= 6:  # Solo per le prime 6 colonne
            break
    
    print(f"Intestazioni trovate: {headers}")
    print(f"Colonne visibili: {visible_columns}")
    
    # Leggere i dati dalle righe successive
    row = 2
    while True:
        # Controllare se la riga è vuota
        if sheet.cell(row=row, column=3).value is None:
            break
            
        record = {}
        for i, header in enumerate(headers):
            col = visible_columns[i]  # Usare l'indice corretto della colonna visibile
            cell_value = sheet.cell(row=row, column=col).value
            
            if header == "Importo":
                # Convertire l'importo in valore positivo
                if cell_value is not None:
                    try:
                        # Rimuovere il segno negativo se presente
                        importo = abs(float(cell_value))
                        record[header] = importo
                    except (ValueError, TypeError):
                        record[header] = 0.0
                else:
                    record[header] = 0.0
            elif header == "Data":
                # Gestire la data
                if cell_value is not None:
                    if isinstance(cell_value, datetime):
                        record[header] = cell_value.strftime("%d/%m/%Y")
                    else:
                        record[header] = str(cell_value)
                else:
                    record[header] = ""
            else:
                # Altri campi come stringa
                record[header] = str(cell_value) if cell_value is not None else ""
        
        data.append(record)
        row += 1
    
    print(f"Trovati {len(data)} record")
    return data

def create_json_body(data: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Crea il corpo JSON per la chiamata API"""
    
    # Calcolare il totale degli importi
    total_amount = sum(record.get("Importo", 0) for record in data)
    
    json_body = {
        "transactions": data,
        "summary": {
            "total_records": len(data),
            "total_amount": round(total_amount, 2),
            "processed_date": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        }
    }
    
    return json_body

def save_json_to_file(data: Dict[str, Any], filename: str) -> None:
    """Salva il JSON su file"""
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(f"JSON salvato in: {filename}")
    except Exception as e:
        print(f"Errore nel salvare il JSON: {e}")

def login_api() -> bool:
    """Effettua il login e ottiene il token di autenticazione"""
    global AUTH_TOKEN
    
    if not API_CONFIG:
        print("Configurazione API non caricata")
        return False
    
    login_url = API_CONFIG['base_url'] + API_CONFIG['login_endpoint']
    login_data = {
        "username": API_CONFIG['username'],
        "password": API_CONFIG['password']
    }
    
    try:
        print(f"Tentativo di login su: {login_url}")
        response = requests.post(
            login_url,
            json=login_data,
            headers={'Content-Type': 'application/json'},
            timeout=30
        )
        
        if response.status_code == 200:
            response_data = response.json()
            # Assumendo che il token sia nel campo 'token' o 'access_token'
            AUTH_TOKEN = response_data.get('token') or response_data.get('access_token')
            
            if AUTH_TOKEN:
                print("Login effettuato con successo")
                return True
            else:
                print("Token non trovato nella risposta del login")
                return False
        else:
            print(f"Errore nel login: {response.status_code} - {response.text}")
            return False
            
    except requests.exceptions.RequestException as e:
        print(f"Errore di connessione durante il login: {e}")
        return False

def api_get_request(endpoint_override: Optional[str] = None) -> Optional[Dict[str, Any]]:
    """Effettua una chiamata GET autenticata"""
    global AUTH_TOKEN
    
    if not AUTH_TOKEN:
        print("Token di autenticazione non disponibile. Effettuare prima il login.")
        return None
    
    endpoint = endpoint_override or API_CONFIG['get_endpoint']
    url = API_CONFIG['base_url'] + endpoint
    
    headers = {
        'Authorization': f'Bearer {AUTH_TOKEN}',
        'Content-Type': 'application/json'
    }
    
    try:
        print(f"Chiamata GET a: {url}")
        response = requests.get(url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            print("Chiamata GET completata con successo")
            return response.json()
        elif response.status_code == 401:
            print("Token scaduto o non valido. Rieffettuare il login.")
            AUTH_TOKEN = None
            return None
        else:
            print(f"Errore nella chiamata GET: {response.status_code} - {response.text}")
            return None
            
    except requests.exceptions.RequestException as e:
        print(f"Errore di connessione nella chiamata GET: {e}")
        return None

def api_post_request(data: Dict[str, Any], endpoint_override: Optional[str] = None) -> bool:
    """Effettua una chiamata POST autenticata"""
    global AUTH_TOKEN
    
    if not AUTH_TOKEN:
        print("Token di autenticazione non disponibile. Effettuare prima il login.")
        return False
    
    endpoint = endpoint_override or API_CONFIG['post_endpoint']
    url = API_CONFIG['base_url'] + endpoint
    
    headers = {
        'Authorization': f'Bearer {AUTH_TOKEN}',
        'Content-Type': 'application/json'
    }
    
    try:
        print(f"Chiamata POST a: {url}")
        response = requests.post(url, json=data, headers=headers, timeout=30)
        
        if response.status_code in [200, 201]:
            print("Chiamata POST completata con successo")
            return True
        elif response.status_code == 401:
            print("Token scaduto o non valido. Rieffettuare il login.")
            AUTH_TOKEN = None
            return False
        else:
            print(f"Errore nella chiamata POST: {response.status_code} - {response.text}")
            return False
            
    except requests.exceptions.RequestException as e:
        print(f"Errore di connessione nella chiamata POST: {e}")
        return False

def execute_api_workflow(json_data: Dict[str, Any]) -> None:
    """Esegue il workflow completo: login, GET, POST"""
    print("\n=== Inizio workflow API ===")
    
    # 1. Login
    if not login_api():
        print("Impossibile effettuare il login. Workflow interrotto.")
        return
    
    # 2. Chiamata GET (opzionale - per verificare dati esistenti)
    print("\n--- Chiamata GET ---")
    existing_data = api_get_request()
    if existing_data:
        print(f"Dati esistenti trovati: {len(existing_data)} record")
    
    # 3. Chiamata POST con i dati del JSON
    print("\n--- Chiamata POST ---")
    if api_post_request(json_data):
        print("Dati inviati con successo!")
    else:
        print("Errore nell'invio dei dati")
    
    print("=== Fine workflow API ===\n")

# open an excel file
def open_excel_file(file_path: str) -> Optional[openpyxl.Workbook]:
    try:
        # data_only=True fa sì che vengano letti i valori calcolati invece delle formule
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        print("Excel file opened successfully.")
        return workbook
    except Exception as e:
        print(f"Error opening Excel file: {e}")
        return None
    
def setup() -> None:
    global DEFAULT_FILE_PATH, API_CONFIG
    config_file_path = "file_da_caricare.ini"
    
    if os.path.exists(config_file_path):
        config = configparser.ConfigParser()
        config.read(config_file_path)
        
        # Leggere il percorso del file
        if 'DEFAULT' in config and 'file_path' in config['DEFAULT']:
            DEFAULT_FILE_PATH = config['DEFAULT']['file_path']
            print(f"Using file path from config: {DEFAULT_FILE_PATH}")
        
        # Leggere la configurazione API
        if 'API' in config:
            API_CONFIG = {
                'base_url': config['API'].get('base_url', ''),
                'login_endpoint': config['API'].get('login_endpoint', ''),
                'get_endpoint': config['API'].get('get_endpoint', ''),
                'post_endpoint': config['API'].get('post_endpoint', ''),
                'username': config['API'].get('username', ''),
                'password': config['API'].get('password', '')
            }
            print(f"API Config loaded: {API_CONFIG['base_url']}")
        else:
            print("Sezione [API] non trovata nel file di configurazione")
    else:
        print("File di configurazione non trovato")
        
if __name__ == "__main__":
    setup()
    main()