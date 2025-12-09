import os
import pandas as pd
from flask import Flask, render_template, request, send_file, redirect, url_for, session
from weasyprint import HTML
from tempfile import NamedTemporaryFile
from urllib.parse import urljoin
from zipfile import ZipFile
from io import BytesIO
import shutil
import locale

# Aggiungi questi import all'inizio del file app_altaformazione.py
import locale
import pandas as pd
# ... (altri import) ...

# Imposta la locale italiana una sola volta all'inizio del tuo script
# Questo deve avvenire PRIMA della definizione delle route Flask
try:
    # Tenta di impostare la locale italiana su sistemi Windows
    locale.setlocale(locale.LC_TIME, 'ita')
except locale.Error:
    try:
        # Tenta di impostare la locale italiana su sistemi Linux/macOS
        locale.setlocale(locale.LC_TIME, 'it_IT.UTF-8')
    except locale.Error:
        # Fallback nel caso in cui le precedenti non funzionino
        print("Attenzione: Locale Italiana non impostata correttamente. I mesi saranno in inglese.")
# ... (continua con il resto del tuo script) ...

# --- Configurazione Iniziale ---
app = Flask(__name__)
# CRUCIALE: Chiave segreta necessaria per usare le sessioni (dove salviamo i percorsi dei file)
app.secret_key = 'una_chiave_segreta_molto_forte_e_casuale_qui' 

app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024 
ALLOWED_EXTENSIONS = {'xlsx', 'xls'} 
# Directory temporanea per salvare i PDF generati
TEMP_PDF_DIR = os.path.join(app.root_path, 'temp_pdfs_af')

# Funzione per il controllo estensione (omissis)
# def allowed_file(filename): ...
# Funzione parse_excel_data (omissis)
# ...

# --- FUNZIONE DI PARSING DA EXCEL ---

def format_name_with_exceptions(name_str):
    """
    Formatta una stringa di nome/cognome:
    1. Capitalizza le parole standard (es. Rossi -> Rossi).
    2. Converte in minuscolo le parole precedute dal prefisso '%' (es. %DE -> de).
    3. Rimuove il prefisso '%'.
    """
    if not name_str:
        return ""

    # Split la stringa in parole
    words = name_str.split()
    formatted_parts = []
    
    for word in words:
        if word.startswith('%'):
            # Caso Eccezione: Rimuovi '%' e converti il resto in minuscolo
            # Esempio: "%DE" -> "de"
            cleaned_word = word[1:].lower()
            formatted_parts.append(cleaned_word)
        else:
            # Caso Standard: Converte la parola in minuscolo e capitalizza la prima lettera
            # Esempio: "MARIO" -> "Mario"
            formatted_parts.append(word.lower().capitalize())
            
    return ' '.join(formatted_parts)




def parse_excel_data(file_stream):
    import pandas as pd
    from io import BytesIO # Assicurati che BytesIO sia importato

    try:
        # 1. Copia il contenuto del file stream in un buffer in memoria
        file_stream.seek(0) # Torna all'inizio del file, fondamentale!
        data = file_stream.read()
        excel_data_buffer = BytesIO(data)

        # 2. Leggi dal buffer in memoria
        # Usa sheet_name=0 per leggere il primo foglio
        # Usa header=1 per saltare la riga vuota iniziale (indice 0)
        df = pd.read_excel(excel_data_buffer, header=0, sheet_name=0,engine='openpyxl',dtype={'data': str, 'data_di_nascita': str}) 
        
        # 3. Pulisce i nomi delle colonne
        df.columns = df.columns.str.strip().str.replace('\n', ' ').str.replace(' ', '_')

        if df.empty:
            print("PANDAS DEBUG: DataFrame è vuoto dopo la lettura.")
            return None 

        # 4. Processa i dati (il tuo codice di mappatura)
        students_data = df.fillna('').to_dict('records')
        
        normalized_data = []
        for student in students_data:
            # Crea il dizionario in minuscolo
            student_dict = {k.lower(): str(v).strip() for k, v in student.items()}
            
            # --- 1. Campi Composti e Derivati ---
            
            # Nome e Cognome uniti
            nome_completo = f"{student_dict.get('nome', '')} {student_dict.get('cognome', '')}"
            student_dict['nom_cog'] = format_name_with_exceptions(nome_completo.strip())           
            # Formattazione per "nato a" / "nata a"
            sesso = student_dict.get('sesso', '').upper()
            student_dict['sesso_formattato'] = 'nato a' if sesso == 'M' else 'nata a'
            student_dict['decreto'] = student_dict.get('decreto', '')
            student_dict['data_decreto'] = student_dict.get('del', '')

            # Elisione (A/Ad): Non è richiesta dal testo del diploma fornito, 
            # ma la mantengo se necessaria per altri testi:
            first_letter = student_dict.get('nom_cog', ' ')[0].upper()
            student_dict['intro_nome'] = 'Ad' if first_letter in ['A', 'E', 'I', 'O', 'U'] else 'A'

            # --- LOGICA COMPLESSA LUOGO DI NASCITA ---
            comune = student_dict.get('nato_a', '').strip()
            provincia = student_dict.get('provincia_di_nascita', '').strip()
            stato = student_dict.get('stato_di_nascita', '').strip()

            luogo_nascita = comune
            
            # Normalizzazione per confronti (ignora maiuscole/minuscole e spazi)
            comune_upper = comune.upper().replace(' ', '')
            provincia_upper = provincia.upper().replace(' ', '')
            stato_upper = stato.upper().replace(' ', '')
            
            # Determina se la nascita è in Italia (stato vuoto o Italia/IT)
            # Spesso i campi italiani hanno stato vuoto, o "Italia", o "IT".
            is_italy = not stato_upper or stato_upper in ['ITALIA', 'IT', 'I']

            if is_italy:
                # Gestione Nascita in Italia (Casi 1 & 2)
                if provincia_upper and comune_upper == provincia_upper:
                    # Caso 1: Città = Provincia (es. ROMA) -> Scrivo solo la città
                    luogo_nascita = comune.capitalize()
                elif provincia_upper:
                    # Caso 2: Città != Provincia (es. Frascati (Roma)) -> Scrivo Città (Provincia)
                    luogo_nascita = f"{comune.capitalize()} ({provincia.capitalize()})"
                # Se non c'è provincia, resta solo il comune (luogo_nascita = comune)
            
            elif stato:
                # Caso 3: Nascita Estera (es. Suceava (ROMANIA))
                luogo_nascita = f"{comune.capitalize()} ({stato.capitalize()})"

            # Salva il risultato finale nel nuovo campo
            student_dict['luogo_nascita_formattato'] = luogo_nascita           
            # Formattazione della data di stampa (data del diploma)
            # Assumo che 'data_diploma' sia una data leggibile o un oggetto datetime
            data_diploma_raw = str(student_dict.get('data', '').strip())
            data_nascita_raw = str(student_dict.get('data_di_nascita', '').strip())
            try:
                # Tenta di convertire e formattare se è una data valida
                if data_diploma_raw:
                    data_obj = pd.to_datetime(data_diploma_raw, format='%Y-%m-%d %H:%M:%S', errors='coerce')
                    if pd.isna(data_obj): # Se la data non è valida dopo il tentativo
                        raise ValueError(f"Formato data non valido: {data_nascita_raw}")                   
                    # Formattazione come "1 dicembre 2025"
                    student_dict['datastampa'] = data_obj.strftime('%#d %B %Y')
                else:
                    student_dict['datastampa'] = 'Data non disponibile'
            except Exception as e:
                print(f"Errore formattazione data: {e}") # Debugging
                student_dict['datastampa'] = data_diploma_raw # Mantiene il testo originale

            try:
                # Tenta di convertire e formattare se è una data valida
                if data_nascita_raw:
                    data_obj = pd.to_datetime(data_nascita_raw, format='%Y-%m-%d %H:%M:%S', errors='coerce')

                    if pd.isna(data_obj): # Se la data non è valida dopo il tentativo
                        raise ValueError(f"Formato data non valido: {data_nascita_raw}")
                    
                    data_formattata= data_obj.strftime('%#d %B %Y')
                    # Logica per la maiuscola sul mese (che avevi commentato ma è necessaria)
                    parti = data_formattata.split()
                    if len(parti) > 1:
                        parti[1] = parti[1]
                    student_dict['data_di_nascita'] = ' '.join(parti)
                else:
                    student_dict['data_di_nascita'] = 'Data non disponibile'
            except Exception as e:
                print(f"Errore formattazione data: {e}") # Debugging
                student_dict['data_di_nascita'] = data_nascita_raw # Mantiene il testo originale

            # --- 2. Mappatura Campi Semplici/Ridenominati per il Template ---
            #student_dict['comune_nascita'] = student_dict.get('nato_a', '').capitalize()
            #student_dict['provincia'] = student_dict.get('provincia_di_nascita', '').capitalize()
            #student_dict['stato'] = student_dict.get('stato_di_nascita', '').capitalize()
            #student_dict['dipartimento'] = student_dict.get('dipartimento_di', '')
            student_dict['master'] = student_dict.get('master', '').strip().lower()
            student_dict['corso_laurea'] = student_dict.get('tipologia', '').lower()
            student_dict['tipologia_corso'] = student_dict.get('classe_accademica', '').lower()
            # Il campo 'cfu' è già presente come cfu nel template
            # I campi anno_accademico, comune_nascita, data_di_nascita sono già mappati correttamente.
            
            normalized_data.append(student_dict)
        return normalized_data
        
    except Exception as e:
        # Stampa l'errore completo nel log della console per vedere l'errore esatto
        print(f"ERRORE GRAVE DI PANDAS: {e}") 
        return None
# --- FINE FUNZIONE DI PARSING ---

ALLOWED_EXTENSIONS = {'xlsx', 'xls'} 
# Directory temporanea per salvare i PDF generati
TEMP_PDF_DIR = os.path.join(app.root_path, 'temp_pdfs_af')

# --- DEFINIZIONE NECESSARIA DELLA FUNZIONE ALLOWED_FILE ---
def allowed_file(filename):
    """
    Controlla se l'estensione del file è tra quelle permesse (xlsx o xls).
    """
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
# -----------------------------------------------------------

# --- ROUTE 1: Pagina Principale / Upload ---
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Reindirizza alla funzione di upload
        return redirect(url_for('upload_excel'))
    
    # Per richieste GET, mostra il form di upload
    return render_template('upload_excel.html')

# --- ROUTE 2: Gestione dell'Upload e della Generazione PDF ---
@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    NOME_FILE_FIRMA = 'M_mgcrespi.png' # Esempio di file (maschile)
    if 'file' not in request.files or request.files['file'].filename == '':
        return render_template('error.html', message="Nessun file selezionato."), 400
    
    file = request.files['file']
        
    if file and allowed_file(file.filename):
        original_filename = file.filename
        session['original_excel_filename'] = original_filename
        students_data = parse_excel_data(file) 
        
        if students_data is None or not students_data:
            return render_template('error.html', message="Impossibile leggere i dati o file vuoto."), 400
            
        # 1. Preparazione Ambiente di Lavoro Temporaneo
        if os.path.exists(TEMP_PDF_DIR):
            shutil.rmtree(TEMP_PDF_DIR) # Pulisce la cartella precedente
        os.makedirs(TEMP_PDF_DIR)
        
        STATIC_ROOT_PATH = os.path.join(app.root_path, app.static_folder)
        base_url_for_weasyprint = urljoin('file:///', STATIC_ROOT_PATH.replace('\\', '/'))

        # AGGIUNGI QUESTO: Garantisce uno slash finale, essenziale per WeasyPrint
        if not base_url_for_weasyprint.endswith('/'):
            base_url_for_weasyprint += '/' # ⬅️ NUOVA RIGA

        print(f"DEBUG URL Base: {base_url_for_weasyprint}") #

        generated_pdf_paths = []
        
        for i, student in enumerate(students_data):
            # ... (logica di preparazione dei dati e elisione AD/A) ...
            
            template_filename = 'diploma_ssas.html' 
            student['firma_filename'] = 'M_mgcrespi.png'
            # Utilizza il nome del corso e matricola per il nome file
            output_filename = f"{student.get('cognome','Sconosciuto')}_{student.get('matricola', i)}.pdf"
            output_path = os.path.join(TEMP_PDF_DIR, output_filename)
            
            rendered_html = render_template(
                template_filename,
                **student
            )

            try:
                html_doc = HTML(string=rendered_html, base_url=base_url_for_weasyprint)
                html_doc.write_pdf(output_path)
                generated_pdf_paths.append(output_path)

            except Exception as e:
                # Se un file fallisce, interrompi e segnala l'errore
                return render_template('error.html', message=f"Errore generazione PDF per {student.get('nom_cog')}: {e}"), 500

        # 2. Salva il percorso dei file generati nella sessione
        session['generated_pdf_paths'] = generated_pdf_paths
        
        # 3. Reindirizza alla pagina di download
        return redirect(url_for('download_page', count=len(generated_pdf_paths)))
    
    return render_template('error.html', message="Estensione file non permessa. Usa .xlsx o .xls"), 400

# --- ROUTE 3: Pagina di Download ---
@app.route('/download')
def download_page():
    count = request.args.get('count', 0)
    return render_template('download_page.html', pdf_count=count)

# --- ROUTE 4: Generazione e Invio dello ZIP ---
@app.route('/download-batch')
def download_batch():
    pdf_paths = session.pop('generated_pdf_paths', None) # Preleva i percorsi e pulisci la sessione
    original_filename = session.pop('original_excel_filename', 'diplomi_alta_formazione.zip')
    if not pdf_paths:
        return render_template('error.html', message="Nessun file PDF da scaricare. Riprova con un nuovo upload."), 404

    # Genera il nome del file ZIP
    if original_filename.endswith('.xlsx'):
        zip_name = original_filename.replace('.xlsx', '.zip')
    elif original_filename.endswith('.xls'):
        zip_name = original_filename.replace('.xls', '.zip')
    else:
        # Fallback se l'estensione non è stata trovata o era un caso limite
        zip_name = 'diplomi_alta_formazione.zip'

    # Crea un file ZIP in memoria (non su disco)
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, 'w') as zf:
        for pdf_path in pdf_paths:
            # Aggiungi il file allo zip. Usa os.path.basename per il nome all'interno dello zip.
            zf.write(pdf_path, os.path.basename(pdf_path))

    # Pulizia: Rimuovi la directory temporanea dopo aver creato lo zip
    if os.path.exists(TEMP_PDF_DIR):
        shutil.rmtree(TEMP_PDF_DIR)
        
    zip_buffer.seek(0)
    
    # Invia il file ZIP all'utente
    return send_file(
        zip_buffer,
        mimetype='application/zip',
        as_attachment=True,
        download_name=zip_name
    )

if __name__ == '__main__': app.run(debug=True)