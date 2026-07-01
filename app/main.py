from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
import os
import uuid
import mimetypes
import re
from typing import Optional
from io import BytesIO
from openpyxl import load_workbook
from fastapi import FastAPI, Request, Form, Depends, HTTPException, UploadFile, File
from fastapi.responses import RedirectResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session
from sqlalchemy import or_, text
from passlib.context import CryptContext
from starlette.middleware.sessions import SessionMiddleware
from supabase import create_client

SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")

print("SUPABASE_URL =", SUPABASE_URL)
print("SUPABASE_KEY cargada =", bool(SUPABASE_KEY))
supabase = create_client(SUPABASE_URL, SUPABASE_KEY) if SUPABASE_URL and SUPABASE_KEY else None
from .database import Base, engine, get_db
from .models import User, Owner, Patient, ClinicalEvent, EventAttachment, Appointment, Product, Sale, SaleItem, SalePayment, WaitingListEntry, Hospitalization, HospitalizationMedication, HospitalizationFluid
# Base.metadata.create_all(bind=engine)
# =====================================
# Zona horaria Argentina
# =====================================

ARG_TZ = ZoneInfo("America/Argentina/Buenos_Aires")

def argentina_now():
    return datetime.now(ARG_TZ).replace(tzinfo=None)
# ==========================================================
# IA CLÍNICA
# ==========================================================

def ai_clinical_summary(event):
    import json
    import urllib.request
    import urllib.error

    api_key = os.getenv("OPENAI_API_KEY", "").strip()

    fallback = {
        "summary": "Asistente IA no disponible todavía.",
        "owner_explanation": "",
        "priority": "Normal",
        "problems": [],
        "differentials": [],
        "recommended_tests": [],
        "treatment": [],
        "alerts": []
    }

    patient = getattr(event, "patient", None)
    previous_events_text = ""

    try:
        previous_events = (
            event.patient.events
            if getattr(event, "patient", None) and getattr(event.patient, "events", None)
            else []
        )

        previous_events = sorted(
            previous_events,
            key=lambda e: e.event_date or datetime.min,
            reverse=True
        )

        previous_events = [
            e for e in previous_events
            if e.id != event.id
        ][:8]

        previous_events_lines = []

        for previous in previous_events:
            previous_events_lines.append(
                f"""
Fecha: {previous.event_date.strftime('%d/%m/%Y') if previous.event_date else ''}
Tipo: {previous.event_type or ''}
Título: {previous.title or ''}
Descripción: {previous.description or ''}
Anamnesis: {previous.anamnesis or ''}
Examen físico: {previous.physical_exam or ''}
Diagnóstico: {previous.diagnosis or ''}
Tratamiento: {previous.treatment or ''}
"""
            )

        previous_events_text = "\n---\n".join(previous_events_lines)

    except Exception:
        previous_events_text = ""
    age_text = ""

    try:
        if patient and getattr(patient, "birth_date", None):
            today = argentina_now().date()
            birth = patient.birth_date
            years = today.year - birth.year - ((today.month, today.day) < (birth.month, birth.day))
            age_text = f"{years} años"
    except Exception:
        age_text = ""
    clinical_text = f"""
Paciente: {getattr(patient, "name", "") if patient else ""}
Especie: {getattr(patient, "species", "") if patient else ""}
Raza: {getattr(patient, "breed", "") if patient else ""}
Sexo: {getattr(patient, "sex", "") if patient else ""}
Edad: {age_text}
Castrado: {getattr(patient, "neutered", "") if patient else ""}
Fecha del evento: {event.event_date.strftime('%d/%m/%Y %H:%M') if event.event_date else ""}
Peso paciente: {getattr(patient, "weight", "") if patient else ""}
Peso evento: {getattr(event, "weight", "") or ""}

Tipo de evento: {getattr(event, "event_type", "") or ""}
Título: {getattr(event, "title", "") or ""}
Descripción: {getattr(event, "description", "") or ""}
Anamnesis: {getattr(event, "anamnesis", "") or ""}
Examen físico: {getattr(event, "physical_exam", "") or ""}
Diagnóstico cargado: {getattr(event, "diagnosis", "") or ""}
Tratamiento cargado: {getattr(event, "treatment", "") or ""}

Constantes:
Temperatura: {getattr(event, "temperature", "") or ""}
FC: {getattr(event, "heart_rate", "") or ""}
FR: {getattr(event, "respiratory_rate", "") or ""}
Mucosas: {getattr(event, "mucous_membranes", "") or ""}
TLLC/CRT: {getattr(event, "crt", "") or ""}
Hidratación: {getattr(event, "hydration", "") or ""}

ECG:
FC ECG: {getattr(event, "ecg_hr", "") or ""}
Ritmo: {getattr(event, "ecg_rhythm", "") or ""}
PR: {getattr(event, "ecg_pr", "") or ""}
QRS: {getattr(event, "ecg_qrs", "") or ""}
QT: {getattr(event, "ecg_qt", "") or ""}
Eje: {getattr(event, "ecg_axis", "") or ""}
Interpretación ECG: {getattr(event, "ecg_interpretation", "") or ""}

Eco:
AI/Ao: {getattr(event, "eco_aiao", "") or ""}
FS: {getattr(event, "eco_fs", "") or ""}
ACVIM: {getattr(event, "eco_acvim", "") or ""}
Diagnóstico eco: {getattr(event, "eco_diagnosis", "") or ""}
Tratamiento eco: {getattr(event, "eco_treatment", "") or ""}

RX:
VHS: {getattr(event, "rx_vhs", "") or ""}
VLAS: {getattr(event, "rx_vlas", "") or ""}
Patrón pulmonar: {getattr(event, "rx_lung_pattern", "") or ""}
Edema: {getattr(event, "rx_edema", "") or ""}
Congestión: {getattr(event, "rx_congestion", "") or ""}
Observaciones RX: {getattr(event, "rx_observations", "") or ""}
Historia clínica previa relevante:
{previous_events_text}
"""

    if not api_key:
        fallback["summary"] = "Falta configurar OPENAI_API_KEY en Render."
        fallback["alerts"] = ["No se encontró OPENAI_API_KEY."]
        return fallback
    event_type_text = (getattr(event, "event_type", "") or "").lower()

    study_instruction = ""

    if "radiografía" in event_type_text or "rx" in event_type_text:
        study_instruction = """
Este evento corresponde a RADIOGRAFÍA.
Si hay imágenes adjuntas, analizalas como radiografías veterinarias.
Describí hallazgos radiológicos con prudencia.
No afirmes lesiones si la calidad o proyección no lo permite.
Correlacioná siempre con la clínica.
"""
    elif "ecg" in event_type_text:
        study_instruction = """
Este evento corresponde a ELECTROCARDIOGRAMA.
Priorizá ritmo, frecuencia, intervalos, eje eléctrico, alteraciones de conducción y signos compatibles con agrandamiento de cámaras.
Si hay imagen adjunta, interpretala junto con los valores cargados.
"""
    elif "ecocardiografía" in event_type_text:
        study_instruction = """
Este evento corresponde a ECOCARDIOGRAFÍA.
Priorizá AI/Ao, FS, FE, EPSS, ACVIM, válvulas, cámaras cardíacas, función sistólica y signos de congestión o hipertensión pulmonar.
"""
    elif "laboratorio" in event_type_text:
        study_instruction = """
Este evento corresponde a LABORATORIO.
Interpretá alteraciones hematológicas y bioquímicas en contexto clínico.
Indicá patrones compatibles, diferenciales y estudios complementarios si faltan datos.
"""
    image_count = 0

    content_blocks = []

    try:
        for attachment in getattr(event, "attachments", []) or []:
            file_url = getattr(attachment, "file_path", "") or ""
            file_name = (getattr(attachment, "filename", "") or "").lower()

            if (
                file_url.startswith("http")
                and (
                    file_name.endswith(".jpg")
                    or file_name.endswith(".jpeg")
                    or file_name.endswith(".png")
                    or file_name.endswith(".webp")
                )
            ):
                content_blocks.append({
                    "type": "input_image",
                    "image_url": file_url
                })

                image_count += 1

                if image_count >= 4:
                    break

    except Exception:
        pass

    prompt = f"""
Sos un veterinario clínico experto en pequeños animales.

{study_instruction}

Analizá TODA la información clínica disponible.
Si hay imágenes adjuntas, utilizalas junto con los datos clínicos.
Nunca inventes hallazgos.

Respondé únicamente en JSON con esta estructura:

{{
  "summary":"",
  "owner_explanation":"",
  "priority":"Normal",
  "problems":[],
  "differentials":[],
  "recommended_tests":[],
  "treatment":[],
  "alerts":[]
}}

Información clínica:

{clinical_text}

Cantidad de imágenes adjuntas enviadas para análisis:
{image_count}
"""

    content_blocks.insert(0, {
        "type": "input_text",
        "text": prompt
    })
    payload = {
        "model": os.getenv("OPENAI_MODEL", "gpt-5.4-mini"),
        "input": [
            {
                "role": "user",
                "content": content_blocks
            }
        ],
        "max_output_tokens": 1400
    }
    try:
        req = urllib.request.Request(
            "https://api.openai.com/v1/responses",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            },
            method="POST"
        )

        with urllib.request.urlopen(req, timeout=35) as response:
            data = json.loads(response.read().decode("utf-8"))

        output_text = data.get("output_text", "")

        if not output_text:
            parts = []
            for item in data.get("output", []):
                for content in item.get("content", []):
                    if content.get("type") in ["output_text", "text"]:
                        parts.append(content.get("text", ""))
            output_text = "\n".join(parts).strip()

        cleaned = output_text.strip()
        if cleaned.startswith("```"):
            cleaned = cleaned.replace("```json", "").replace("```", "").strip()

        result = json.loads(cleaned)

        for key in fallback:
            if key not in result:
                result[key] = fallback[key]

        if result.get("priority") not in ["Normal", "Alta", "Crítica"]:
            result["priority"] = "Normal"

        return result

    except Exception as e:
        print("ERROR IA CLINICA:", str(e))
        fallback["summary"] = "No se pudo generar el análisis IA en este momento."
        fallback["alerts"] = [f"Error IA: {str(e)}"]
        return fallback

app = FastAPI(title='Los Aromos Cloud')
app.add_middleware(SessionMiddleware, secret_key=os.getenv('SECRET_KEY', 'dev-secret-change-me'))
app.mount('/static', StaticFiles(directory='app/static'), name='static')
os.makedirs("app/uploads", exist_ok=True)
os.makedirs('app/uploads', exist_ok=True)
app.mount('/uploads', StaticFiles(directory='app/uploads'), name='uploads')
templates = Jinja2Templates(directory='app/templates')
def get_pending_count():
    return 0

def get_waiting_count():
    return 0

templates.env.globals['get_waiting_count'] = get_waiting_count
templates.env.globals['get_pending_count'] = get_pending_count
pwd_context = CryptContext(schemes=['bcrypt'], deprecated='auto')

EVENT_TYPES = ['Consulta clínica','Control','Vacuna','Desparasitación','Radiografía','ECG','Ecocardiografía','Ecografía','Laboratorio','Cirugía','Anestesia','Internación','Alta','Otro procedimiento']
DUE_ACTIVE_MARKER = '[DEUDA_ACTIVA]'
DUE_CLOSED_MARKER = '[DEUDA_CERRADA]'


def pending_action_matches_event(pending_event, real_event):
    pending_type = (pending_event.event_type or '').lower()
    real_type = (real_event.event_type or '').lower()

    pending_text = f'{pending_event.title or ""} {pending_event.description or ""}'.lower()
    real_text = f'{real_event.title or ""} {real_event.description or ""} {real_event.vaccine_name or ""} {real_event.dewormer_product or ""}'.lower()

    if pending_type == real_type:
        return True

    if 'vacun' in pending_text and real_type == 'vacuna':
        return True

    if (
        'despar' in pending_text
        or 'antiparas' in pending_text
        or 'aprax' in pending_text
        or 'pipeta' in pending_text
    ) and real_type == 'desparasitación':
        return True

    if 'control' in pending_text and real_type in ['control', 'consulta clínica']:
        return True

    return False


def close_pending_actions_after_real_event(db, patient, real_event):
    pending_actions = (
        db.query(ClinicalEvent)
        .filter(ClinicalEvent.patient_id == patient.id)
        .filter(ClinicalEvent.description.ilike(f'%{DUE_ACTIVE_MARKER}%'))
        .all()
    )

    for pending_event in pending_actions:
        if pending_action_matches_event(pending_event, real_event):
            pending_event.description = (pending_event.description or '').replace(
                DUE_ACTIVE_MARKER,
                DUE_CLOSED_MARKER
            )
            pending_event.reminder_date = None
def init_db():
    # Base.metadata.create_all(bind=engine)
    db = next(get_db())

    with engine.begin() as conn:
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS weight FLOAT"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS temperature FLOAT"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS heart_rate INTEGER"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS respiratory_rate INTEGER"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS mucous_membranes TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS crt TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS hydration TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS anamnesis TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS physical_exam TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS vaccine_name TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS vaccine_lot TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS vaccine_expiration TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS next_vaccine_date TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS dewormer_product TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS dewormer_drug TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS dewormer_dose TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS next_deworming_date TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_hr TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_rhythm TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_p TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_pr TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_qrs TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_st TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_t TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_qt TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_axis TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_interpretation TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_aiao TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_fs TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_acvim TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_diagnosis TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_treatment TEXT DEFAULT ''"))
        # ECG ampliado
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_p_mv TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_qrs_mv TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_t_mv TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_qtc TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_polarity TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_arrhythmia TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_conduction TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS ecg_notes TEXT DEFAULT ''"))

        # Ecocardiografía ampliada
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_epss TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_lvidd TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_lvids TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_ivsd TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_ivss TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_lvpwd TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_lvpws TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_fe TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_la_size TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_lv_size TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_rv_size TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_ra_size TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_mitral TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_tricuspid TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_aortic TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_pulmonary TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_pulmonary_htn TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_pericardium TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_doppler TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS eco_observations TEXT DEFAULT ''"))

        # Radiografía cardiológica
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS rx_vhs TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS rx_vlas TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS rx_heart_size TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS rx_left_atrium TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS rx_left_heart TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS rx_right_heart TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS rx_pulmonary_vessels TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS rx_lung_pattern TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS rx_edema TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS rx_congestion TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS rx_trachea TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE clinical_events ADD COLUMN IF NOT EXISTS rx_observations TEXT DEFAULT ''"))
        conn.execute(text("ALTER TABLE sales ADD COLUMN IF NOT EXISTS patient_id INTEGER REFERENCES patients(id)"))
        conn.execute(text("ALTER TABLE sales ADD COLUMN IF NOT EXISTS owner_id INTEGER REFERENCES owners(id)"))
        conn.execute(text("ALTER TABLE sales ADD COLUMN IF NOT EXISTS discount_percent FLOAT DEFAULT 0"))
        conn.execute(text("ALTER TABLE sales ADD COLUMN IF NOT EXISTS discount_amount FLOAT DEFAULT 0"))
        conn.execute(text("ALTER TABLE sales ADD COLUMN IF NOT EXISTS credit_surcharge_percent FLOAT DEFAULT 0"))
        conn.execute(text("ALTER TABLE sales ADD COLUMN IF NOT EXISTS credit_surcharge_amount FLOAT DEFAULT 0"))
        conn.execute(text("ALTER TABLE sales ADD COLUMN IF NOT EXISTS cost_total FLOAT DEFAULT 0"))
        conn.execute(text("ALTER TABLE sales ADD COLUMN IF NOT EXISTS profit_amount FLOAT DEFAULT 0"))
        conn.execute(text("ALTER TABLE sales ADD COLUMN IF NOT EXISTS margin_percent FLOAT DEFAULT 0"))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS vademecum_active_ingredients (
                id SERIAL PRIMARY KEY,
                name VARCHAR(200) DEFAULT '',
                category VARCHAR(120) DEFAULT '',
                species VARCHAR(120) DEFAULT '',
                dog_dose TEXT DEFAULT '',
                cat_dose TEXT DEFAULT '',
                route VARCHAR(120) DEFAULT '',
                frequency VARCHAR(120) DEFAULT '',
                indications TEXT DEFAULT '',
                contraindications TEXT DEFAULT '',
                interactions TEXT DEFAULT '',
                warnings TEXT DEFAULT '',
                observations TEXT DEFAULT '',
                active BOOLEAN DEFAULT TRUE,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """))

        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS vademecum_brands (
                id SERIAL PRIMARY KEY,
                active_ingredient_id INTEGER REFERENCES vademecum_active_ingredients(id),
                brand_name VARCHAR(200) DEFAULT '',
                laboratory VARCHAR(180) DEFAULT '',
                presentation VARCHAR(200) DEFAULT '',
                concentration VARCHAR(120) DEFAULT '',
                active BOOLEAN DEFAULT TRUE,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS vademecum_drugs (
                id SERIAL PRIMARY KEY,
                commercial_name VARCHAR(200) DEFAULT '',
                active_ingredient VARCHAR(200) DEFAULT '',
                category VARCHAR(120) DEFAULT '',
                species VARCHAR(120) DEFAULT '',
                dose TEXT DEFAULT '',
                route VARCHAR(120) DEFAULT '',
                frequency VARCHAR(120) DEFAULT '',
                indications TEXT DEFAULT '',
                contraindications TEXT DEFAULT '',
                observations TEXT DEFAULT '',
                active BOOLEAN DEFAULT TRUE,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """))       
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS event_attachments (
                id SERIAL PRIMARY KEY,
                event_id INTEGER REFERENCES clinical_events(id),
                filename VARCHAR(255) NOT NULL,
                file_path VARCHAR(500) NOT NULL
            )
        """))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS products (
                id SERIAL PRIMARY KEY,
                rubro VARCHAR(120) DEFAULT '',
                tipo VARCHAR(120) DEFAULT '',
                name VARCHAR(200) DEFAULT '',
                code VARCHAR(100) DEFAULT '',
                barcode VARCHAR(120) DEFAULT '',
                cost_price FLOAT,
                sale_price FLOAT,
                margin_percent FLOAT,
                stock FLOAT,
                min_stock FLOAT,
                expiration_date DATE,
                provider VARCHAR(180) DEFAULT '',
                manufacturer VARCHAR(180) DEFAULT '',
                active BOOLEAN DEFAULT TRUE,
                notes TEXT DEFAULT '',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """))
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS sale_payments (
            id SERIAL PRIMARY KEY,
            sale_id INTEGER REFERENCES sales(id),
            method VARCHAR(50),
            amount FLOAT DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """))
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS hospitalizations (
            id SERIAL PRIMARY KEY,
            patient_id INTEGER REFERENCES patients(id),
            clinical_event_id INTEGER REFERENCES clinical_events(id),
            admission_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            discharge_date TIMESTAMP NULL,
            status VARCHAR(40) DEFAULT 'Internado',
            cage VARCHAR(80) DEFAULT '',
            responsible_vet VARCHAR(120) DEFAULT '',
            reason TEXT DEFAULT '',
            diagnosis TEXT DEFAULT '',
            treatment_plan TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            initial_weight FLOAT,
            initial_temperature FLOAT,
            initial_heart_rate INTEGER,
            initial_respiratory_rate INTEGER,
            initial_mucous_membranes VARCHAR(100) DEFAULT '',
            initial_crt VARCHAR(50) DEFAULT '',
            initial_hydration VARCHAR(100) DEFAULT '',
            discharge_summary TEXT DEFAULT '',
            discharge_indications TEXT DEFAULT '',
            created_by VARCHAR(100) DEFAULT 'admin',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """)) 
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS hospitalization_medications (
            id SERIAL PRIMARY KEY,
            hospitalization_id INTEGER REFERENCES hospitalizations(id),
            medication_name VARCHAR(200) DEFAULT '',
            dose VARCHAR(120) DEFAULT '',
            route VARCHAR(80) DEFAULT '',
            frequency VARCHAR(120) DEFAULT '',
            scheduled_time VARCHAR(50) DEFAULT '',
            status VARCHAR(40) DEFAULT 'Pendiente',
            notes TEXT DEFAULT '',
            created_by VARCHAR(100) DEFAULT 'admin',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """))
        conn.execute(text("ALTER TABLE hospitalization_medications ADD COLUMN IF NOT EXISTS applied_at TIMESTAMP NULL"))
        conn.execute(text("ALTER TABLE hospitalization_fluids ADD COLUMN IF NOT EXISTS finished_at TIMESTAMP NULL"))
        conn.execute(text("ALTER TABLE hospitalization_fluids ADD COLUMN IF NOT EXISTS finished_by VARCHAR(100) DEFAULT ''"))
        conn.execute(text("ALTER TABLE hospitalization_medications ADD COLUMN IF NOT EXISTS applied_by VARCHAR(100) DEFAULT ''"))
        conn.execute(text("""
        CREATE TABLE IF NOT EXISTS hospitalization_fluids (
            id SERIAL PRIMARY KEY,
            hospitalization_id INTEGER REFERENCES hospitalizations(id),
            fluid_type VARCHAR(200) DEFAULT '',
            fluid_rate VARCHAR(80) DEFAULT '',
            ml_kg_h VARCHAR(80) DEFAULT '',
            drip_set VARCHAR(80) DEFAULT '',
            notes TEXT DEFAULT '',
            status VARCHAR(40) DEFAULT 'Activo',
            created_by VARCHAR(100) DEFAULT 'admin',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """))

    try:
        admin = db.query(User).filter(User.username == 'admin').first()
        if not admin:
            admin = User(
                username='admin',
                full_name='Carolina',
                password_hash=pwd_context.hash('losaromos')
            )
            db.add(admin)
            db.commit()
    finally:
        db.close()

@app.on_event("startup")
def on_startup():
    pass

def current_user(request: Request, db: Session = Depends(get_db)):
    username = request.session.get('user')
    if not username:
        return None
    return db.query(User).filter(User.username == username, User.is_active == True).first()

def require_user(request: Request, db: Session = Depends(get_db)):
    user = current_user(request, db)
    if not user:
        raise HTTPException(status_code=303, headers={'Location': '/login'})
    return user

@app.get('/login', response_class=HTMLResponse)
def login_form(request: Request):
    return templates.TemplateResponse('login.html', {'request': request, 'error': None})

@app.post('/login')
def login(request: Request, username: str = Form(...), password: str = Form(...), db: Session = Depends(get_db)):
    user = db.query(User).filter(User.username == username).first()
    if not user or not pwd_context.verify(password, user.password_hash):
        return templates.TemplateResponse('login.html', {'request': request, 'error': 'Usuario o clave incorrectos'})
    request.session['user'] = user.username
    return RedirectResponse('/', status_code=303)

@app.get('/logout')
def logout(request: Request):
    request.session.clear()
    return RedirectResponse('/login', status_code=303)

@app.get('/', response_class=HTMLResponse)
def home(request: Request, db: Session = Depends(get_db), user: User = Depends(require_user)):
    today = argentina_now().date()

    stats = {
        'owners': db.query(Owner).count(),
        'patients': db.query(Patient).count(),
        'events': db.query(ClinicalEvent).count(),
        'pending': db.query(ClinicalEvent).filter(
            ClinicalEvent.reminder_date != None
        ).count(),
        'overdue': db.query(ClinicalEvent).filter(
            ClinicalEvent.reminder_date != None,
            ClinicalEvent.reminder_date < today
        ).count(),
        'today': db.query(ClinicalEvent).filter(
            ClinicalEvent.reminder_date == today
        ).count(),
        'vaccines': db.query(ClinicalEvent).filter(
            ClinicalEvent.event_type == 'Vacuna'
        ).count(),

        'rx': db.query(ClinicalEvent).filter(
            ClinicalEvent.event_type == 'Radiografía'
        ).count(),

        'ecg': db.query(ClinicalEvent).filter(
            ClinicalEvent.event_type == 'ECG'
        ).count(),
    }

    latest_events = (
        db.query(ClinicalEvent)
        .order_by(ClinicalEvent.event_date.desc())
        .limit(8)
        .all()
    )

    upcoming = (
        db.query(ClinicalEvent)
        .filter(ClinicalEvent.reminder_date != None)
        .order_by(ClinicalEvent.reminder_date.asc())
        .limit(8)
        .all()
    )

    overdue_events = (
        db.query(ClinicalEvent)
        .filter(
            ClinicalEvent.reminder_date != None,
            ClinicalEvent.reminder_date < today
        )
        .order_by(ClinicalEvent.reminder_date.asc())
        .limit(5)
        .all()
    )

    today_events = (
        db.query(ClinicalEvent)
        .filter(ClinicalEvent.reminder_date == today)
        .order_by(ClinicalEvent.reminder_date.asc())
        .limit(5)
        .all()
    )
    upcoming_appointments = (
        db.query(Appointment)
        .filter(Appointment.appointment_date >= datetime.combine(today, datetime.min.time()))
        .filter(Appointment.status.in_(['Pendiente', 'Confirmado']))
        .order_by(Appointment.appointment_date.asc(), Appointment.start_time.asc())
        .limit(8)
        .all()
    )
    recent_patients = (
        db.query(Patient)
        .order_by(Patient.id.desc())
        .limit(8)
        .all()
    )


    try:
        inactive_cutoff = today.replace(year=today.year - 1)
    except ValueError:
        inactive_cutoff = today.replace(year=today.year - 1, month=2, day=28)

    inactive_patients = []

    for p in db.query(Patient).all():
        last_event = (
            db.query(ClinicalEvent)
            .filter(ClinicalEvent.patient_id == p.id)
            .order_by(ClinicalEvent.event_date.desc())
            .first()
        )

        if last_event is None or last_event.event_date.date() < inactive_cutoff:
            inactive_patients.append({
                'patient': p,
                'last_event': last_event
            })

    inactive_patients = inactive_patients[:8]
    return templates.TemplateResponse(
        'home.html',
        {
            'request': request,
            'user': user,
            'stats': stats,
            'latest_events': latest_events,
            'upcoming': upcoming,
            'overdue_events': overdue_events,
            'today_events': today_events,
            'upcoming_appointments': upcoming_appointments,
            'recent_patients': recent_patients,
            'inactive_patients': inactive_patients,
            'today': today
            
        }
    )
@app.get('/quick/rx')
def quick_rx():
    return RedirectResponse('/search?quick=rx', status_code=303)
@app.get('/quick/ecg')
def quick_ecg():
    return RedirectResponse('/search?quick=ecg', status_code=303)


@app.get('/quick/eco')
def quick_eco():
    return RedirectResponse('/search?quick=eco', status_code=303)    
@app.get('/owners', response_class=HTMLResponse)
def owners(request: Request, q: str = '', db: Session = Depends(get_db), user: User = Depends(require_user)):
    query = db.query(Owner)
    if q:
        like = f'%{q}%'
        query = query.filter(or_(Owner.name.ilike(like), Owner.phone.ilike(like), Owner.whatsapp.ilike(like)))
    return templates.TemplateResponse('owners.html', {'request': request, 'owners': query.order_by(Owner.name).limit(200).all(), 'q': q})

@app.post('/owners')
def owner_create(name: str = Form(...), phone: str = Form(''), whatsapp: str = Form(''), email: str = Form(''), address: str = Form(''), notes: str = Form(''), db: Session = Depends(get_db), user: User = Depends(require_user)):
    owner = Owner(name=name, phone=phone, whatsapp=whatsapp, email=email, address=address, notes=notes)
    db.add(owner); db.commit()
    return RedirectResponse('/owners', status_code=303)
@app.get('/owners/{owner_id}/edit', response_class=HTMLResponse)
def owner_edit_form(request: Request, owner_id: int, db: Session = Depends(get_db), user: User = Depends(require_user)):
    owner = db.get(Owner, owner_id)
    if not owner:
        raise HTTPException(status_code=404)

    return templates.TemplateResponse(
        'owner_edit.html',
        {'request': request, 'owner': owner}
    )


@app.post('/owners/{owner_id}/edit')
def owner_edit_save(
    owner_id: int,
    name: str = Form(...),
    phone: str = Form(''),
    whatsapp: str = Form(''),
    email: str = Form(''),
    address: str = Form(''),
    notes: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    owner = db.get(Owner, owner_id)
    if not owner:
        raise HTTPException(status_code=404)

    owner.name = name
    owner.phone = phone
    owner.whatsapp = whatsapp
    owner.email = email
    owner.address = address
    owner.notes = notes

    db.commit()
    return RedirectResponse('/owners', status_code=303)
@app.get('/patients/new', response_class=HTMLResponse)
def patient_new(request: Request, owner_id: Optional[int] = None, db: Session = Depends(get_db), user: User = Depends(require_user)):
    owners = db.query(Owner).order_by(Owner.name).limit(500).all()
    return templates.TemplateResponse('patient_new.html', {'request': request, 'owners': owners, 'owner_id': owner_id})

@app.post('/patients')
def patient_create(
    name: str = Form(...),
    owner_id: str = Form(""),
    owner_name: str = Form(""),
    owner_phone: str = Form(""),
    owner_whatsapp: str = Form(""),
    owner_address: str = Form(""),
    owner_email: str = Form(""),
    species: str = Form(""),
    breed: str = Form(""),
    sex: str = Form(""),
    birth_date: str = Form(""),
    weight: str = Form(""),
    alerts: str = Form(""),
    notes: str = Form(""),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    owner_id_value = int(owner_id) if owner_id and owner_id.strip() else None

    if not owner_id_value:
        if owner_name.strip():
            owner = Owner(
                name=owner_name.strip(),
                phone=owner_phone.strip(),
                whatsapp=owner_whatsapp.strip(),
                address=owner_address.strip(),
                email=owner_email.strip()
            )
            db.add(owner)
            db.commit()
            db.refresh(owner)
            owner_id_value = owner.id
        else:
            owner_id = None

    w = float(weight.replace(',', '.')) if weight.strip() else None

    bd = None
    if birth_date and birth_date.strip():
        try:
            bd = datetime.strptime(birth_date.strip(), '%Y-%m-%d').date()
        except ValueError:
            bd = None

    p = Patient(
        name=name,
        owner_id=owner_id_value,
        species=species,
        breed=breed,
        sex=sex,
        birth_date=bd,
        weight=w,
        alerts=alerts,
        notes=notes
    )

    db.add(p)
    db.commit()

    return RedirectResponse(f'/patients/{p.id}', status_code=303)
@app.get('/agenda', response_class=HTMLResponse)
def agenda(
    request: Request,
    date: str = '',
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    if date:
        selected_day = datetime.strptime(date, '%Y-%m-%d').date()
    else:
        selected_day = argentina_now().date()

    start_dt = datetime.combine(selected_day, datetime.min.time())
    end_dt = start_dt + timedelta(days=1)
    import calendar

    cal = calendar.Calendar(firstweekday=0)
    
    month_days = cal.monthdatescalendar(
        selected_day.year,
        selected_day.month
    )
    
    prev_month = (selected_day.replace(day=1) - timedelta(days=1)).replace(day=1)
    
    if selected_day.month == 12:
        next_month = selected_day.replace(year=selected_day.year + 1, month=1, day=1)
    else:
        next_month = selected_day.replace(month=selected_day.month + 1, day=1)
    appointments = (
        db.query(Appointment)
        .filter(Appointment.appointment_date >= start_dt)
        .filter(Appointment.appointment_date < end_dt)
        .order_by(Appointment.start_time)
        .all()
    )
    month_start = selected_day.replace(day=1)
    month_end = next_month

    month_appointments = (
        db.query(Appointment)
        .filter(Appointment.appointment_date >= month_start)
        .filter(Appointment.appointment_date < month_end)
        .all()
    )

    appointments_count_by_day = {}
    appointments_tooltip_by_day = {}
    for appointment in month_appointments:
        day_key = appointment.appointment_date.day

        appointments_count_by_day[day_key] = (
            appointments_count_by_day.get(day_key, 0) + 1
        )

        patient_name = (
            appointment.patient.name
            if appointment.patient
            else "Sin paciente"
        )

        title = appointment.title or ""

        text = f"• {patient_name}"
        if title:
            text += f" - {title}"

        appointments_tooltip_by_day.setdefault(day_key, []).append(text)
    owners = db.query(Owner).order_by(Owner.name).limit(300).all()
    patients = db.query(Patient).order_by(Patient.name).limit(300).all()

    return templates.TemplateResponse(
        'agenda.html',
        {
            'request': request,
            'selected_day': selected_day,
            'appointments': appointments,
            'owners': owners,
            'patients': patients,
            'month_days': month_days,
            'prev_month': prev_month,
            'next_month': next_month,
            'appointments_count_by_day': appointments_count_by_day,
            'appointments_tooltip_by_day': appointments_tooltip_by_day,
        }
    )


@app.post('/agenda')
def agenda_create(
    service: str = Form(...),
    title: str = Form(''),
    appointment_date: str = Form(...),
    start_time: str = Form(''),
    end_time: str = Form(''),
    owner_id: str = Form(''),
    patient_id: str = Form(''),
    notes: str = Form(''),
    reminder_24h: str = Form(''),
    reminder_12h: str = Form(''),
    contact_whatsapp: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    selected_day = datetime.strptime(appointment_date, '%Y-%m-%d').date()
    appointment_dt = datetime.combine(selected_day, datetime.min.time())

    owner_value = int(owner_id) if owner_id else None
    patient_value = int(patient_id) if patient_id else None

    appointment = Appointment(
        service=service,
        title=title,
        appointment_date=appointment_dt,
        start_time=start_time,
        end_time=end_time,
        owner_id=owner_value,
        patient_id=patient_value,
        reminder_12h=True if reminder_12h else False,
        contact_whatsapp=contact_whatsapp,
        notes=notes,
        reminder_24h=True if reminder_24h else False,
        status='Pendiente'
    )

    db.add(appointment)
    db.commit()

    return RedirectResponse(f'/agenda?date={appointment_date}', status_code=303)
@app.post('/agenda/{appointment_id}/status')
def agenda_update_status(
    appointment_id: int,
    status: str = Form(...),
    date: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    appointment = db.get(Appointment, appointment_id)

    if not appointment:
        raise HTTPException(status_code=404, detail='Turno no encontrado')

    appointment.status = status
    db.commit()

    if date:
        return RedirectResponse(f'/agenda?date={date}', status_code=303)

    return RedirectResponse('/agenda', status_code=303)
@app.post('/agenda/{appointment_id}/arrived')
def agenda_patient_arrived(
    appointment_id: int,
    date: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    appointment = db.get(Appointment, appointment_id)

    if not appointment:
        raise HTTPException(status_code=404, detail='Turno no encontrado')

    existing = (
        db.query(WaitingListEntry)
        .filter(WaitingListEntry.appointment_id == appointment.id)
        .filter(WaitingListEntry.status.in_(['Esperando', 'En consulta']))
        .first()
    )

    if not existing:
        entry = WaitingListEntry(
            appointment_id=appointment.id,
            owner_id=appointment.owner_id,
            patient_id=appointment.patient_id,
            reason=appointment.service or appointment.title or 'Turno agendado',
            notes=appointment.notes or '',
            priority='Normal',
            status='Esperando',
            created_by=user.username
        )

        db.add(entry)

    appointment.status = 'Confirmado'
    db.commit()

    return RedirectResponse('/waitlist', status_code=303)


@app.get('/waitlist', response_class=HTMLResponse)
def waitlist_page(
    request: Request,
    status: str = '',
    q: str = '',
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    today = argentina_now().date()
    day_start = datetime.combine(today, datetime.min.time())
    day_end = datetime.combine(today, datetime.max.time())

    query = db.query(WaitingListEntry).filter(
        WaitingListEntry.arrival_time >= day_start,
        WaitingListEntry.arrival_time <= day_end
    )

    if status:
        query = query.filter(WaitingListEntry.status == status)

    entries = query.order_by(
        WaitingListEntry.finished_at.asc().nullsfirst(),
        WaitingListEntry.arrival_time.asc()
    ).all()

    if q:
        q_lower = q.lower()
        entries = [
            e for e in entries
            if (
                (e.patient and q_lower in e.patient.name.lower()) or
                (e.owner and q_lower in e.owner.name.lower()) or
                q_lower in (e.reason or '').lower() or
                q_lower in (e.notes or '').lower()
            )
        ]

    owners = db.query(Owner).order_by(Owner.name).limit(500).all()
    patients = db.query(Patient).order_by(Patient.name).limit(500).all()

    stats = {
        'total': len(entries),
        'waiting': len([e for e in entries if e.status == 'Esperando']),
        'consulting': len([e for e in entries if e.status == 'En consulta']),
        'finished': len([e for e in entries if e.status == 'Finalizado']),
    }

    return templates.TemplateResponse(
        'waitlist.html',
        {
            'request': request,
            'entries': entries,
            'owners': owners,
            'patients': patients,
            'stats': stats,
            'selected_status': status,
            'q': q,
            'today': today,
            'now': argentina_now().replace(tzinfo=None)
        }
    )


@app.post('/waitlist')
def waitlist_create(
    owner_id: str = Form(''),
    patient_id: str = Form(''),
    reason: str = Form(''),
    priority: str = Form('Normal'),
    notes: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    entry = WaitingListEntry(
        owner_id=int(owner_id) if owner_id else None,
        patient_id=int(patient_id) if patient_id else None,
        reason=reason or 'Consulta',
        priority=priority or 'Normal',
        notes=notes or '',
        status='Esperando',
        created_by=user.username
    )

    db.add(entry)
    db.commit()

    return RedirectResponse('/waitlist', status_code=303)


@app.post('/waitlist/{entry_id}/enter')
def waitlist_enter_hc(
    entry_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    entry = db.get(WaitingListEntry, entry_id)

    if not entry:
        raise HTTPException(status_code=404, detail='Entrada no encontrada')

    entry.status = 'En consulta'
    entry.started_at = argentina_now()
    db.commit()

    if entry.patient_id:
        return RedirectResponse(f'/patients/{entry.patient_id}', status_code=303)

    return RedirectResponse('/waitlist', status_code=303)


@app.post('/waitlist/{entry_id}/status')
def waitlist_update_status(
    entry_id: int,
    status: str = Form(...),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    entry = db.get(WaitingListEntry, entry_id)

    if not entry:
        raise HTTPException(status_code=404, detail='Entrada no encontrada')

    entry.status = status

    if status == 'En consulta':
        entry.started_at = argentina_now()

    if status == 'Finalizado':
        entry.finished_at = argentina_now()

    db.commit()

    return RedirectResponse('/waitlist', status_code=303)
@app.get('/agenda/{appointment_id}/edit', response_class=HTMLResponse)
def agenda_edit(
    request: Request,
    appointment_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    appointment = db.get(Appointment, appointment_id)

    if not appointment:
        raise HTTPException(status_code=404, detail='Turno no encontrado')

    owners = db.query(Owner).order_by(Owner.name).limit(300).all()
    patients = db.query(Patient).order_by(Patient.name).limit(300).all()

    return templates.TemplateResponse(
        'agenda_edit.html',
        {
            'request': request,
            'appointment': appointment,
            'owners': owners,
            'patients': patients
        }
    )
@app.post('/agenda/{appointment_id}/edit')
def agenda_update(
    appointment_id: int,
    service: str = Form(...),
    title: str = Form(''),
    appointment_date: str = Form(...),
    start_time: str = Form(''),
    end_time: str = Form(''),
    owner_id: str = Form(''),
    patient_id: str = Form(''),
    notes: str = Form(''),
    reminder_24h: str = Form(''),
    reminder_12h: str = Form(''),
    contact_whatsapp: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    appointment = db.get(Appointment, appointment_id)

    if not appointment:
        raise HTTPException(status_code=404, detail='Turno no encontrado')

    selected_day = datetime.strptime(appointment_date, '%Y-%m-%d').date()
    appointment_dt = datetime.combine(selected_day, datetime.min.time())

    appointment.service = service
    appointment.title = title
    appointment.appointment_date = appointment_dt
    appointment.start_time = start_time
    appointment.end_time = end_time
    appointment.owner_id = int(owner_id) if owner_id else None
    appointment.patient_id = int(patient_id) if patient_id else None
    appointment.reminder_12h = True if reminder_12h else False
    appointment.contact_whatsapp = contact_whatsapp
    appointment.notes = notes
    appointment.reminder_24h = True if reminder_24h else False

    db.commit()

    return RedirectResponse(f'/agenda?date={appointment_date}', status_code=303)
@app.get('/search', response_class=HTMLResponse)
def search(
    request: Request,
    q: str = '',
    quick: str = '',
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    results = []
    if quick:
        request.session['quick_event_type'] = quick
    if q:
        like = f'%{q}%'
        results = db.query(Patient).join(Owner).filter(or_(Patient.name.ilike(like), Owner.name.ilike(like), Owner.phone.ilike(like), Owner.whatsapp.ilike(like))).order_by(Patient.name).limit(100).all()
    return templates.TemplateResponse(
    'search.html',
    {
        'request': request,
        'q': q,
        'results': results,
        'today': argentina_now().date()
    }
)
@app.post('/patients/{patient_id}/weight')
def patient_update_weight(
    patient_id: int,
    weight: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)
    if not patient:
        raise HTTPException(status_code=404)

    patient.weight = float(weight.replace(',', '.')) if weight.strip() else None
    db.commit()

    return RedirectResponse(f'/patients/{patient.id}', status_code=303)
@app.get('/patients/{patient_id}/cart', response_class=HTMLResponse)
def patient_cart(
    request: Request,
    patient_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)

    if not patient:
        raise HTTPException(
            status_code=404,
            detail='Paciente no encontrado'
        )

    products = (
        db.query(Product)
        .filter(Product.active == True)
        .order_by(Product.name)
        .limit(300)
        .all()
    )
    quick_names = [
        'Consulta clínica',
        'INY CONS',
        'INY FARM GRANDE',
        'INY CONS URGENCIA',
        'INY FARM URGENCIA',
        'DESP COMPLETO CONS',
        'DESP COMPLETO FARM',
        'INT CONS',
        'INT FARM',
        'ECG',
        'RX x1',
        'RX x2',
        'RX x3'
    ]
    
    quick_services = []
    for name in quick_names:
        product = (
            db.query(Product)
            .filter(Product.active == True, Product.name == name)
            .first()
        )
    
        if product:
            quick_services.append(product)
    return templates.TemplateResponse(
        'patient_cart.html',
        {
            'request': request,
            'patient': patient,
            'products': products,
            'quick_services': quick_services
        }
    )
@app.post('/patients/{patient_id}/cart/send')
def patient_cart_send(
    patient_id: int,
    cart_json: str = Form('[]'),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    import json

    patient = db.get(Patient, patient_id)

    if not patient:
        raise HTTPException(status_code=404)

    owner_id = patient.owner_id if patient.owner_id else None

    sale = Sale(
        status='pending',
        total=0,
        payment_method='Pendiente',
        patient_id=patient.id,
        owner_id=owner_id,
        notes='Generado desde carrito clínico'
    )

    db.add(sale)
    db.flush()

    total = 0
    cost_total = 0

    try:
        cart_items = json.loads(cart_json or '[]')
    except Exception:
        cart_items = []

    for cart_item in cart_items:
        product_id = cart_item.get('productId')
        name = cart_item.get('name', '')
        qty = float(cart_item.get('quantity') or 1)
        price = float(cart_item.get('price') or 0)

        if qty <= 0:
            continue

        subtotal = qty * price
        total += subtotal

        if product_id:
            product = db.get(Product, int(product_id))

            if not product:
                continue

            product_cost = product.cost_price or 0
            cost_total += qty * product_cost

            item = SaleItem(
                sale_id=sale.id,
                product_id=product.id,
                quantity=qty,
                unit_price=price,
                subtotal=subtotal
            )

            db.add(item)

            if product.stock is not None:
                product.stock = (product.stock or 0) - qty

        else:
            # Servicio rápido sin producto asociado:
            # por ahora se guarda como nota dentro de la venta
            sale.notes = (sale.notes or '') + f'\n- {name} x {qty} = $ {subtotal:.2f}'

    sale.total = total
    sale.cost_total = cost_total
    sale.profit_amount = total - cost_total
    sale.margin_percent = ((total - cost_total) / total * 100) if total > 0 else 0

    db.commit()
    db.refresh(sale)

    return RedirectResponse(
        url=f'/sales/{sale.id}',
        status_code=303
    )
DUE_ACTIVE_MARKER = '[DEUDA_ACTIVA]'
DUE_CLOSED_MARKER = '[DEUDA_CERRADA]'


def is_managed_due_event(event):
    return event and event.description and DUE_ACTIVE_MARKER in event.description


def event_matches_managed_due(real_event, due_event):
    real_type = (real_event.event_type or '').lower()
    real_text = f'{real_event.title or ""} {real_event.description or ""} {real_event.vaccine_name or ""} {real_event.dewormer_product or ""}'.lower()

    due_type = (due_event.event_type or '').lower()
    due_text = f'{due_event.title or ""} {due_event.description or ""}'.lower()

    if due_type == real_type:
        return True

    if ('vacun' in due_text or 'vacuna' in due_type) and real_type == 'vacuna':
        return True

    if (
        'despar' in due_text
        or 'antiparas' in due_text
        or 'aprax' in due_text
        or 'pipeta' in due_text
        or due_type == 'desparasitación'
    ) and real_type == 'desparasitación':
        return True

    if 'control' in due_text and real_type in ['control', 'consulta clínica']:
        return True

    if 'consulta' in due_text and real_type == 'consulta clínica':
        return True

    return False


def close_managed_due_events(db, patient, real_event, user):
    active_due_events = (
        db.query(ClinicalEvent)
        .filter(ClinicalEvent.patient_id == patient.id)
        .filter(ClinicalEvent.description.ilike(f'%{DUE_ACTIVE_MARKER}%'))
        .all()
    )

    closed_count = 0

    for due_event in active_due_events:
        if event_matches_managed_due(real_event, due_event):
            due_event.description = (due_event.description or '').replace(
                DUE_ACTIVE_MARKER,
                DUE_CLOSED_MARKER
            )

            due_event.description += (
                '\n\n✅ Pendiente cumplido automáticamente al cargar el evento real '
                f'el {argentina_now().strftime("%d/%m/%Y %H:%M")} por {user.username}.'
            )

            due_event.reminder_date = None
            closed_count += 1

    return closed_count
@app.get('/patients/{patient_id}', response_class=HTMLResponse)
def patient_detail(request: Request, patient_id: int, db: Session = Depends(get_db), user: User = Depends(require_user)):
    patient = db.get(Patient, patient_id)

    if not patient:
        raise HTTPException(404)

    events = (
        db.query(ClinicalEvent)
        .filter(ClinicalEvent.patient_id == patient.id)
        .filter(~ClinicalEvent.description.ilike(f'%{DUE_ACTIVE_MARKER}%'))
        .filter(~ClinicalEvent.description.ilike(f'%{DUE_CLOSED_MARKER}%'))
        .order_by(ClinicalEvent.event_date.desc())
        .limit(20)
        .all()
    )

    today = argentina_now().date()
    pending_actions = (
        db.query(ClinicalEvent)
        .filter(ClinicalEvent.patient_id == patient.id)
        .filter(ClinicalEvent.description.ilike(f'%{DUE_ACTIVE_MARKER}%'))
        .order_by(ClinicalEvent.event_date.desc())
        .all()
    )
    upcoming_events = (
        db.query(ClinicalEvent)
        .filter(
            ClinicalEvent.patient_id == patient.id,
            ClinicalEvent.reminder_date != None,
            ClinicalEvent.reminder_date >= today
        )
        .order_by(ClinicalEvent.reminder_date.asc())
        .limit(5)
        .all()
    )
    
    clinical_ai = {}    
    
    last_anesthesia = (
        db.query(ClinicalEvent)
        .filter(
            ClinicalEvent.patient_id == patient.id,
            ClinicalEvent.event_type == 'Anestesia'
        )
        .order_by(ClinicalEvent.event_date.desc())
        .first()
    )

    anesthesia_history = (
        db.query(ClinicalEvent)
        .filter(
            ClinicalEvent.patient_id == patient.id,
            ClinicalEvent.event_type == 'Anestesia'
        )
        .order_by(ClinicalEvent.event_date.desc())
        .all()
    )
    cardiology_ecgs = (
    db.query(ClinicalEvent)
    .filter(
        ClinicalEvent.patient_id == patient.id,
        ClinicalEvent.event_type == "ECG"
    )
    .order_by(ClinicalEvent.event_date.desc())
    .all()
)

    last_ecg = cardiology_ecgs[0] if cardiology_ecgs else None
    ecg_count = len(cardiology_ecgs)

    cardiology_ecos = (
        db.query(ClinicalEvent)
        .filter(
            ClinicalEvent.patient_id == patient.id,
            ClinicalEvent.event_type == "Ecocardiografía"
        )
        .order_by(ClinicalEvent.event_date.desc())
        .all()
    )

    last_eco = cardiology_ecos[0] if cardiology_ecos else None
    eco_count = len(cardiology_ecos)

    next_visit = upcoming_events[0] if upcoming_events else None
    vaccine_names = sorted(
        list(
            {
                e.vaccine_name.strip()
                for e in db.query(ClinicalEvent).all()
                if e.vaccine_name and e.vaccine_name.strip()
            }
        )
    )

    dewormer_names = sorted(
        list(
            {
                e.dewormer_product.strip()
                for e in db.query(ClinicalEvent).all()
                if e.dewormer_product and e.dewormer_product.strip()
            }
        )
    )
    
    active_hospitalization = (
        db.query(Hospitalization)
        .filter(
            Hospitalization.patient_id == patient.id,
            Hospitalization.status == 'Internado'
        )
        .order_by(Hospitalization.admission_date.desc())
        .first()
    )
    return templates.TemplateResponse(
        'patient_detail.html',
        {
            'request': request,
            'patient': patient,
            'today': today,
            'events': events,
            'clinical_ai': clinical_ai,
            'event_types': EVENT_TYPES,
            'upcoming_events': upcoming_events,
            'pending_actions': pending_actions,
            'timedelta': timedelta,
            'last_anesthesia': last_anesthesia,
            'anesthesia_history': anesthesia_history,
            'next_visit': next_visit,
            'last_ecg': last_ecg,
            'ecg_count': ecg_count,
            'last_eco': last_eco,
            'eco_count': eco_count,
            "vaccine_names": vaccine_names,
            "dewormer_names": dewormer_names,
            "active_hospitalization": active_hospitalization
        }
    )
@app.get('/patients/{patient_id}/events/{event_id}/edit', response_class=HTMLResponse)
def edit_clinical_event_form(
    request: Request,
    patient_id: int,
    event_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)
    event = db.get(ClinicalEvent, event_id)

    if not patient or not event or event.patient_id != patient.id:
        raise HTTPException(status_code=404)

    return templates.TemplateResponse(
        'patient_event_edit.html',
        {
            'request': request,
            'patient': patient,
            'event': event,
            'event_types': EVENT_TYPES,
        }
    )


@app.post('/patients/{patient_id}/events/{event_id}/edit')
async def edit_clinical_event_save(
    patient_id: int,
    event_id: int,
    request: Request,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)
    event = db.get(ClinicalEvent, event_id)

    if not patient or not event or event.patient_id != patient.id:
        raise HTTPException(status_code=404)

    form = await request.form()

    fields = [
        'event_type', 'title', 'description', 'anamnesis', 'physical_exam',
        'diagnosis', 'treatment',
        'weight', 'temperature', 'heart_rate', 'respiratory_rate',
        'mucous_membranes', 'crt', 'hydration',

        'ecg_hr', 'ecg_rhythm', 'ecg_p', 'ecg_pr', 'ecg_qrs',
        'ecg_st', 'ecg_t', 'ecg_qt', 'ecg_axis', 'ecg_interpretation',

        'ecg_p_mv', 'ecg_qrs_mv', 'ecg_t_mv', 'ecg_qtc',
        'ecg_polarity', 'ecg_arrhythmia', 'ecg_conduction', 'ecg_notes',

        'eco_aiao', 'eco_fs', 'eco_acvim', 'eco_diagnosis', 'eco_treatment',

        'eco_epss', 'eco_lvidd', 'eco_lvids', 'eco_ivsd', 'eco_ivss',
        'eco_lvpwd', 'eco_lvpws', 'eco_fe',
        'eco_la_size', 'eco_lv_size', 'eco_rv_size', 'eco_ra_size',
        'eco_mitral', 'eco_tricuspid', 'eco_aortic', 'eco_pulmonary',
        'eco_pulmonary_htn', 'eco_pericardium',
        'eco_doppler', 'eco_observations',

        'rx_vhs', 'rx_vlas', 'rx_heart_size', 'rx_left_atrium',
        'rx_left_heart', 'rx_right_heart',
        'rx_pulmonary_vessels', 'rx_lung_pattern',
        'rx_edema', 'rx_congestion', 'rx_trachea',
        'rx_observations',

        'vaccine_name', 'vaccine_lot', 'vaccine_expiration', 'next_vaccine_date',
        'dewormer_product', 'dewormer_drug', 'dewormer_dose', 'next_deworming_date',
        'reminder_date'
    ]
    for field in fields:
        if hasattr(event, field):
            value = form.get(field)

            if value == '':
                value = None

            if field in ['diagnosis', 'treatment', 'description', 'anamnesis', 'physical_exam']:
                value = value or ''

            if field in ['next_vaccine_date', 'next_deworming_date', 'reminder_date'] and value:
                value = datetime.strptime(value, '%Y-%m-%d').date()

            setattr(event, field, value)
    attachments = form.getlist("attachments")
    upload_dir = "/opt/render/project/src/app/uploads"
    os.makedirs(upload_dir, exist_ok=True)

    for file in attachments:
        if file and file.filename:
            safe_filename = os.path.basename(file.filename)
            full_path = os.path.join(upload_dir, safe_filename)

            with open(full_path, "wb") as buffer:
                buffer.write(await file.read())

            attachment = EventAttachment(
                event_id=event.id,
                filename=safe_filename,
                file_path="/uploads/" + safe_filename
            )

            db.add(attachment)

  
    db.commit()

    return RedirectResponse(
        url=f'/patients/{patient_id}',
        status_code=303
    )
@app.post('/patients/{patient_id}/events/{event_id}/attachments/{attachment_id}/delete')
def delete_event_attachment_from_edit(
    patient_id: int,
    event_id: int,
    attachment_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    attachment = db.get(EventAttachment, attachment_id)

    if not attachment:
        raise HTTPException(status_code=404)

    db.delete(attachment)
    db.commit()

    return RedirectResponse(
        url=f'/patients/{patient_id}/events/{event_id}/edit',
        status_code=303
    )
@app.get('/patients/{patient_id}/cardiology/ecg', response_class=HTMLResponse)
def patient_ecg_list(
    request: Request,
    patient_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)

    if not patient:
        raise HTTPException(status_code=404)

    ecg_events = (
        db.query(ClinicalEvent)
        .filter(
            ClinicalEvent.patient_id == patient.id,
            ClinicalEvent.event_type == "ECG"
        )
        .order_by(ClinicalEvent.event_date.desc())
        .all()
    )

    return templates.TemplateResponse(
        'patient_ecg_list.html',
        {
            'request': request,
            'patient': patient,
            'ecg_events': ecg_events
        }
    )
def build_cardio_ai(last_ecg, previous_ecg, last_eco, previous_eco, last_rx, previous_rx):
    def to_float(value):
        try:
            if value is None:
                return None
            value = str(value).replace(",", ".").strip()
            if not value:
                return None
            return float(value)
        except Exception:
            return None

    score = 100
    alerts = []
    sections = []
    recommendations = []

    hr = to_float(getattr(last_ecg, "ecg_hr", None)) if last_ecg else None
    axis = to_float(getattr(last_ecg, "ecg_axis", None)) if last_ecg else None
    rhythm = ((getattr(last_ecg, "ecg_rhythm", "") or "") + " " + (getattr(last_ecg, "ecg_arrhythmia", "") or "")).lower() if last_ecg else ""

    aiao = to_float(getattr(last_eco, "eco_aiao", None)) if last_eco else None
    fs = to_float(getattr(last_eco, "eco_fs", None)) if last_eco else None
    epss = to_float(getattr(last_eco, "eco_epss", None)) if last_eco else None
    fe = to_float(getattr(last_eco, "eco_fe", None)) if last_eco else None
    acvim = (getattr(last_eco, "eco_acvim", "") or "").upper() if last_eco else ""

    vhs = to_float(getattr(last_rx, "rx_vhs", None)) if last_rx else None
    vlas = to_float(getattr(last_rx, "rx_vlas", None)) if last_rx else None
    rx_text = " ".join([
        getattr(last_rx, "rx_edema", "") or "",
        getattr(last_rx, "rx_congestion", "") or "",
        getattr(last_rx, "rx_pulmonary_vessels", "") or "",
        getattr(last_rx, "rx_lung_pattern", "") or "",
        getattr(last_rx, "rx_observations", "") or "",
        getattr(last_rx, "description", "") or "",
    ]).lower() if last_rx else ""

    if last_ecg:
        ecg_notes = []

        if hr is not None:
            if hr > 180:
                score -= 10
                ecg_notes.append(f"FC elevada ({hr:g} lpm).")
                alerts.append("FC elevada en el último ECG.")
            elif hr < 60:
                score -= 10
                ecg_notes.append(f"FC baja ({hr:g} lpm).")
                alerts.append("FC baja en el último ECG.")
            else:
                ecg_notes.append(f"FC dentro de rango clínico orientativo ({hr:g} lpm).")

        if "fibril" in rhythm or "bloqueo" in rhythm or "taqui" in rhythm or "extra" in rhythm:
            score -= 12
            ecg_notes.append("Se describe arritmia o alteración de conducción.")
            alerts.append("Revisar arritmia/conducción descrita en ECG.")
        elif rhythm.strip():
            ecg_notes.append("Ritmo sin alteraciones mayores descritas.")

        if axis is not None:
            if axis < 40 or axis > 100:
                score -= 6
                ecg_notes.append(f"Eje eléctrico fuera del rango esperado ({axis:g}°).")
            else:
                ecg_notes.append(f"Eje eléctrico conservado ({axis:g}°).")

        sections.append({"title": "ECG", "items": ecg_notes or ["ECG cargado sin datos suficientes para interpretación automática."]})
    else:
        score -= 8
        sections.append({"title": "ECG", "items": ["Sin ECG cargado."]})

    if last_eco:
        eco_notes = []

        if aiao is not None:
            if aiao >= 1.9:
                score -= 18
                eco_notes.append(f"AI/Ao aumentado de forma importante ({aiao:g}).")
                alerts.append("AI/Ao elevado: sugiere dilatación auricular izquierda.")
            elif aiao >= 1.6:
                score -= 10
                eco_notes.append(f"AI/Ao aumentado ({aiao:g}).")
            else:
                eco_notes.append(f"AI/Ao dentro de rango orientativo ({aiao:g}).")

        if fs is not None:
            if fs < 20:
                score -= 18
                eco_notes.append(f"FS baja ({fs:g}%). Posible disfunción sistólica.")
                alerts.append("FS baja: revisar función sistólica.")
            elif fs > 45:
                eco_notes.append(f"FS aumentada/hiperdinámica ({fs:g}%).")
            else:
                eco_notes.append(f"FS conservada ({fs:g}%).")

        if epss is not None:
            if epss > 7:
                score -= 8
                eco_notes.append(f"EPSS aumentado ({epss:g} mm).")
            else:
                eco_notes.append(f"EPSS sin aumento relevante ({epss:g} mm).")

        if fe is not None:
            if fe < 45:
                score -= 12
                eco_notes.append(f"FE disminuida ({fe:g}%).")
            else:
                eco_notes.append(f"FE conservada ({fe:g}%).")

        if acvim:
            if acvim == "B2":
                score -= 12
                eco_notes.append("Clasificación ACVIM B2: requiere seguimiento cardiológico.")
            elif acvim in ["C", "D"]:
                score -= 25
                eco_notes.append(f"Clasificación ACVIM {acvim}: paciente de control estricto.")
                alerts.append(f"ACVIM {acvim}: riesgo cardiológico elevado.")

        sections.append({"title": "Ecocardiografía", "items": eco_notes or ["Eco cargada sin datos suficientes para interpretación automática."]})
    else:
        score -= 8
        sections.append({"title": "Ecocardiografía", "items": ["Sin ecocardiografía cargada."]})

    if last_rx:
        rx_notes = []

        if vhs is not None:
            if vhs > 11.5:
                score -= 10
                rx_notes.append(f"VHS aumentado ({vhs:g}). Compatible con cardiomegalia.")
                alerts.append("VHS elevado en RX.")
            else:
                rx_notes.append(f"VHS sin aumento relevante ({vhs:g}).")

        if vlas is not None:
            if vlas > 3:
                score -= 8
                rx_notes.append(f"VLAS aumentado ({vlas:g}). Sugiere aumento auricular izquierdo.")
            else:
                rx_notes.append(f"VLAS sin aumento relevante ({vlas:g}).")

        if "edema" in rx_text:
            score -= 20
            rx_notes.append("Se menciona edema pulmonar.")
            alerts.append("RX con mención de edema pulmonar.")
        if "congest" in rx_text:
            score -= 12
            rx_notes.append("Se menciona congestión vascular/pulmonar.")
        if "normal" in rx_text and not rx_notes:
            rx_notes.append("RX descrita sin alteraciones cardiopulmonares relevantes.")

        sections.append({"title": "Radiografía", "items": rx_notes or ["RX cargada sin datos suficientes para interpretación automática."]})
    else:
        sections.append({"title": "Radiografía", "items": ["Sin RX cardíaca cargada."]})

    if aiao is not None and vhs is not None:
        if aiao >= 1.6 and vhs <= 11.5:
            alerts.append("Eco sugiere AI aumentada, pero VHS no acompaña: revisar concordancia RX/Eco.")
        if aiao < 1.6 and vhs > 11.5:
            alerts.append("VHS elevado con AI/Ao normal: revisar otras causas de cardiomegalia.")

    if acvim in ["C", "D"] or "edema" in rx_text:
        recommendations.append("Control cardiológico estricto y reevaluar signos de insuficiencia cardíaca congestiva.")
    elif acvim == "B2" or (aiao is not None and aiao >= 1.6):
        recommendations.append("Seguimiento cardiológico periódico y control ecocardiográfico.")
    else:
        recommendations.append("Continuar controles periódicos según evolución clínica.")

    if not last_ecg:
        recommendations.append("Cargar ECG para completar interpretación eléctrica.")
    if not last_eco:
        recommendations.append("Cargar ecocardiografía para estadificación cardiológica.")
    if not last_rx:
        recommendations.append("Cargar RX torácica/cardiológica si hay tos, disnea, soplo o sospecha de congestión.")

    score = max(0, min(100, int(score)))

    if score >= 85:
        label = "🟢 Estable"
        conclusion = "Paciente cardiológicamente estable con los datos disponibles."
    elif score >= 60:
        label = "🟡 Requiere seguimiento"
        conclusion = "Paciente con hallazgos que justifican seguimiento cardiológico."
    elif score >= 40:
        label = "🟠 Riesgo moderado"
        conclusion = "Paciente con alteraciones cardiológicas relevantes; conviene control cercano."
    else:
        label = "🔴 Alto riesgo"
        conclusion = "Paciente con indicadores de alto riesgo cardiológico o posible descompensación."

    return {
        "score": score,
        "label": label,
        "conclusion": conclusion,
        "sections": sections,
        "alerts": alerts,
        "recommendations": recommendations,
    }
@app.get('/patients/{patient_id}/cardiology', response_class=HTMLResponse)
def patient_cardiology(
    request: Request,
    patient_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)

    if not patient:
        raise HTTPException(status_code=404)

    cardiology_events = (
        db.query(ClinicalEvent)
        .filter(
            ClinicalEvent.patient_id == patient.id,
            ClinicalEvent.event_type.in_(
                [
                    "ECG",
                    "Ecocardiografía",
                    "Radiografía"
                ]
            )
        )
        .order_by(ClinicalEvent.event_date.desc())
        .all()
    )

    ecg_events = [
        e for e in cardiology_events
        if e.event_type == "ECG"
    ]

    eco_events = [
        e for e in cardiology_events
        if e.event_type == "Ecocardiografía"
    ]

    rx_events = [
        e for e in cardiology_events
        if e.event_type == "Radiografía"
    ]

    last_ecg = ecg_events[0] if ecg_events else None
    previous_ecg = ecg_events[1] if len(ecg_events) > 1 else None

    last_eco = eco_events[0] if eco_events else None
    previous_eco = eco_events[1] if len(eco_events) > 1 else None

    last_rx = rx_events[0] if rx_events else None
    previous_rx = rx_events[1] if len(rx_events) > 1 else None

    def to_float(value):
        try:
            if value is None:
                return None

            value = str(value).replace(",", ".").strip()

            if value == "":
                return None

            return float(value)

        except Exception:
            return None

    def delta(current, previous):

        current_value = to_float(current)
        previous_value = to_float(previous)

        if current_value is None:
            return None

        if previous_value is None:
            return None

        return round(current_value - previous_value, 2)

    def trend_width(value, maximum):

        number = to_float(value)

        if number is None:
            return 0

        width = int((number / maximum) * 100)

        if width < 4:
            width = 4

        if width > 100:
            width = 100

        return width

    fc_delta = delta(
        last_ecg.ecg_hr if last_ecg else None,
        previous_ecg.ecg_hr if previous_ecg else None
    )

    axis_delta = delta(
        last_ecg.ecg_axis if last_ecg else None,
        previous_ecg.ecg_axis if previous_ecg else None
    )

    aiao_delta = delta(
        last_eco.eco_aiao if last_eco else None,
        previous_eco.eco_aiao if previous_eco else None
    )

    fs_delta = delta(
        last_eco.eco_fs if last_eco else None,
        previous_eco.eco_fs if previous_eco else None
    )
    acvim = (last_eco.eco_acvim if last_eco else "") or ""

    cardio_status = {
        "label": "Estable",
        "class": "ok",
        "icon": "🟢"
    }

    if acvim in ["B2"]:
        cardio_status = {
            "label": "En seguimiento",
            "class": "warn",
            "icon": "🟡"
        }

    if acvim in ["C", "D"]:
        cardio_status = {
            "label": "Control estricto",
            "class": "danger",
            "icon": "🔴"
        }

    active_treatment = ""

    if last_eco:
        active_treatment = (
            getattr(last_eco, "eco_treatment", "")
            or ""
        )

    if not active_treatment and last_ecg:
        active_treatment = (
            getattr(last_ecg, "treatment", "")
            or ""
        )

    if not active_treatment:
        active_treatment = "Sin tratamiento registrado"

    next_control = None

    if last_eco:
        stage = (last_eco.eco_acvim or "").upper()

        if stage in ["C", "D"]:
            next_control = 30

        elif stage == "B2":
            next_control = 90

        else:
            next_control = 180

    elif last_ecg:
        next_control = 180

    else:
        next_control = None
    fc_trend = [
        {
            "label": e.event_date.strftime("%d/%m"),
            "value": e.ecg_hr,
            "width": trend_width(e.ecg_hr, 220),
            "event_id": e.id
        }
        for e in reversed(ecg_events[:8])
        if e.ecg_hr
    ]

    aiao_trend = [
        {
            "label": e.event_date.strftime("%d/%m"),
            "value": e.eco_aiao,
            "width": trend_width(e.eco_aiao, 3),
            "event_id": e.id
        }
        for e in reversed(eco_events[:8])
        if e.eco_aiao
    ]

    fs_trend = [
        {
            "label": e.event_date.strftime("%d/%m"),
            "value": e.eco_fs,
            "width": trend_width(e.eco_fs, 60),
            "event_id": e.id
        }
        for e in reversed(eco_events[:8])
        if e.eco_fs
    ]

    comparison_notes = []

    if fc_delta is not None:

        if fc_delta < 0:
            comparison_notes.append(
                f"FC disminuyó {abs(fc_delta):g} lpm respecto al ECG previo."
            )

        elif fc_delta > 0:
            comparison_notes.append(
                f"FC aumentó {fc_delta:g} lpm respecto al ECG previo."
            )

        else:
            comparison_notes.append(
                "FC sin cambios respecto al ECG previo."
            )

    if axis_delta is not None:

        if abs(axis_delta) <= 10:
            comparison_notes.append(
                "Eje eléctrico sin cambios clínicamente relevantes."
            )

        elif axis_delta > 0:
            comparison_notes.append(
                f"Eje eléctrico aumentó {axis_delta:g}°."
            )

        else:
            comparison_notes.append(
                f"Eje eléctrico disminuyó {abs(axis_delta):g}°."
            )

    if aiao_delta is not None:

        if abs(aiao_delta) < 0.05:
            comparison_notes.append(
                "AI/Ao sin variación significativa."
            )

        elif aiao_delta > 0:
            comparison_notes.append(
                f"AI/Ao aumentó {aiao_delta:g}."
            )

        else:
            comparison_notes.append(
                f"AI/Ao disminuyó {abs(aiao_delta):g}."
            )

    if fs_delta is not None:

        if fs_delta > 0:
            comparison_notes.append(
                f"FS aumentó {fs_delta:g}%."
            )

        elif fs_delta < 0:
            comparison_notes.append(
                f"FS disminuyó {abs(fs_delta):g}%."
            )

        else:
            comparison_notes.append(
                "FS sin cambios."
            )

    if not comparison_notes:

        comparison_notes.append(
            "Todavía no hay controles suficientes para comparar la evolución."
        )

    cardio_ai = build_cardio_ai(
        last_ecg,
        previous_ecg,
        last_eco,
        previous_eco,
        last_rx,
        previous_rx
    )

    return templates.TemplateResponse(
        "patient_cardiology.html",
        {
            "request": request,
            "patient": patient,

            "cardiology_events": cardiology_events,

            "ecg_events": ecg_events,
            "eco_events": eco_events,
            "rx_events": rx_events,

            "last_ecg": last_ecg,
            "previous_ecg": previous_ecg,

            "last_eco": last_eco,
            "previous_eco": previous_eco,

            "last_rx": last_rx,
            "previous_rx": previous_rx,

            "fc_delta": fc_delta,
            "axis_delta": axis_delta,
            "aiao_delta": aiao_delta,
            "fs_delta": fs_delta,

            "cardio_status": cardio_status,

            "active_treatment": active_treatment,

            "next_control": next_control,

            "fc_trend": fc_trend,
            "aiao_trend": aiao_trend,
            "fs_trend": fs_trend,

            "comparison_notes": comparison_notes,

            "cardio_score": cardio_ai["score"],
            "cardio_score_label": cardio_ai["label"],
            "cardio_ai_conclusion": cardio_ai["conclusion"],
            "cardio_ai_sections": cardio_ai["sections"],
            "cardio_ai_alerts": cardio_ai["alerts"],
            "cardio_ai_recommendations": cardio_ai["recommendations"],

            "today": argentina_now().date()
        }
    )
@app.get('/patients/{patient_id}/edit', response_class=HTMLResponse)
def patient_edit_form(request: Request, patient_id: int, db: Session = Depends(get_db), user: User = Depends(require_user)):
    patient = db.get(Patient, patient_id)
    if not patient:
        raise HTTPException(status_code=404)
    owners = db.query(Owner).order_by(Owner.name).limit(500).all()
    return templates.TemplateResponse(
        'patient_edit.html',
        {'request': request, 'patient': patient, 'owners': owners}
    )


@app.post('/patients/{patient_id}/edit')
def patient_edit_save(
    patient_id: int,
    name: str = Form(...),
    owner_id: int = Form(...),
    species: str = Form(''),
    breed: str = Form(''),
    sex: str = Form(''),
    birth_date: str = Form(''),
    weight: str = Form(''),
    alerts: str = Form(''),
    notes: str = Form(''),
    owner_name: str = Form(''),
    owner_phone: str = Form(''),
    owner_whatsapp: str = Form(''),
    owner_email: str = Form(''),
    owner_address: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)
    if not patient:
        raise HTTPException(status_code=404)

    patient.name = name
    patient.owner_id = owner_id
    patient.species = species
    patient.breed = breed
    patient.sex = sex
    patient.birth_date = datetime.strptime(birth_date, '%Y-%m-%d').date() if birth_date.strip() else None
    patient.weight = float(weight.replace(',', '.')) if weight.strip() else None
    patient.alerts = alerts
    patient.notes = notes
    owner = db.get(Owner, owner_id)

    if owner:
        owner.name = owner_name
        owner.phone = owner_phone
        owner.whatsapp = owner_whatsapp
        owner.email = owner_email
        owner.address = owner_address

    db.commit()
    return RedirectResponse(f'/patients/{patient.id}', status_code=303)
@app.get('/patients/{patient_id}/events/{event_id}', response_class=HTMLResponse)
def event_detail(
    request: Request,
    patient_id: int,
    event_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)
    event = db.get(ClinicalEvent, event_id)
    previous_ecg = None

    if event and event.event_type == "ECG":
        previous_ecg = (
            db.query(ClinicalEvent)
            .filter(
                ClinicalEvent.patient_id == patient_id,
                ClinicalEvent.event_type == "ECG",
                ClinicalEvent.id != event.id,
                ClinicalEvent.event_date < event.event_date
            )
            .order_by(ClinicalEvent.event_date.desc())
            .first()
        )
    if not patient or not event:
        raise HTTPException(status_code=404)

    return templates.TemplateResponse(
        'event_detail.html',
{
    'request': request,
    'patient': patient,
    'event': event,
    'previous_ecg': previous_ecg,
    'clinical_ai': None
}
    )
@app.post('/patients/{patient_id}/events/{event_id}/ai')
def event_ai_analyze(
    patient_id: int,
    event_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)
    event = db.get(ClinicalEvent, event_id)

    if not patient or not event or event.patient_id != patient.id:
        raise HTTPException(status_code=404)

    result = ai_clinical_summary(event)

    return JSONResponse(result)
@app.post('/patients/{patient_id}/events/{event_id}/ai/save')
def event_ai_save_to_event(
    patient_id: int,
    event_id: int,
    ai_text: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)
    event = db.get(ClinicalEvent, event_id)

    if not patient or not event or event.patient_id != patient.id:
        raise HTTPException(status_code=404)

    if not ai_text.strip():
        return JSONResponse({
            "ok": False,
            "message": "No hay análisis IA para guardar."
        })

    separator = "\n\n────────────────────────\n\n"
    block = "🧠 ANÁLISIS IA\n\n" + ai_text.strip()

    current_description = event.description or ""

    if "🧠 ANÁLISIS IA" not in current_description:
        event.description = current_description + separator + block
    else:
        event.description = current_description + "\n\n" + block

    db.commit()

    return JSONResponse({
        "ok": True,
        "message": "Análisis IA agregado al evento."
    })
@app.post('/patients/{patient_id}/events/{event_id}/delete')
def event_delete(
    patient_id: int,
    event_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)
    event = db.get(ClinicalEvent, event_id)

    if not patient or not event or event.patient_id != patient.id:
        raise HTTPException(status_code=404)

    db.delete(event)
    db.commit()

    return RedirectResponse(f'/patients/{patient.id}', status_code=303)
@app.post('/patients/{patient_id}/events')
def event_create(
    request: Request,
    patient_id: int,
    event_type: str = Form(...),
    title: str = Form(''),
    description: str = Form(''),
    anamnesis: str = Form(''),
    physical_exam: str = Form(''),
    diagnosis: str = Form(''),
    treatment: str = Form(''),
    vaccine_name: str = Form(''),
    vaccine_lot: str = Form(''),
    vaccine_expiration: str = Form(''),
    next_vaccine_date: str = Form(''),
    dewormer_product: str = Form(''),
    dewormer_drug: str = Form(''),
    dewormer_dose: str = Form(''),
    next_deworming_date: str = Form(''),
    reminder_date: str = Form(''),
    reminder_type: str = Form(''),
    reminder_vaccine_name: str = Form(''),
    reminder_dewormer_name: str = Form(''),
    reminder_notes: str = Form(''),
    event_date: str = Form(''),
    weight: str = Form(''),
    temperature: str = Form(''),
    heart_rate: str = Form(''),
    respiratory_rate: str = Form(''),
    mucous_membranes: str = Form(''),
    crt: str = Form(''),
    hydration: str = Form(''),
    ecg_hr: str = Form(''),
    ecg_rhythm: str = Form(''),

    ecg_p: str = Form(''),
    ecg_pr: str = Form(''),

    ecg_qrs: str = Form(''),
    ecg_st: str = Form(''),

    ecg_t: str = Form(''),
    ecg_qt: str = Form(''),

    ecg_axis: str = Form(''),
    ecg_interpretation: str = Form(''),
    eco_aiao: str = Form(''),
    eco_fs: str = Form(''),
    eco_acvim: str = Form(''),

    eco_diagnosis: str = Form(''),
    eco_treatment: str = Form(''),

    # ECG ampliado
    ecg_p_mv: str = Form(''),
    ecg_qrs_mv: str = Form(''),
    ecg_t_mv: str = Form(''),
    ecg_qtc: str = Form(''),
    ecg_polarity: str = Form(''),
    ecg_arrhythmia: str = Form(''),
    ecg_conduction: str = Form(''),
    ecg_notes: str = Form(''),

    # Ecocardiografía ampliada
    eco_epss: str = Form(''),
    eco_lvidd: str = Form(''),
    eco_lvids: str = Form(''),
    eco_ivsd: str = Form(''),
    eco_ivss: str = Form(''),
    eco_lvpwd: str = Form(''),
    eco_lvpws: str = Form(''),
    eco_fe: str = Form(''),
    eco_la_size: str = Form(''),
    eco_lv_size: str = Form(''),
    eco_rv_size: str = Form(''),
    eco_ra_size: str = Form(''),
    eco_mitral: str = Form(''),
    eco_tricuspid: str = Form(''),
    eco_aortic: str = Form(''),
    eco_pulmonary: str = Form(''),
    eco_pulmonary_htn: str = Form(''),
    eco_pericardium: str = Form(''),
    eco_doppler: str = Form(''),
    eco_observations: str = Form(''),

    # Radiografía cardiológica
    rx_vhs: str = Form(''),
    rx_vlas: str = Form(''),
    rx_heart_size: str = Form(''),
    rx_left_atrium: str = Form(''),
    rx_left_heart: str = Form(''),
    rx_right_heart: str = Form(''),
    rx_pulmonary_vessels: str = Form(''),
    rx_lung_pattern: str = Form(''),
    rx_edema: str = Form(''),
    rx_congestion: str = Form(''),
    rx_trachea: str = Form(''),
    rx_observations: str = Form(''),

    attachments: list[UploadFile] = File(default=[]),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)
    if not patient:
        raise HTTPException(status_code=404, detail="Paciente no encontrado")
    quick_event_type = request.session.pop('quick_event_type', '')

    quick_map = {
        'rx': 'Radiografía',
        'ecg': 'ECG',
        'eco': 'Ecografía',
        'lab': 'Laboratorio',
        'vacuna': 'Vacuna'
        }

    if quick_event_type in quick_map:
        event_type = quick_map[quick_event_type]
    def to_float(value):
        try:
            return float(value.replace(',', '.')) if value and value.strip() else None
        except ValueError:
            return None

    def to_int(value):
        try:
            return int(float(value.replace(',', '.'))) if value and value.strip() else None
        except ValueError:
            return None 



    rd = None

    if event_type == 'Vacuna':
        if next_vaccine_date and next_vaccine_date.strip():
            try:
                rd = datetime.strptime(next_vaccine_date.strip(), '%Y-%m-%d').date()
            except ValueError:
                rd = None

        dewormer_product = ''
        dewormer_drug = ''
        dewormer_dose = ''
        next_deworming_date = ''

    elif event_type == 'Desparasitación':
        if next_deworming_date and next_deworming_date.strip():
            try:
                rd = datetime.strptime(next_deworming_date.strip(), '%Y-%m-%d').date()
            except ValueError:
                rd = None

        vaccine_name = ''
        vaccine_lot = ''
        vaccine_expiration = ''
        next_vaccine_date = ''

    else:
        vaccine_name = ''
        vaccine_lot = ''
        vaccine_expiration = ''
        next_vaccine_date = ''
        dewormer_product = ''
        dewormer_drug = ''
        dewormer_dose = ''
        next_deworming_date = ''

        if reminder_date and reminder_date.strip():
            try:
                rd = datetime.strptime(reminder_date.strip(), '%Y-%m-%d').date()
            except ValueError:
                rd = None
    managed_reminder_date = None

    if reminder_type and reminder_date and reminder_date.strip():
        try:
            managed_reminder_date = datetime.strptime(reminder_date.strip(), '%Y-%m-%d').date()
            rd = None
        except ValueError:
            managed_reminder_date = None    
    event_created_at = argentina_now()
    event_date = event_date or ""
    
    if event_date.strip():
        try:
            event_created_at = datetime.strptime(
                event_date.strip(),
                "%Y-%m-%d"
            )
        except ValueError:
            pass

    event = ClinicalEvent(
        patient_id=patient_id,
        event_date=event_created_at,
        event_type=event_type,
        title=title or event_type,
        description=description or '',
        anamnesis=anamnesis or '',
        physical_exam=physical_exam or '',
        diagnosis=diagnosis or '',
        treatment=treatment or '',
        vaccine_name=vaccine_name or '',
        vaccine_lot=vaccine_lot or '',
        vaccine_expiration=vaccine_expiration or '',
        next_vaccine_date=next_vaccine_date or '',
        dewormer_product=dewormer_product or '',
        dewormer_drug=dewormer_drug or '',
        dewormer_dose=dewormer_dose or '',
        next_deworming_date=next_deworming_date or '',
        reminder_date=rd,
        weight=to_float(weight),
        temperature=to_float(temperature),
        heart_rate=to_int(heart_rate),
        respiratory_rate=to_int(respiratory_rate),
        mucous_membranes=mucous_membranes or '',
        crt=crt or '',
        hydration=hydration or '',
        ecg_hr=ecg_hr or '',
        ecg_rhythm=ecg_rhythm or '',
        ecg_p=ecg_p or '',
        ecg_pr=ecg_pr or '',
        ecg_qrs=ecg_qrs or '',
        ecg_st=ecg_st or '',
        ecg_t=ecg_t or '',
        ecg_qt=ecg_qt or '',
        ecg_axis=ecg_axis or '',
        ecg_interpretation=ecg_interpretation or '',
        eco_aiao=eco_aiao or '',
        eco_fs=eco_fs or '',
        eco_acvim=eco_acvim or '',
        eco_diagnosis=eco_diagnosis or '',
        eco_treatment=eco_treatment or '',

        ecg_p_mv=ecg_p_mv or '',
        ecg_qrs_mv=ecg_qrs_mv or '',
        ecg_t_mv=ecg_t_mv or '',
        ecg_qtc=ecg_qtc or '',
        ecg_polarity=ecg_polarity or '',
        ecg_arrhythmia=ecg_arrhythmia or '',
        ecg_conduction=ecg_conduction or '',
        ecg_notes=ecg_notes or '',

        eco_epss=eco_epss or '',
        eco_lvidd=eco_lvidd or '',
        eco_lvids=eco_lvids or '',
        eco_ivsd=eco_ivsd or '',
        eco_ivss=eco_ivss or '',
        eco_lvpwd=eco_lvpwd or '',
        eco_lvpws=eco_lvpws or '',
        eco_fe=eco_fe or '',
        eco_la_size=eco_la_size or '',
        eco_lv_size=eco_lv_size or '',
        eco_rv_size=eco_rv_size or '',
        eco_ra_size=eco_ra_size or '',
        eco_mitral=eco_mitral or '',
        eco_tricuspid=eco_tricuspid or '',
        eco_aortic=eco_aortic or '',
        eco_pulmonary=eco_pulmonary or '',
        eco_pulmonary_htn=eco_pulmonary_htn or '',
        eco_pericardium=eco_pericardium or '',
        eco_doppler=eco_doppler or '',
        eco_observations=eco_observations or '',

        rx_vhs=rx_vhs or '',
        rx_vlas=rx_vlas or '',
        rx_heart_size=rx_heart_size or '',
        rx_left_atrium=rx_left_atrium or '',
        rx_left_heart=rx_left_heart or '',
        rx_right_heart=rx_right_heart or '',
        rx_pulmonary_vessels=rx_pulmonary_vessels or '',
        rx_lung_pattern=rx_lung_pattern or '',
        rx_edema=rx_edema or '',
        rx_congestion=rx_congestion or '',
        rx_trachea=rx_trachea or '',
        rx_observations=rx_observations or '',

        created_by=user.username
    )

    db.add(event)
    db.commit()
    db.refresh(event)
    close_managed_due_events(db, patient, event, user)
    db.commit()
    if managed_reminder_date:
        reminder_title = 'Recordatorio pendiente'
        reminder_event_type = 'Control'

        if reminder_type == 'Vacuna':
            reminder_event_type = 'Vacuna'
            reminder_title = f'Vacuna pendiente: {reminder_vaccine_name or "Vacuna"}'

        elif reminder_type == 'Desparasitación':
            reminder_event_type = 'Desparasitación'
            reminder_title = f'Desparasitación pendiente: {reminder_dewormer_name or "Desparasitación"}'

        elif reminder_type == 'Control clínico':
            reminder_event_type = 'Control'
            reminder_title = 'Control clínico pendiente'

        reminder_description = (
            f'{DUE_ACTIVE_MARKER}\n'
            f'Tipo de recordatorio: {reminder_type}\n'
            f'Vacuna: {reminder_vaccine_name or "-"}\n'
            f'Desparasitación: {reminder_dewormer_name or "-"}\n'
            f'Notas: {reminder_notes or "-"}'
        )

        pending_event = ClinicalEvent(
            patient_id=patient.id,
            event_date=argentina_now(),
            event_type=reminder_event_type,
            title=reminder_title,
            description=reminder_description,
            reminder_date=managed_reminder_date,
            created_by=user.username
        )

        db.add(pending_event)
        db.commit()    
    if weight and str(weight).strip():
        try:
            patient.weight = float(str(weight).replace(',', '.'))
            db.commit()
        except ValueError:
            pass
    active_waiting_entries = (
        db.query(WaitingListEntry)
        .filter(WaitingListEntry.patient_id == patient.id)
        .filter(WaitingListEntry.status.in_(['Esperando', 'En consulta']))
        .all()
    )

    for waiting_entry in active_waiting_entries:
        waiting_entry.status = 'Finalizado'
        waiting_entry.finished_at = argentina_now()

    if active_waiting_entries:
        db.commit()
    for file in attachments:
        if not file or not file.filename:
            continue

        if supabase is None:
            raise HTTPException(
                status_code=500,
                detail="Supabase Storage no configurado. Faltan SUPABASE_URL o SUPABASE_KEY en Render."
            )

        original_name = os.path.basename(file.filename)
        safe_name = original_name.replace(" ", "_")
        unique_name = f"{uuid.uuid4().hex}_{safe_name}"

        content = file.file.read()

        if not content:
            continue

        content_type = file.content_type or mimetypes.guess_type(original_name)[0] or "application/octet-stream"
        storage_path = f"patient_{patient_id}/event_{event.id}/{unique_name}"

        try:
            supabase.storage.from_("adjuntos").upload(
                path=storage_path,
                file=content,
                file_options={
                    "content-type": content_type,
                    "upsert": "false"
                }
            )

            public_url = supabase.storage.from_("adjuntos").get_public_url(storage_path)

        except Exception as e:
            import traceback
        
            print("=" * 80)
            print("ERROR SUBIENDO A SUPABASE")
            traceback.print_exc()
            print("=" * 80)
        
            db.rollback()
        
            raise HTTPException(
                status_code=500,
                detail=str(e)
            )

        attachment = EventAttachment(
            event_id=event.id,
            filename=original_name,
            file_path=public_url
        )

        db.add(attachment)

    db.commit()

    return RedirectResponse(
        url=f"/patients/{patient_id}",
        status_code=303
    )
    

@app.get('/products', response_class=HTMLResponse)
def products_page(
    request: Request,
    q: str = '',
    rubro: str = '',
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    query = db.query(Product).filter(Product.active == True)

    if q:
        like = f"%{q}%"
        query = query.filter(
            or_(
                Product.name.ilike(like),
                Product.code.ilike(like),
                Product.barcode.ilike(like),
                Product.provider.ilike(like),
                Product.manufacturer.ilike(like),
            )
        )

    if rubro:
        query = query.filter(Product.rubro == rubro)

    

    total_products = db.query(Product).filter(Product.active == True).count()
    low_stock = db.query(Product).filter(
        Product.active == True,
        Product.stock != None,
        Product.min_stock != None,
        Product.stock <= Product.min_stock
    ).count()

    today = argentina_now().date()
    soon = today + timedelta(days=60)
    products = query.order_by(Product.name).limit(300).all()
    for p in products:
        p.row_color = ''

        if p.expiration_date:
            if p.expiration_date < today:
                p.row_color = '#ffd6d6'
            elif (p.expiration_date - today).days <= 60:
                p.row_color = '#fff6cc'

        if (
            p.stock is not None and
            p.min_stock is not None and
            p.min_stock > 0 and
            p.stock <= p.min_stock
        ):
                p.row_color = '#ffe5e5'
    expired_or_soon = db.query(Product).filter(
        Product.active == True,
        Product.expiration_date != None,
        Product.expiration_date <= soon
    ).count()

    total_cost = 0
    total_sale = 0

    for p in db.query(Product).filter(Product.active == True).all():
        stock = p.stock or 0
        total_cost += stock * (p.cost_price or 0)
        total_sale += stock * (p.sale_price or 0)

    avg_margin = 0
    if total_cost > 0:
        avg_margin = ((total_sale - total_cost) / total_cost) * 100

    rubros = [
        r[0] for r in db.query(Product.rubro)
        .filter(Product.rubro != '')
        .distinct()
        .order_by(Product.rubro)
        .all()
    ]

    return templates.TemplateResponse(
        'products.html',
        {
            'request': request,
            'products': products,
            'q': q,
            'rubro': rubro,
            'rubros': rubros,
            'total_products': total_products,
            'low_stock': low_stock,
            'expired_or_soon': expired_or_soon,
            'total_cost': total_cost,
            'total_sale': total_sale,
            'avg_margin': avg_margin,
            'today': today
        }
    )
@app.post('/products/import')
async def import_products(
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    content = await file.read()

    wb = load_workbook(BytesIO(content), data_only=True)
    ws = wb.active

    headers = []
    for cell in ws[1]:
        headers.append(str(cell.value).strip() if cell.value else "")

    def get_value(row, column_name):
        if column_name not in headers:
            return None
        idx = headers.index(column_name)
        return row[idx] if idx < len(row) else None

    def to_float(value):
        try:
            if value is None:
                return None
            value = str(value).strip()
            if value == "" or value == "---":
                return None
            return float(value.replace(",", "."))
        except ValueError:
            return None

    def to_date(value):
        try:
            if value is None:
                return None
            if hasattr(value, "date"):
                return value.date()
            value = str(value).strip()
            if value == "" or value.lower() == "sin asignar":
                return None
            return datetime.strptime(value, "%d/%m/%Y").date()
        except ValueError:
            return None

    imported = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        name = str(get_value(row, "Nombre") or "").strip()
        if not name:
            continue

        cost = to_float(get_value(row, "Costo"))
        sale = to_float(get_value(row, "Precio"))
        margin = to_float(get_value(row, "Margen (%)"))

        if margin is None and cost and sale:
            margin = ((sale - cost) / cost) * 100

        code = str(get_value(row, "Código") or "").strip()
        barcode = str(get_value(row, "Código de Barras") or "").strip()
        
        existing = None
        
        if code:
            existing = db.query(Product).filter(Product.code == code).first()
        
        if existing is None and barcode:
            existing = db.query(Product).filter(Product.barcode == barcode).first()
        
        if existing is None:
            existing = db.query(Product).filter(Product.name == name).first()
        
        if existing:
            existing.rubro = str(get_value(row, "Rubro") or "").strip()
            existing.name = name
            existing.code = code
            existing.barcode = barcode
            existing.sale_price = sale
            existing.cost_price = cost
            existing.margin_percent = margin
            existing.stock = to_float(get_value(row, "Stock"))
            existing.min_stock = to_float(get_value(row, "Stock Min"))
            existing.expiration_date = to_date(get_value(row, "Vencimiento"))
            existing.manufacturer = str(get_value(row, "Elaborador") or "").strip()
            existing.provider = str(get_value(row, "Proveedores") or "").strip()
            existing.notes = str(get_value(row, "Nota") or "").strip()
            existing.active = True
        else:
            product = Product(
                rubro=str(get_value(row, "Rubro") or "").strip(),
                name=name,
                code=code,
                barcode=barcode,
                sale_price=sale,
                cost_price=cost,
                margin_percent=margin,
                stock=to_float(get_value(row, "Stock")),
                min_stock=to_float(get_value(row, "Stock Min")),
                expiration_date=to_date(get_value(row, "Vencimiento")),
                manufacturer=str(get_value(row, "Elaborador") or "").strip(),
                provider=str(get_value(row, "Proveedores") or "").strip(),
                notes=str(get_value(row, "Nota") or "").strip()
            )
        
            db.add(product)
        
        imported += 1

    db.commit()

    return RedirectResponse(
        url=f"/products?imported={imported}",
        status_code=303
    )
@app.post('/products')
def product_create(
    name: str = Form(''),
    rubro: str = Form(''),
    tipo: str = Form(''),
    code: str = Form(''),
    barcode: str = Form(''),
    cost_price: str = Form(''),
    sale_price: str = Form(''),
    stock: str = Form(''),
    min_stock: str = Form(''),
    expiration_date: str = Form(''),
    provider: str = Form(''),
    manufacturer: str = Form(''),
    notes: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    def to_float(value):
        try:
            return float(value.replace(',', '.')) if value and value.strip() else None
        except ValueError:
            return None

    exp = None
    if expiration_date and expiration_date.strip():
        try:
            exp = datetime.strptime(expiration_date.strip(), "%Y-%m-%d").date()
        except ValueError:
            exp = None

    cost = to_float(cost_price)
    sale = to_float(sale_price)
    margin = None
    if cost and sale:
        margin = ((sale - cost) / cost) * 100

    product = Product(
        name=name or '',
        rubro=rubro or '',
        tipo=tipo or '',
        code=code or '',
        barcode=barcode or '',
        cost_price=cost,
        sale_price=sale,
        margin_percent=margin,
        stock=to_float(stock),
        min_stock=to_float(min_stock),
        expiration_date=exp,
        provider=provider or '',
        manufacturer=manufacturer or '',
        notes=notes or ''
    )

    db.add(product)
    db.commit()

    return RedirectResponse('/products', status_code=303)
@app.get('/products/{product_id}/edit', response_class=HTMLResponse)
def product_edit_page(
    request: Request,
    product_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    product = db.get(Product, product_id)

    if not product:
        raise HTTPException(
            status_code=404,
            detail="Producto no encontrado"
        )

    return templates.TemplateResponse(
        'product_edit.html',
        {
            'request': request,
            'product': product
        }
    )
@app.post('/products/{product_id}/edit')
def product_edit_save(
    product_id: int,
    name: str = Form(''),
    rubro: str = Form(''),
    tipo: str = Form(''),
    code: str = Form(''),
    barcode: str = Form(''),
    cost_price: str = Form(''),
    sale_price: str = Form(''),
    margin_percent: str = Form(''),
    stock: str = Form(''),
    min_stock: str = Form(''),
    provider: str = Form(''),
    manufacturer: str = Form(''),
    notes: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    product = db.get(Product, product_id)

    if not product:
        raise HTTPException(
            status_code=404,
            detail="Producto no encontrado"
        )

    def to_float(value):
        try:
            return float(value.replace(',', '.')) if value and value.strip() else None
        except ValueError:
            return None

    product.name = name
    product.rubro = rubro
    product.tipo = tipo
    product.code = code
    product.barcode = barcode
    product.cost_price = to_float(cost_price)
    product.sale_price = to_float(sale_price)
    product.margin_percent = to_float(margin_percent)
    product.stock = to_float(stock)
    product.min_stock = to_float(min_stock)
    product.provider = provider
    product.manufacturer = manufacturer
    product.notes = notes

    

    db.commit()

    return RedirectResponse(
        url='/products',
        status_code=303
    )
@app.post('/products/{product_id}/adjust')
def product_adjust_stock(
    product_id: int,
    qty: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    product = db.get(Product, product_id)

    if not product:
        raise HTTPException(
            status_code=404,
            detail="Producto no encontrado"
        )

    try:
        amount = float(str(qty).replace(',', '.'))
    except ValueError:
        amount = 0

    current_stock = product.stock or 0
    product.stock = current_stock + amount

    db.commit()

    return RedirectResponse(
        url='/products',
        status_code=303
    )
# ===== VENTAS =====

@app.get('/sales', response_class=HTMLResponse)
def sales_page(
    request: Request,
    sale_date: str = '',
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    products = (
        db.query(Product)
        .filter(Product.active == True)
        .order_by(Product.name)
        .all()
    )
    patients = db.query(Patient).order_by(Patient.name).all()
    owners = db.query(Owner).order_by(Owner.name).all()
    patient_owner_map = {}

    for patient in patients:
        patient_owner_map[str(patient.id)] = patient.owner_id if patient.owner_id else ''
    selected_date = datetime.strptime(sale_date, "%Y-%m-%d").date() if sale_date else argentina_now().date()
    
    day_start = datetime.combine(selected_date, datetime.min.time())
    day_end = datetime.combine(selected_date, datetime.max.time())
    
    sales = (
        db.query(Sale)
        .filter(Sale.date >= day_start, Sale.date <= day_end)
        .order_by(Sale.date.desc())
        .all()
    )

    for sale in sales:
        sale.items_count = (
            db.query(SaleItem)
            .filter(SaleItem.sale_id == sale.id)
            .count()
        )
        sale.patient_name = ''
        sale.owner_name = ''
        sale.items_detail = []
        sale.total_paid = (
            sum(
                p.amount or 0
                for p in db.query(SalePayment)
                .filter(SalePayment.sale_id == sale.id)
                .all()
                if p.method != 'Cuenta corriente'
            )
        )

        sale.balance_due = (sale.total or 0) - sale.total_paid

        items = (
            db.query(SaleItem)
            .filter(SaleItem.sale_id == sale.id)
            .all()
        )
    
        for item in items:
            product = db.get(Product, item.product_id)
            sale.items_detail.append({
                'product_name': product.name if product else 'Producto eliminado',
                'quantity': item.quantity,
                'unit_price': item.unit_price,
                'subtotal': item.subtotal
            })

        if sale.patient_id:
            patient = db.get(Patient, sale.patient_id)
            if patient:
                sale.patient_name = patient.name

        if sale.owner_id:
            owner = db.get(Owner, sale.owner_id)
            if owner:
                sale.owner_name = owner.name
    today = argentina_now().date()
    month_start = today.replace(day=1)

    today_sales_total = 0
    month_sales_total = 0
    today_products_count = 0
    month_profit_total = 0
    month_cost_total = 0
    month_sales_count = 0
    ticket_average = 0
    all_sales = db.query(Sale).all()

    for sale in all_sales:
        sale_date = sale.date.date() if sale.date else None

        if sale_date == today:
            today_sales_total += sale.total or 0

        if sale_date and sale_date >= month_start:
            month_sales_total += sale.total or 0
            month_cost_total += sale.cost_total or 0
            month_profit_total += sale.profit_amount or 0
            month_sales_count += 1
        if month_sales_count > 0:
            ticket_average = month_sales_total / month_sales_count
        today_items = (
        db.query(SaleItem)
        .join(Sale, SaleItem.sale_id == Sale.id)
        .filter(Sale.date >= datetime.combine(today, datetime.min.time()))
        .all()
    )

    for item in today_items:
        today_products_count += item.quantity or 0
    month_items = (
        db.query(SaleItem)
        .join(Sale, SaleItem.sale_id == Sale.id)
        .filter(Sale.date >= datetime.combine(month_start, datetime.min.time()))
        .all()
    )
    
    product_stats = {}
    
    for item in month_items:
        product = db.get(Product, item.product_id)
        if not product:
            continue
    
        name = product.name
    
        if name not in product_stats:
            product_stats[name] = 0
    
        product_stats[name] += item.quantity or 0
    
    top_product_name = "-"
    top_product_qty = 0
    top_products_sold = sorted(
        product_stats.items(),
        key=lambda item: item[1],
        reverse=True
    )[:10]
    if product_stats:
        top_product_name = max(product_stats, key=product_stats.get)
        top_product_qty = product_stats[top_product_name]
    return templates.TemplateResponse(
        'sales.html',
        {
            'request': request,
            'products': products,
            'patients': patients,
            'patient_owner_map': patient_owner_map,
            'owners': owners,
            'sales': sales,
            'selected_date': selected_date,
            'today_sales_total': today_sales_total,
            'month_sales_total': month_sales_total,
            'today_products_count': today_products_count,
            'month_cost_total': month_cost_total,
            'month_profit_total': month_profit_total,
            'month_sales_count': month_sales_count,
            'ticket_average': ticket_average,
            'top_product_name': top_product_name,
            'top_product_qty': top_product_qty,
            'top_products_sold': top_products_sold
            
            
        }
    )
@app.post('/sales')
def sales_create(
    patient_id: str = Form(''),
    owner_id: str = Form(''),
    product_id: list[str] = Form([]),
    quantity: list[str] = Form([]),
    unit_price: list[str] = Form([]),
    notes: str = Form(''),

    pay_efectivo: str = Form('0'),
    pay_debito: str = Form('0'),
    pay_credito: str = Form('0'),
    pay_transferencia: str = Form('0'),
    pay_cuenta_corriente: str = Form('0'),
    save_as_quote: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    def to_float(value):
        try:
            return float(str(value).replace(',', '.')) if value and str(value).strip() else 0
        except ValueError:
            return 0
    has_product = False

    for pid in product_id:
        if pid and str(pid).strip():
            has_product = True
            break

    if not has_product:
        return RedirectResponse(
            url='/sales',
            status_code=303
        )
    sale = Sale(
        status='quote' if save_as_quote else 'paid',
        total=0,
        payment_method='Mixto',
        patient_id=int(patient_id) if patient_id else None,
        owner_id=int(owner_id) if owner_id else None,
        notes=notes or ''
    )
    db.add(sale)
    db.flush()
    
    total = 0
    cost_total = 0
    for pid, qty_raw, price_raw in zip(product_id, quantity, unit_price):

        if not pid:
            continue

        product = db.get(Product, int(pid))

        if not product:
            continue

        qty = to_float(qty_raw)
        price = to_float(price_raw)

        if qty <= 0:
            continue

        subtotal = qty * price
        product_cost = product.cost_price or 0
        cost_total += qty * product_cost
        total += subtotal

        item = SaleItem(
            sale_id=sale.id,
            product_id=product.id,
            quantity=qty,
            unit_price=price,
            subtotal=subtotal
        )

        db.add(item)

        current_stock = product.stock or 0
        if not save_as_quote:
            product.stock = current_stock - qty

    sale.total = total
    sale.cost_total = cost_total
    sale.profit_amount = total - cost_total
    sale.margin_percent = ((total - cost_total) / total * 100) if total > 0 else 0

    payments_to_create = [
        ('Efectivo', to_float(pay_efectivo)),
        ('Débito', to_float(pay_debito)),
        ('Crédito', to_float(pay_credito)),
        ('Transferencia', to_float(pay_transferencia)),
        ('Cuenta corriente', to_float(pay_cuenta_corriente)),
    ]

    total_paid = 0
    sum_payments = sum(
        amount
        for method, amount in payments_to_create
        if method != 'Cuenta corriente'
    )

    has_account_debt = any(
        method == 'Cuenta corriente' and amount > 0
        for method, amount in payments_to_create
    )
    
    if total > 0 and sum_payments <= 0 and not has_account_debt and not save_as_quote:
        payments_to_create = [
            ('Efectivo', total)
        ]
    if not save_as_quote:
        for method, amount in payments_to_create:
            if amount <= 0:
                continue
    
            payment = SalePayment(
                sale_id=sale.id,
                method=method,
                amount=amount
            )
    
            db.add(payment)
    
            if method != 'Cuenta corriente':
                total_paid += amount

    if save_as_quote:
        sale.status = 'quote'
    elif total_paid >= total:
        sale.status = 'paid'
    else:
        sale.status = 'pending'

    db.commit()

    return RedirectResponse(
        url='/sales',
        status_code=303
    )
@app.post('/sales/{sale_id}/convert')
def sales_convert(
    sale_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    sale = db.get(Sale, sale_id)

    if not sale:
        raise HTTPException(status_code=404, detail='Venta no encontrada')

    if sale.status != 'quote':
        return RedirectResponse(
            url=f'/sales/{sale_id}',
            status_code=303
        )

    items = (
        db.query(SaleItem)
        .filter(SaleItem.sale_id == sale.id)
        .all()
    )

    for item in items:
        product = db.get(Product, item.product_id)

        if product:
            product.stock = (product.stock or 0) - (item.quantity or 0)

    sale.status = 'pending'

    db.commit()

    return RedirectResponse(
        url=f'/sales/{sale_id}',
        status_code=303
    )
@app.post('/sales/{sale_id}/cancel')
def sales_cancel(
    sale_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    sale = db.get(Sale, sale_id)

    if not sale:
        raise HTTPException(
            status_code=404,
            detail='Venta no encontrada'
        )

    if sale.status == 'cancelled':
        return RedirectResponse(
            url='/sales',
            status_code=303
        )

    items = (
        db.query(SaleItem)
        .filter(SaleItem.sale_id == sale.id)
        .all()
    )

    for item in items:
        product = db.get(Product, item.product_id)
        if product:
            current_stock = product.stock or 0
            product.stock = current_stock + (item.quantity or 0)

    sale.status = 'cancelled'

    db.commit()

    return RedirectResponse(
        url='/sales',
        status_code=303
    )
@app.get('/sales/{sale_id}', response_class=HTMLResponse)
def sales_detail(
    sale_id: int,
    request: Request,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    sale = db.get(Sale, sale_id)

    if not sale:
        raise HTTPException(status_code=404, detail='Venta no encontrada')

    patient = db.get(Patient, sale.patient_id) if sale.patient_id else None
    owner = db.get(Owner, sale.owner_id) if sale.owner_id else None

    items = (
        db.query(SaleItem)
        .filter(SaleItem.sale_id == sale.id)
        .all()
    )

    item_details = []

    for item in items:
        product = db.get(Product, item.product_id)

        item_details.append({
            'product_name': product.name if product else 'Producto eliminado',
            'quantity': item.quantity,
            'unit_price': item.unit_price,
            'subtotal': item.subtotal
        })
    products = (
        db.query(Product)
        .filter(Product.active == True)
        .order_by(Product.name)
        .all()
    )
    payments = (
        db.query(SalePayment)
        .filter(SalePayment.sale_id == sale.id)
        .all()
    )
    
    total_paid = 0
    
    for p in payments:
        method_clean = (p.method or '').strip().lower()
    
        if 'cuenta' in method_clean and 'corriente' in method_clean:
            continue
    
        total_paid += p.amount or 0
    
    balance_due = (sale.total or 0) - total_paid
    return templates.TemplateResponse(
        'sale_detail.html',
    {
        'request': request,
        'sale': sale,
        'patient': patient,
        'owner': owner,
        'items': item_details,
        'products': products,
        'payments': payments,
        'balance_due': balance_due
    }
    )
@app.post('/sales/{sale_id}/add-item')
def sale_add_item(
    sale_id: int,
    product_id: str = Form(''),
    quantity: str = Form('1'),
    unit_price: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    def to_float(value):
        try:
            return float(str(value).replace(',', '.')) if value and str(value).strip() else 0
        except ValueError:
            return 0

    sale = db.get(Sale, sale_id)

    if not sale:
        raise HTTPException(status_code=404, detail='Venta no encontrada')

    if not product_id:
        return RedirectResponse(
            url=f'/sales/{sale.id}',
            status_code=303
        )

    product = db.get(Product, int(product_id))

    if not product:
        raise HTTPException(status_code=404, detail='Producto no encontrado')

    qty = to_float(quantity)
    price = to_float(unit_price)

    if price <= 0:
        price = product.sale_price or 0

    subtotal = qty * price

    item = SaleItem(
        sale_id=sale.id,
        product_id=product.id,
        quantity=qty,
        unit_price=price,
        subtotal=subtotal
    )

    db.add(item)

    sale.total = (sale.total or 0) + subtotal

    if product.stock is not None:
        product.stock = (product.stock or 0) - qty

    db.commit()

    return RedirectResponse(
        url=f'/sales/{sale.id}',
        status_code=303
    )
@app.post('/sales/{sale_id}/confirm')
def sale_confirm(
    sale_id: int,
    payment_method: str = Form('Efectivo'),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    sale = db.get(Sale, sale_id)

    if not sale:
        raise HTTPException(status_code=404, detail='Venta no encontrada')
    sale.payment_method = payment_method or 'Efectivo'
    sale.status = 'paid'

    db.commit()

    return RedirectResponse(
        url=f'/sales/{sale.id}',
        status_code=303
    )
@app.post('/sales/{sale_id}/add-payment')
def sale_add_payment(
    sale_id: int,
    method: str = Form('Efectivo'),
    amount: str = Form('0'),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    def to_float(value):
        try:
            return float(str(value).replace(',', '.')) if value and str(value).strip() else 0
        except ValueError:
            return 0

    sale = db.get(Sale, sale_id)

    if not sale:
        raise HTTPException(status_code=404, detail='Venta no encontrada')

    payment_amount = to_float(amount)

    if payment_amount <= 0:
        return RedirectResponse(
            url=f'/sales/{sale.id}',
            status_code=303
        )

    payment = SalePayment(
        sale_id=sale.id,
        method=method or 'Efectivo',
        amount=payment_amount
    )

    db.add(payment)

    payments = (
        db.query(SalePayment)
        .filter(SalePayment.sale_id == sale.id)
        .all()
    )

    total_paid = sum(p.amount or 0 for p in payments) + payment_amount
    balance_due = (sale.total or 0) - total_paid
    if balance_due > 0 and sale.status != 'cancelled':
        sale.status = 'pending'
    elif balance_due <= 0 and sale.status != 'cancelled':
        sale.status = 'paid'
    if balance_due <= 0:
        sale.status = 'paid'
    else:
        sale.status = 'pending'
    if sale.status == 'paid':
        sale.payment_method = method
    else:
        sale.payment_method = 'Cuenta corriente'
    db.commit()

    return RedirectResponse(
        url=f'/sales/{sale.id}',
        status_code=303
    )
@app.get('/migration', response_class=HTMLResponse)
def migration(request: Request, user: User = Depends(require_user)):
    return templates.TemplateResponse(
        'migration.html',
        {
            'request': request,
            'result': None
        }
    )


@app.post('/migration/agenda-pendientes', response_class=HTMLResponse)
async def migration_agenda_pendientes(
    request: Request,
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    import csv
    import io

    def clean_text(value):
        if value is None:
            return ''
        value = str(value).strip()
        try:
            value = value.encode('latin1').decode('utf-8')
        except Exception:
            pass
        return value

    def pick(row, names):
        normalized = {clean_text(k).lower(): clean_text(v) for k, v in row.items()}
        for name in names:
            key = name.lower()
            if key in normalized:
                return normalized[key]
        return ''

    def parse_date(value):
        if value is None:
            return None
        if hasattr(value, 'date'):
            return value.date()

        value = clean_text(value)

        for fmt in ['%d/%m/%Y', '%Y-%m-%d', '%Y-%m-%d %H:%M:%S', '%d-%m-%Y', '%d/%m/%y']:
            try:
                return datetime.strptime(value, fmt).date()
            except Exception:
                pass

        return None

    content = await file.read()
    filename = (file.filename or '').lower()

    rows = []

    if filename.endswith('.xlsx'):
        wb = load_workbook(BytesIO(content), data_only=True)
        ws = wb.active
        headers = [clean_text(c.value) for c in ws[1]]

        for row in ws.iter_rows(min_row=2, values_only=True):
            rows.append(dict(zip(headers, row)))
    else:
        text_content = content.decode('utf-8-sig', errors='replace')
        sample = text_content[:1000]
        delimiter = ';' if sample.count(';') >= sample.count(',') else ','
        reader = csv.DictReader(io.StringIO(text_content), delimiter=delimiter)
        rows = list(reader)

    imported = 0
    skipped = 0
    created_owners = 0
    created_patients = 0

    for row in rows:
        owner_name = pick(row, ['Propietario / Cliente', 'Propietario', 'Cliente', 'Dueño', 'Responsable'])
        patient_name = pick(row, ['Paciente / Mascota', 'Paciente', 'Mascota', 'Animal'])
        phone = pick(row, ['Teléfono / WhatsApp', 'Teléfono', 'Telefono', 'WhatsApp', 'Whatsapp', 'Celular'])
        date_value = pick(row, ['Fecha', 'Fecha pendiente', 'Fecha agenda', 'Fecha turno'])
        service = pick(row, ['Motivo / Servicio', 'Motivo', 'Servicio', 'Detalle'])
        notes = pick(row, ['Notas', 'Observaciones', 'Comentario'])

        reminder_date = parse_date(date_value)

        if not reminder_date:
            skipped += 1
            continue

        if not owner_name:
            owner_name = 'Sin propietario'

        if not patient_name:
            patient_name = 'Sin paciente'

        owner = db.query(Owner).filter(Owner.name.ilike(owner_name)).first()

        if not owner:
            owner = Owner(
                name=owner_name,
                phone=phone,
                whatsapp=phone
            )
            db.add(owner)
            db.flush()
            created_owners += 1
        else:
            if phone and not owner.phone:
                owner.phone = phone
            if phone and not owner.whatsapp:
                owner.whatsapp = phone

        patient = (
            db.query(Patient)
            .filter(Patient.name.ilike(patient_name))
            .filter(Patient.owner_id == owner.id)
            .first()
        )

        if not patient:
            patient = Patient(
                name=patient_name,
                owner_id=owner.id
            )
            db.add(patient)
            db.flush()
            created_patients += 1

        service_text = f'{service} {notes}'.lower()

        if 'vacun' in service_text:
            event_type = 'Vacuna'
            title = service or 'Vacuna pendiente'
        elif 'despar' in service_text:
            event_type = 'Desparasitación'
            title = service or 'Desparasitación pendiente'
        elif 'control' in service_text:
            event_type = 'Control'
            title = service or 'Control pendiente'
        else:
            event_type = 'Consulta clínica'
            title = service or 'Pendiente importado MyVete'

        existing = (
            db.query(ClinicalEvent)
            .filter(ClinicalEvent.patient_id == patient.id)
            .filter(ClinicalEvent.reminder_date == reminder_date)
            .filter(ClinicalEvent.event_type == event_type)
            .filter(ClinicalEvent.title == title)
            .first()
        )

        if existing:
            skipped += 1
            continue

        description = (
            'Importado desde pendientes MyVete\n'
            f'Fecha: {reminder_date.strftime("%d/%m/%Y")}\n'
            f'Propietario: {owner.name}\n'
            f'Paciente: {patient.name}\n'
            f'Motivo/servicio: {service}\n'
            f'Notas: {notes}'
        )

        event = ClinicalEvent(
            patient_id=patient.id,
            event_date=datetime.combine(reminder_date, datetime.min.time()),
            event_type=event_type,
            title=title,
            description=description,
            reminder_date=reminder_date,
            created_by=user.username
        )

        db.add(event)
        imported += 1

    db.commit()

    result = {
        'type': 'agenda_pendientes',
        'imported': imported,
        'skipped': skipped,
        'created_owners': created_owners,
        'created_patients': created_patients
    }

    return templates.TemplateResponse(
        'migration.html',
        {
            'request': request,
            'result': result
        }
    )

@app.post('/migration/clientes-pacientes', response_class=HTMLResponse)
async def migration_clientes_pacientes(
    request: Request,
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    import csv
    import io

    def clean_text(value):
        if value is None:
            return ''
        value = str(value).strip()
        try:
            value = value.encode('latin1').decode('utf-8')
        except Exception:
            pass
        return value

    def pick(row, names):
        normalized = {clean_text(k).lower(): clean_text(v) for k, v in row.items()}
        for name in names:
            if name.lower() in normalized:
                return normalized[name.lower()]
        return ''

    content = await file.read()
    filename = (file.filename or '').lower()

    rows = []

    if filename.endswith('.xlsx'):
        wb = load_workbook(BytesIO(content), data_only=True)
        ws = wb.active
        headers = [clean_text(c.value) for c in ws[1]]

        for row in ws.iter_rows(min_row=2, values_only=True):
            rows.append(dict(zip(headers, row)))
    else:
        text_content = content.decode('utf-8-sig', errors='replace')
        sample = text_content[:1000]
        delimiter = ';' if sample.count(';') > sample.count(',') else ','
        reader = csv.DictReader(io.StringIO(text_content), delimiter=delimiter)
        rows = list(reader)

    created_owners = 0
    created_patients = 0
    updated_owners = 0
    skipped = 0

    for row in rows:
        owner_name = pick(row, ['Propietario', 'Cliente', 'Dueño', 'Responsable', 'Nombre propietario', 'Apellido y Nombre'])
        phone = pick(row, ['Teléfono', 'Telefono', 'Celular', 'WhatsApp', 'Whatsapp'])
        email = pick(row, ['Email', 'Mail', 'Correo'])
        address = pick(row, ['Dirección', 'Direccion', 'Domicilio'])
        patient_name = pick(row, ['Paciente', 'Mascota', 'Animal', 'Nombre paciente', 'Nombre mascota'])
        species = pick(row, ['Especie'])
        breed = pick(row, ['Raza'])
        sex = pick(row, ['Sexo'])
        color = pick(row, ['Color'])
        notes = pick(row, ['Notas', 'Observaciones', 'Comentario'])

        if not owner_name and not patient_name:
            skipped += 1
            continue

        if not owner_name:
            owner_name = 'Sin propietario'

        owner = (
            db.query(Owner)
            .filter(Owner.name.ilike(owner_name))
            .first()
        )

        if not owner:
            owner = Owner(
                name=owner_name,
                phone=phone,
                whatsapp=phone,
                email=email,
                address=address,
                notes=notes
            )
            db.add(owner)
            db.flush()
            created_owners += 1
        else:
            if phone and not owner.phone:
                owner.phone = phone
            if phone and not owner.whatsapp:
                owner.whatsapp = phone
            if email and not owner.email:
                owner.email = email
            if address and not owner.address:
                owner.address = address
            updated_owners += 1

        if patient_name:
            patient = (
                db.query(Patient)
                .filter(Patient.name.ilike(patient_name))
                .filter(Patient.owner_id == owner.id)
                .first()
            )

            if not patient:
                patient = Patient(
                    name=patient_name,
                    owner_id=owner.id,
                    species=species,
                    breed=breed,
                    sex=sex,
                    color=color,
                    notes=notes
                )
                db.add(patient)
                created_patients += 1

    db.commit()

    result = {
        'type': 'clientes_pacientes',
        'created_owners': created_owners,
        'updated_owners': updated_owners,
        'created_patients': created_patients,
        'skipped': skipped
    }
@app.post('/migration/visitas', response_class=HTMLResponse)
async def migration_visitas(
    request: Request,
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    import csv
    import io

    def clean_text(value):
        if value is None:
            return ''
        value = str(value).strip()
        try:
            value = value.encode('latin1').decode('utf-8')
        except Exception:
            pass
        return value

    def pick(row, names):
        normalized = {clean_text(k).lower(): clean_text(v) for k, v in row.items()}
        for name in names:
            key = name.lower()
            if key in normalized:
                return normalized[key]
        return ''

    def parse_datetime(value):
        value = clean_text(value)
        for fmt in [
            '%d/%m/%Y %H:%M',
            '%d/%m/%Y',
            '%Y-%m-%d %H:%M',
            '%Y-%m-%d'
        ]:
            try:
                return datetime.strptime(value, fmt)
            except Exception:
                pass
        return None

    content = await file.read()
    filename = (file.filename or '').lower()

    rows = []

    if filename.endswith('.xlsx'):
        wb = load_workbook(BytesIO(content), data_only=True)
        ws = wb.active
        headers = [clean_text(c.value) for c in ws[1]]

        for row in ws.iter_rows(min_row=2, values_only=True):
            rows.append(dict(zip(headers, row)))
    else:
        text_content = content.decode('utf-8-sig', errors='replace')
        sample = text_content[:1000]
        delimiter = ';' if sample.count(';') >= sample.count(',') else ','
        reader = csv.DictReader(io.StringIO(text_content), delimiter=delimiter)
        rows = list(reader)

    imported = 0
    skipped = 0
    created_owners = 0
    created_patients = 0

    for row in rows:
        fecha = pick(row, ['Fecha'])
        owner_name = pick(row, ['Cliente', 'Propietario', 'Dueño'])
        patient_name = pick(row, ['Paciente', 'Mascota', 'Animal'])
        email = pick(row, ['Email', 'Mail'])
        usuario = pick(row, ['Usuario'])

        event_dt = parse_datetime(fecha)

        if not event_dt:
            skipped += 1
            continue

        if not owner_name:
            owner_name = 'Sin propietario'

        if not patient_name:
            patient_name = 'Sin paciente'

        owner = (
            db.query(Owner)
            .filter(Owner.name.ilike(owner_name))
            .first()
        )

        if not owner:
            owner = Owner(
                name=owner_name,
                email=email
            )
            db.add(owner)
            db.flush()
            created_owners += 1
        else:
            if email and not owner.email:
                owner.email = email

        patient = (
            db.query(Patient)
            .filter(Patient.name.ilike(patient_name))
            .filter(Patient.owner_id == owner.id)
            .first()
        )

        if not patient:
            patient = Patient(
                name=patient_name,
                owner_id=owner.id
            )
            db.add(patient)
            db.flush()
            created_patients += 1

        existing = (
            db.query(ClinicalEvent)
            .filter(ClinicalEvent.patient_id == patient.id)
            .filter(ClinicalEvent.event_date == event_dt)
            .filter(ClinicalEvent.title == 'Visita importada MyVete')
            .first()
        )

        if existing:
            skipped += 1
            continue

        description = (
            'Importado desde Reporte de visitas MyVete\n'
            f'Fecha original: {event_dt.strftime("%d/%m/%Y %H:%M")}\n'
            f'Cliente: {owner.name}\n'
            f'Paciente: {patient.name}\n'
            f'Usuario original: {usuario}'
        )

        event = ClinicalEvent(
            patient_id=patient.id,
            event_date=event_dt,
            event_type='Consulta clínica',
            title='Visita importada MyVete',
            description=description,
            created_by=user.username
        )

        db.add(event)
        imported += 1

    db.commit()

    result = {
        'type': 'visitas',
        'imported': imported,
        'created_owners': created_owners,
        'created_patients': created_patients,
        'skipped': skipped
    }

    return templates.TemplateResponse(
        'migration.html',
        {
            'request': request,
            'result': result
        }
    )
    return templates.TemplateResponse(
        'migration.html',
        {
            'request': request,
            'result': result
        }
    )
@app.post('/attachment/{attachment_id}/delete')
def delete_attachment(
    attachment_id: int,
    db: Session = Depends(get_db)
):
    attachment = db.query(EventAttachment).filter(
        EventAttachment.id == attachment_id
    ).first()

    if not attachment:
        raise HTTPException(status_code=404, detail="Adjunto no encontrado")

    patient_id = attachment.event.patient_id

    try:
        path_parts = attachment.file_path.split("/storage/v1/object/public/adjuntos/")
        if len(path_parts) > 1:
            storage_path = path_parts[1]
            supabase.storage.from_("adjuntos").remove([storage_path])
    except Exception:
        pass

    db.delete(attachment)
    db.commit()

    return RedirectResponse(f"/patients/{patient_id}", status_code=303)
@app.get('/pendientes', response_class=HTMLResponse)
def pendientes(request: Request, db: Session = Depends(get_db), user: User = Depends(require_user)):
    today = argentina_now().date()

    eventos = (
        db.query(ClinicalEvent)
        .filter(ClinicalEvent.reminder_date != None)
        .order_by(ClinicalEvent.reminder_date.asc())
        .all()
    )

    return templates.TemplateResponse(
        'pendientes.html',
       {
    'request': request,
    'eventos': eventos,
    'today': today,
    'pending_count': len(eventos)
}
    )
@app.get('/patients/{patient_id}/quick-pendiente')
def quick_pendiente(
    patient_id: int,
    type: str,
    days: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)

    if not patient:
        raise HTTPException(
            status_code=404,
            detail="Paciente no encontrado"
        )

    reminder_date = argentina_now().date() + timedelta(days=days)

    existing_event = (
        db.query(ClinicalEvent)
        .filter(ClinicalEvent.patient_id == patient.id)
        .filter(ClinicalEvent.event_type == type)
        .filter(ClinicalEvent.title == type)
        .filter(ClinicalEvent.reminder_date == reminder_date)
        .first()
    )

    if existing_event:
        return RedirectResponse(
            f"/patients/{patient.id}",
            status_code=303
        )

    event = ClinicalEvent(
        patient_id=patient.id,
        event_type=type,
        title=type,
        reminder_date=reminder_date
    )

    db.add(event)
    db.commit()

    return RedirectResponse(
        f"/patients/{patient.id}",
        status_code=303
    )
@app.post('/events/{event_id}/done')
def event_done(
    event_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    event = db.get(ClinicalEvent, event_id)

    if not event:
        raise HTTPException(status_code=404, detail="Evento no encontrado")

    original_date = event.reminder_date

    event.reminder_date = None

    if DUE_ACTIVE_MARKER not in (event.description or ''):
        event.description = (
            (event.description or '').strip()
            + '\n\n'
            + DUE_ACTIVE_MARKER
            + '\n'
            + '⚠ Pendiente gestionado administrativamente.\n'
            + 'Esto NO significa que el acto clínico se haya realizado.\n'
            + f'Gestionado por: {user.username}\n'
            + f'Fecha de gestión: {argentina_now().strftime("%d/%m/%Y %H:%M")}\n'
            + f'Fecha pendiente original: {original_date.strftime("%d/%m/%Y") if original_date else "-"}'
        ).strip()

    db.commit()

    return RedirectResponse('/pendientes', status_code=303)
@app.post('/events/{event_id}/delete')
def delete_event(
    event_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    event = db.get(ClinicalEvent, event_id)

    if not event:
        raise HTTPException(status_code=404, detail="Evento no encontrado")

    patient_id = event.patient_id

    db.delete(event)
    db.commit()

    return RedirectResponse(
        f"/patients/{patient_id}",
        status_code=303
    )
@app.get('/events/{event_id}/whatsapp')
def event_whatsapp(
    event_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    event = db.get(ClinicalEvent, event_id)

    if not event:
        raise HTTPException(status_code=404, detail="Evento no encontrado")

    patient = event.patient
    owner = patient.owner

    number = (owner.whatsapp or owner.phone or "").strip()

    if not number:
        return RedirectResponse(f"/patients/{patient.id}", status_code=303)

    number = ''.join(c for c in number if c.isdigit())

    if not number.startswith("54"):
        number = "549" + number

    if event.event_type == "Vacuna" or "vacun" in (event.title or "").lower():
        vacuna = event.title if event.title else "correspondiente"
    
        message = (
            f"Buenos días *{owner.name}*.\n\n"
            f"Te recordamos que *{patient.name}* debe recibir hoy su *{vacuna}*.\n\n"
            f"¡Los esperamos!"
        )
    
    elif event.event_type == "Desparasitación" or "despar" in (event.title or "").lower() or "antiparas" in (event.title or "").lower() or "aprax" in (event.title or "").lower() or "pipeta" in (event.title or "").lower():
            message = (
                f"Buenos días *{owner.name}*.\n\n"
                f"Te recordamos que *{patient.name}* debe desparasitarse hoy.\n\n"
                f"Es solo un comprimido o una pipeta, por lo que pueden:\n"
                f"• Traerlo a la clínica para realizar la desparasitación.\n"
                f"• Comprar la medicación y administrarla en casa.\n"
                f"• Solicitar envío a domicilio.\n\n"
                f"¿Qué preferís?"
            )
    else:
        message = (
            f"Hola *{owner.name}*.\n\n"
            f"Te recordamos que *{patient.name}* tiene pendiente: *{event.title or event.event_type}*.\n\n"
            f"Clínica Veterinaria Los Aromos."
        )

    import urllib.parse
    url = f"https://wa.me/{number}?text={urllib.parse.quote(message)}"

    return RedirectResponse(url, status_code=303)
@app.get('/patients/{patient_id}/history', response_class=HTMLResponse)
def patient_history(
    request: Request,
    patient_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)

    if not patient:
        raise HTTPException(status_code=404)

    events = (
        db.query(ClinicalEvent)
        .filter(ClinicalEvent.patient_id == patient.id)
        .order_by(ClinicalEvent.event_date.desc())
        .all()
    )

    return templates.TemplateResponse(
    'patient_history.html',
    {
        'request': request,
        'patient': patient,
        'events': events,
        'timedelta': timedelta
    }
)
@app.get('/patients/{patient_id}/print', response_class=HTMLResponse)
def patient_print(
    request: Request,
    patient_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)

    if not patient:
        raise HTTPException(status_code=404, detail="Paciente no encontrado")

    events = (
        db.query(ClinicalEvent)
        .filter(ClinicalEvent.patient_id == patient.id)
        .order_by(ClinicalEvent.event_date.desc())
        .all()
    )

    return templates.TemplateResponse(
        'patient_print.html',
        {
            'request': request,
            'patient': patient,
            'events': events
        }
    )
def calculate_anesthesia_doses(weight: float):
    try:
        weight = float(weight)
    except:
        weight = 0

    return {
        "acepromacina_mg": round(weight * 0.03, 2),
        "acepromacina_ml": round((weight * 0.03) / 1, 2),
        "dexmedetomidina_mcg_baja": round(weight * 5, 2),
        "dexmedetomidina_mcg_alta": round(weight * 10, 2),
        "dexmedetomidina_ml_baja": round((weight * 5 / 1000) / 0.5, 2),
        "dexmedetomidina_ml_alta": round((weight * 10 / 1000) / 0.5, 2),
        "ketamina_premed_mg": round(weight * 3, 2),
        "ketamina_premed_ml": round((weight * 3) / 50, 2),
        "midazolam_mg": round(weight * 0.2, 2),
        "midazolam_ml": round((weight * 0.2) / 5, 2),
        "nalbufina_mg": round(weight * 0.5, 2),
        "butorfanol_mg": round(weight * 0.3, 2),

        "propofol_ind_ml": round((weight * 3) / 10, 2),
        "ketamina_ind_ml": round((weight * 5) / 50, 2),
        "midazolam_ind_ml": round((weight * 0.2) / 5, 2),

        "meloxicam_mg": round(weight * 0.2, 2),
        "fentanilo_bolo_mcg": round(weight * 2, 2),
        "fentanilo_bolo_ml": round((weight * 2) / 50, 2),

        "flk_velocidad_ml_h": round(weight * 3, 2),
"flk_macro_gtt_min": round((weight * 3 * 20) / 60, 1),
"flk_micro_gtt_min": round((weight * 3 * 60) / 60, 1),

"flk_fentanilo_mcg_h": round(weight * 5, 2),
"flk_lidocaina_mg_h": round(weight * 0.6, 2),
"flk_ketamina_mg_h": round(weight * 0.6, 2),

"flk_fentanilo_ml_100": round(((weight * 5 * 1.6667) / 50), 2),
"flk_lidocaina_ml_100": round(((weight * 0.6 * 1.6667) / 20), 2),
"flk_ketamina_ml_100": round(((weight * 0.6 * 1.6667) / 50), 2),
"flk_total_drogas_ml_100": round(
    ((weight * 5 * 1.6667) / 50)
    + ((weight * 0.6 * 1.6667) / 20)
    + ((weight * 0.6 * 1.6667) / 50),
    2
),
"flk_ssf_ml_100": round(
    100
    - (
        ((weight * 5 * 1.6667) / 50)
        + ((weight * 0.6 * 1.6667) / 20)
        + ((weight * 0.6 * 1.6667) / 50)
    ),
    2
),
"flk_fentanilo_ml_200": round(((weight * 5 * 3.3334) / 50), 2),
"flk_lidocaina_ml_200": round(((weight * 0.6 * 3.3334) / 20), 2),
"flk_ketamina_ml_200": round(((weight * 0.6 * 3.3334) / 50), 2),
"flk_total_drogas_ml_200": round(
    ((weight * 5 * 3.3334) / 50)
    + ((weight * 0.6 * 3.3334) / 20)
    + ((weight * 0.6 * 3.3334) / 50),
    2
),
"flk_ssf_ml_200": round(
    200
    - (
        ((weight * 5 * 3.3334) / 50)
        + ((weight * 0.6 * 3.3334) / 20)
        + ((weight * 0.6 * 3.3334) / 50)
    ),
    2
),

"flk_fentanilo_ml_250": round(((weight * 5 * 4.16675) / 50), 2),
"flk_lidocaina_ml_250": round(((weight * 0.6 * 4.16675) / 20), 2),
"flk_ketamina_ml_250": round(((weight * 0.6 * 4.16675) / 50), 2),
"flk_total_drogas_ml_250": round(
    ((weight * 5 * 4.16675) / 50)
    + ((weight * 0.6 * 4.16675) / 20)
    + ((weight * 0.6 * 4.16675) / 50),
    2
),
"flk_ssf_ml_250": round(
    250
    - (
        ((weight * 5 * 4.16675) / 50)
        + ((weight * 0.6 * 4.16675) / 20)
        + ((weight * 0.6 * 4.16675) / 50)
    ),
    2
),
"rl_ml_h": round(weight * 5, 2),
"rl_macro_gtt_min": round((weight * 5 * 20) / 60, 1),
"rl_micro_gtt_min": round((weight * 5 * 60) / 60, 1),

"dobutamina_mcg_min_baja": round(weight * 5, 2),
"dobutamina_mcg_min_alta": round(weight * 10, 2),
"dobutamina_ml_h_baja": round(((weight * 5 * 60) / 1000) / 12.5, 2),
"dobutamina_ml_h_alta": round(((weight * 10 * 60) / 1000) / 12.5, 2),
"dobutamina_macro_gtt_min_baja": round(((((weight * 5 * 60) / 1000) / 12.5) * 20) / 60, 1),
"dobutamina_macro_gtt_min_alta": round(((((weight * 10 * 60) / 1000) / 12.5) * 20) / 60, 1),
"dobutamina_micro_gtt_min_baja": round(((((weight * 5 * 60) / 1000) / 12.5) * 60) / 60, 1),
"dobutamina_micro_gtt_min_alta": round(((((weight * 10 * 60) / 1000) / 12.5) * 60) / 60, 1),
"dobutamina_macro_seg_gota_baja": round(
    60 / ((((weight * 5 * 60) / 1000) / 12.5) * 20 / 60),
    0
) if weight > 0 else 0,

"dobutamina_macro_seg_gota_alta": round(
    60 / ((((weight * 10 * 60) / 1000) / 12.5) * 20 / 60),
    0
) if weight > 0 else 0,

"dobutamina_micro_seg_gota_baja": round(
    60 / ((((weight * 5 * 60) / 1000) / 12.5) * 60 / 60),
    0
) if weight > 0 else 0,

"dobutamina_micro_seg_gota_alta": round(
    60 / ((((weight * 10 * 60) / 1000) / 12.5) * 60 / 60),
    0
) if weight > 0 else 0,

"atropina_mg": round(weight * 0.04, 2),
"atropina_ml": round((weight * 0.04) / 0.5, 2),

"adrenalina_mg_baja": round(weight * 0.01, 2),
"adrenalina_mg_alta": round(weight * 0.02, 2),
"adrenalina_ml_baja": round((weight * 0.01) / 1, 2),
"adrenalina_ml_alta": round((weight * 0.02) / 1, 2),

"diazepam_mg": round(weight * 0.5, 2),
"diazepam_ml": round((weight * 0.5) / 5, 2),

"dexametasona_mg": round(weight * 0.2, 2),
"dexametasona_ml": round((weight * 0.2) / 4, 2)

}
@app.get('/patients/{patient_id}/anesthesia', response_class=HTMLResponse)
def patient_anesthesia(
    request: Request,
    patient_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)

    if not patient:
        raise HTTPException(status_code=404, detail="Paciente no encontrado")

    weight = patient.weight or 0

    try:
        weight = float(weight)
    except:
        weight = 0

    doses = calculate_anesthesia_doses(weight)

    return templates.TemplateResponse(
        'anesthesia.html',
        {
            'request': request,
            'patient': patient,
            'doses': doses,
            'user': user
        }
    )
@app.post('/patients/{patient_id}/anesthesia/recalculate', response_class=HTMLResponse)
async def patient_anesthesia_recalculate(
    request: Request,
    patient_id: int,
    weight: float = Form(...),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)

    if not patient:
        raise HTTPException(status_code=404, detail="Paciente no encontrado")

    doses = calculate_anesthesia_doses(weight)

    return templates.TemplateResponse(
        'anesthesia.html',
        {
            'request': request,
            'patient': patient,
            'doses': doses,
            'user': user,
            'recalculated': True,
            'recalculated_weight': weight
        }
    )
@app.get('/patients/{patient_id}/anesthesia/print', response_class=HTMLResponse)
async def patient_anesthesia_print(
    request: Request,
    patient_id: int,
    event_id: int | None = None,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)

    if not patient:
        raise HTTPException(status_code=404, detail="Paciente no encontrado")
    last_anesthesia = (
        db.query(ClinicalEvent)
        .filter(ClinicalEvent.patient_id == patient.id)
        .filter(ClinicalEvent.event_type == 'Anestesia')
        .order_by(ClinicalEvent.event_date.desc())
        .first()
    )
    selected_anesthesia = last_anesthesia

    if event_id:
        selected_anesthesia = db.get(ClinicalEvent, event_id)

        if (
            not selected_anesthesia
            or selected_anesthesia.patient_id != patient.id
            or selected_anesthesia.event_type != 'Anestesia'
        ):
            raise HTTPException(status_code=404, detail="Protocolo anestésico no encontrado")
    weight = patient.weight or 0

    try:
        weight = float(weight)
    except:
        weight = 0

    doses = calculate_anesthesia_doses(weight)
    return templates.TemplateResponse(
        'anesthesia_print.html',
        {
            'request': request,
            'patient': patient,
            'doses': doses,
            'last_anesthesia': selected_anesthesia,
            'user': user
        }
    )    
@app.post('/patients/{patient_id}/anesthesia')
async def save_patient_anesthesia(
    request: Request,
    patient_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)

    if not patient:
        raise HTTPException(status_code=404, detail="Paciente no encontrado")

    form = await request.form()

    weight = form.get('weight', '')

    if form.get('update_patient_weight') and weight:
        patient.weight = float(weight)
    procedure = form.get('procedure', '')
    asa = form.get('asa', '')
    estimated_duration = form.get('estimated_duration', '')
    drip_set = form.get('drip_set', '')
    fluid_rate = form.get('fluid_rate', '')
    
    premedication = []

    if form.get('premed_acepromazine'):
        premedication.append('Acepromacina')

    if form.get('premed_dexmedetomidine'):
        premedication.append('Dexmedetomidina')

    if form.get('premed_ketamine'):
        premedication.append('Ketamina IM')

    if form.get('premed_midazolam'):
        premedication.append('Midazolam')

    if form.get('premed_nalbuphine'):
        premedication.append('Nalbufina')

    if form.get('premed_butorphanol'):
        premedication.append('Butorfanol')


    induction = []

    if form.get('ind_propofol'):
        induction.append('Propofol')

    if form.get('ind_ketamine'):
        induction.append('Ketamina')

    if form.get('ind_midazolam'):
        induction.append('Midazolam') 
    maintenance = []

    if form.get('maint_isoflurane'):
        maintenance.append('Isofluorano')

    if form.get('maint_propofol'):
        maintenance.append('Propofol CRI')

    if form.get('maint_ketamine'):
        maintenance.append('Ketamina CRI')

    if form.get('maint_flk'):
        maintenance.append('FLK')


    intraop_analgesia = []

    if form.get('intra_fentanyl'):
        intraop_analgesia.append('Fentanilo')

    if form.get('intra_lidocaine'):
        intraop_analgesia.append('Lidocaína')

    if form.get('intra_ketamine'):
        intraop_analgesia.append('Ketamina')

    if form.get('intra_dobutamine'):
        intraop_analgesia.append('Dobutamina')
    flk_detail = []

    if form.get('maint_flk') or form.get('intra_fentanyl') or form.get('intra_lidocaine') or form.get('intra_ketamine'):
        flk_detail.append('FLK en 100 ml SSF')
        flk_detail.append('Fentanilo + Lidocaína + Ketamina')
        flk_detail.append('Velocidad según cálculo automático')
    postop_analgesia = []
    if form.get('meloxicam'):
        postop_analgesia.append('Meloxicam')
    if form.get('dipyrone'):
        postop_analgesia.append('Dipirona')
    if form.get('tramadol'):
        postop_analgesia.append('Tramadol')
    if form.get('gabapentin'):
        postop_analgesia.append('Gabapentina')
    if form.get('pregabalin'):
        postop_analgesia.append('Pregabalina')
    if form.get('nalbuphine'):
        postop_analgesia.append('Nalbufina')
    if form.get('butorphanol'):
        postop_analgesia.append('Butorfanol')
    if form.get('dexamethasone'):
        postop_analgesia.append('Dexametasona')

    antibiotics = []
    if form.get('cefazolin'):
        antibiotics.append('Cefazolina')
    if form.get('amoxiclav'):
        antibiotics.append('Amoxicilina clavulánico')
    if form.get('cephalexin'):
        antibiotics.append('Cefalexina')
    if form.get('enrofloxacin'):
        antibiotics.append('Enrofloxacina')
    if form.get('metronidazole'):
        antibiotics.append('Metronidazol')
    if form.get('clindamycin'):
        antibiotics.append('Clindamicina')

    description = f"""
Protocolo anestésico

Procedimiento: {procedure}
Peso: {weight} kg
ASA: {asa}
Duración estimada: {estimated_duration}
Premedicación:
{', '.join(premedication) if premedication else 'No especificada'}

Inducción:
{', '.join(induction) if induction else 'No especificada'}
Mantenimiento:
{', '.join(maintenance) if maintenance else 'No especificado'}

Analgesia intraoperatoria / soporte:
{', '.join(intraop_analgesia) if intraop_analgesia else 'No especificada'}
Detalle FLK:
{chr(10).join(flk_detail) if flk_detail else 'No especificado'}
Analgesia posoperatoria:
{', '.join(postop_analgesia) if postop_analgesia else 'No especificada'}

Observaciones analgesia:
{form.get('postop_notes', '')}

Antibioticoterapia:
{', '.join(antibiotics) if antibiotics else 'No especificada'}

Observaciones antibioticoterapia:
{form.get('antibiotic_notes', '')}

Fluidoterapia:
Equipo: {drip_set}
Velocidad: {fluid_rate} ml/kg/h
"""

    event = ClinicalEvent(
        patient_id=patient.id,
        event_type='Anestesia',
        description=description,
        event_date=argentina_now()
    )

    db.add(event)
    db.commit()

    return RedirectResponse(
            url=f"/patients/{patient.id}",
            status_code=303
    )
@app.get('/stats', response_class=HTMLResponse)
def stats_page(
    request: Request,
    period: str = 'month',
    start: str = '',
    end: str = '',
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    from collections import defaultdict
    import calendar

    today = argentina_now().date()

    if start and end:
        try:
            start_date = datetime.strptime(start, '%Y-%m-%d').date()
            end_date = datetime.strptime(end, '%Y-%m-%d').date()
            period = 'custom'
        except ValueError:
            start_date = today.replace(day=1)
            end_date = today
            period = 'month'
    elif period == 'today':
        start_date = today
        end_date = today
    elif period == 'week':
        start_date = today - timedelta(days=today.weekday())
        end_date = today
    elif period == 'year':
        start_date = today.replace(month=1, day=1)
        end_date = today
    else:
        start_date = today.replace(day=1)
        end_date = today
        period = 'month'

    start_dt = datetime.combine(start_date, datetime.min.time())
    end_dt = datetime.combine(end_date, datetime.max.time())

    sales = (
        db.query(Sale)
        .filter(Sale.date >= start_dt, Sale.date <= end_dt)
        .all()
    )

    sales_total = sum(s.total or 0 for s in sales if s.status != 'cancelled')
    sales_count = len([s for s in sales if s.status != 'cancelled'])
    profit_total = sum(s.profit_amount or 0 for s in sales if s.status != 'cancelled')
    ticket_average = (sales_total / sales_count) if sales_count else 0

    sale_ids = [s.id for s in sales if s.status != 'cancelled']

    payments_by_method = defaultdict(float)
    account_pending = 0
    clients_with_debt = set()

    if sale_ids:
        payments = db.query(SalePayment).filter(SalePayment.sale_id.in_(sale_ids)).all()
        for payment in payments:
            method = payment.method or 'Sin método'
            payments_by_method[method] += payment.amount or 0

        for sale in sales:
            paid = sum(
                p.amount or 0
                for p in db.query(SalePayment).filter(SalePayment.sale_id == sale.id).all()
                if 'cuenta' not in (p.method or '').lower()
            )
            balance = (sale.total or 0) - paid
            if balance > 0:
                account_pending += balance
                if sale.owner_id:
                    clients_with_debt.add(sale.owner_id)

    top_products = []
    top_profit_products = []
    product_stats = {}

    if sale_ids:
        items = db.query(SaleItem).filter(SaleItem.sale_id.in_(sale_ids)).all()

        for item in items:
            product = db.get(Product, item.product_id)
            name = product.name if product else 'Producto eliminado'
            cost = product.cost_price or 0 if product else 0
            qty = item.quantity or 0
            subtotal = item.subtotal or 0
            profit = subtotal - (qty * cost)

            if name not in product_stats:
                product_stats[name] = {'qty': 0, 'profit': 0}

            product_stats[name]['qty'] += qty
            product_stats[name]['profit'] += profit

        top_products = sorted(
            [{'name': k, 'qty': v['qty']} for k, v in product_stats.items()],
            key=lambda x: x['qty'],
            reverse=True
        )[:10]

        top_profit_products = sorted(
            [{'name': k, 'profit': v['profit']} for k, v in product_stats.items()],
            key=lambda x: x['profit'],
            reverse=True
        )[:10]

    events = (
        db.query(ClinicalEvent)
        .filter(ClinicalEvent.event_date >= start_dt, ClinicalEvent.event_date <= end_dt)
        .all()
    )

    event_counts = defaultdict(int)
    patient_visit_counts = defaultdict(int)
    owner_visit_counts = defaultdict(int)

    for event in events:
        event_counts[event.event_type or 'Sin tipo'] += 1
        if event.patient:
            patient_visit_counts[event.patient.name] += 1
            if event.patient.owner:
                owner_visit_counts[event.patient.owner.name] += 1

    event_type_cards = [
        {'label': 'Consultas', 'value': event_counts.get('Consulta clínica', 0) + event_counts.get('Control', 0), 'icon': '🩺'},
        {'label': 'Vacunas', 'value': event_counts.get('Vacuna', 0), 'icon': '💉'},
        {'label': 'Desparasitaciones', 'value': event_counts.get('Desparasitación', 0), 'icon': '🪱'},
        {'label': 'ECG', 'value': event_counts.get('ECG', 0), 'icon': '💗'},
        {'label': 'Radiografías', 'value': event_counts.get('Radiografía', 0), 'icon': '📷'},
        {'label': 'Ecografías', 'value': event_counts.get('Ecografía', 0) + event_counts.get('Ecocardiografía', 0), 'icon': '🖥️'},
        {'label': 'Cirugías', 'value': event_counts.get('Cirugía', 0), 'icon': '🔪'},
    ]

    top_patients = sorted(
        [{'name': k, 'visits': v} for k, v in patient_visit_counts.items()],
        key=lambda x: x['visits'],
        reverse=True
    )[:5]

    top_owners = sorted(
        [{'name': k, 'visits': v} for k, v in owner_visit_counts.items()],
        key=lambda x: x['visits'],
        reverse=True
    )[:5]

    products = db.query(Product).filter(Product.active == True).all()

    critical_stock = []
    negative_stock = []
    expiring_products = []
    stock_value_sale = 0
    stock_value_cost = 0

    soon = today + timedelta(days=60)

    for product in products:
        stock = product.stock or 0
        stock_value_sale += stock * (product.sale_price or 0)
        stock_value_cost += stock * (product.cost_price or 0)

        if product.stock is not None and product.stock < 0:
            negative_stock.append(product)

        if (
            product.stock is not None and
            product.min_stock is not None and
            product.min_stock > 0 and
            product.stock <= product.min_stock
        ):
            critical_stock.append(product)

        if product.expiration_date and product.expiration_date <= soon:
            expiring_products.append(product)

    appointments = (
        db.query(Appointment)
        .filter(Appointment.appointment_date >= start_dt, Appointment.appointment_date <= end_dt)
        .all()
    )

    appointment_stats = {
        'total': len(appointments),
        'confirmed': len([a for a in appointments if a.status == 'Confirmado']),
        'done': len([a for a in appointments if a.status == 'Realizado']),
        'cancelled': len([a for a in appointments if a.status == 'Cancelado']),
        'pending': len([a for a in appointments if a.status == 'Pendiente']),
    }

    waiting_active = db.query(WaitingListEntry).filter(
        WaitingListEntry.status.in_(['Esperando', 'En consulta'])
    ).count()

    new_patients_estimated = len(set(e.patient_id for e in events if e.patient_id))
    patients_seen = len(set(e.patient_id for e in events if e.patient_id))

    alerts = []
    if critical_stock:
        alerts.append(f'Tenés {len(critical_stock)} productos en stock crítico.')
    if expiring_products:
        alerts.append(f'Hay {len(expiring_products)} productos próximos a vencer.')
    if account_pending > 0:
        alerts.append(f'La cuenta corriente pendiente suma ${account_pending:,.0f}.')
    if waiting_active > 0:
        alerts.append(f'Hay {waiting_active} pacientes activos en lista de espera.')
    if not alerts:
        alerts.append('Todo se ve ordenado para el período seleccionado.')

    chart_days = []
    current = start_date
    while current <= end_date:
        day_start = datetime.combine(current, datetime.min.time())
        day_end = datetime.combine(current, datetime.max.time())
        day_total = sum(
            s.total or 0
            for s in sales
            if s.date and day_start <= s.date <= day_end and s.status != 'cancelled'
        )
        chart_days.append({
            'label': current.strftime('%d/%m'),
            'value': day_total
        })
        current += timedelta(days=1)

    max_chart_value = max([d['value'] for d in chart_days], default=1) or 1

    return templates.TemplateResponse(
        'stats.html',
        {
            'request': request,
            'period': period,
            'start_date': start_date,
            'end_date': end_date,
            'sales_total': sales_total,
            'sales_count': sales_count,
            'profit_total': profit_total,
            'ticket_average': ticket_average,
            'patients_seen': patients_seen,
            'new_patients_estimated': new_patients_estimated,
            'payments_by_method': dict(payments_by_method),
            'account_pending': account_pending,
            'clients_with_debt': len(clients_with_debt),
            'top_products': top_products,
            'top_profit_products': top_profit_products,
            'event_type_cards': event_type_cards,
            'critical_stock': critical_stock[:5],
            'negative_stock_count': len(negative_stock),
            'critical_stock_count': len(critical_stock),
            'expiring_products': sorted(expiring_products, key=lambda p: p.expiration_date or today)[:5],
            'expiring_count': len(expiring_products),
            'stock_value_sale': stock_value_sale,
            'stock_value_cost': stock_value_cost,
            'appointment_stats': appointment_stats,
            'waiting_active': waiting_active,
            'top_patients': top_patients,
            'top_owners': top_owners,
            'alerts': alerts,
            'chart_days': chart_days,
            'max_chart_value': max_chart_value,
        }
    )
# ===== VADEMÉCUM =====

@app.get('/vademecum', response_class=HTMLResponse)
def vademecum_page(
    request: Request,
    q: str = "",
    category: str = "",
    species: str = "",
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    active_ingredients = db.execute(
            text("""
                SELECT DISTINCT a.*
                FROM vademecum_active_ingredients a
                LEFT JOIN vademecum_brands b
                    ON b.active_ingredient_id = a.id
                WHERE a.active = TRUE
                AND (
                    :q = ''
                    OR a.name ILIKE :like
                    OR a.category ILIKE :like
                    OR a.species ILIKE :like
                    OR a.indications ILIKE :like
                    OR a.contraindications ILIKE :like
                    OR a.observations ILIKE :like
                    OR b.brand_name ILIKE :like
                    OR b.laboratory ILIKE :like
                )
                ORDER BY a.name
            """),
            {
                "q": q,
                "like": f"%{q}%"
            }
        ).mappings().all()
    
    brands = db.execute(
            text("""
                SELECT
                    b.*,
                    a.name AS active_name
                FROM vademecum_brands b
                JOIN vademecum_active_ingredients a
                    ON a.id = b.active_ingredient_id
                WHERE b.active = TRUE
                ORDER BY a.name, b.brand_name
            """)
        ).mappings().all()
    
    return templates.TemplateResponse(
            "vademecum.html",
            {
                "request": request,
                "active_ingredients": active_ingredients,
                "brands": brands,
                "q": q,
                "category": category,
                "species": species
            }
        )
@app.get('/vademecum/{active_id}', response_class=HTMLResponse)
def vademecum_detail(
    active_id: int,
    request: Request,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    active = db.execute(
        text("""
            SELECT *
            FROM vademecum_active_ingredients
            WHERE id = :active_id AND active = TRUE
        """),
        {"active_id": active_id}
    ).mappings().first()

    if not active:
        raise HTTPException(status_code=404, detail="Principio activo no encontrado")

    brands = db.execute(
        text("""
            SELECT *
            FROM vademecum_brands
            WHERE active_ingredient_id = :active_id
            AND active = TRUE
            ORDER BY brand_name
        """),
        {"active_id": active_id}
    ).mappings().all()

    return templates.TemplateResponse(
        "vademecum_detail.html",
        {
            "request": request,
            "active": active,
            "brands": brands
        }
    )
@app.get("/vademecum/{active_id}/edit", response_class=HTMLResponse)
def edit_vademecum(
    active_id: int,
    request: Request,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    active = db.execute(
        text("""
            SELECT *
            FROM vademecum_active_ingredients
            WHERE id = :id
        """),
        {"id": active_id}
    ).mappings().first()

    if not active:
        raise HTTPException(status_code=404, detail="Principio activo no encontrado")

    return templates.TemplateResponse(
        "vademecum_edit.html",
        {
            "request": request,
            "active": active
        }
    )


@app.post("/vademecum/{active_id}/edit")
def save_vademecum(
    active_id: int,
    name: str = Form(...),
    category: str = Form(""),
    species: str = Form(""),
    dog_dose: str = Form(""),
    cat_dose: str = Form(""),
    route: str = Form(""),
    frequency: str = Form(""),
    indications: str = Form(""),
    contraindications: str = Form(""),
    interactions: str = Form(""),
    warnings: str = Form(""),
    observations: str = Form(""),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    db.execute(
        text("""
            UPDATE vademecum_active_ingredients
            SET
                name=:name,
                category=:category,
                species=:species,
                dog_dose=:dog_dose,
                cat_dose=:cat_dose,
                route=:route,
                frequency=:frequency,
                indications=:indications,
                contraindications=:contraindications,
                interactions=:interactions,
                warnings=:warnings,
                observations=:observations
            WHERE id=:id
        """),
        {
            "id": active_id,
            "name": name,
            "category": category,
            "species": species,
            "dog_dose": dog_dose,
            "cat_dose": cat_dose,
            "route": route,
            "frequency": frequency,
            "indications": indications,
            "contraindications": contraindications,
            "interactions": interactions,
            "warnings": warnings,
            "observations": observations
        }
    )

    db.commit()

    return RedirectResponse(
        f"/vademecum/{active_id}",
        status_code=303
    )
@app.post("/vademecum/{active_id}/delete")
def delete_vademecum(
    active_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    db.execute(
        text("""
            UPDATE vademecum_active_ingredients
            SET active = FALSE
            WHERE id = :id
        """),
        {"id": active_id}
    )

    db.execute(
        text("""
            UPDATE vademecum_brands
            SET active = FALSE
            WHERE active_ingredient_id = :id
        """),
        {"id": active_id}
    )

    db.commit()

    return RedirectResponse(
        "/vademecum",
        status_code=303
    )
@app.post('/vademecum')
def vademecum_create(
    commercial_name: str = Form(''),
    active_ingredient: str = Form(''),
    category: str = Form(''),
    species: str = Form(''),
    dose: str = Form(''),
    route: str = Form(''),
    frequency: str = Form(''),
    indications: str = Form(''),
    contraindications: str = Form(''),
    observations: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    db.execute(
        text("""
            INSERT INTO vademecum_drugs (
                commercial_name,
                active_ingredient,
                category,
                species,
                dose,
                route,
                frequency,
                indications,
                contraindications,
                observations,
                active
            )
            VALUES (
                :commercial_name,
                :active_ingredient,
                :category,
                :species,
                :dose,
                :route,
                :frequency,
                :indications,
                :contraindications,
                :observations,
                TRUE
            )
        """),
        {
            "commercial_name": commercial_name,
            "active_ingredient": active_ingredient,
            "category": category,
            "species": species,
            "dose": dose,
            "route": route,
            "frequency": frequency,
            "indications": indications,
            "contraindications": contraindications,
            "observations": observations
        }
    )

    db.commit()

    return RedirectResponse('/vademecum', status_code=303)


@app.post('/vademecum/{drug_id}/delete')
def vademecum_delete(
    drug_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    db.execute(
        text("""
            UPDATE vademecum_drugs
            SET active = FALSE
            WHERE id = :drug_id
        """),
        {"drug_id": drug_id}
    )

    db.commit()

    return RedirectResponse('/vademecum', status_code=303)
@app.post('/vademecum/active-ingredient')
def vademecum_active_ingredient_create(
    name: str = Form(''),
    category: str = Form(''),
    species: str = Form(''),
    dog_dose: str = Form(''),
    cat_dose: str = Form(''),
    route: str = Form(''),
    frequency: str = Form(''),
    indications: str = Form(''),
    contraindications: str = Form(''),
    interactions: str = Form(''),
    warnings: str = Form(''),
    observations: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    db.execute(
        text("""
            INSERT INTO vademecum_active_ingredients (
                name, category, species, dog_dose, cat_dose, route,
                frequency, indications, contraindications, interactions,
                warnings, observations, active
            )
            VALUES (
                :name, :category, :species, :dog_dose, :cat_dose, :route,
                :frequency, :indications, :contraindications, :interactions,
                :warnings, :observations, TRUE
            )
        """),
        {
            "name": name,
            "category": category,
            "species": species,
            "dog_dose": dog_dose,
            "cat_dose": cat_dose,
            "route": route,
            "frequency": frequency,
            "indications": indications,
            "contraindications": contraindications,
            "interactions": interactions,
            "warnings": warnings,
            "observations": observations
        }
    )

    db.commit()

    return RedirectResponse('/vademecum', status_code=303)


@app.post('/vademecum/brand')
def vademecum_brand_create(
    active_ingredient_id: str = Form(''),
    brand_name: str = Form(''),
    laboratory: str = Form(''),
    presentation: str = Form(''),
    concentration: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    if not active_ingredient_id:
        return RedirectResponse('/vademecum', status_code=303)

    db.execute(
        text("""
            INSERT INTO vademecum_brands (
                active_ingredient_id, brand_name, laboratory,
                presentation, concentration, active
            )
            VALUES (
                :active_ingredient_id, :brand_name, :laboratory,
                :presentation, :concentration, TRUE
            )
        """),
        {
            "active_ingredient_id": int(active_ingredient_id),
            "brand_name": brand_name,
            "laboratory": laboratory,
            "presentation": presentation,
            "concentration": concentration
        }
    )

    db.commit()

    return RedirectResponse('/vademecum', status_code=303)
@app.get("/vademecum/brand/{brand_id}/edit", response_class=HTMLResponse)
def edit_vademecum_brand(
    brand_id: int,
    request: Request,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    brand = db.execute(
        text("""
            SELECT *
            FROM vademecum_brands
            WHERE id = :brand_id
            AND active = TRUE
        """),
        {"brand_id": brand_id}
    ).mappings().first()

    if not brand:
        raise HTTPException(status_code=404, detail="Marca comercial no encontrada")

    return templates.TemplateResponse(
        "vademecum_brand_edit.html",
        {
            "request": request,
            "brand": brand
        }
    )


@app.post("/vademecum/brand/{brand_id}/edit")
def save_vademecum_brand(
    brand_id: int,
    brand_name: str = Form(""),
    laboratory: str = Form(""),
    presentation: str = Form(""),
    concentration: str = Form(""),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    brand = db.execute(
        text("""
            SELECT active_ingredient_id
            FROM vademecum_brands
            WHERE id = :brand_id
        """),
        {"brand_id": brand_id}
    ).mappings().first()

    if not brand:
        raise HTTPException(status_code=404, detail="Marca comercial no encontrada")

    db.execute(
        text("""
            UPDATE vademecum_brands
            SET
                brand_name = :brand_name,
                laboratory = :laboratory,
                presentation = :presentation,
                concentration = :concentration
            WHERE id = :brand_id
        """),
        {
            "brand_id": brand_id,
            "brand_name": brand_name,
            "laboratory": laboratory,
            "presentation": presentation,
            "concentration": concentration
        }
    )

    db.commit()

    return RedirectResponse(
        f"/vademecum/{brand.active_ingredient_id}",
        status_code=303
    )


@app.post("/vademecum/brand/{brand_id}/delete")
def delete_vademecum_brand(
    brand_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    brand = db.execute(
        text("""
            SELECT active_ingredient_id
            FROM vademecum_brands
            WHERE id = :brand_id
        """),
        {"brand_id": brand_id}
    ).mappings().first()

    if not brand:
        raise HTTPException(status_code=404, detail="Marca comercial no encontrada")

    db.execute(
        text("""
            UPDATE vademecum_brands
            SET active = FALSE
            WHERE id = :brand_id
        """),
        {"brand_id": brand_id}
    )

    db.commit()

    return RedirectResponse(
        f"/vademecum/{brand.active_ingredient_id}",
        status_code=303
    )
@app.get("/vademecum/{active_id}/print", response_class=HTMLResponse)
def vademecum_print(
    active_id: int,
    request: Request,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    active = db.execute(
        text("""
            SELECT *
            FROM vademecum_active_ingredients
            WHERE id = :active_id
            AND active = TRUE
        """),
        {"active_id": active_id}
    ).mappings().first()

    if not active:
        raise HTTPException(status_code=404, detail="Principio activo no encontrado")

    brands = db.execute(
        text("""
            SELECT *
            FROM vademecum_brands
            WHERE active_ingredient_id = :active_id
            AND active = TRUE
            ORDER BY brand_name
        """),
        {"active_id": active_id}
    ).mappings().all()

    return templates.TemplateResponse(
        "vademecum_print.html",
        {
            "request": request,
            "active": active,
            "brands": brands
        }
    )
# ===== INTERNACIÓN =====

@app.get('/hospitalizations', response_class=HTMLResponse)
def hospitalizations_page(
    request: Request,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    active_hospitalizations = (
        db.query(Hospitalization)
        .filter(Hospitalization.status == 'Internado')
        .order_by(Hospitalization.admission_date.desc())
        .all()
    )

    closed_hospitalizations = (
        db.query(Hospitalization)
        .filter(Hospitalization.status != 'Internado')
        .order_by(Hospitalization.admission_date.desc())
        .limit(20)
        .all()
    )

    return templates.TemplateResponse(
        'hospitalizations.html',
        {
            'request': request,
            'active_hospitalizations': active_hospitalizations,
            'closed_hospitalizations': closed_hospitalizations,
            'today': argentina_now().date()
        }
    )


@app.get('/patients/{patient_id}/hospitalize', response_class=HTMLResponse)
def hospitalization_form(
    request: Request,
    patient_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)

    if not patient:
        raise HTTPException(status_code=404, detail='Paciente no encontrado')

    active_hospitalization = (
        db.query(Hospitalization)
        .filter(
            Hospitalization.patient_id == patient.id,
            Hospitalization.status == 'Internado'
        )
        .first()
    )

    if active_hospitalization:
        return RedirectResponse(
            f'/hospitalizations/{active_hospitalization.id}',
            status_code=303
        )

    return templates.TemplateResponse(
        'hospitalization_form.html',
        {
            'request': request,
            'patient': patient,
            'today': argentina_now().date()
        }
    )


@app.post('/patients/{patient_id}/hospitalize')
def hospitalization_create(
    patient_id: int,
    cage: str = Form(''),
    responsible_vet: str = Form(''),
    reason: str = Form(''),
    diagnosis: str = Form(''),
    treatment_plan: str = Form(''),
    notes: str = Form(''),
    initial_weight: str = Form(''),
    initial_temperature: str = Form(''),
    initial_heart_rate: str = Form(''),
    initial_respiratory_rate: str = Form(''),
    initial_mucous_membranes: str = Form(''),
    initial_crt: str = Form(''),
    initial_hydration: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    patient = db.get(Patient, patient_id)

    if not patient:
        raise HTTPException(status_code=404, detail='Paciente no encontrado')

    active_hospitalization = (
        db.query(Hospitalization)
        .filter(
            Hospitalization.patient_id == patient.id,
            Hospitalization.status == 'Internado'
        )
        .first()
    )

    if active_hospitalization:
        return RedirectResponse(
            f'/hospitalizations/{active_hospitalization.id}',
            status_code=303
        )

    def to_float(value):
        try:
            return float(str(value).replace(',', '.')) if value and str(value).strip() else None
        except ValueError:
            return None

    def to_int(value):
        try:
            return int(float(str(value).replace(',', '.'))) if value and str(value).strip() else None
        except ValueError:
            return None

    event = ClinicalEvent(
        patient_id=patient.id,
        event_type='Internación',
        title='Ingreso a internación',
        description=reason or '',
        diagnosis=diagnosis or '',
        treatment=treatment_plan or '',
        weight=to_float(initial_weight),
        temperature=to_float(initial_temperature),
        heart_rate=to_int(initial_heart_rate),
        respiratory_rate=to_int(initial_respiratory_rate),
        mucous_membranes=initial_mucous_membranes or '',
        crt=initial_crt or '',
        hydration=initial_hydration or '',
        created_by=user.username,
        event_date=argentina_now()
    )

    db.add(event)
    db.flush()

    hospitalization = Hospitalization(
        patient_id=patient.id,
        clinical_event_id=event.id,
        cage=cage or '',
        responsible_vet=responsible_vet or '',
        reason=reason or '',
        diagnosis=diagnosis or '',
        treatment_plan=treatment_plan or '',
        notes=notes or '',
        initial_weight=to_float(initial_weight),
        initial_temperature=to_float(initial_temperature),
        initial_heart_rate=to_int(initial_heart_rate),
        initial_respiratory_rate=to_int(initial_respiratory_rate),
        initial_mucous_membranes=initial_mucous_membranes or '',
        initial_crt=initial_crt or '',
        initial_hydration=initial_hydration or '',
        created_by=user.username,
        admission_date=argentina_now(),
        status='Internado'
    )

    db.add(hospitalization)

    if initial_weight and str(initial_weight).strip():
        patient.weight = to_float(initial_weight)

    db.commit()
    db.refresh(hospitalization)

    return RedirectResponse(
        f'/hospitalizations/{hospitalization.id}',
        status_code=303
    )
def build_hospitalization_ai(hospitalization, patient, related_events, medications, fluids):
    now = argentina_now().replace(tzinfo=None)

    def to_float(value):
        try:
            if value is None:
                return None
            value = str(value).replace(',', '.').strip()
            if not value:
                return None
            return float(value)
        except Exception:
            return None

    def hours_from(dt):
        if not dt:
            return None
        return (now - dt.replace(tzinfo=None)).total_seconds() / 3600

    def clean_text(value):
        return (value or '').strip().lower()

    def med_minutes(time_text):
        try:
            if not time_text:
                return None
            h, m = str(time_text).split(':')[:2]
            return int(h) * 60 + int(m)
        except Exception:
            return None

    alerts = []
    score = 0

    vitals_events = [
        e for e in related_events
        if e.temperature or e.heart_rate or e.respiratory_rate or e.weight or e.mucous_membranes or e.crt or e.hydration
    ]

    last_control = related_events[0] if related_events else None
    last_vital = vitals_events[0] if vitals_events else None
    previous_vital = vitals_events[1] if len(vitals_events) > 1 else None

    temp = to_float(getattr(last_vital, 'temperature', None)) if last_vital else to_float(hospitalization.initial_temperature)
    hr = to_float(getattr(last_vital, 'heart_rate', None)) if last_vital else to_float(hospitalization.initial_heart_rate)
    rr = to_float(getattr(last_vital, 'respiratory_rate', None)) if last_vital else to_float(hospitalization.initial_respiratory_rate)
    weight = to_float(getattr(last_vital, 'weight', None)) if last_vital else to_float(hospitalization.initial_weight or patient.weight)

    prev_temp = to_float(getattr(previous_vital, 'temperature', None)) if previous_vital else None
    prev_hr = to_float(getattr(previous_vital, 'heart_rate', None)) if previous_vital else None
    prev_rr = to_float(getattr(previous_vital, 'respiratory_rate', None)) if previous_vital else None
    prev_weight = to_float(getattr(previous_vital, 'weight', None)) if previous_vital else None

    mucosas = (getattr(last_vital, 'mucous_membranes', '') if last_vital else hospitalization.initial_mucous_membranes) or ''
    crt = (getattr(last_vital, 'crt', '') if last_vital else hospitalization.initial_crt) or ''
    hidratacion = (getattr(last_vital, 'hydration', '') if last_vital else hospitalization.initial_hydration) or ''

    clinical_text = f"{mucosas} {crt} {hidratacion}".lower()
    all_recent_text = " ".join([clean_text(e.description) for e in related_events[:8]])

    danger_words = ['pálida', 'palida', 'cianótica', 'cianotica', 'gris', 'shock', 'colapso', 'débil', 'debil', 'estupor', 'convuls']
    warning_words = ['ictérica', 'icterica', 'congestiva', 'seca', 'deshidrat', 'pegajosa', 'deca', 'dolor', 'hiporexia']

    if temp is not None:
        if temp >= 40 or temp <= 36.5:
            score += 35
            alerts.append({'level': 'danger', 'icon': '🌡', 'title': 'Temperatura crítica', 'detail': f'Última temperatura: {temp:g} °C.'})
        elif temp >= 39.5 or temp <= 37:
            score += 22
            alerts.append({'level': 'warning', 'icon': '🌡', 'title': 'Temperatura fuera de rango', 'detail': f'Última temperatura: {temp:g} °C.'})

    if prev_temp is not None and temp is not None:
        delta = temp - prev_temp
        if abs(delta) >= 1:
            score += 12
            alerts.append({'level': 'warning', 'icon': '📈', 'title': 'Cambio térmico importante', 'detail': f'Variación de temperatura: {delta:+.1f} °C respecto al control previo.'})

    if hr is not None:
        if hr >= 200 or hr <= 50:
            score += 30
            alerts.append({'level': 'danger', 'icon': '❤️', 'title': 'Frecuencia cardíaca crítica', 'detail': f'Última FC: {hr:g} lpm.'})
        elif hr >= 180 or hr <= 60:
            score += 18
            alerts.append({'level': 'warning', 'icon': '❤️', 'title': 'Frecuencia cardíaca alterada', 'detail': f'Última FC: {hr:g} lpm.'})

    if prev_hr is not None and hr is not None and abs(hr - prev_hr) >= 35:
        score += 10
        alerts.append({'level': 'warning', 'icon': '📈', 'title': 'Cambio de FC', 'detail': f'FC varió {hr - prev_hr:+.0f} lpm respecto al control previo.'})

    if rr is not None:
        if rr >= 80 or rr <= 8:
            score += 30
            alerts.append({'level': 'danger', 'icon': '🫁', 'title': 'Frecuencia respiratoria crítica', 'detail': f'Última FR: {rr:g} rpm.'})
        elif rr >= 60 or rr <= 10:
            score += 18
            alerts.append({'level': 'warning', 'icon': '🫁', 'title': 'Frecuencia respiratoria alterada', 'detail': f'Última FR: {rr:g} rpm.'})

    if prev_rr is not None and rr is not None and abs(rr - prev_rr) >= 20:
        score += 10
        alerts.append({'level': 'warning', 'icon': '📈', 'title': 'Cambio de FR', 'detail': f'FR varió {rr - prev_rr:+.0f} rpm respecto al control previo.'})

    if prev_weight is not None and weight is not None and prev_weight > 0:
        weight_delta_percent = ((weight - prev_weight) / prev_weight) * 100
        if abs(weight_delta_percent) >= 5:
            score += 12
            alerts.append({'level': 'warning', 'icon': '⚖', 'title': 'Cambio de peso', 'detail': f'Peso varió {weight_delta_percent:+.1f}% respecto al control previo.'})

    if any(word in clinical_text for word in danger_words):
        score += 30
        alerts.append({'level': 'danger', 'icon': '👄', 'title': 'Perfusión preocupante', 'detail': 'Mucosas/TRC/hidratación con términos de riesgo.'})
    elif any(word in clinical_text for word in warning_words):
        score += 16
        alerts.append({'level': 'warning', 'icon': '👄', 'title': 'Perfusión a revisar', 'detail': 'Mucosas o hidratación requieren control.'})

    if 'vómitos / diarrea: sí' in all_recent_text or 'vomitos / diarrea: sí' in all_recent_text:
        score += 18
        alerts.append({'level': 'warning', 'icon': '🤢', 'title': 'Vómitos o diarrea registrados', 'detail': 'Checklist reciente con vómitos/diarrea positivos.'})

    if 'comió: no' in all_recent_text or 'comio: no' in all_recent_text:
        score += 10
        alerts.append({'level': 'warning', 'icon': '🍽', 'title': 'No comió', 'detail': 'Checklist reciente indica que no comió.'})

    if 'orinó: no' in all_recent_text or 'orino: no' in all_recent_text:
        score += 12
        alerts.append({'level': 'warning', 'icon': '🚽', 'title': 'No orinó', 'detail': 'Checklist reciente indica que no orinó.'})

    hours_since_control = hours_from(last_control.event_date) if last_control else None

    if hours_since_control is None:
        score += 20
        alerts.append({'level': 'warning', 'icon': '🕓', 'title': 'Sin controles registrados', 'detail': 'Todavía no hay controles/evoluciones en esta internación.'})
    elif hours_since_control >= 8:
        score += 25
        alerts.append({'level': 'danger', 'icon': '🕓', 'title': 'Control muy demorado', 'detail': f'Último control hace {int(hours_since_control)} horas.'})
    elif hours_since_control >= 4:
        score += 12
        alerts.append({'level': 'warning', 'icon': '🕓', 'title': 'Control próximo', 'detail': f'Último control hace {int(hours_since_control)} horas.'})

    now_minutes = now.hour * 60 + now.minute
    pending_medications = [m for m in medications if m.status != 'Aplicada']
    overdue_medications = []

    for m in pending_medications:
        mins = med_minutes(m.scheduled_time)
        if mins is not None and mins <= now_minutes:
            overdue_medications.append(m)

    if overdue_medications:
        score += min(35, len(overdue_medications) * 12)
        alerts.append({'level': 'danger', 'icon': '💊', 'title': 'Medicación vencida', 'detail': f'{len(overdue_medications)} medicación/es ya deberían haberse aplicado.'})
    elif pending_medications:
        score += min(20, len(pending_medications) * 7)
        alerts.append({'level': 'warning', 'icon': '💊', 'title': 'Medicación pendiente', 'detail': f'{len(pending_medications)} medicación/es pendiente/s.'})

    active_fluids = [f for f in fluids if f.status == 'Activo']

    if active_fluids:
        alerts.append({'level': 'ok', 'icon': '💧', 'title': 'Fluidoterapia activa', 'detail': f'{len(active_fluids)} plan/es activo/s.'})
    else:
        if hospitalization.status == 'Internado' and ('deshidrat' in clinical_text or 'shock' in clinical_text):
            score += 18
            alerts.append({'level': 'warning', 'icon': '💧', 'title': 'Revisar fluidoterapia', 'detail': 'Hay datos de hidratación/perfusión a revisar y no figura fluido activo.'})

    if hospitalization.status != 'Internado':
        priority_label = 'Alta / Cerrado'
        priority_class = 'closed'
        priority_icon = '⚪'
    elif score >= 70:
        priority_label = 'Crítico'
        priority_class = 'danger'
        priority_icon = '🔴'
    elif score >= 40:
        priority_label = 'Alta prioridad'
        priority_class = 'warning'
        priority_icon = '🟠'
    elif score >= 15:
        priority_label = 'Control cercano'
        priority_class = 'watch'
        priority_icon = '🟡'
    else:
        priority_label = 'Estable'
        priority_class = 'ok'
        priority_icon = '🟢'

    if not alerts:
        alerts.append({'level': 'ok', 'icon': '🟢', 'title': 'Sin alertas activas', 'detail': 'No se detectan alertas automáticas por ahora.'})

    nursing_tasks = []

    for m in overdue_medications:
        nursing_tasks.append({
            'icon': '💊',
            'title': m.medication_name or 'Medicación vencida',
            'detail': f'{m.scheduled_time or "Sin horario"} · {m.dose or "-"} · {m.route or "-"}',
            'priority': 'Urgente',
            'class': 'danger',
            'hospitalization_id': hospitalization.id
        })

    for m in pending_medications:
        if m not in overdue_medications:
            nursing_tasks.append({
                'icon': '💊',
                'title': m.medication_name or 'Medicación pendiente',
                'detail': f'{m.scheduled_time or "Sin horario"} · {m.dose or "-"} · {m.route or "-"}',
                'priority': 'Pendiente',
                'class': 'warning',
                'hospitalization_id': hospitalization.id
            })

    for f in active_fluids:
        nursing_tasks.append({
            'icon': '💧',
            'title': f.fluid_type or 'Fluidoterapia',
            'detail': f'{f.fluid_rate or "-"} ml/h · {f.drip_set or "-"}',
            'priority': 'Activa',
            'class': 'ok',
            'hospitalization_id': hospitalization.id
        })

    nursing_tasks.append({
        'icon': '📋',
        'title': 'Próximo control clínico',
        'detail': 'Registrar T°, FC, FR, mucosas, TRC, hidratación y evolución.',
        'priority': 'Prioritario' if priority_class in ['danger', 'warning'] else 'Rutina',
        'class': priority_class,
        'hospitalization_id': hospitalization.id
    })

    uti_monitor = [
        {'label': 'Temperatura', 'value': temp if temp is not None else '-', 'unit': '°C', 'class': 'danger' if temp is not None and (temp >= 40 or temp <= 36.5) else 'warning' if temp is not None and (temp >= 39.5 or temp <= 37) else 'ok'},
        {'label': 'FC', 'value': hr if hr is not None else '-', 'unit': 'lpm', 'class': 'danger' if hr is not None and (hr >= 200 or hr <= 50) else 'warning' if hr is not None and (hr >= 180 or hr <= 60) else 'ok'},
        {'label': 'FR', 'value': rr if rr is not None else '-', 'unit': 'rpm', 'class': 'danger' if rr is not None and (rr >= 80 or rr <= 8) else 'warning' if rr is not None and (rr >= 60 or rr <= 10) else 'ok'},
        {'label': 'Peso', 'value': weight if weight is not None else '-', 'unit': 'kg', 'class': 'warning' if prev_weight is not None and weight is not None and prev_weight > 0 and abs(((weight - prev_weight) / prev_weight) * 100) >= 5 else 'ok'},
        {'label': 'Mucosas', 'value': mucosas or '-', 'unit': '', 'class': 'danger' if any(word in clinical_text for word in danger_words) else 'warning' if any(word in clinical_text for word in warning_words) else 'ok'},
        {'label': 'Hidratación', 'value': hidratacion or '-', 'unit': '', 'class': 'warning' if any(word in clinical_text for word in warning_words) else 'ok'}
    ]

    clinical_summary = []
    clinical_summary.append(f"{priority_icon} Prioridad: {priority_label}.")
    if temp is not None:
        clinical_summary.append(f"Temperatura actual: {temp:g} °C.")
    if hr is not None:
        clinical_summary.append(f"FC actual: {hr:g} lpm.")
    if rr is not None:
        clinical_summary.append(f"FR actual: {rr:g} rpm.")
    if pending_medications:
        clinical_summary.append(f"Medicación pendiente: {len(pending_medications)}.")
    if active_fluids:
        clinical_summary.append(f"Fluidoterapia activa: {len(active_fluids)} plan/es.")

    return {
        'alerts': alerts,
        'score': min(score, 100),
        'priority_label': priority_label,
        'priority_class': priority_class,
        'priority_icon': priority_icon,
        'uti_monitor': uti_monitor,
        'nursing_tasks': nursing_tasks,
        'last_vital': last_vital,
        'hours_since_control': hours_since_control,
        'clinical_summary': clinical_summary,
        'pending_medications_count': len(pending_medications),
        'overdue_medications_count': len(overdue_medications),
        'active_fluids_count': len(active_fluids)
    }

@app.post('/hospitalizations/{hospitalization_id}/ai')
def hospitalization_ai_real(
    hospitalization_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    hospitalization = db.get(Hospitalization, hospitalization_id)

    if not hospitalization:
        raise HTTPException(status_code=404, detail='Internación no encontrada')

    patient = hospitalization.patient

    related_events = (
        db.query(ClinicalEvent)
        .filter(ClinicalEvent.patient_id == patient.id)
        .order_by(ClinicalEvent.event_date.desc())
        .limit(12)
        .all()
    )

    medications = (
        db.query(HospitalizationMedication)
        .filter(HospitalizationMedication.hospitalization_id == hospitalization.id)
        .order_by(HospitalizationMedication.scheduled_time.asc())
        .all()
    )

    fluids = (
        db.query(HospitalizationFluid)
        .filter(HospitalizationFluid.hospitalization_id == hospitalization.id)
        .order_by(HospitalizationFluid.created_at.desc())
        .all()
    )

    fake_event = ClinicalEvent(
        patient_id=patient.id,
        patient=patient,
        event_type='Internación',
        title='Análisis IA de internación',
        description=f"""
Motivo de internación:
{hospitalization.reason or ''}

Diagnóstico inicial:
{hospitalization.diagnosis or ''}

Plan terapéutico:
{hospitalization.treatment_plan or ''}

Notas:
{hospitalization.notes or ''}

Medicación:
{chr(10).join([(m.scheduled_time or '-') + ' · ' + (m.medication_name or '-') + ' · ' + (m.dose or '-') + ' · ' + (m.route or '-') + ' · ' + (m.status or '-') for m in medications])}

Fluidoterapia:
{chr(10).join([(f.fluid_type or '-') + ' · ' + (f.fluid_rate or '-') + ' ml/h · ' + (f.status or '-') for f in fluids])}
""",
        diagnosis=hospitalization.diagnosis or '',
        treatment=hospitalization.treatment_plan or '',
        weight=hospitalization.initial_weight or patient.weight,
        temperature=hospitalization.initial_temperature,
        heart_rate=hospitalization.initial_heart_rate,
        respiratory_rate=hospitalization.initial_respiratory_rate,
        mucous_membranes=hospitalization.initial_mucous_membranes or '',
        crt=hospitalization.initial_crt or '',
        hydration=hospitalization.initial_hydration or '',
        event_date=argentina_now()
    )

    result = ai_clinical_summary(fake_event)

    return JSONResponse(result)
@app.get('/nursing', response_class=HTMLResponse)
def nursing_view(
    request: Request,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    active_hospitalizations = (
        db.query(Hospitalization)
        .filter(Hospitalization.status == 'Internado')
        .order_by(Hospitalization.admission_date.asc())
        .all()
    )

    nursing_patients = []
    global_tasks = []
    stats = {
        'total': 0,
        'critical': 0,
        'warning': 0,
        'watch': 0,
        'stable': 0,
        'overdue_meds': 0,
        'pending_meds': 0,
        'active_fluids': 0
    }

    for hospitalization in active_hospitalizations:
        patient = hospitalization.patient

        related_events = (
            db.query(ClinicalEvent)
            .filter(ClinicalEvent.patient_id == patient.id)
            .filter(ClinicalEvent.event_type.in_(['Internación', 'Control', 'Consulta clínica', 'Alta']))
            .order_by(ClinicalEvent.event_date.desc())
            .limit(30)
            .all()
        )

        medications = (
            db.query(HospitalizationMedication)
            .filter(HospitalizationMedication.hospitalization_id == hospitalization.id)
            .order_by(HospitalizationMedication.scheduled_time.asc(), HospitalizationMedication.created_at.desc())
            .all()
        )

        fluids = (
            db.query(HospitalizationFluid)
            .filter(HospitalizationFluid.hospitalization_id == hospitalization.id)
            .order_by(HospitalizationFluid.created_at.desc())
            .all()
        )

        hospital_ai = build_hospitalization_ai(
            hospitalization,
            patient,
            related_events,
            medications,
            fluids
        )

        stats['total'] += 1
        stats['overdue_meds'] += hospital_ai['overdue_medications_count']
        stats['pending_meds'] += hospital_ai['pending_medications_count']
        stats['active_fluids'] += hospital_ai['active_fluids_count']

        if hospital_ai['priority_class'] == 'danger':
            stats['critical'] += 1
        elif hospital_ai['priority_class'] == 'warning':
            stats['warning'] += 1
        elif hospital_ai['priority_class'] == 'watch':
            stats['watch'] += 1
        else:
            stats['stable'] += 1

        patient_card = {
            'hospitalization': hospitalization,
            'patient': patient,
            'owner': patient.owner if patient else None,
            'ai': hospital_ai,
            'tasks': hospital_ai['nursing_tasks'],
            'alerts': hospital_ai['alerts'],
            'last_event': related_events[0] if related_events else None,
            'medications': medications,
            'fluids': fluids
        }

        nursing_patients.append(patient_card)

        for task in hospital_ai['nursing_tasks']:
            global_tasks.append({
                'patient_name': patient.name if patient else '-',
                'cage': hospitalization.cage or '-',
                'hospitalization_id': hospitalization.id,
                'task': task,
                'priority_class': task.get('class') or hospital_ai['priority_class']
            })

    order_map = {'danger': 0, 'warning': 1, 'watch': 2, 'ok': 3, 'closed': 4}
    nursing_patients = sorted(
        nursing_patients,
        key=lambda item: order_map.get(item['ai']['priority_class'], 9)
    )

    global_tasks = sorted(
        global_tasks,
        key=lambda item: order_map.get(item['priority_class'], 9)
    )

    return templates.TemplateResponse(
        'nursing.html',
        {
            'request': request,
            'nursing_patients': nursing_patients,
            'global_tasks': global_tasks,
            'stats': stats,
            'now': argentina_now().replace(tzinfo=None),
            'today': argentina_now().date()
        }
    )


@app.get('/hospitalizations/nursing', response_class=HTMLResponse)
def hospitalizations_nursing_alias(
    request: Request,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    return RedirectResponse('/nursing', status_code=303)


@app.get('/hospitalizations/{hospitalization_id}', response_class=HTMLResponse)
def hospitalization_detail(
    hospitalization_id: int,
    request: Request,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    hospitalization = db.get(Hospitalization, hospitalization_id)

    if not hospitalization:
        raise HTTPException(status_code=404, detail='Internación no encontrada')

    patient = hospitalization.patient

    related_events = (
        db.query(ClinicalEvent)
        .filter(ClinicalEvent.patient_id == patient.id)
        .filter(ClinicalEvent.event_type.in_(['Internación', 'Control', 'Consulta clínica', 'Alta']))
        .order_by(ClinicalEvent.event_date.desc())
        .limit(30)
        .all()
    )

    medications = (
        db.query(HospitalizationMedication)
        .filter(HospitalizationMedication.hospitalization_id == hospitalization.id)
        .order_by(HospitalizationMedication.scheduled_time.asc(), HospitalizationMedication.created_at.desc())
        .all()
    )

    fluids = (
        db.query(HospitalizationFluid)
        .filter(HospitalizationFluid.hospitalization_id == hospitalization.id)
        .order_by(HospitalizationFluid.created_at.desc())
        .all()
    )

    hospital_ai = build_hospitalization_ai(
        hospitalization,
        patient,
        related_events,
        medications,
        fluids
    )

    return templates.TemplateResponse(
        'hospitalization_detail.html',
        {
            'request': request,
            'hospitalization': hospitalization,
            'patient': patient,
            'related_events': related_events,
            'medications': medications,
            'fluids': fluids,
            'hospital_alerts': hospital_ai['alerts'],
            'priority_label': hospital_ai['priority_label'],
            'priority_class': hospital_ai['priority_class'],
            'priority_icon': hospital_ai['priority_icon'],
            'priority_score': hospital_ai['score'],
            'uti_monitor': hospital_ai['uti_monitor'],
            'nursing_tasks': hospital_ai['nursing_tasks'],
            'last_vital': hospital_ai['last_vital'],
            'hours_since_control': hospital_ai['hours_since_control'],
            'clinical_summary': hospital_ai['clinical_summary'],
            'today': argentina_now().date()
        }
    )

@app.post('/hospitalizations/{hospitalization_id}/evolution')
def hospitalization_add_evolution(
    hospitalization_id: int,
    title: str = Form('Evolución de internación'),
    description: str = Form(''),
    temperature: str = Form(''),
    heart_rate: str = Form(''),
    respiratory_rate: str = Form(''),
    weight: str = Form(''),
    mucous_membranes: str = Form(''),
    crt: str = Form(''),
    hydration: str = Form(''),
    treatment: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    hospitalization = db.get(Hospitalization, hospitalization_id)

    if not hospitalization:
        raise HTTPException(status_code=404, detail='Internación no encontrada')

    patient = hospitalization.patient

    def to_float(value):
        try:
            return float(str(value).replace(',', '.')) if value and str(value).strip() else None
        except ValueError:
            return None

    def to_int(value):
        try:
            return int(float(str(value).replace(',', '.'))) if value and str(value).strip() else None
        except ValueError:
            return None

    evolution_lines = []

    evolution_lines.append('Constantes / control')
    evolution_lines.append('')

    if temperature and str(temperature).strip():
        evolution_lines.append(f'🌡 Temperatura: {temperature} °C')

    if heart_rate and str(heart_rate).strip():
        evolution_lines.append(f'❤️ FC: {heart_rate} lpm')

    if respiratory_rate and str(respiratory_rate).strip():
        evolution_lines.append(f'🫁 FR: {respiratory_rate} rpm')

    if weight and str(weight).strip():
        evolution_lines.append(f'⚖ Peso: {weight} kg')

    if mucous_membranes and str(mucous_membranes).strip():
        evolution_lines.append(f'👄 Mucosas: {mucous_membranes}')

    if crt and str(crt).strip():
        evolution_lines.append(f'⏱ TRC: {crt}')

    if hydration and str(hydration).strip():
        evolution_lines.append(f'💧 Hidratación: {hydration}')

    if description and description.strip():
        evolution_lines.append('')
        evolution_lines.append('Evolución:')
        evolution_lines.append(description.strip())

    if treatment and treatment.strip():
        evolution_lines.append('')
        evolution_lines.append('Tratamiento / indicaciones:')
        evolution_lines.append(treatment.strip())

    final_description = '\n'.join(evolution_lines)

    event = ClinicalEvent(
        patient_id=patient.id,
        event_type='Control',
        title=title or 'Evolución de internación',
        description=final_description,
        treatment=treatment or '',
        temperature=to_float(temperature),
        heart_rate=to_int(heart_rate),
        respiratory_rate=to_int(respiratory_rate),
        weight=to_float(weight),
        mucous_membranes=mucous_membranes or '',
        crt=crt or '',
        hydration=hydration or '',
        created_by=user.username,
        event_date=argentina_now()
    )

    db.add(event)

    if weight and str(weight).strip():
        patient.weight = to_float(weight)

    db.commit()

    return RedirectResponse(
        f'/hospitalizations/{hospitalization.id}',
        status_code=303
    )
@app.post('/hospitalizations/{hospitalization_id}/checklist')
def hospitalization_checklist(
    hospitalization_id: int,
    ate: str = Form(''),
    drank: str = Form(''),
    urinated: str = Form(''),
    defecated: str = Form(''),
    vomiting_diarrhea: str = Form(''),
    cage_cleaned: str = Form(''),
    checklist_notes: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    hospitalization = db.get(Hospitalization, hospitalization_id)

    if not hospitalization:
        raise HTTPException(status_code=404, detail='Internación no encontrada')

    patient = hospitalization.patient

    checklist_lines = []

    checklist_lines.append('Checklist de enfermería')
    checklist_lines.append('')
    checklist_lines.append(f"Comió: {'Sí' if ate else 'No'}")
    checklist_lines.append(f"Tomó agua: {'Sí' if drank else 'No'}")
    checklist_lines.append(f"Orinó: {'Sí' if urinated else 'No'}")
    checklist_lines.append(f"Defecó: {'Sí' if defecated else 'No'}")
    checklist_lines.append(f"Vómitos / diarrea: {'Sí' if vomiting_diarrhea else 'No'}")
    checklist_lines.append(f"Limpieza de jaula: {'Sí' if cage_cleaned else 'No'}")

    if checklist_notes and checklist_notes.strip():
        checklist_lines.append('')
        checklist_lines.append('Observaciones:')
        checklist_lines.append(checklist_notes.strip())

    event = ClinicalEvent(
        patient_id=patient.id,
        event_type='Control',
        title='Checklist de enfermería',
        description='\n'.join(checklist_lines),
        created_by=user.username,
        event_date=argentina_now()
    )

    db.add(event)
    db.commit()

    return RedirectResponse(
        f'/hospitalizations/{hospitalization.id}',
        status_code=303
    )
@app.post('/hospitalizations/{hospitalization_id}/fluids')
def hospitalization_fluids(
    hospitalization_id: int,
    fluid_type: str = Form(''),
    fluid_rate: str = Form(''),
    ml_kg_h: str = Form(''),
    drip_set: str = Form(''),
    fluid_notes: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    hospitalization = db.get(Hospitalization, hospitalization_id)

    if not hospitalization:
        raise HTTPException(status_code=404, detail='Internación no encontrada')

    patient = hospitalization.patient

    fluid = HospitalizationFluid(
        hospitalization_id=hospitalization.id,
        fluid_type=fluid_type or '',
        fluid_rate=fluid_rate or '',
        ml_kg_h=ml_kg_h or '',
        drip_set=drip_set or '',
        notes=fluid_notes or '',
        status='Activo',
        created_by=user.username
    )

    db.add(fluid)

    lines = [
        'Fluidoterapia',
        '',
        f'Tipo de fluido: {fluid_type or "-"}',
        f'Velocidad: {fluid_rate or "-"} ml/h',
        f'Ml/kg/h: {ml_kg_h or "-"}',
        f'Equipo: {drip_set or "-"}'
    ]

    if fluid_notes and fluid_notes.strip():
        lines.append('')
        lines.append('Observaciones:')
        lines.append(fluid_notes.strip())

    event = ClinicalEvent(
        patient_id=patient.id,
        event_type='Control',
        title='Fluidoterapia',
        description='\n'.join(lines),
        created_by=user.username,
        event_date=argentina_now()
    )

    db.add(event)
    db.commit()

    return RedirectResponse(
        f'/hospitalizations/{hospitalization.id}',
        status_code=303
    )
@app.post('/hospitalizations/{hospitalization_id}/fluids/{fluid_id}/finish')
def hospitalization_fluid_finish(
    hospitalization_id: int,
    fluid_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    hospitalization = db.get(Hospitalization, hospitalization_id)
    fluid = db.get(HospitalizationFluid, fluid_id)

    if not hospitalization or not fluid:
        raise HTTPException(status_code=404, detail='Fluidoterapia no encontrada')

    if fluid.hospitalization_id != hospitalization.id:
        raise HTTPException(status_code=404, detail='Fluidoterapia no corresponde a esta internación')

    fluid.status = 'Finalizado'

    if hasattr(fluid, 'finished_at'):
        fluid.finished_at = argentina_now()

    if hasattr(fluid, 'finished_by'):
        fluid.finished_by = user.username
    event = ClinicalEvent(
        patient_id=hospitalization.patient_id,
        event_type='Control',
        title='Fluidoterapia finalizada',
        description=(
            f'Fluido: {fluid.fluid_type or "-"}\n'
            f'Velocidad: {fluid.fluid_rate or "-"} ml/h\n'
            f'Ml/kg/h: {fluid.ml_kg_h or "-"}\n'
            f'Equipo: {fluid.drip_set or "-"}'
        ),
        created_by=user.username,
        event_date=argentina_now()
    )

    db.add(event)
    db.commit()

    return RedirectResponse(
        f'/hospitalizations/{hospitalization.id}',
        status_code=303
    )
@app.post('/hospitalizations/{hospitalization_id}/fluids/{fluid_id}/delete')
def hospitalization_fluid_delete(
    hospitalization_id: int,
    fluid_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    hospitalization = db.get(Hospitalization, hospitalization_id)
    fluid = db.get(HospitalizationFluid, fluid_id)

    if not hospitalization or not fluid:
        raise HTTPException(status_code=404, detail='Fluidoterapia no encontrada')

    if fluid.hospitalization_id != hospitalization.id:
        raise HTTPException(status_code=404, detail='Fluidoterapia no corresponde a esta internación')

    db.delete(fluid)
    db.commit()

    return RedirectResponse(
        f'/hospitalizations/{hospitalization.id}',
        status_code=303
    )
@app.post('/hospitalizations/{hospitalization_id}/medications')
def hospitalization_medication_create(
    hospitalization_id: int,
    medication_name: str = Form(''),
    dose: str = Form(''),
    route: str = Form(''),
    frequency: str = Form(''),
    scheduled_time: str = Form(''),
    notes: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    hospitalization = db.get(Hospitalization, hospitalization_id)

    if not hospitalization:
        raise HTTPException(status_code=404, detail='Internación no encontrada')

    medication = HospitalizationMedication(
        hospitalization_id=hospitalization.id,
        medication_name=medication_name or '',
        dose=dose or '',
        route=route or '',
        frequency=frequency or '',
        scheduled_time=scheduled_time or '',
        notes=notes or '',
        created_by=user.username,
        status='Pendiente'
    )

    db.add(medication)
    db.commit()

    return RedirectResponse(
        f'/hospitalizations/{hospitalization.id}',
        status_code=303
    )
@app.post('/hospitalizations/{hospitalization_id}/medications/{medication_id}/done')
def hospitalization_medication_done(
    hospitalization_id: int,
    medication_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    hospitalization = db.get(Hospitalization, hospitalization_id)
    medication = db.get(HospitalizationMedication, medication_id)

    if not hospitalization or not medication:
        raise HTTPException(status_code=404, detail='Medicación no encontrada')

    if medication.hospitalization_id != hospitalization.id:
        raise HTTPException(status_code=404, detail='Medicación no corresponde a esta internación')

    medication.status = 'Aplicada'
    medication.applied_at = argentina_now()
    medication.applied_by = user.username

    event = ClinicalEvent(
        patient_id=hospitalization.patient_id,
        event_type='Control',
        title='Medicación aplicada',
        description=(
            f'Medicación aplicada: {medication.medication_name or "-"}\n'
            f'Dosis: {medication.dose or "-"}\n'
            f'Vía: {medication.route or "-"}\n'
            f'Horario programado: {medication.scheduled_time or "-"}'
        ),
        created_by=user.username,
        event_date=argentina_now()
    )

    db.add(event)
    db.commit()

    return RedirectResponse(
        f'/hospitalizations/{hospitalization.id}',
        status_code=303
    )
@app.post('/hospitalizations/{hospitalization_id}/medications/{medication_id}/delete')
def hospitalization_medication_delete(
    hospitalization_id: int,
    medication_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    hospitalization = db.get(Hospitalization, hospitalization_id)
    medication = db.get(HospitalizationMedication, medication_id)

    if not hospitalization or not medication:
        raise HTTPException(status_code=404, detail='Medicación no encontrada')

    if medication.hospitalization_id != hospitalization.id:
        raise HTTPException(status_code=404, detail='Medicación no corresponde a esta internación')

    db.delete(medication)
    db.commit()

    return RedirectResponse(
        f'/hospitalizations/{hospitalization.id}',
        status_code=303
    )
@app.post('/hospitalizations/{hospitalization_id}/discharge')
def hospitalization_discharge(
    hospitalization_id: int,
    discharge_summary: str = Form(''),
    discharge_indications: str = Form(''),
    db: Session = Depends(get_db),
    user: User = Depends(require_user)
):
    hospitalization = db.get(Hospitalization, hospitalization_id)

    if not hospitalization:
        raise HTTPException(status_code=404, detail='Internación no encontrada')

    patient = hospitalization.patient

    hospitalization.status = 'Alta'
    hospitalization.discharge_date = argentina_now()
    hospitalization.discharge_summary = discharge_summary or ''
    hospitalization.discharge_indications = discharge_indications or ''

    event = ClinicalEvent(
        patient_id=patient.id,
        event_type='Alta',
        title='Alta de internación',
        description=discharge_summary or '',
        treatment=discharge_indications or '',
        created_by=user.username,
        event_date=argentina_now()
    )

    db.add(event)
    db.commit()

    return RedirectResponse(
        f'/hospitalizations/{hospitalization.id}',
        status_code=303
    )
@app.get('/health')
def health():
    return {'status': 'ok', 'app': 'Los Aromos Cloud'}
    
