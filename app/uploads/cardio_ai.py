from dataclasses import dataclass


def to_float(value):
    try:
        if value is None:
            return None

        value = str(value).strip().replace(",", ".")

        if value == "":
            return None

        return float(value)

    except Exception:
        return None


@dataclass
class AISection:
    title: str
    items: list


class CardioAI:

    def __init__(self):

        self.score = 100

        self.sections = []

        self.alerts = []

        self.recommendations = []

        self.owner_summary = []

        self.report = []

        self.conclusion = ""



    def penalize(self, points):

        self.score -= points

        if self.score < 0:
            self.score = 0



    def add_alert(self, text):

        if text not in self.alerts:

            self.alerts.append(text)



    def add_recommendation(self, text):

        if text not in self.recommendations:

            self.recommendations.append(text)



    def add_owner(self, text):

        if text not in self.owner_summary:

            self.owner_summary.append(text)



    def add_report(self, text):

        self.report.append(text)



    def add_section(self, title, items):

        self.sections.append(
            AISection(
                title=title,
                items=items
            )
        )

def analyze_ecg(event, ai):

    if event is None:

        ai.add_section(
            "ECG",
            [
                "No existe ECG cargado."
            ]
        )

        ai.penalize(8)

        return

    notes = []

    hr = to_float(getattr(event, "ecg_hr", None))
    axis = to_float(getattr(event, "ecg_axis", None))

    rhythm = (
        getattr(event, "ecg_rhythm", "") or ""
    ).lower()

    arrhythmia = (
        getattr(event, "ecg_arrhythmia", "") or ""
    ).lower()

    if hr is not None:

        if hr < 60:

            ai.penalize(10)

            notes.append(
                f"Bradicardia ({hr:.0f} lpm)."
            )

            ai.add_alert(
                "Frecuencia cardíaca baja."
            )

        elif hr > 180:

            ai.penalize(10)

            notes.append(
                f"Taquicardia ({hr:.0f} lpm)."
            )

            ai.add_alert(
                "Frecuencia cardíaca elevada."
            )

        else:

            notes.append(
                f"Frecuencia cardíaca normal ({hr:.0f} lpm)."
            )

    if axis is not None:

        if axis < 40:

            ai.penalize(5)

            notes.append(
                f"Eje izquierdo ({axis:.0f}°)."
            )

        elif axis > 100:

            ai.penalize(5)

            notes.append(
                f"Eje derecho ({axis:.0f}°)."
            )

        else:

            notes.append(
                f"Eje normal ({axis:.0f}°)."
            )

    if "bloqueo" in rhythm:

        ai.penalize(12)

        notes.append(
            "Se describe bloqueo de conducción."
        )

    if "fibril" in arrhythmia:

        ai.penalize(15)

        notes.append(
            "Compatible con fibrilación."
        )

    if "extras" in arrhythmia:

        ai.penalize(8)

        notes.append(
            "Extrasístoles registradas."
        )

    if len(notes) == 0:

        notes.append(
            "Sin alteraciones relevantes."
        )

    ai.add_section(
        "Electrocardiografía",
        notes
    )
def analyze_echo(event, ai):

    if event is None:

        ai.add_section(
            "Ecocardiografía",
            [
                "No existe ecocardiografía cargada."
            ]
        )

        ai.penalize(8)

        return

    notes = []

    aiao = to_float(getattr(event, "eco_aiao", None))
    fs = to_float(getattr(event, "eco_fs", None))
    fe = to_float(getattr(event, "eco_fe", None))
    epss = to_float(getattr(event, "eco_epss", None))

    acvim = (
        getattr(event, "eco_acvim", "") or ""
    ).upper()

    if aiao is not None:

        if aiao >= 1.9:

            ai.penalize(18)

            notes.append(
                f"AI/Ao severamente aumentado ({aiao})."
            )

            ai.add_alert(
                "Dilatación auricular izquierda importante."
            )

        elif aiao >= 1.6:

            ai.penalize(10)

            notes.append(
                f"AI/Ao aumentado ({aiao})."
            )

        else:

            notes.append(
                f"AI/Ao normal ({aiao})."
            )

    if fs is not None:

        if fs < 20:

            ai.penalize(18)

            notes.append(
                f"FS disminuida ({fs}%)."
            )

            ai.add_alert(
                "Posible disfunción sistólica."
            )

        elif fs > 45:

            notes.append(
                f"FS elevada ({fs}%)."
            )

        else:

            notes.append(
                f"FS conservada ({fs}%)."
            )

    if fe is not None:

        if fe < 45:

            ai.penalize(10)

            notes.append(
                f"Fracción de eyección disminuida ({fe}%)."
            )

    if epss is not None:

        if epss > 7:

            ai.penalize(8)

            notes.append(
                f"EPSS aumentado ({epss})."
            )

    if acvim:

        if acvim == "B2":

            ai.penalize(10)

            notes.append(
                "Clasificación ACVIM B2."
            )

        elif acvim == "C":

            ai.penalize(22)

            notes.append(
                "Clasificación ACVIM C."
            )

            ai.add_alert(
                "Paciente con insuficiencia cardíaca."
            )

        elif acvim == "D":

            ai.penalize(35)

            notes.append(
                "Clasificación ACVIM D."
            )

            ai.add_alert(
                "Insuficiencia cardíaca avanzada."
            )

    if len(notes) == 0:

        notes.append(
            "Sin alteraciones ecocardiográficas relevantes."
        )

    ai.add_section(
        "Ecocardiografía",
        notes
    )
def analyze_rx(event, ai):

    if event is None:

        ai.add_section(
            "Radiografía",
            [
                "No existe radiografía cardíaca cargada."
            ]
        )

        return

    notes = []

    vhs = to_float(getattr(event, "rx_vhs", None))
    vlas = to_float(getattr(event, "rx_vlas", None))

    edema = (
        getattr(event, "rx_edema", "") or ""
    ).lower()

    congestion = (
        getattr(event, "rx_congestion", "") or ""
    ).lower()

    pattern = (
        getattr(event, "rx_lung_pattern", "") or ""
    ).lower()

    if vhs is not None:

        if vhs > 11.5:

            ai.penalize(10)

            notes.append(
                f"VHS aumentado ({vhs})."
            )

            ai.add_alert(
                "Cardiomegalia radiográfica."
            )

        else:

            notes.append(
                f"VHS normal ({vhs})."
            )

    if vlas is not None:

        if vlas > 3:

            ai.penalize(8)

            notes.append(
                f"VLAS aumentado ({vlas})."
            )

        else:

            notes.append(
                f"VLAS normal ({vlas})."
            )

    if "edema" in edema:

        ai.penalize(20)

        notes.append(
            "Compatible con edema pulmonar."
        )

        ai.add_alert(
            "Edema pulmonar."
        )

    if "congest" in congestion:

        ai.penalize(12)

        notes.append(
            "Congestión vascular pulmonar."
        )

    if "intersticial" in pattern:

        ai.penalize(5)

        notes.append(
            "Patrón intersticial."
        )

    if "alveolar" in pattern:

        ai.penalize(10)

        notes.append(
            "Patrón alveolar."
        )

    if len(notes) == 0:

        notes.append(
            "Sin alteraciones radiográficas relevantes."
        )

    ai.add_section(
        "Radiografía",
        notes
    )
def calculate_score(ai):

    if ai.score < 0:
        ai.score = 0

    if ai.score > 100:
        ai.score = 100

    if ai.score >= 85:

        ai.conclusion = (
            "Paciente cardiológicamente estable."
        )

    elif ai.score >= 60:

        ai.conclusion = (
            "Paciente estable pero requiere seguimiento."
        )

    elif ai.score >= 40:

        ai.conclusion = (
            "Paciente con riesgo cardiológico moderado."
        )

    else:

        ai.conclusion = (
            "Paciente con alto riesgo cardiológico."
        )

    return ai.score
def build_recommendations(ai):

    if ai.score >= 85:

        ai.add_recommendation(
            "Continuar controles cardiológicos periódicos."
        )

    elif ai.score >= 60:

        ai.add_recommendation(
            "Realizar control ecocardiográfico según evolución."
        )

    elif ai.score >= 40:

        ai.add_recommendation(
            "Control cardiológico cercano y ajuste terapéutico si corresponde."
        )

    else:

        ai.add_recommendation(
            "Reevaluación inmediata y tratamiento intensivo según criterio clínico."
        )

    if len(ai.alerts):

        ai.add_recommendation(
            "Correlacionar los hallazgos con la clínica del paciente."
        )
def build_owner_summary(ai):

    ai.owner_summary.clear()

    if ai.score >= 85:

        ai.add_owner(
            "Actualmente el corazón se encuentra estable según los estudios disponibles."
        )

    elif ai.score >= 60:

        ai.add_owner(
            "Se observaron algunos cambios que requieren controles periódicos."
        )

    elif ai.score >= 40:

        ai.add_owner(
            "El paciente presenta alteraciones cardíacas que necesitan seguimiento cercano."
        )

    else:

        ai.add_owner(
            "Los estudios muestran alteraciones cardíacas importantes que requieren tratamiento y controles frecuentes."
        )
def analyze_complete_case(last_ecg, last_eco, last_rx):

    ai = CardioAI()

    analyze_ecg(last_ecg, ai)

    analyze_echo(last_eco, ai)

    analyze_rx(last_rx, ai)

    calculate_score(ai)

    build_recommendations(ai)

    build_owner_summary(ai)

    report = []

    report.append(ai.conclusion)

    for section in ai.sections:

        report.append("")

        report.append(section.title)

        for item in section.items:

            report.append(f"- {item}")

    ai.report = "\n".join(report)

    return {

        "score": ai.score,

        "label": (
            "🟢 Estable"
            if ai.score >= 85 else
            "🟡 Seguimiento"
            if ai.score >= 60 else
            "🟠 Riesgo moderado"
            if ai.score >= 40 else
            "🔴 Alto riesgo"
        ),

        "conclusion": ai.conclusion,

        "sections": [

            {

                "title": s.title,

                "items": s.items

            }

            for s in ai.sections

        ],

        "alerts": ai.alerts,

        "recommendations": ai.recommendations,

        "owner_summary": ai.owner_summary,

        "report": ai.report

    }        
