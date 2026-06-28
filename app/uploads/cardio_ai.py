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

        getattr(event, "ecg_rhythm", "")

        or ""

    ).lower()



    arrhythmia = (

        getattr(event, "ecg_arrhythmia", "")

        or ""

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
  
