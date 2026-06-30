from datetime import datetime, date
from zoneinfo import ZoneInfo
from typing import Optional

from sqlalchemy import String, Text, Date, DateTime, Float, ForeignKey, Boolean
from sqlalchemy.orm import Mapped, mapped_column, relationship
from .database import Base

ARG_TZ = ZoneInfo("America/Argentina/Buenos_Aires")

def argentina_now():
    return datetime.now(ARG_TZ)


class User(Base):
    __tablename__ = 'users'

    id: Mapped[int] = mapped_column(primary_key=True)
    username: Mapped[str] = mapped_column(String(80), unique=True, index=True)
    password_hash: Mapped[str] = mapped_column(String(255))
    full_name: Mapped[str] = mapped_column(String(120), default='')
    is_active: Mapped[bool] = mapped_column(Boolean, default=True)


class Owner(Base):
    __tablename__ = 'owners'

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(180), index=True)
    phone: Mapped[str] = mapped_column(String(80), default='', index=True)
    whatsapp: Mapped[str] = mapped_column(String(80), default='')
    email: Mapped[str] = mapped_column(String(180), default='')
    address: Mapped[str] = mapped_column(String(255), default='')
    notes: Mapped[str] = mapped_column(Text, default='')
    created_at: Mapped[datetime] = mapped_column(DateTime, default=argentina_now)

    patients: Mapped[list['Patient']] = relationship(back_populates='owner')


class Patient(Base):
    __tablename__ = 'patients'

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(140), index=True)
    species: Mapped[str] = mapped_column(String(80), default='')
    breed: Mapped[str] = mapped_column(String(120), default='')
    sex: Mapped[str] = mapped_column(String(50), default='')
    birth_date: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    weight: Mapped[Optional[float]] = mapped_column(Float, nullable=True)
    color: Mapped[str] = mapped_column(String(80), default='')
    microchip: Mapped[str] = mapped_column(String(100), default='')
    alerts: Mapped[str] = mapped_column(Text, default='')
    notes: Mapped[str] = mapped_column(Text, default='')

    owner_id: Mapped[int] = mapped_column(ForeignKey('owners.id'))
    owner: Mapped['Owner'] = relationship(back_populates='patients')
    events: Mapped[list['ClinicalEvent']] = relationship(
        back_populates='patient',
        order_by='desc(ClinicalEvent.event_date)'
    )

    hospitalizations: Mapped[list['Hospitalization']] = relationship(
        back_populates='patient',
        order_by='desc(Hospitalization.admission_date)'
    )


class ClinicalEvent(Base):
    __tablename__ = 'clinical_events'

    id: Mapped[int] = mapped_column(primary_key=True)
    patient_id: Mapped[int] = mapped_column(ForeignKey('patients.id'), index=True)
    event_date: Mapped[datetime] = mapped_column(DateTime, default=argentina_now, index=True)
    event_type: Mapped[str] = mapped_column(String(80), default='Consulta')
    title: Mapped[str] = mapped_column(String(200), default='')
    description: Mapped[str] = mapped_column(Text, default='')
    diagnosis: Mapped[str] = mapped_column(Text, default='')
    treatment: Mapped[str] = mapped_column(Text, default='')
    reminder_date: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    created_by: Mapped[str] = mapped_column(String(100), default='admin')

    weight: Mapped[Optional[float]] = mapped_column(Float, nullable=True)
    temperature: Mapped[Optional[float]] = mapped_column(Float, nullable=True)
    heart_rate: Mapped[Optional[int]] = mapped_column(nullable=True)
    respiratory_rate: Mapped[Optional[int]] = mapped_column(nullable=True)

    mucous_membranes: Mapped[str] = mapped_column(String(100), default='')
    crt: Mapped[str] = mapped_column(String(50), default='')
    hydration: Mapped[str] = mapped_column(String(100), default='')
    anamnesis: Mapped[str] = mapped_column(Text, default='')
    physical_exam: Mapped[str] = mapped_column(Text, default='')

    ecg_hr: Mapped[str] = mapped_column(String(50), default='')
    ecg_rhythm: Mapped[str] = mapped_column(String(100), default='')
    ecg_p: Mapped[str] = mapped_column(String(50), default='')
    ecg_pr: Mapped[str] = mapped_column(String(50), default='')
    ecg_qrs: Mapped[str] = mapped_column(String(50), default='')
    ecg_st: Mapped[str] = mapped_column(String(50), default='')
    ecg_t: Mapped[str] = mapped_column(String(50), default='')
    ecg_qt: Mapped[str] = mapped_column(String(50), default='')
    ecg_axis: Mapped[str] = mapped_column(String(50), default='')
    ecg_interpretation: Mapped[str] = mapped_column(Text, default='')

    ecg_p_mv: Mapped[str] = mapped_column(String(50), default='')
    ecg_qrs_mv: Mapped[str] = mapped_column(String(50), default='')
    ecg_t_mv: Mapped[str] = mapped_column(String(50), default='')
    ecg_qtc: Mapped[str] = mapped_column(String(50), default='')
    ecg_polarity: Mapped[str] = mapped_column(String(100), default='')
    ecg_arrhythmia: Mapped[str] = mapped_column(String(150), default='')
    ecg_conduction: Mapped[str] = mapped_column(String(150), default='')
    ecg_notes: Mapped[str] = mapped_column(Text, default='')

    eco_aiao: Mapped[str] = mapped_column(String(50), default='')
    eco_fs: Mapped[str] = mapped_column(String(50), default='')
    eco_acvim: Mapped[str] = mapped_column(String(50), default='')
    eco_diagnosis: Mapped[str] = mapped_column(Text, default='')
    eco_treatment: Mapped[str] = mapped_column(Text, default='')

    eco_epss: Mapped[str] = mapped_column(String(50), default='')
    eco_lvidd: Mapped[str] = mapped_column(String(50), default='')
    eco_lvids: Mapped[str] = mapped_column(String(50), default='')
    eco_ivsd: Mapped[str] = mapped_column(String(50), default='')
    eco_ivss: Mapped[str] = mapped_column(String(50), default='')
    eco_lvpwd: Mapped[str] = mapped_column(String(50), default='')
    eco_lvpws: Mapped[str] = mapped_column(String(50), default='')
    eco_fe: Mapped[str] = mapped_column(String(50), default='')
    eco_la_size: Mapped[str] = mapped_column(String(100), default='')
    eco_lv_size: Mapped[str] = mapped_column(String(100), default='')
    eco_rv_size: Mapped[str] = mapped_column(String(100), default='')
    eco_ra_size: Mapped[str] = mapped_column(String(100), default='')
    eco_mitral: Mapped[str] = mapped_column(String(150), default='')
    eco_tricuspid: Mapped[str] = mapped_column(String(150), default='')
    eco_aortic: Mapped[str] = mapped_column(String(150), default='')
    eco_pulmonary: Mapped[str] = mapped_column(String(150), default='')
    eco_pulmonary_htn: Mapped[str] = mapped_column(String(150), default='')
    eco_pericardium: Mapped[str] = mapped_column(String(150), default='')
    eco_doppler: Mapped[str] = mapped_column(Text, default='')
    eco_observations: Mapped[str] = mapped_column(Text, default='')

    rx_vhs: Mapped[str] = mapped_column(String(50), default='')
    rx_vlas: Mapped[str] = mapped_column(String(50), default='')
    rx_heart_size: Mapped[str] = mapped_column(String(150), default='')
    rx_left_atrium: Mapped[str] = mapped_column(String(150), default='')
    rx_left_heart: Mapped[str] = mapped_column(String(150), default='')
    rx_right_heart: Mapped[str] = mapped_column(String(150), default='')
    rx_pulmonary_vessels: Mapped[str] = mapped_column(String(150), default='')
    rx_lung_pattern: Mapped[str] = mapped_column(String(150), default='')
    rx_edema: Mapped[str] = mapped_column(String(150), default='')
    rx_congestion: Mapped[str] = mapped_column(String(150), default='')
    rx_trachea: Mapped[str] = mapped_column(String(150), default='')
    rx_observations: Mapped[str] = mapped_column(Text, default='')

    vaccine_name: Mapped[str] = mapped_column(String(150), default='')
    vaccine_lot: Mapped[str] = mapped_column(String(100), default='')
    vaccine_expiration: Mapped[str] = mapped_column(String(100), default='')
    next_vaccine_date: Mapped[str] = mapped_column(String(100), default='')

    dewormer_product: Mapped[str] = mapped_column(String(150), default='')
    dewormer_drug: Mapped[str] = mapped_column(String(150), default='')
    dewormer_dose: Mapped[str] = mapped_column(String(100), default='')
    next_deworming_date: Mapped[str] = mapped_column(String(100), default='')

    patient: Mapped['Patient'] = relationship(back_populates='events')
    attachments: Mapped[list['EventAttachment']] = relationship(
        back_populates='event',
        cascade='all, delete-orphan'
    )


class Hospitalization(Base):
    __tablename__ = 'hospitalizations'

    id: Mapped[int] = mapped_column(primary_key=True)

    patient_id: Mapped[int] = mapped_column(ForeignKey('patients.id'), index=True)
    clinical_event_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey('clinical_events.id'),
        nullable=True,
        index=True
    )

    admission_date: Mapped[datetime] = mapped_column(DateTime, default=argentina_now, index=True)
    discharge_date: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)

    status: Mapped[str] = mapped_column(String(40), default='Internado', index=True)

    cage: Mapped[str] = mapped_column(String(80), default='')
    responsible_vet: Mapped[str] = mapped_column(String(120), default='')

    reason: Mapped[str] = mapped_column(Text, default='')
    diagnosis: Mapped[str] = mapped_column(Text, default='')
    treatment_plan: Mapped[str] = mapped_column(Text, default='')
    notes: Mapped[str] = mapped_column(Text, default='')

    initial_weight: Mapped[Optional[float]] = mapped_column(Float, nullable=True)
    initial_temperature: Mapped[Optional[float]] = mapped_column(Float, nullable=True)
    initial_heart_rate: Mapped[Optional[int]] = mapped_column(nullable=True)
    initial_respiratory_rate: Mapped[Optional[int]] = mapped_column(nullable=True)
    initial_mucous_membranes: Mapped[str] = mapped_column(String(100), default='')
    initial_crt: Mapped[str] = mapped_column(String(50), default='')
    initial_hydration: Mapped[str] = mapped_column(String(100), default='')

    discharge_summary: Mapped[str] = mapped_column(Text, default='')
    discharge_indications: Mapped[str] = mapped_column(Text, default='')

    created_by: Mapped[str] = mapped_column(String(100), default='admin')
    created_at: Mapped[datetime] = mapped_column(DateTime, default=argentina_now)

    patient: Mapped['Patient'] = relationship(back_populates='hospitalizations')
    clinical_event: Mapped[Optional['ClinicalEvent']] = relationship()
class HospitalizationMedication(Base):
    __tablename__ = 'hospitalization_medications'

    id: Mapped[int] = mapped_column(primary_key=True)

    hospitalization_id: Mapped[int] = mapped_column(
        ForeignKey('hospitalizations.id'),
        index=True
    )

    medication_name: Mapped[str] = mapped_column(String(200), default='')
    dose: Mapped[str] = mapped_column(String(120), default='')
    route: Mapped[str] = mapped_column(String(80), default='')
    frequency: Mapped[str] = mapped_column(String(120), default='')
    scheduled_time: Mapped[str] = mapped_column(String(50), default='')

    status: Mapped[str] = mapped_column(String(40), default='Pendiente')
    applied_at: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)
    applied_by: Mapped[str] = mapped_column(String(100), default='')
    notes: Mapped[str] = mapped_column(Text, default='')

    created_by: Mapped[str] = mapped_column(String(100), default='admin')
    created_at: Mapped[datetime] = mapped_column(DateTime, default=argentina_now)

    hospitalization: Mapped['Hospitalization'] = relationship()
class HospitalizationFluid(Base):
    __tablename__ = 'hospitalization_fluids'

    id: Mapped[int] = mapped_column(primary_key=True)

    hospitalization_id: Mapped[int] = mapped_column(
        ForeignKey('hospitalizations.id'),
        index=True
    )

    fluid_type: Mapped[str] = mapped_column(String(200), default='')
    fluid_rate: Mapped[str] = mapped_column(String(80), default='')
    ml_kg_h: Mapped[str] = mapped_column(String(80), default='')
    drip_set: Mapped[str] = mapped_column(String(80), default='')
    notes: Mapped[str] = mapped_column(Text, default='')

    status: Mapped[str] = mapped_column(String(40), default='Activo')
    finished_at: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)
    finished_by: Mapped[str] = mapped_column(String(100), default='')
    created_by: Mapped[str] = mapped_column(String(100), default='admin')
    created_at: Mapped[datetime] = mapped_column(DateTime, default=argentina_now)

    hospitalization: Mapped['Hospitalization'] = relationship()
class Appointment(Base):
    __tablename__ = 'appointments'

    id: Mapped[int] = mapped_column(primary_key=True)

    service: Mapped[str] = mapped_column(String(120), default='', index=True)
    title: Mapped[str] = mapped_column(String(200), default='')
    notes: Mapped[str] = mapped_column(Text, default='')

    appointment_date: Mapped[datetime] = mapped_column(DateTime, index=True)
    start_time: Mapped[str] = mapped_column(String(20), default='')
    end_time: Mapped[str] = mapped_column(String(20), default='')

    owner_id: Mapped[Optional[int]] = mapped_column(ForeignKey('owners.id'), nullable=True)
    patient_id: Mapped[Optional[int]] = mapped_column(ForeignKey('patients.id'), nullable=True)

    status: Mapped[str] = mapped_column(String(40), default='Pendiente')
    reminder_24h: Mapped[bool] = mapped_column(Boolean, default=True)
    reminder_12h: Mapped[bool] = mapped_column(Boolean, default=False)
    contact_whatsapp: Mapped[str] = mapped_column(String(80), default='')
    created_at: Mapped[datetime] = mapped_column(DateTime, default=argentina_now)

    owner: Mapped[Optional['Owner']] = relationship()
    patient: Mapped[Optional['Patient']] = relationship()


class WaitingListEntry(Base):
    __tablename__ = 'waiting_list'

    id: Mapped[int] = mapped_column(primary_key=True)

    arrival_time: Mapped[datetime] = mapped_column(DateTime, default=argentina_now, index=True)
    status: Mapped[str] = mapped_column(String(40), default='Esperando', index=True)
    priority: Mapped[str] = mapped_column(String(40), default='Normal')
    reason: Mapped[str] = mapped_column(String(200), default='')
    notes: Mapped[str] = mapped_column(Text, default='')

    owner_id: Mapped[Optional[int]] = mapped_column(ForeignKey('owners.id'), nullable=True)
    patient_id: Mapped[Optional[int]] = mapped_column(ForeignKey('patients.id'), nullable=True)
    appointment_id: Mapped[Optional[int]] = mapped_column(ForeignKey('appointments.id'), nullable=True)

    started_at: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)
    finished_at: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)

    created_by: Mapped[str] = mapped_column(String(100), default='admin')

    owner: Mapped[Optional['Owner']] = relationship()
    patient: Mapped[Optional['Patient']] = relationship()
    appointment: Mapped[Optional['Appointment']] = relationship()


class Product(Base):
    __tablename__ = 'products'

    id: Mapped[int] = mapped_column(primary_key=True)

    rubro: Mapped[str] = mapped_column(String(120), default='', index=True)
    tipo: Mapped[str] = mapped_column(String(120), default='')
    name: Mapped[str] = mapped_column(String(200), default='', index=True)
    code: Mapped[str] = mapped_column(String(100), default='', index=True)
    barcode: Mapped[str] = mapped_column(String(120), default='', index=True)

    cost_price: Mapped[Optional[float]] = mapped_column(Float, nullable=True)
    sale_price: Mapped[Optional[float]] = mapped_column(Float, nullable=True)
    margin_percent: Mapped[Optional[float]] = mapped_column(Float, nullable=True)

    stock: Mapped[Optional[float]] = mapped_column(Float, nullable=True)
    min_stock: Mapped[Optional[float]] = mapped_column(Float, nullable=True)

    expiration_date: Mapped[Optional[date]] = mapped_column(Date, nullable=True)

    provider: Mapped[str] = mapped_column(String(180), default='')
    manufacturer: Mapped[str] = mapped_column(String(180), default='')

    active: Mapped[bool] = mapped_column(Boolean, default=True)
    notes: Mapped[str] = mapped_column(Text, default='')

    created_at: Mapped[datetime] = mapped_column(DateTime, default=argentina_now)


class Sale(Base):
    __tablename__ = 'sales'

    id: Mapped[int] = mapped_column(primary_key=True)
    date: Mapped[datetime] = mapped_column(DateTime, default=argentina_now)
    status: Mapped[str] = mapped_column(String(20), default='paid')
    total: Mapped[float] = mapped_column(Float, default=0)
    payment_method: Mapped[str] = mapped_column(String(50), default='Efectivo')

    patient_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey('patients.id'),
        nullable=True,
        index=True
    )

    owner_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey('owners.id'),
        nullable=True,
        index=True
    )

    notes: Mapped[str] = mapped_column(Text, default='')

    discount_percent: Mapped[float] = mapped_column(Float, default=0)
    discount_amount: Mapped[float] = mapped_column(Float, default=0)
    credit_surcharge_percent: Mapped[float] = mapped_column(Float, default=0)
    credit_surcharge_amount: Mapped[float] = mapped_column(Float, default=0)

    cost_total: Mapped[float] = mapped_column(Float, default=0)
    profit_amount: Mapped[float] = mapped_column(Float, default=0)
    margin_percent: Mapped[float] = mapped_column(Float, default=0)


class SaleItem(Base):
    __tablename__ = 'sale_items'

    id: Mapped[int] = mapped_column(primary_key=True)

    sale_id: Mapped[int] = mapped_column(
        ForeignKey('sales.id'),
        index=True
    )

    product_id: Mapped[int] = mapped_column(
        ForeignKey('products.id'),
        index=True
    )

    quantity: Mapped[float] = mapped_column(Float, default=1)
    unit_price: Mapped[float] = mapped_column(Float, default=0)
    subtotal: Mapped[float] = mapped_column(Float, default=0)


class SalePayment(Base):
    __tablename__ = 'sale_payments'

    id: Mapped[int] = mapped_column(primary_key=True)

    sale_id: Mapped[int] = mapped_column(
        ForeignKey('sales.id'),
        index=True
    )

    method: Mapped[str] = mapped_column(String(50), default='Efectivo')
    amount: Mapped[float] = mapped_column(Float, default=0)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=argentina_now)


class EventAttachment(Base):
    __tablename__ = 'event_attachments'

    id: Mapped[int] = mapped_column(primary_key=True)
    event_id: Mapped[int] = mapped_column(ForeignKey('clinical_events.id'), index=True)
    filename: Mapped[str] = mapped_column(String(255))
    file_path: Mapped[str] = mapped_column(String(500))

    event: Mapped['ClinicalEvent'] = relationship(back_populates='attachments')
