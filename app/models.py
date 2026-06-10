from datetime import datetime, date
from typing import Optional

from sqlalchemy import String, Text, Date, DateTime, Float, ForeignKey, Boolean
from sqlalchemy.orm import Mapped, mapped_column, relationship
from .database import Base

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
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
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
    events: Mapped[list['ClinicalEvent']] = relationship(back_populates='patient', order_by='desc(ClinicalEvent.event_date)')

class ClinicalEvent(Base):
    __tablename__ = 'clinical_events'

    id: Mapped[int] = mapped_column(primary_key=True)
    patient_id: Mapped[int] = mapped_column(ForeignKey('patients.id'), index=True)
    event_date: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, index=True)
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

    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
class EventAttachment(Base):
    __tablename__ = 'event_attachments'

    id: Mapped[int] = mapped_column(primary_key=True)
    event_id: Mapped[int] = mapped_column(ForeignKey('clinical_events.id'), index=True)
    filename: Mapped[str] = mapped_column(String(255))
    file_path: Mapped[str] = mapped_column(String(500))

    event: Mapped['ClinicalEvent'] = relationship(back_populates='attachments')
   
