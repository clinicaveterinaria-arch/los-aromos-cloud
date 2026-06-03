# Los Aromos Cloud

Sistema de gestión veterinaria para Clínica Veterinaria Los Aromos.

## Versión base V1
Incluye:
- Login básico
- Propietarios
- Pacientes
- Historia clínica cronológica
- Módulo inicial Migración MyVete
- Diseño rosa/gris con logo

## Render
Build command:
```bash
pip install -r requirements.txt
```
Start command:
```bash
uvicorn app.main:app --host 0.0.0.0 --port $PORT
```

Variable requerida:
- DATABASE_URL
- SECRET_KEY

Usuario inicial:
- admin / losaromos
