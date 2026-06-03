# Sistema de Gestión de Vacaciones 2026
## Guía de despliegue paso a paso

### LO QUE NECESITAS (todo gratis)
- Cuenta en github.com
- Cuenta en share.streamlit.io

---

### PASO 1 — Preparar archivos (5 min)
1. Abre `.streamlit/secrets_TEMPLATE.toml`
2. Cambia las contraseñas por contraseñas reales y seguras
3. Agrega todos los usuarios con sus correos y contraseñas
4. Renómbralo a `secrets.toml`
5. Copia los archivos de datos a esta carpeta:
   - `CONSOLIDADO_GENERADO.xlsx`
   - `Vacaciones_-_Dias_solicitados__28_.xlsx`
   - `person_access.json`
   - `area_to_admin.json`
   - `cargo_to_cat.json`
   - `jerarquia.csv`
   - `areas.csv`

### PASO 2 — Subir a GitHub (5 min)
1. Ve a github.com → New repository
2. Nombre: `vacaciones-2026` → Private → Create
3. Sube todos los archivos de esta carpeta EXCEPTO `secrets.toml`
   (los secretos nunca van a GitHub)

### PASO 3 — Publicar en Streamlit (5 min)
1. Ve a share.streamlit.io → New app
2. Conecta tu GitHub → selecciona el repo `vacaciones-2026`
3. Main file: `app.py`
4. Advanced settings → Secrets → pega el contenido de `secrets.toml`
5. Deploy!

### PASO 4 — Dar acceso a los jefes
1. Una vez publicada, copia la URL (ej: vacaciones-empresa.streamlit.app)
2. Envía la URL + su contraseña a cada jefe por correo
3. Para dar/quitar acceso: edita el secrets en Streamlit Cloud

### ACTUALIZACIÓN MENSUAL
1. Exporta los 4 archivos de Visma
2. Ejecuta: `python consolidador.py` (genera CONSOLIDADO_GENERADO.xlsx)
3. Sube el CONSOLIDADO_GENERADO.xlsx nuevo a GitHub
4. Streamlit se actualiza automáticamente en segundos

### CUANDO ROTA UN JEFE
1. Actualiza `Jerarquia_Sistema_Vacaciones_v2.xlsx`
2. Ejecuta de nuevo el consolidador
3. Si cambia quién accede: edita secrets en Streamlit Cloud
