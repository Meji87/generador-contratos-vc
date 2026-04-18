# Generador de Contratos de Financiación

App Streamlit para generar contratos Word rellenos automáticamente a partir de datos de clientes en Excel.

## Uso

1. Abre la app en https://generador-contratos-vc.streamlit.app
2. Introduce la contraseña de acceso
3. Sube la plantilla Word (`.docx`) y el Excel de la producción (`.xlsx`)
4. Selecciona el cliente y pulsa **Generar Contrato**
5. Descarga el `.docx` generado

---

## Estructura del repositorio

```
generador-contratos-vc/
├── app.py                  ← código principal
├── requirements.txt        ← dependencias Python
├── README.md               ← este archivo
└── plantillas/
    ├── Contrato_Plantilla_MEJORADA.docx   ← plantilla Word con marcadores
    └── Plantilla_Contratos_Financiacion.xlsx  ← Excel vacío para rellenar
```

---

## Cambiar la contraseña de acceso

**Paso 1 — Genera el hash de tu nueva contraseña**

Ve a: https://emn178.github.io/online-tools/sha256.html

- Escribe tu nueva contraseña en el campo **Input**
- Copia el texto del campo **Output** (es el hash)

O alternativamente ejecuta en Python:
```python
import hashlib
print(hashlib.sha256("tu_nueva_contraseña".encode()).hexdigest())
```

**Paso 2 — Actualiza el Secret en Streamlit Cloud**

1. Ve a https://share.streamlit.io
2. Pulsa los **⋮ tres puntos** de la app → **Settings** → **Secrets**
3. Borra el contenido actual y pega esto (en dos líneas exactas):

```
[auth]
password_hash = "PEGA_AQUI_EL_NUEVO_HASH"
```

4. Sustituye `PEGA_AQUI_EL_NUEVO_HASH` por el hash que copiaste en el Paso 1
5. Pulsa **Save changes**

La app se reinicia automáticamente con la nueva contraseña. No hay que tocar el código.

---

## Actualizar el código

Cualquier cambio en el código se despliega automáticamente al hacer push a GitHub:

```powershell
git add .
git commit -m "descripción del cambio"
git push
```

---

## Añadir o quitar usuarios con acceso

En la app (una vez dentro) pulsa **Share** → **Invite** → introduce el email del nuevo usuario.

Para quitar acceso: **Share** → pulsa la X junto al email.

---

## Estructura del Excel

El Excel de cada producción tiene 4 pestañas:

| Pestaña | Contenido |
|---|---|
| `PRODUCTORA` | Datos de la empresa productora (una sola fila) |
| `PERSONAS_FISICAS` | Clientes persona física (una fila por cliente) |
| `PERSONAS_JURIDICAS` | Clientes persona jurídica (una fila por cliente) |
| `INSTRUCCIONES` | Referencia de campos y marcadores |

Descarga la plantilla vacía en la carpeta `plantillas/` de este repositorio.
