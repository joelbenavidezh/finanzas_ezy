"""
Módulo de Extracción Inteligente con IA (Google Gemini)
Extrae información estructurada a partir de fotos de tickets o resúmenes de tarjeta (PDF/Imagen).
Lee la API Key de forma segura y automática desde el archivo .env
"""

import os
import io
import json
import base64
import requests
from datetime import datetime
from dotenv import load_dotenv
from PIL import Image, ImageOps

# Cargar automáticamente las variables de entorno de .env
ENV_FILE = os.path.join(os.path.dirname(__file__), ".env")
load_dotenv(ENV_FILE, override=True)

CATEGORIAS_DISPONIBLES = [
    "Supermercado",
    "Restaurante y Bares",
    "Servicios y Facturas",
    "Farmacia y Salud",
    "Transporte y Combustible",
    "Ocio y Entretenimiento",
    "Ropa y Calzado",
    "Hogar y Mantenimiento",
    "Suscripciones y Apps",
    "Educación",
    "Otros"
]

# Lista de modelos compatibles con fallback automático si hay alta demanda
MODELOS_IA = [
    "gemini-3.5-flash",
    "gemini-3.8-flash",
    "gemini-3.7-flash",
    "gemini-flash-latest"
]

def optimizar_imagen(image_bytes, max_dim=1600, quality=85):
    """
    Corrige la orientación de la foto (cámara del móvil) y comprime el peso
    (de 3-10 MB a ~250 KB) para que la subida a Gemini sea casi instantánea
    sin perder nada de nitidez en el texto del ticket.
    """
    try:
        img = Image.open(io.BytesIO(image_bytes))
        img = ImageOps.exif_transpose(img)
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        
        w, h = img.size
        if max(w, h) > max_dim:
            scale = max_dim / max(w, h)
            img = img.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)
            
        out = io.BytesIO()
        img.save(out, format="JPEG", quality=quality, optimize=True)
        return out.getvalue(), "image/jpeg"
    except Exception:
        return image_bytes, "image/jpeg"

def obtener_api_key(override_key=None):
    """Obtiene la clave de Gemini leyendo directamente del archivo .env para evitar problemas de caché."""
    if override_key and override_key.strip():
        return override_key.strip()
    
    from dotenv import dotenv_values
    valores = dotenv_values(ENV_FILE)
    key = valores.get("GEMINI_API_KEY", "") or os.getenv("GEMINI_API_KEY", "")
    key = key.strip()
    if not key:
        raise Exception(
            "No se encontró la clave GEMINI_API_KEY en el archivo .env. "
            "Asegúrate de guardarla en el archivo .env de tu carpeta."
        )
    return key

def llamar_gemini(payload, api_key=None):
    """Realiza la llamada HTTP a la API de Gemini Flash con fallback automático."""
    key = obtener_api_key(api_key)
    headers = {"Content-Type": "application/json"}
    
    ultimo_error = None
    for modelo in MODELOS_IA:
        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{modelo}:generateContent?key={key}"
            response = requests.post(url, headers=headers, json=payload, timeout=60)
            
            if response.status_code == 200:
                data = response.json()
                candidates = data.get("candidates", [])
                if not candidates:
                    continue
                content = candidates[0].get("content", {})
                parts = content.get("parts", [])
                if not parts:
                    continue
                raw_text = parts[0].get("text", "")
                return json.loads(raw_text)
            elif response.status_code in [503, 429, 404]:
                ultimo_error = response.text
                continue
            else:
                raise Exception(f"Error ({response.status_code}): {response.text}")
        except Exception as e:
            ultimo_error = str(e)
            continue
            
    raise Exception(f"No se pudo procesar con ningún modelo disponible: {ultimo_error}")

def extraer_ticket(image_bytes, mime_type, api_key=None):
    """
    Analiza una foto de un ticket o recibo de compra.
    Devuelve un diccionario con los datos del gasto.
    """
    # Optimizar imagen para subida ultra rápida
    img_optimizada, mime_opt = optimizar_imagen(image_bytes)
    base64_data = base64.b64encode(img_optimizada).decode("utf-8")
    hoy = datetime.now().strftime("%Y-%m-%d")
    
    prompt = f"""
Eres un asistente contable y financiero experto. Analiza la imagen de este ticket/factura/recibo de compra.
Ten en cuenta que el ticket puede estar orientado en cualquier dirección (horizontal o vertical).
Extrae estrictamente un objeto JSON con las siguientes claves:
- "fecha": Fecha de la compra en formato YYYY-MM-DD (si no es legible o no figura, usa "{hoy}").
- "comercio": Nombre comercial de la tienda, negocio o empresa.
- "monto": Importe total final pagado (número flotante positivo, sin símbolos de moneda).
- "moneda": Código de 3 letras de la moneda (ARS, USD, EUR, etc. Por defecto "ARS").
- "categoria": La categoría más precisa eligiendo estrictamente entre: {CATEGORIAS_DISPONIBLES}.
- "medio_pago": Si figura, elige entre ['Efectivo', 'Tarjeta de Débito', 'Tarjeta de Crédito', 'Transferencia / QR']. Si no es claro, pon 'Efectivo'.
- "cuotas": Número de cuotas si fue en tarjeta (ej: '1', '03/06', etc. Por defecto '1').
- "notas": Breve resumen de los principales productos o detalles comprados (máximo 1 renglón).

Responde únicamente con el objeto JSON válido.
"""

    payload = {
        "contents": [{
            "parts": [
                {"text": prompt},
                {
                    "inline_data": {
                        "mime_type": mime_opt,
                        "data": base64_data
                    }
                }
            ]
        }],
        "generationConfig": {
            "response_mime_type": "application/json"
        }
    }
    
    return llamar_gemini(payload, api_key)

def extraer_resumen_tarjeta(file_bytes, mime_type, api_key=None):
    """
    Analiza un resumen de tarjeta de crédito (PDF o imagen).
    Extrae la lista de todos los consumos detallados.
    """
    if mime_type.startswith("image/"):
        file_bytes, mime_type = optimizar_imagen(file_bytes)
        
    base64_data = base64.b64encode(file_bytes).decode("utf-8")

    prompt = f"""
Eres un auditor financiero experto. Analiza este resumen/extracto de tarjeta de crédito (Visa, Mastercard, Amex, etc.).
Debes ubicar la tabla de consumos o liquidación del período y extraer CADA UNO de los consumos realizados.
Ignora pagos recibidos, saldo anterior, débitos de intereses/impuestos de la tarjeta si no son consumos directos (o inclúyelos como 'Servicios y Facturas' si corresponden a gastos del usuario).

Devuelve estrictamente un objeto JSON con una clave "consumos" que contenga una lista de objetos:
{{
  "consumos": [
    {{
      "fecha": "YYYY-MM-DD",
      "comercio": "Nombre del comercio o establecimiento",
      "monto": 12500.50,
      "moneda": "ARS" o "USD",
      "cuotas": "ej: 02/06 o 1",
      "categoria": "Elegir estrictamente una de: {CATEGORIAS_DISPONIBLES}",
      "medio_pago": "Tarjeta de Crédito",
      "notas": "Detalle breve si está disponible"
    }}
  ]
}}

Reglas:
- Asegúrate de que "monto" sea siempre un número decimal positivo.
- En "categoria", infiere inteligentemente el rubro según el nombre del comercio (ej: 'Carrefour' -> 'Supermercado', 'YPF' o 'Shell' -> 'Transporte y Combustible', 'Netflix' o 'Spotify' -> 'Suscripciones y Apps', 'Farmacity' -> 'Farmacia y Salud', etc.).
- Si una fecha no incluye el año, usa el año actual ({datetime.now().year}).

Responde únicamente con el JSON.
"""

    payload = {
        "contents": [{
            "parts": [
                {"text": prompt},
                {
                    "inline_data": {
                        "mime_type": mime_type,
                        "data": base64_data
                    }
                }
            ]
        }],
        "generationConfig": {
            "response_mime_type": "application/json"
        }
    }
    
    resultado = llamar_gemini(payload, api_key)
    return resultado.get("consumos", [])
