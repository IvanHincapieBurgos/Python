"""
Test: Análisis de llamadas IA
--------------------------------------------------------------
Sube un archivo MP3, lo analiza y devuelve un JSON con el resumen de la llamada + métricas de consumo de tokens.

Requisitos:
    pip install google-generativeai

Uso:
    1. Reemplaza API_KEY en el archivo .env con tu clave de Google AI Studio
    2. Reemplaza AUDIO_FILE en el archivo .env con la ruta a tu archivo MP3
    3. Ejecuta: python analizar_llamada.py
"""

import google.generativeai as genai
import json
import time
import os
import dotenv 

# ─────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────
dotenv.load_dotenv()
API_KEY= os.getenv("API_KEY")
AUDIO_FILE = os.getenv("AUDIO_FILE")
# ─────────────────────────────────────────
 
PROMPT = """
You're a customer service phone quality analyst, bilingual and who listens calls in spanish.
Listen the entire call, omit silence time. Return ONLY a valid JSON object, with no additional text, no markdown code blocks.

JSON keys must be exactly as specified, and values must be one of the allowed options or a brief text as indicated.
{
  "resumen": {
    "general_descripcion": "Describe shortly what the call was about"
    ,"contact_reason": "What was the main reason for the customer to call?"
    ,"resolution": "Was the issue resolved? (Yes/No/Partially)"
    ,"resolution_description": "What was done or left pending?"
    ,"estimated_duration": "Approximate duration in minutes"
    ,"agent_tone": "Professional / Neutral / Negative"
    ,"customer_tone": "Satisfied / Neutral / Unsatisfied"
  }
}
"""

def analizar_llamada(api_key: str, audio_path: str) -> dict:
    # Validar que el archivo existe
    if not os.path.exists(audio_path):
        raise FileNotFoundError(f"No se encontró el archivo: {audio_path}")

    print(f"📁 Archivo encontrado: {audio_path}")
    file_size_mb = os.path.getsize(audio_path) / (1024 * 1024)
    print(f"   Tamaño: {file_size_mb:.2f} MB\n")

    # Configurar cliente
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.5-flash")

    # Subir el audio a la API de Google
    print("⬆️  Subiendo audio a Google AI...")
    t_inicio_upload = time.time()

    audio_file = genai.upload_file(
        path=audio_path,
        mime_type="audio/mpeg"
    )

    # Esperar a que el archivo esté listo
    while audio_file.state.name == "PROCESSING":
        time.sleep(2)
        audio_file = genai.get_file(audio_file.name)

    if audio_file.state.name == "FAILED":
        raise RuntimeError("El archivo falló al procesarse en Google.")

    t_fin_upload = time.time()
    print(f"   ✅ Audio listo en {t_fin_upload - t_inicio_upload:.1f}s\n")

    # Llamar al modelo
    print("🤖 Analizando llamada con Gemini 2.5 Flash...")
    t_inicio_llm = time.time()

    response = model.generate_content(
        [audio_file, PROMPT],
        generation_config=genai.GenerationConfig(
            temperature=0.2,   # Bajo para respuestas más consistentes
        )
    )

    t_fin_llm = time.time()
    print(f"   ✅ Análisis completado en {t_fin_llm - t_inicio_llm:.1f}s\n")

    # Limpiar y parsear JSON
    raw_text = response.text.strip()
    # Remover posibles bloques markdown si el modelo los incluye
    if raw_text.startswith("```"):
        raw_text = raw_text.split("```")[1]
        if raw_text.startswith("json"):
            raw_text = raw_text[4:]
    raw_text = raw_text.strip()

    resultado = json.loads(raw_text)

    # ── Métricas de consumo ──────────────────────────────────────────
    usage = response.usage_metadata
    tokens_audio  = usage.prompt_token_count      # incluye audio + prompt
    tokens_output = usage.candidates_token_count
    tokens_total  = usage.total_token_count

    # Precios Gemini 2.5 Flash (USD por 1M tokens, Marzo 2026)
    PRECIO_INPUT_POR_MILLON  = 1.00   # audio/texto input
    PRECIO_OUTPUT_POR_MILLON = 3.50   # output

    costo_input  = (tokens_audio  / 1_000_000) * PRECIO_INPUT_POR_MILLON
    costo_output = (tokens_output / 1_000_000) * PRECIO_OUTPUT_POR_MILLON
    costo_total  = costo_input + costo_output

    consumo = {
        "tokens_input":  tokens_audio,
        "tokens_output": tokens_output,
        "tokens_total":  tokens_total,
        "costo_usd": {
            "input":  round(costo_input,  6),
            "output": round(costo_output, 6),
            "total":  round(costo_total,  6)
        },
        "proyeccion_150_llamadas_usd": round(costo_total * 150, 4)
    }

    return {
        "archivo":  os.path.basename(audio_path),
        "modelo":   "gemini-2.5-flash",
        "analisis": resultado,
        "consumo":  consumo
    }


if __name__ == "__main__":
    try:
        resultado = analizar_llamada(API_KEY, AUDIO_FILE)

        # Mostrar resultado en consola
        print("─" * 50)
        print("📊 RESULTADO:")
        print("─" * 50)
        print(json.dumps(resultado, ensure_ascii=False, indent=2))

        # Guardar en archivo
        output_file = "resultado_qa.json"
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(resultado, f, ensure_ascii=False, indent=2)

        print(f"\n💾 Resultado guardado en: {output_file}")

        # Resumen de costos en consola
        consumo = resultado["consumo"]
        print(f"\n💰 CONSUMO ESTIMADO:")
        print(f"   Tokens usados:     {consumo['tokens_total']:,}")
        print(f"   Costo esta llamada: ${consumo['costo_usd']['total']:.6f} USD")
        print(f"   Proyección 150 llamadas: ${consumo['proyeccion_150_llamadas_usd']:.4f} USD/mes")

    except FileNotFoundError as e:
        print(f"\n❌ Error: {e}")
        print("   Verifica que AUDIO_FILE apunte al MP3 correcto.")
    except json.JSONDecodeError:
        print("\n❌ El modelo no devolvió JSON válido. Intenta correr el script de nuevo.")
    except Exception as e:
        print(f"\n❌ Error inesperado: {e}")