import sys
import multiprocessing
import os
import yaml
from loguru import logger
from voice_assistant import VoiceAssistant

def load_config():
    config_path = "config.yaml"
    default_config = {
        "logging": {"level": "INFO"},
        "tts": {"language_search": "english"}
    }
    
    if not os.path.exists(config_path):
        logger.warning(f"Konfigurationsdatei {config_path} nicht gefunden. Nutze Standardwerte.")
        return default_config

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
            return config if config else default_config
    except Exception as e:
        logger.error(f"Fehler beim Laden der Konfiguration: {e}")
        return default_config

def main():
    multiprocessing.freeze_support()
    
    # Konfiguration laden
    config = load_config()
    log_level = config.get("logging", {}).get("level", "INFO")
    tts_lang_search = config.get("tts", {}).get("language_search", "english")

    # Logger für den Hauptprozess konfigurieren
    logger.remove()
    logger.add(sys.stderr, level=log_level)

    # Assistant starten und Config übergeben
    assistant = VoiceAssistant(log_level=log_level, language_search=tts_lang_search)
    assistant.run()

if __name__ == '__main__':
    main()
