# python
import pyttsx3
from loguru import logger
import sys

logger.add(sys.stderr, level="INFO")

def main():
    engine = pyttsx3.init('sapi5')  # explizit SAPI5 auf Windows
    logger.info("TTS initialisiert")
    voices = engine.getProperty('voices')
    for i, v in enumerate(voices):
        logger.info("Voice {}: id='{}' name='{}' langs='{}'", i, getattr(v, 'id', None), getattr(v, 'name', None), getattr(v, 'languages', None))
    if voices:
        engine.setProperty('voice', voices[0].id)  # ggf. Index ändern
    engine.setProperty('volume', 1.0)  # 0.0 .. 1.0
    engine.setProperty('rate', 150)
    engine.say("Test. Wenn Sie das hören, funktioniert TTS.")
    engine.runAndWait()

if __name__ == '__main__':
    main()
