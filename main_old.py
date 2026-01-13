# python
from loguru import logger
import pyttsx3
import sys
import threading
import queue

# Logger: entferne vorhandene Sinks, dann eigenen hinzufügen (vermeidet doppelte Einträge)
logger.remove()
logger.add(sys.stderr, level="INFO")

try:
    import comtypes
except Exception:
    comtypes = None
    logger.debug("comtypes nicht verfügbar; COM-Initialisierung im Worker wird übersprungen")

class VoiceAssistant:
    def __init__(self):
        self._queue = queue.Queue()
        self._stop_event = threading.Event()
        self._ready_event = threading.Event()
        self.running = False
        self._worker = threading.Thread(target=self._tts_worker, daemon=False)
        self._worker.start()

    def _tts_worker(self):
        # COM für diesen Thread im STA initialisieren
        if comtypes is not None:
            try:
                comtypes.CoInitializeEx(2)
            except Exception as e:
                logger.debug("comtypes.CoInitializeEx fehlgeschlagen: {}", e)

        # kurzer hörbarer Test mit temporärer Engine und Ermittlung einer bevorzugten Stimme
        preferred_voice_id = None
        try:
            try:
                test_engine = pyttsx3.init('sapi5')
            except Exception as e:
                logger.error("Test-TTS-Engine konnte nicht initialisiert werden: {}", e)
                test_engine = None

            if test_engine is not None:
                try:
                    voices = test_engine.getProperty('voices')
                    logger.info("Test-Engine voices: {}", [getattr(v, 'name', None) for v in voices])
                    # bevorzugte deutsche Stimme suchen
                    for v in voices:
                        name = getattr(v, 'name', '') or ''
                        langs = getattr(v, 'languages', '') or ''
                        if 'Hedda' in name or 'German' in name or 'de' in str(langs):
                            preferred_voice_id = v.id
                            break
                    if voices and preferred_voice_id is None:
                        preferred_voice_id = voices[0].id

                    test_engine.setProperty('volume', 1.0)
                    test_engine.setProperty('rate', 150)

                    logger.info("Worker: initialer Test-Say")
                    test_engine.say("T T S Worker bereit")
                    test_engine.runAndWait()
                except Exception:
                    logger.exception("Fehler beim initialen Test-Sprechen (Test-Engine)")
                finally:
                    try:
                        test_engine.stop()
                    except Exception:
                        pass
                    threading.Event().wait(0.15)
                    try:
                        del test_engine
                    except Exception:
                        pass
        finally:
            # Signalisieren, dass der Worker bereit ist (Test durch)
            self._ready_event.set()

        try:
            # Für jede Ansage: neue Engine erstellen, konfigurieren, sprechen, zerstören
            while not self._stop_event.is_set():
                try:
                    text = self._queue.get(timeout=0.1)
                except queue.Empty:
                    continue
                if text is None:  # Sentinel zum Beenden
                    break
                try:
                    logger.info("Speaking (worker): {}", text)
                    try:
                        engine = pyttsx3.init('sapi5')
                    except Exception as e:
                        logger.error("Per-Utterance-Engine konnte nicht initialisiert werden: {}", e)
                        continue

                    try:
                        # Stimme/Parameter setzen
                        voices = engine.getProperty('voices')
                        if preferred_voice_id:
                            try:
                                engine.setProperty('voice', preferred_voice_id)
                            except Exception:
                                # fallback: falls id ungültig, wähle erste
                                if voices:
                                    engine.setProperty('voice', voices[0].id)
                        else:
                            if voices:
                                engine.setProperty('voice', voices[0].id)
                        engine.setProperty('volume', 1.0)
                        engine.setProperty('rate', 150)
                    except Exception:
                        logger.exception("Fehler beim Setzen der TTS-Parameter (per-utterance)")

                    try:
                        engine.say(text)
                        engine.runAndWait()
                    except Exception:
                        logger.exception("Fehler beim Sprechen (per-utterance)")
                    finally:
                        try:
                            engine.stop()
                        except Exception:
                            pass
                        try:
                            del engine
                        except Exception:
                            pass
                        # kurze Pause, um SAPI intern zu beruhigen
                        threading.Event().wait(0.05)
                except Exception:
                    logger.exception("Fehler im Worker-Hauptloop")
        finally:
            if comtypes is not None:
                try:
                    comtypes.CoUninitialize()
                except Exception:
                    pass

    def speak(self, text: str):
        logger.info("Enqueue speak: {}", text)
        self._queue.put(text)

    def run(self):
        logger.info("Starting Voice Assistant")
        self.running = True
        try:
            if not self._ready_event.wait(timeout=5):
                logger.warning("TTS-Worker nicht rechtzeitig bereit; erste Ansage könnte fehlen")
            self.speak("Ich bin bereit")
            while self.running:
                try:
                    text = input("Say (type) something (or 'quit')> ").strip()
                except EOFError:
                    logger.info("EOF on input, stopping")
                    break
                if not text:
                    continue
                if text.lower() in ("quit", "exit"):
                    logger.info("Stopping on user request")
                    break
                self.speak(text)
        except KeyboardInterrupt:
            logger.info("Interrupted by user")
        finally:
            self.shutdown()

    def shutdown(self):
        logger.info("Shutting down TTS")
        self.running = False
        self._stop_event.set()
        self._queue.put(None)
        self._worker.join(timeout=5)



def main():
    assistant = VoiceAssistant()
    assistant.run()

if __name__ == '__main__':
    main()
