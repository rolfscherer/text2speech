import sys
import multiprocessing
import queue
import time
import pyttsx3
from loguru import logger

# Plattform-Check
IS_WINDOWS = sys.platform == 'win32'

# COM nur auf Windows importieren
if IS_WINDOWS:
    try:
        import comtypes
    except ImportError:
        comtypes = None
else:
    comtypes = None

class VoiceAssistant:
    def __init__(self, log_level="INFO", language_search="english"):
        self.log_level = log_level
        self.language_search = language_search
        
        self._queue = multiprocessing.Queue()
        self._stop_event = multiprocessing.Event()
        self._ready_event = multiprocessing.Event()
        self.running = False

        self._process = multiprocessing.Process(target=self._tts_worker, daemon=False)
        self._process.start()

    def _tts_worker(self):
        """
        Plattformunabhängiger Worker-Prozess.
        """
        # Logger im neuen Prozess konfigurieren
        logger.remove()
        logger.add(sys.stderr, level=self.log_level)

        # Windows-spezifische Initialisierung
        if IS_WINDOWS and comtypes is not None:
            try:
                comtypes.CoInitializeEx(2)
            except Exception as e:
                logger.debug("WIN: comtypes.CoInitializeEx fehlgeschlagen: {}", e)

        preferred_voice_id = None

        # --- TEIL 1: Initialer Test & Stimmensuche ---
        try:
            try:
                # Automatische Treiberwahl (sapi5 auf Win, espeak auf Linux)
                test_engine = pyttsx3.init()
            except Exception as e:
                logger.error("TTS-Engine Init fehlgeschlagen: {}", e)
                test_engine = None

            if test_engine is not None:
                try:
                    voices = test_engine.getProperty('voices')
                    # Logging verkürzen
                    voice_names = [getattr(v, 'name', 'Unknown') for v in voices]
                    logger.debug("Verfügbare Stimmen: {}", voice_names)

                    # Simpler Algo für Stimme basierend auf Config
                    search_term = self.language_search.lower()
                    
                    for v in voices:
                        v_name = getattr(v, 'name', '') or ''
                        v_langs = getattr(v, 'languages', []) or []
                        v_id = getattr(v, 'id', '') or ''

                        # Check: Name oder ID enthält Suchbegriff oder Sprache ist gesetzt
                        s_check = str(v_name).lower() + str(v_id).lower() + str(v_langs)
                        
                        # Spezifische Suche nach Config-Wert
                        if search_term in s_check:
                            preferred_voice_id = v.id
                            break
                        
                        # Fallback für 'german'/'de' falls Config das sagt
                        if (search_term == 'german' or search_term == 'de') and ('hedda' in s_check or 'de' in s_check):
                             preferred_voice_id = v.id
                             break

                    # Fallback
                    if voices and preferred_voice_id is None:
                        preferred_voice_id = voices[0].id
                        logger.warning("Keine passende Stimme für '{}' gefunden. Nutze Standard: {}", search_term, getattr(voices[0], 'name', 'Unknown'))
                    else:
                        logger.info("Stimme ausgewählt: {}", preferred_voice_id)

                    test_engine.setProperty('volume', 1.0)
                    test_engine.setProperty('rate', 150)

                    if preferred_voice_id:
                        try:
                            test_engine.setProperty('voice', preferred_voice_id)
                        except:
                            pass

                    logger.info("Worker: Initialer Test-Say")
                    test_engine.say("System bereit")
                    test_engine.runAndWait()
                except Exception:
                    logger.exception("Fehler beim Test-Sprechen")
                finally:
                    try:
                        test_engine.stop()
                    except:
                        pass
                    time.sleep(0.1)
                    try:
                        del test_engine
                    except:
                        pass
        finally:
            self._ready_event.set()

        # --- TEIL 2: Hauptschleife ---
        try:
            while not self._stop_event.is_set():
                try:
                    text = self._queue.get(timeout=0.1)
                except queue.Empty:
                    continue

                if text is None:
                    break

                logger.info("Speaking: {}", text)

                engine = None
                try:
                    # Neue Engine pro Satz (für Isolation)
                    engine = pyttsx3.init()

                    if preferred_voice_id:
                        try:
                            engine.setProperty('voice', preferred_voice_id)
                        except:
                            pass

                    engine.setProperty('volume', 1.0)
                    engine.setProperty('rate', 150)

                    engine.say(text)
                    engine.runAndWait()
                except Exception:
                    logger.exception("Fehler beim Sprechen")
                finally:
                    if engine:
                        try:
                            engine.stop()
                        except:
                            pass
                        try:
                            del engine
                        except:
                            pass
                    time.sleep(0.05)

        finally:
            if IS_WINDOWS and comtypes is not None:
                try:
                    comtypes.CoUninitialize()
                except:
                    pass
            logger.info("TTS Worker beendet")

    def speak(self, text: str):
        self._queue.put(text)

    def run(self):
        logger.info("Starting Voice Assistant (Multi-Platform)")
        self.running = True
        try:
            if not self._ready_event.wait(timeout=10):
                logger.warning("TTS-Prozess nicht rechtzeitig bereit")

            self.speak("Ich bin bereit")

            while self.running:
                try:
                    text = input("Say > ").strip()
                except EOFError:
                    break

                if not text:
                    continue
                if text.lower() in ("quit", "exit"):
                    break
                self.speak(text)
        except KeyboardInterrupt:
            pass
        finally:
            self.shutdown()

    def shutdown(self):
        logger.info("Shutting down...")
        self.running = False
        self._stop_event.set()
        self._queue.put(None)
        self._process.join(timeout=5)
        if self._process.is_alive():
            self._process.terminate()
