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

# Logger initial konfigurieren
logger.remove()
logger.add(sys.stderr, level="INFO")


class VoiceAssistant:
    def __init__(self):
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
        logger.remove()
        logger.add(sys.stderr, level="INFO")

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
                    # Logging verkürzen, um Konsole nicht zu fluten
                    voice_names = [getattr(v, 'name', 'Unknown') for v in voices]
                    logger.info("Verfügbare Stimmen: {}", voice_names)

                    # Simpler Algo für deutsche Stimme (Plattformübergreifend)
                    for v in voices:
                        v_name = getattr(v, 'name', '') or ''
                        v_langs = getattr(v, 'languages', []) or []
                        v_id = getattr(v, 'id', '') or ''

                        # Check: Name oder ID enthält 'german'/'de' oder Sprache ist gesetzt
                        s_check = str(v_name).lower() + str(v_id).lower() + str(v_langs)
                        if 'hedda' in s_check or 'german' in s_check or 'de' in s_check:
                            preferred_voice_id = v.id
                            break

                    # Fallback
                    if voices and preferred_voice_id is None:
                        preferred_voice_id = voices[0].id

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
                    # Kleiner Sleep hilft oft beim Aufräumen von Treibern
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
                            # Fallback auf Standard, wenn ID fehlschlägt
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
                    # Kurze Pause damit Audio-Buffer geleert werden kann
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


def main():
    multiprocessing.freeze_support()
    assistant = VoiceAssistant()
    assistant.run()


if __name__ == '__main__':
    main()
