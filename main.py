# python
from loguru import logger
import pyttsx3
import sys

logger.add(sys.stderr, level="INFO")

class VoiceAssistant:
    def __init__(self):
        # Kein extra Leerzeichen vor dem Argument
        self.tts = pyttsx3.init(debug=True)
        self.running = False

    def speak(self, text: str):
        logger.info("Speaking: {}", text)
        self.tts.say(text)
        self.tts.runAndWait()

    def run(self):
        logger.info("Starting Voice Assistant")
        self.running = True
        try:
            self.speak("Voice assistant ready.")
            while self.running:
                text = input("Say (type) something (or 'quit')> ").strip()
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
        try:
            self.tts.stop()
        except Exception:
            pass

def main():
    assistant = VoiceAssistant()
    assistant.run()

if __name__ == '__main__':
    main()
