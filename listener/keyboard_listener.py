import logging
import threading

from pynput import keyboard

logger = logging.getLogger("OfficeAgent")


class KeyboardListener:
    def __init__(self, on_command_callback, cmd_buf):
        self.on_command = on_command_callback
        self.cmd_buf = cmd_buf
        self._typed_chars = []

    def start(self):
        with keyboard.Listener(on_press=self._on_press) as listener:
            listener.join()

    def _current_line(self):
        return "".join(self._typed_chars).strip()

    def _reset_line(self):
        self._typed_chars.clear()

    def _on_press(self, key):
        try:
            if key == keyboard.Key.enter:
                typed_candidate = self._current_line()
                candidate = typed_candidate if typed_candidate.lower().startswith("agent:") else self.cmd_buf.get_candidate()
                self._reset_line()
                if candidate:
                    self.cmd_buf.clear()
                    logger.info(f"Keyboard trigger: {candidate}")
                    threading.Thread(
                        target=self.on_command,
                        args=(candidate,),
                        daemon=True,
                    ).start()
                return

            if key == keyboard.Key.backspace:
                if self._typed_chars:
                    self._typed_chars.pop()
                return

            if key == keyboard.Key.space:
                self._typed_chars.append(" ")
                return

            if key in (keyboard.Key.esc, keyboard.Key.tab):
                self._reset_line()
                return

            char = getattr(key, "char", None)
            if char and char.isprintable():
                self._typed_chars.append(char)
                if len(self._typed_chars) > 500:
                    self._typed_chars = self._typed_chars[-500:]
        except Exception as e:
            logger.error(f"Keyboard listener error: {e}")
