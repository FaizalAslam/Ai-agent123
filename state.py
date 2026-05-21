from utils.command_buffer import CommandBuffer
from ai.openai_handler import OpenAIHandler

try:
    from listener.clipboard_listener import ClipboardListener
    CLIPBOARD_MODULE_AVAILABLE = True
except Exception:
    ClipboardListener = None
    CLIPBOARD_MODULE_AVAILABLE = False

try:
    from listener.keyboard_listener import KeyboardListener
    KEYBOARD_MODULE_AVAILABLE = True
except Exception:
    KeyboardListener = None
    KEYBOARD_MODULE_AVAILABLE = False

try:
    from listener.voice_listener import VoiceListener
    VOICE_MODULE_AVAILABLE = True
except Exception:
    VoiceListener = None
    VOICE_MODULE_AVAILABLE = False

# Shared OCR text store (written by hotkeys, polled by /ocr/poll route)
last_ocr = {"text": "", "pending": False}

# Voice toggle state used by /voice/* routes
voice_state = {"enabled": False}

# Single OpenAI handler instance shared across all request threads
openai_handler = OpenAIHandler()

# Command buffer for keyboard/clipboard listeners
_cmd_buf = CommandBuffer()

# Listeners are initialized with None callback; patched in server.py after
# _handle_global_command is defined.
_clipboard_listener = ClipboardListener(_cmd_buf) if CLIPBOARD_MODULE_AVAILABLE else None
_keyboard_listener  = KeyboardListener(None, _cmd_buf) if KEYBOARD_MODULE_AVAILABLE else None
_voice_listener     = VoiceListener(None) if VOICE_MODULE_AVAILABLE else None
