import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, Toplevel
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import json
import os
import threading
import soundfile as sf
import numpy as np
import pyaudio
import time
import shutil
import subprocess
import platform
import logging
import traceback
from collections import deque
from threading import Lock, Event
import pydub
import importlib.metadata
import re
import sys
import uuid
import resampy
import sounddevice as sd
from pynput import keyboard, mouse
import pygame
import enum

# --- Configuration and Constants ---
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
SOUNDS_DIR = os.path.join(ROOT_DIR, "sounds")
CONFIG_DIR = os.path.join(ROOT_DIR, "config")
CONFIG_FILE = os.path.join(CONFIG_DIR, "soundboard_config.json")
APP_SETTINGS_FILE = os.path.join(CONFIG_DIR, "app_settings.json")
LOG_FILE = os.path.join(ROOT_DIR, "warpboard.log")

VB_CABLE_URL = "https://download.vb-audio.com/Download_CABLE/VBCABLE_Driver_Pack43.zip"

# --- Audio Settings ---
SUPPORTED_FORMATS = [("Audio Files", "*.wav *.mp3 *.ogg *.flac")]
VIRTUAL_MIC_NAME_PARTIAL = "CABLE Input"
SAMPLE_RATE = 44100
CHANNELS = 2
FRAME_SIZE = 1024

# --- UI and Validation ---
INVALID_FILENAME_CHARS = r'[<>:"/\\|?*]'
MAX_DISPLAY_NAME_LENGTH = 25
MAX_DEVICE_NAME_LENGTH = 35

# --- Setup Logging ---
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

current_playing_sound_details = {}
current_playing_sound_details_lock = threading.Lock()

# --- Hotkey Enums and Classes ---
class KeyModifier(enum.Flag):
    NONE = 0
    SHIFT = enum.auto()
    CTRL = enum.auto()
    ALT = enum.auto()
    SUPER = enum.auto()

class KeyCode(enum.Enum):
    A = 'a'
    B = 'b'
    C = 'c'
    D = 'd'
    E = 'e'
    F = 'f'
    G = 'g'
    H = 'h'
    I = 'i'
    J = 'j'
    K = 'k'
    L = 'l'
    M = 'm'
    N = 'n'
    O = 'o'
    P = 'p'
    Q = 'q'
    R = 'r'
    S = 's'
    T = 't'
    U = 'u'
    V = 'v'
    W = 'w'
    X = 'x'
    Y = 'y'
    Z = 'z'
    
    NUM_0 = '0'
    NUM_1 = '1'
    NUM_2 = '2'
    NUM_3 = '3'
    NUM_4 = '4'
    NUM_5 = '5'
    NUM_6 = '6'
    NUM_7 = '7'
    NUM_8 = '8'
    NUM_9 = '9'

    F1 = 'f1'
    F2 = 'f2'
    F3 = 'f3'
    F4 = 'f4'
    F5 = 'f5'
    F6 = 'f6'
    F7 = 'f7'
    F8 = 'f8'
    F9 = 'f9'
    F10 = 'f10'
    F11 = 'f11'
    F12 = 'f12'

    SPACE = 'space'
    ENTER = 'enter'
    ESC = 'esc'
    TAB = 'tab'
    BACKSPACE = 'backspace'
    DELETE = 'delete'
    INSERT = 'insert'
    HOME = 'home'
    END = 'end'
    PAGE_UP = 'page_up'
    PAGE_DOWN = 'page_down'
    ARROW_UP = 'up'
    ARROW_DOWN = 'down'
    ARROW_LEFT = 'left'
    ARROW_RIGHT = 'right'
    
    MOUSE_MIDDLE = 'mouse_middle'
    MOUSE_X1 = 'mouse_x1'
    MOUSE_X2 = 'mouse_x2'
    UNKNOWN = 'unknown'

class Hotkey:
    def __init__(self, key_code: KeyCode, modifiers: KeyModifier = KeyModifier.NONE, raw_key: str = None):
        if not isinstance(key_code, KeyCode):
            raise TypeError("key_code must be an instance of KeyCode enum.")
        if not isinstance(modifiers, KeyModifier):
            raise TypeError("modifiers must be an instance of KeyModifier enum.")

        self.key_code = key_code
        self.modifiers = modifiers
        self.raw_key = raw_key if key_code == KeyCode.UNKNOWN else None

    def __hash__(self):
        return hash((self.key_code, self.modifiers, self.raw_key))

    def __eq__(self, other):
        if not isinstance(other, Hotkey):
            return NotImplemented
        return (self.key_code == other.key_code and 
                self.modifiers == other.modifiers and 
                self.raw_key == other.raw_key)

    def __repr__(self):
        modifier_parts = []
        if KeyModifier.CTRL in self.modifiers:
            modifier_parts.append('Ctrl')
        if KeyModifier.SHIFT in self.modifiers:
            modifier_parts.append('Shift')
        if KeyModifier.ALT in self.modifiers:
            modifier_parts.append('Alt')
        if KeyModifier.SUPER in self.modifiers:
            modifier_parts.append('Super')

        key_name = self.raw_key if self.key_code == KeyCode.UNKNOWN else self.key_code.value
        key_name = key_name.replace('_', ' ').title()
        if key_name.startswith('mouse_'):
            key_name = key_name.replace('Mouse ', 'Mouse-')
        elif key_name.startswith('<vk-') and key_name.endswith('>'):
            key_name = f"VK-{key_name[4:-1]}"

        if modifier_parts:
            return f"{'+'.join(modifier_parts)}+{key_name}"
        return key_name

    def to_json_serializable(self):
        parts = []
        if KeyModifier.CTRL in self.modifiers:
            parts.append('ctrl')
        if KeyModifier.SHIFT in self.modifiers:
            parts.append('shift')
        if KeyModifier.ALT in self.modifiers:
            parts.append('alt')
        if KeyModifier.SUPER in self.modifiers:
            parts.append('cmd')
        
        key_value = self.raw_key if self.key_code == KeyCode.UNKNOWN else self.key_code.value
        if key_value.startswith('mouse_button_') or key_value.startswith('mouse_') or key_value.startswith('<vk-'):
            parts.append(key_value)
        elif len(key_value) == 1 and key_value.isalpha():
            parts.append(key_value.lower())
        else:
            parts.append(key_value)
        
        return sorted(parts)

    @classmethod
    def from_json_serializable(cls, hotkey_list_strings: list[str]):
        modifiers = KeyModifier.NONE
        key_code = None
        raw_key = None
        
        for s in hotkey_list_strings:
            if s == 'ctrl':
                modifiers |= KeyModifier.CTRL
            elif s == 'shift':
                modifiers |= KeyModifier.SHIFT
            elif s == 'alt':
                modifiers |= KeyModifier.ALT
            elif s == 'cmd':
                modifiers |= KeyModifier.SUPER
            else:
                try:
                    key_code = KeyCode(s)
                    raw_key = None

                except ValueError:
                    if (s.startswith('<vk-') and s.endswith('>')) or s.startswith('mouse_button_') or s in ['mouse_middle', 'mouse_x1', 'mouse_x2']:
                        key_code = KeyCode.UNKNOWN
                        raw_key = s
                    elif s and all(c in map(chr, range(32, 127)) for c in s):  # printable chars
                        key_code = KeyCode.UNKNOWN
                        raw_key = s
                    else:
                        raise ValueError(f"Unknown key string encountered: {s}")

        
        if key_code is None:
            raise ValueError("Hotkey must contain at least one non-modifier key/button.")
        
        return cls(key_code=key_code, modifiers=modifiers, raw_key=raw_key)

# --- Utility Functions ---
def get_hotkey_display_string(hotkey_list_strings):
    if not hotkey_list_strings:
        return "Not Assigned"
    try:
        hotkey_obj = Hotkey.from_json_serializable(hotkey_list_strings)
        return str(hotkey_obj)
    except ValueError as e:
        logging.warning(f"Failed to convert hotkey strings {hotkey_list_strings} to Hotkey object for display: {e}")
        display_map = {
            'ctrl': 'Ctrl', 'alt': 'Alt', 'shift': 'Shift', 'cmd': 'Cmd',
            'space': 'Space', 'enter': 'Enter', 'backspace': 'Backspace',
            'caps_lock': 'Caps Lock', 'tab': 'Tab', 'esc': 'Esc',
            'up': 'Up', 'down': 'Down', 'left': 'Left', 'right': 'Right',
            'insert': 'Insert', 'Insert': 'Ins', 'delete': 'Del', 'home': 'Home', 'end': 'End',
            'page_up': 'PgUp', 'page_down': 'PgDn',
            'f1': 'F1', 'f2': 'F2', 'f3': 'F3', 'f4': 'F4', 'f5': 'F5', 'f6': 'F6',
            'f7': 'F7', 'f8': 'F8', 'f9': 'F9', 'f10': 'F10', 'f11': 'F11', 'f12': 'F12',
            'mouse_middle': 'Mouse-Middle', 'mouse_x1': 'Mouse-X1', 'mouse_x2': 'Mouse-X2',
        }
        parts = []
        for key_str in sorted(hotkey_list_strings):
            if key_str in display_map:
                parts.append(display_map[key_str])
            elif key_str.startswith('<vk-') and key_str.endswith('>'):
                parts.append(f"VK-{key_str[4:-1]}")
            elif key_str.startswith('mouse_button_'):
                parts.append(f"Mouse-Btn-{key_str[13:]}")
            elif len(key_str) == 1 and key_str.isalpha():
                parts.append(key_str.upper())
            else:
                parts.append(key_str.replace('_', ' ').title())
        return "+".join(parts) if parts else "Not Assigned"

def get_pynput_key_string(key):
    if isinstance(key, keyboard.Key):
        return str(key).split('.')[-1]
    elif isinstance(key, keyboard.KeyCode):
        if key.char:
            char_map = {
                '\x00': 'space',
                '\r': 'enter',
                '\t': 'tab',
                '\x1b': 'esc',
                '\x08': 'backspace',
                '\x7f': 'delete',
            }
            return char_map.get(key.char, key.char)
        else:
            return f"<vk-{key.vk}>"
    elif isinstance(key, mouse.Button):
        return {
            mouse.Button.left: "mouse_left",
            mouse.Button.right: "mouse_right",
            mouse.Button.middle: "mouse_middle",
            mouse.Button.x1: "mouse_x1",
            mouse.Button.x2: "mouse_x2",
        }.get(key, f"<mouse-{key.value}>")
    return str(key)

# --- ToolTip Class ---
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.id = None
        self.x = 0
        self.y = 0
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)

    def enter(self, event=None):
        self.schedule()

    def unschedule(self):
        id_ = self.id
        self.id = None
        if id_:
            self.widget.after_cancel(id_)

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(500, self.show)

    def show(self):
        if self.tip_window or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 20
        y += self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.tip_window = tk.Toplevel(self.widget)
        self.tip_window.wm_overrideredirect(True)
        self.tip_window.wm_geometry("+%d+%d" % (x, y))

        label = tk.Label(self.tip_window, text=self.text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def leave(self, event=None):
        self.unschedule()
        self.hide()

    def hide(self):
        if self.tip_window:
            self.tip_window.destroy()
        self.tip_window = None

# --- HotkeyRecorder Class ---
class HotkeyRecorder(Toplevel):
    def __init__(self, parent, target_name, target_hotkey_var, on_hotkey_recorded_callback):
        super().__init__(parent)
        self.parent = parent
        self.target_name = target_name
        self.target_hotkey_var = target_hotkey_var
        self.on_hotkey_recorded_callback = on_hotkey_recorded_callback
        
        self.transient(parent)
        self.title("Assign Hotkey")
        self.geometry("300x100")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

        self.hotkey_combination_strings = set()
        self.hotkey_lock = Lock()

        self.label = ttk.Label(self, text=f"Press keys for '{target_name}'...", font=("Helvetica", 10))
        self.label.pack(pady=10)

        self.display_label = ttk.Label(self, textvariable=self.target_hotkey_var, font=("Helvetica", 12, "bold"), bootstyle="info")
        self.display_label.pack(pady=5)

        self.keyboard_listener = None
        self.mouse_listener = None
        self.timeout_id = None

        self._start_listening()
        self._reset_timeout()

    def _get_pynput_key_string(self, key_or_button):
        if isinstance(key_or_button, keyboard.Key):
            if key_or_button in [keyboard.Key.ctrl_l, keyboard.Key.ctrl_r]:
                return 'ctrl'
            elif key_or_button in [keyboard.Key.alt_l, keyboard.Key.alt_r]:
                return 'alt'
            elif key_or_button in [keyboard.Key.shift_l, keyboard.Key.shift_r]:
                return 'shift'
            elif key_or_button in [keyboard.Key.cmd_l, keyboard.Key.cmd_r]:
                return 'cmd'
            key_name = str(key_or_button).split('.')[-1]
            try:
                KeyCode(key_name)
                return key_name
            except ValueError:
                return f"<vk-{key_or_button.value}>" if hasattr(key_or_button, 'value') else key_name
        elif isinstance(key_or_button, keyboard.KeyCode):
            if key_or_button.char:
                char = key_or_button.char.lower()
                try:
                    KeyCode(char)
                    return char
                except ValueError:
                    return char
            else:
                return f"<vk-{key_or_button.vk}>"
        elif isinstance(key_or_button, mouse.Button):
            button_map = {
                mouse.Button.middle: "mouse_middle",
                mouse.Button.x1: "mouse_x1",
                mouse.Button.x2: "mouse_x2",
            }
            if key_or_button in button_map:
                return button_map[key_or_button]
            else:
                return f"mouse_button_{id(key_or_button)}"
        return str(key_or_button)

    def _on_key_press(self, key):
        with self.hotkey_lock:
            if key == keyboard.Key.esc:
                self._on_cancel()
                return False
            key_str = self._get_pynput_key_string(key)
            if key_str and key_str not in self.hotkey_combination_strings:
                self.hotkey_combination_strings.add(key_str)
                self._update_display()
            self._reset_timeout()
        return True

    def _on_key_release(self, key):
        with self.hotkey_lock:
            self._reset_timeout()
        return True

    def _on_mouse_click(self, x, y, button, pressed):
        with self.hotkey_lock:
            if pressed:
                if button in [mouse.Button.left, mouse.Button.right]:
                    logging.info(f"HotkeyRecorder: Ignoring disallowed mouse button: {button}")
                    return True
                key_str = self._get_pynput_key_string(button)
                if key_str and key_str not in self.hotkey_combination_strings:
                    self.hotkey_combination_strings.add(key_str)
                    self._update_display()
                self._reset_timeout()
            return True

    def _update_display(self):
        display_text = get_hotkey_display_string(list(self.hotkey_combination_strings))
        self.target_hotkey_var.set(display_text)

    def _reset_timeout(self):
        if self.timeout_id:
            self.after_cancel(self.timeout_id)
        self.timeout_id = self.after(1500, self._finalize_hotkey)

    def _start_listening(self):
        self.keyboard_listener = keyboard.Listener(
            on_press=self._on_key_press,
            on_release=self._on_key_release
        )
        self.mouse_listener = mouse.Listener(
            on_click=self._on_mouse_click
        )
        self.keyboard_listener.start()
        self.mouse_listener.start()
        logging.info("HotkeyRecorder: Listeners started.")

    def _stop_listening(self):
        if self.keyboard_listener and self.keyboard_listener.is_alive():
            self.keyboard_listener.stop()
            self.keyboard_listener = None
        if self.mouse_listener and self.mouse_listener.is_alive():
            self.mouse_listener.stop()
            self.mouse_listener = None
        logging.info("HotkeyRecorder: Listeners stopped.")

    def _finalize_hotkey(self):
        self._stop_listening()
        if self.timeout_id:
            self.after_cancel(self.timeout_id)
            self.timeout_id = None
        
        non_modifiers = [k for k in self.hotkey_combination_strings if k not in ['ctrl', 'shift', 'alt', 'cmd']]
        if not non_modifiers:
            logging.warning("HotkeyRecorder: No non-modifier key captured, cancelling hotkey assignment.")
            self.parent.after(0, self.on_hotkey_recorded_callback, [])
            self.destroy()
            return
        
        self.parent.after(0, self.on_hotkey_recorded_callback, sorted(list(self.hotkey_combination_strings)))
        self.destroy()

    def _on_cancel(self):
        self.hotkey_combination_strings.clear()
        self._stop_listening()
        if self.timeout_id:
            self.after_cancel(self.timeout_id)
            self.timeout_id = None
        self.parent.after(0, self.on_hotkey_recorded_callback, [])
        self.destroy()
        logging.info("HotkeyRecorder: Assignment cancelled.")

# --- Check Dependencies ---
def check_dependencies():
    required_python_libs = ["pydub", "pynput", "sounddevice", "pygame", "resampy"]
    for lib in required_python_libs:
        try:
            importlib.metadata.version(lib)
            logging.info(f"{lib} library installed")
        except importlib.metadata.PackageNotFoundError:
            logging.error(f"{lib} library not installed")
            messagebox.showerror("Dependency Error", f"Please install the '{lib}' library: pip install {lib}")
            sys.exit(1)

    try:
        subprocess.run(['ffmpeg', '-version'], capture_output=True, check=True)
        logging.info("FFmpeg detected")
    except (subprocess.CalledProcessError, FileNotFoundError):
        logging.error("FFmpeg not installed or not in PATH")
        messagebox.showerror("Dependency Error", "FFmpeg is required. Install it and add to PATH.")
        sys.exit(1)

check_dependencies()

# --- Ensure Folder Structure ---
def ensure_folders():
    for folder in [SOUNDS_DIR, CONFIG_DIR]:
        try:
            os.makedirs(folder, exist_ok=True)
            logging.info(f"Ensured folder exists: {folder}")
        except Exception as e:
            logging.error(f"Failed to create folder {folder}: {e}")
            messagebox.showerror("Setup Error", f"Failed to create folder {folder}: {e}")

# --- MixingBuffer Class ---
class MixingBuffer:
    def __init__(self):
        self.sounds = deque()
        self.lock = Lock()
        self.single_sound_mode = False

    def set_single_sound_mode(self, enabled):
        with self.lock:
            self.single_sound_mode = enabled
            logging.info(f"Single sound mode {'enabled' if enabled else 'disabled'}")

    def add_sound(self, data, volume, loop, sound_id, sound_name):
        with self.lock:
            if self.single_sound_mode:
                self.clear_sounds()
                logging.info(f"Single sound mode active: Cleared mixer before adding {sound_name}")
            self.sounds.append({"id": sound_id, "data": data, "volume": volume, "loop": loop, "index": 0, "name": sound_name})
            logging.debug(f"Added sound to mixer: {sound_name}")

    def mix_audio(self, frames):
        with self.lock:
            mixed = np.zeros((frames, CHANNELS), dtype=np.float32)
            sounds_to_remove = []
            currently_playing_names = []
            
            for sound in list(self.sounds):
                start = sound["index"]
                end = start + frames
                
                if end > len(sound["data"]):
                    if sound["loop"]:
                        sound["index"] = 0
                        start = 0
                        end = frames
                        logging.debug(f"Looping sound: {sound['name']}")
                    else:
                        sounds_to_remove.append(sound)
                        continue

                chunk = sound["data"][start:end] * sound["volume"]
                if len(chunk) < frames:
                    chunk = np.pad(chunk, ((0, frames - len(chunk)), (0, 0)), mode='constant')
                
                mixed += chunk
                sound["index"] = end
                currently_playing_names.append(sound["name"])

            for sound in sounds_to_remove:
                try:
                    self.sounds.remove(sound)
                    logging.debug(f"Removed finished sound from mixer: {sound['name']}")
                except ValueError:
                    pass
            
            np.clip(mixed, -1.0, 1.0, out=mixed)
            
            return mixed, currently_playing_names

    def clear_sounds(self):
        with self.lock:
            self.sounds.clear()
            logging.info("Mixer sounds cleared")

    def remove_sound_by_id(self, sound_id):
        with self.lock:
            original_len = len(self.sounds)
            self.sounds = deque([s for s in self.sounds if s["id"] != sound_id])
            if len(self.sounds) < original_len:
                logging.info(f"Removed sound with ID {sound_id} from mixer queue.")
                return True
            logging.warning(f"Sound with ID {sound_id} not found in mixer queue to remove.")
            return False

# --- SoundManager Class ---
class SoundManager:
    def __init__(self):
        self.sounds = []
        self.global_hotkeys = {
            "stop_all": [],
            "toggle_mic_to_mixer": [],
        }
        ensure_folders()
        self.load_config()

    def _get_next_id(self):
        return str(uuid.uuid4())

    def add_sound(self, file_path, custom_name=None):
        try:
            audio = pydub.AudioSegment.from_file(file_path)
            
            if audio.frame_rate != SAMPLE_RATE:
                logging.info(f"Resampling audio from {audio.frame_rate}Hz to {SAMPLE_RATE}Hz for {os.path.basename(file_path)}")
                audio = audio.set_frame_rate(SAMPLE_RATE)
            if audio.channels != CHANNELS:
                logging.info(f"Converting audio from {audio.channels} channels to {CHANNELS} channels for {os.path.basename(file_path)}")
                audio = audio.set_channels(CHANNELS)
            
            sound_name = custom_name or os.path.splitext(os.path.basename(file_path))[0]
            sound_name = re.sub(INVALID_FILENAME_CHARS, '_', sound_name)

            output_path = os.path.join(SOUNDS_DIR, f"{sound_name}.wav")
            base_name, ext = os.path.splitext(output_path)
            counter = 1
            while os.path.exists(output_path):
                output_path = f"{base_name}_{counter}{ext}"
                counter += 1
            
            audio.export(output_path, format="wav")

            info = sf.info(output_path)
            new_id = self._get_next_id()
            
            new_sound = {
                "id": new_id,
                "name": sound_name,
                "path": output_path,
                "volume": 1.0,
                "hotkeys": [],
                "loop": False,
                "duration": info.duration,
            }

            self.sounds.append(new_sound)
            self.save_config()
            logging.info(f"Added sound: {sound_name} at {output_path}")
            return new_sound
        except Exception as e:
            logging.error(f"Failed to add sound {file_path}: {e}")
            messagebox.showerror("Add Sound Error", f"Failed to add sound: {e}")
            raise

    def rename_sound(self, sound_id, new_name):
        try:
            new_name = re.sub(INVALID_FILENAME_CHARS, '_', new_name.strip())
            if not new_name:
                raise ValueError("Sound name cannot be empty")
            
            sound = next((s for s in self.sounds if s["id"] == sound_id), None)
            if not sound:
                raise ValueError("Sound not found")
            
            if any(s["name"] == new_name for s in self.sounds if s["id"] != sound_id):
                raise ValueError("Sound name already exists")
            
            old_path = sound["path"]
            base, ext = os.path.splitext(os.path.basename(old_path))
            new_path = os.path.join(SOUNDS_DIR, f"{new_name}{ext}")
            
            counter = 1
            while os.path.exists(new_path) and new_path != old_path:
                new_path = f"{new_name}_{counter}{ext}"
                counter += 1
            
            if old_path != new_path:
                os.rename(old_path, new_path)
            
            sound["name"] = new_name
            sound["path"] = new_path
            self.save_config()
            logging.info(f"Renamed sound {sound_id} to {new_name}")
            return sound
        except Exception as e:
            logging.error(f"Failed to rename sound {sound_id}: {e}")
            messagebox.showerror("Rename Error", f"Failed to rename sound: {e}")
            raise

    def remove_sounds(self, sound_ids):
        try:
            sounds_to_keep = []
            for sound in self.sounds:
                if sound["id"] in sound_ids:
                    if os.path.exists(sound["path"]):
                        try:
                            os.remove(sound["path"])
                            logging.info(f"Deleted sound file: {sound['path']}")
                        except Exception as e:
                            logging.error(f"Failed to delete sound file {sound['path']}: {e}")
                else:
                    sounds_to_keep.append(sound)
            self.sounds = sounds_to_keep
            self.save_config()
            logging.info(f"Removed sounds with IDs: {sound_ids}")
        except Exception as e:
            logging.error(f"Failed to remove sounds: {e}")
            messagebox.showerror("Remove Error", f"Failed to remove sounds: {e}")

    def update_sound_hotkeys(self, sound_id, hotkey_list_strings):
        sound = next((s for s in self.sounds if s["id"] == sound_id), None)
        if sound:
            sound["hotkeys"] = hotkey_list_strings
            self.save_config()
            logging.info(f"Updated hotkeys for sound {sound_id}: {hotkey_list_strings}")
        else:
            logging.warning(f"Attempted to update hotkeys for non-existent sound: {sound_id}")

    def set_global_hotkey(self, action, hotkey_list_strings):
        if action not in self.global_hotkeys:
            logging.warning(f"Attempted to set hotkey for unknown action: {action}")
            return
        self.global_hotkeys[action] = hotkey_list_strings
        self.save_config()
        logging.info(f"Set global hotkey for {action}: {hotkey_list_strings}")

    def set_sound_loop(self, sound_id, loop_status):
        sound = next((s for s in self.sounds if s["id"] == sound_id), None)
        if sound:
            sound["loop"] = loop_status
            self.save_config()
            logging.info(f"Updated loop status for sound {sound['id']} to {loop_status}")
        else:
            logging.warning(f"Attempted to update loop status for non-existent sound: {sound_id}")

    def get_all_assigned_hotkeys(self):
        all_hotkeys = set()
        for sound in self.sounds:
            if sound["hotkeys"]:
                all_hotkeys.add(tuple(sorted(sound["hotkeys"])))
        
        for action, hotkeys in self.global_hotkeys.items():
            if hotkeys:
                all_hotkeys.add(tuple(sorted(hotkeys)))
        return all_hotkeys

    def save_config(self):
        try:
            serializable_sounds = []
            for sound in self.sounds:
                s_copy = sound.copy()
                s_copy.pop("_volume_var_ui", None)
                s_copy.pop("_loop_var_ui", None)
                s_copy.pop("_hotkey_label_var_ui", None)
                serializable_sounds.append(s_copy)

            config_data = {
                "sounds": serializable_sounds,
                "global_hotkeys": self.global_hotkeys
            }
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config_data, f, indent=4)
            logging.info("Soundboard config saved successfully")
        except Exception as e:
            logging.error(f"Error saving soundboard config: {e}")
            messagebox.showerror("Config Error", f"Failed to save soundboard config: {e}")

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    data = json.load(f)
                self.sounds = []
                for s in data.get("sounds", []):
                    if os.path.exists(s["path"]):
                        try:
                            info = sf.info(s["path"])
                            sound = {
                                "id": s.get("id", str(uuid.uuid4())),
                                "name": s.get("name", os.path.splitext(os.path.basename(s["path"]))[0]),
                                "path": s["path"],
                                "volume": s.get("volume", 1.0),
                                "hotkeys": s.get("hotkeys", []),
                                "loop": s.get("loop", False),
                                "duration": info.duration,
                            }
                            self.sounds.append(sound)
                        except Exception as inner_e:
                            logging.warning(f"Skipping corrupted sound entry {s.get('name', s.get('path'))}: {inner_e}")
                    else:
                        logging.warning(f"Skipping sound entry with missing file: {s.get('path', 'Unknown Path')}")

                for key in ["stop_all", "toggle_mic_to_mixer"]:
                    self.global_hotkeys[key] = data.get("global_hotkeys", {}).get(key, [])

                logging.info("Soundboard config loaded successfully")
            except Exception as e:
                logging.error(f"Error loading soundboard config: {e}")
                messagebox.showwarning("Config Error", f"Failed to load soundboard config: {e}. Starting with empty config.")
                self.sounds = []
                self.global_hotkeys = {"stop_all": [], "toggle_mic_to_mixer": []}
        else:
            logging.info("No soundboard config file found, starting with empty config")

# --- AudioOutputManager Class ---
class AudioOutputManager:
    def __init__(self, app_instance):
        self.app_instance = app_instance
        self.p = None
        self.output_devices = []
        self.input_devices = []
        self.virtual_mic_device_id = None
        self.current_output_device_id_pyaudio = None
        self.current_input_device_id_pyaudio = None
        self.current_soundboard_monitor_device_id_pyaudio = None
        self.current_mic_monitor_device_id_pyaudio = None

        self.mixer = MixingBuffer()
        self.output_stream_pyaudio = None
        self.mic_input_stream_pyaudio = None
        self.soundboard_monitor_stream_pyaudio = None
        self.mic_monitor_stream_pyaudio = None

        self._mic_input_thread = None
        self._mic_input_thread_stop_event = Event()
        self._mic_buffer = deque(maxlen=int(SAMPLE_RATE * CHANNELS * 1 / FRAME_SIZE))
        self._mic_buffer_lock = Lock()

        self._soundboard_monitor_buffer = deque(maxlen=int(SAMPLE_RATE * CHANNELS * 1 / FRAME_SIZE))
        self._soundboard_monitor_buffer_lock = Lock()

        self.include_mic_in_mix = False
        self.soundboard_monitor_enabled = False
        self.mic_monitor_enabled = False

        self.master_volume = 1.0
        self.soundboard_monitor_volume = 1.0
        self.mic_monitor_volume = 1.0

        self._initialize_pyaudio()
        self._enumerate_devices()
        self._initialize_pygame_mixer()

    def _initialize_pyaudio(self):
        if self.p is None:
            try:
                self.p = pyaudio.PyAudio()
                logging.info("PyAudio initialized successfully.")
            except Exception as e:
                logging.critical(f"Failed to initialize PyAudio: {e}", exc_info=True)
                messagebox.showerror("Audio Error", f"Failed to initialize PyAudio. Audio features will be disabled: {e}")
                self.p = None

    def _enumerate_devices(self):
        if self.p is None:
            logging.warning("PyAudio not initialized, skipping device enumeration.")
            return

        try:
            self.output_devices = []
            self.input_devices = []
            
            for i in range(self.p.get_device_count()):
                device = self.p.get_device_info_by_index(i)
                if device.get('maxOutputChannels', 0) >= CHANNELS and abs(device.get('defaultSampleRate', 0) - SAMPLE_RATE) < 1.0:
                    self.output_devices.append({"name": device['name'], "index": i, "api_name": self.p.get_host_api_info_by_index(device['hostApi'])['name']})
                if device.get('maxInputChannels', 0) >= CHANNELS and abs(device.get('defaultSampleRate', 0) - SAMPLE_RATE) < 1.0 and VIRTUAL_MIC_NAME_PARTIAL.lower() not in device['name'].lower():
                    self.input_devices.append({"name": device['name'], "index": i, "api_name": self.p.get_host_api_info_by_index(device['hostApi'])['name']})

            vb_cable_devices = [dev for dev in self.output_devices if VIRTUAL_MIC_NAME_PARTIAL.lower() in dev['name'].lower()]
            if vb_cable_devices:
                wasapi_vb_cable = next((dev for dev in vb_cable_devices if 'wasapi' in dev['api_name'].lower()), None)
                if wasapi_vb_cable:
                    self.virtual_mic_device_id = wasapi_vb_cable['index']
                else:
                    self.virtual_mic_device_id = vb_cable_devices[0]['index']
                logging.info(f"Virtual Mic (VB-CABLE) detected at index: {self.virtual_mic_device_id}")
            else:
                logging.warning("Virtual Mic (VB-CABLE) not detected. Please ensure it is installed and enabled.")

            self.current_output_device_id_pyaudio = self.virtual_mic_device_id
            if self.current_output_device_id_pyaudio is None and self.output_devices:
                self.current_output_device_id_pyaudio = self.output_devices[0]['index']
                logging.warning(f"No VB-CABLE detected, defaulting output to: {self.output_devices[0]['name']}")
            elif self.current_output_device_id_pyaudio is None:
                logging.error("No output devices found at all. Audio playback will not work.")

            self.current_soundboard_monitor_device_id_pyaudio = None
            self.current_mic_monitor_device_id_pyaudio = None
            if self.output_devices:
                self.current_soundboard_monitor_device_id_pyaudio = self.output_devices[0]['index']
                self.current_mic_monitor_device_id_pyaudio = self.output_devices[0]['index']
            else:
                logging.error("No output devices found for monitoring.")

            self.current_input_device_id_pyaudio = None
            try:
                default_input_info = self.p.get_default_input_device_info()
                input_device_candidates = [dev for dev in self.input_devices if dev['name'] == default_input_info['name']]
                wasapi_input = next((dev for dev in input_device_candidates if 'WASAPI' in dev['api_name'].lower()), None)
                if wasapi_input:
                    self.current_input_device_id_pyaudio = wasapi_input['index']
                elif input_device_candidates:
                    self.current_input_device_id_pyaudio = input_device_candidates[0]['index']
                else:
                    self.current_input_device_id_pyaudio = default_input_info['index']
            except IOError:
                logging.warning("No default input device found, attempting to use first available.")
                if self.input_devices:
                    self.current_input_device_id_pyaudio = self.input_devices[0]['index']
                else:
                    logging.error("No input devices found at all.")
                    self.current_input_device_id_pyaudio = None
            
            logging.info(f"PyAudio devices: Output={self.get_pyaudio_device_name_by_index(self.current_output_device_id_pyaudio)}, "
                         f"Input={self.get_pyaudio_device_name_by_index(self.current_input_device_id_pyaudio)}, "
                         f"Virtual Mic={self.virtual_mic_device_id}, "
                         f"Soundboard Monitor={self.get_pyaudio_device_name_by_index(self.current_soundboard_monitor_device_id_pyaudio)}, "
                         f"Mic Monitor={self.get_pyaudio_device_name_by_index(self.current_mic_monitor_device_id_pyaudio)}")

        except Exception as e:
            logging.error(f"Error enumerating devices: {e}", exc_info=True)
            traceback.print_exc()
            messagebox.showerror("Audio Device Error", f"Failed to enumerate audio devices: {e}\nTry restarting the app.")

    def _initialize_pygame_mixer(self):
        if 'pygame' in sys.modules and not pygame.mixer.get_init():
            try:
                pygame.mixer.init(frequency=SAMPLE_RATE, channels=CHANNELS)
                logging.info(f"Pygame mixer re-initialized at {SAMPLE_RATE}Hz, {CHANNELS} channels.")
            except pygame.error as e:
                logging.error(f"Pygame mixer re-initialization error: {e}")

    def get_pyaudio_device_name_by_index(self, index):
        if index is None or self.p is None:
            return "None"
        try:
            return self.p.get_device_info_by_index(index)['name']
        except Exception:
            return f"Unknown Device ({index})"

    def start_mic_input_stream(self):
        if self.mic_input_stream_pyaudio and self.mic_input_stream_pyaudio.is_active():
            logging.info("Mic input stream already active.")
            return

        if self.p is None or self.current_input_device_id_pyaudio is None:
            logging.warning("Cannot start mic input stream: PyAudio not initialized or no input device selected.")
            return
        
        if self._mic_input_thread and self._mic_input_thread.is_alive():
            logging.info("Mic input reader thread already running, not starting again.")
            return

        try:
            self.mic_input_stream_pyaudio = self.p.open(
                format=pyaudio.paFloat32,
                channels=CHANNELS,
                rate=SAMPLE_RATE,
                input=True,
                frames_per_buffer=FRAME_SIZE,
                input_device_index=self.current_input_device_id_pyaudio,
            )
            logging.info(f"PyAudio mic input stream opened from Mic ({self.current_input_device_id_pyaudio}).")

            self._mic_input_thread_stop_event.clear()
            self._mic_input_thread = threading.Thread(target=self._mic_input_reader_thread, daemon=True)
            self._mic_input_thread.start()
            logging.info("Mic input reader thread started.")

        except Exception as e:
            logging.error(f"Failed to start mic input stream: {e}", exc_info=True)
            traceback.print_exc()
            messagebox.showerror("Audio Error", f"Failed to start mic input stream: {e}")

    def stop_mic_input_stream(self):
        if self._mic_input_thread and self._mic_input_thread.is_alive():
            self._mic_input_thread_stop_event.set()
            self._mic_input_thread.join(timeout=1)
            if self._mic_input_thread.is_alive():
                logging.warning("Mic input reader thread did not stop gracefully.")
            self._mic_input_thread = None
        
        if self.mic_input_stream_pyaudio:
            try:
                if self.mic_input_stream_pyaudio.is_active():
                    self.mic_input_stream_pyaudio.stop_stream()
                self.mic_input_stream_pyaudio.close()
            except OSError as e:
                logging.warning(f"Error stopping/closing mic input stream: {e}")
                traceback.print_exc()
            finally:
                self.mic_input_stream_pyaudio = None
            logging.info("PyAudio mic input stream stopped and closed.")
        
        with self._mic_buffer_lock:
            self._mic_buffer.clear()
            logging.debug("Mic buffer cleared.")

    def _mic_input_reader_thread(self):
        logging.info("Mic input reader thread started.")
        while not self._mic_input_thread_stop_event.is_set():
            try:
                if self.mic_input_stream_pyaudio and self.mic_input_stream_pyaudio.is_active():
                    raw_mic_data = self.mic_input_stream_pyaudio.read(FRAME_SIZE, exception_on_overflow=False)
                    mic_data_np = np.frombuffer(raw_mic_data, dtype=np.float32).reshape(-1, CHANNELS)
                    
                    with self._mic_buffer_lock:
                        self._mic_buffer.append(mic_data_np)
                else:
                    time.sleep(0.01)
            except IOError as e:
                logging.warning(f"Mic input stream read error in reader thread: {e}")
                traceback.print_exc()
                self._mic_input_thread_stop_event.set()
            except Exception as e:
                logging.error(f"Unexpected error in mic input reader thread: {e}", exc_info=True)
                traceback.print_exc()
                self._mic_input_thread_stop_event.set()
        logging.info("Mic input reader thread stopped.")

    def _get_mic_data_from_buffer(self, frame_count):
        mic_data = np.zeros((frame_count, CHANNELS), dtype=np.float32)
        with self._mic_buffer_lock:
            if not self._mic_buffer:
                return mic_data

            current_frames = 0
            while current_frames < frame_count and self._mic_buffer:
                chunk = self._mic_buffer.popleft()
                frames_to_copy = min(frame_count - current_frames, len(chunk))
                mic_data[current_frames : current_frames + frames_to_copy] = chunk[:frames_to_copy]
                current_frames += frames_to_copy
                
                if frames_to_copy < len(chunk):
                    self._mic_buffer.appendleft(chunk[frames_to_copy:])
        return mic_data

    def start_main_audio_output_stream(self):
        if self.output_stream_pyaudio and self.output_stream_pyaudio.is_active():
            logging.info("Main audio output stream already active.")
            return

        if self.p is None or self.current_output_device_id_pyaudio is None:
            logging.warning("Cannot start main audio output stream: PyAudio not initialized or no output device selected.")
            return

        try:
            self.output_stream_pyaudio = self.p.open(
                format=pyaudio.paFloat32,
                channels=CHANNELS,
                rate=SAMPLE_RATE,
                output=True,
                frames_per_buffer=FRAME_SIZE,
                output_device_index=self.current_output_device_id_pyaudio,
                stream_callback=self._main_output_callback
            )
            logging.info(f"PyAudio main output stream opened to Virtual Mic ({self.current_output_device_id_pyaudio}).")
        except Exception as e:
            logging.error(f"Failed to start main audio output stream: {e}", exc_info=True)
            traceback.print_exc()
            messagebox.showerror("Audio Error", f"Failed to start main audio output stream: {e}")

    def _main_output_callback(self, in_data, frame_count, time_info, status):
        mixed_audio, playing_names = self.mixer.mix_audio(frame_count)
        
        with self._soundboard_monitor_buffer_lock:
            self._soundboard_monitor_buffer.append(mixed_audio.copy())
        
        with current_playing_sound_details_lock:
            current_playing_sound_details["names"] = playing_names
            current_playing_sound_details["active"] = len(playing_names) > 0 or self.include_mic_in_mix
        
        if self.include_mic_in_mix:
            mic_data = self._get_mic_data_from_buffer(frame_count)
            mixed_audio += mic_data
            np.clip(mixed_audio, -1.0, 1.0, out=mixed_audio)
        
        mixed_audio *= self.master_volume
        return (mixed_audio.tobytes(), pyaudio.paContinue)

    def start_soundboard_monitor_stream(self):
        if self.soundboard_monitor_stream_pyaudio and self.soundboard_monitor_stream_pyaudio.is_active():
            logging.info("Soundboard monitor stream already active.")
            return

        if self.p is None or self.current_soundboard_monitor_device_id_pyaudio is None:
            logging.warning("Cannot start soundboard monitor stream: PyAudio not initialized or no monitor device selected.")
            return

        try:
            self.soundboard_monitor_stream_pyaudio = self.p.open(
                format=pyaudio.paFloat32,
                channels=CHANNELS,
                rate=SAMPLE_RATE,
                output=True,
                frames_per_buffer=FRAME_SIZE,
                output_device_index=self.current_soundboard_monitor_device_id_pyaudio,
                stream_callback=self._soundboard_monitor_callback
            )
            logging.info(f"PyAudio soundboard monitor stream opened to device ({self.current_soundboard_monitor_device_id_pyaudio}).")
        except Exception as e:
            logging.error(f"Failed to start soundboard monitor stream: {e}", exc_info=True)
            traceback.print_exc()
            messagebox.showerror("Audio Error", f"Failed to start soundboard monitor stream: {e}")

    def _soundboard_monitor_callback(self, in_data, frame_count, time_info, status):
        with self._soundboard_monitor_buffer_lock:
            if self._soundboard_monitor_buffer:
                data = self._soundboard_monitor_buffer.popleft()
                if len(data) < frame_count:
                    data = np.pad(data, ((0, frame_count - len(data)), (0, 0)), mode='constant')
                elif len(data) > frame_count:
                    self._soundboard_monitor_buffer.appendleft(data[frame_count:])
                    data = data[:frame_count]
                data *= self.soundboard_monitor_volume
                return (data.tobytes(), pyaudio.paContinue)
            return (np.zeros((frame_count, CHANNELS), dtype=np.float32).tobytes(), pyaudio.paContinue)

    def start_mic_monitor_stream(self):
        if self.mic_monitor_stream_pyaudio and self.mic_monitor_stream_pyaudio.is_active():
            logging.info("Mic monitor stream already active.")
            return

        if self.p is None or self.current_mic_monitor_device_id_pyaudio is None:
            logging.warning("Cannot start mic monitor stream: PyAudio not initialized or no mic monitor device selected.")
            return

        try:
            self.mic_monitor_stream_pyaudio = self.p.open(
                format=pyaudio.paFloat32,
                channels=CHANNELS,
                rate=SAMPLE_RATE,
                output=True,
                frames_per_buffer=FRAME_SIZE,
                output_device_index=self.current_mic_monitor_device_id_pyaudio,
                stream_callback=self._mic_monitor_callback
            )
            logging.info(f"PyAudio mic monitor stream opened to device ({self.current_mic_monitor_device_id_pyaudio}).")
        except Exception as e:
            logging.error(f"Failed to start mic monitor stream: {e}", exc_info=True)
            traceback.print_exc()
            messagebox.showerror("Audio Error", f"Failed to start mic monitor stream: {e}")

    def _mic_monitor_callback(self, in_data, frame_count, time_info, status):
        mic_data = self._get_mic_data_from_buffer(frame_count)
        mic_data *= self.mic_monitor_volume
        return (mic_data.tobytes(), pyaudio.paContinue)

    def stop_main_audio_output_stream(self):
        if self.output_stream_pyaudio:
            try:
                if self.output_stream_pyaudio.is_active():
                    self.output_stream_pyaudio.stop_stream()
                self.output_stream_pyaudio.close()
            except OSError as e:
                logging.warning(f"Error stopping/closing main output stream: {e}")
                traceback.print_exc()
            finally:
                self.output_stream_pyaudio = None
            logging.info("PyAudio main output stream stopped and closed.")

    def stop_soundboard_monitor_stream(self):
        if self.soundboard_monitor_stream_pyaudio:
            try:
                if self.soundboard_monitor_stream_pyaudio.is_active():
                    self.soundboard_monitor_stream_pyaudio.stop_stream()
                self.soundboard_monitor_stream_pyaudio.close()
            except OSError as e:
                logging.warning(f"Error stopping/closing soundboard monitor stream: {e}")
                traceback.print_exc()
            finally:
                self.soundboard_monitor_stream_pyaudio = None
            logging.info("PyAudio soundboard monitor stream stopped and closed.")

    def stop_mic_monitor_stream(self):
        if self.mic_monitor_stream_pyaudio:
            try:
                if self.mic_monitor_stream_pyaudio.is_active():
                    self.mic_monitor_stream_pyaudio.stop_stream()
                self.mic_monitor_stream_pyaudio.close()
            except OSError as e:
                logging.warning(f"Error stopping/closing mic monitor stream: {e}")
                traceback.print_exc()
            finally:
                self.mic_monitor_stream_pyaudio = None
            logging.info("PyAudio mic monitor stream stopped and closed.")

    def restart_audio_streams(self):
        self.stop_mic_input_stream()
        self.stop_main_audio_output_stream()
        self.stop_soundboard_monitor_stream()
        self.stop_mic_monitor_stream()

        self.start_main_audio_output_stream()

        if self.include_mic_in_mix:
            self.start_mic_input_stream()

        if self.soundboard_monitor_enabled:
            self.start_soundboard_monitor_stream()

        if self.mic_monitor_enabled:
            self.start_mic_monitor_stream()

        logging.info("Audio streams restarted with current settings.")

    def close(self):
        self.stop_mic_input_stream()
        self.stop_main_audio_output_stream()
        self.stop_soundboard_monitor_stream()
        self.stop_mic_monitor_stream()

        with self._soundboard_monitor_buffer_lock:
            self._soundboard_monitor_buffer.clear()
            logging.debug("Soundboard monitor buffer cleared.")

        if self.p is not None:
            try:
                self.p.terminate()
                logging.info("PyAudio terminated.")
            except Exception as e:
                logging.error(f"Error terminating PyAudio: {e}", exc_info=True)
            self.p = None

        if 'pygame' in sys.modules and pygame.mixer.get_init():
            pygame.mixer.quit()
            logging.info("Pygame mixer quit on close.")

# --- KeybindManager Class ---
class KeybindManager:
    def __init__(self, app):
        self.app = app
        self.hotkey_registry = {}
        self.keyboard_listener = None
        self.mouse_listener = None
        self.current_hotkey_assignment_target_id = None
        self.active_keys = set()

    def update_hotkeys(self):
        self.stop()
        self.hotkey_registry = {}
        logging.debug("Updating hotkeys...")

        for sound in self.app.sound_manager.sounds:
            if sound.get("hotkeys"):
                try:
                    hotkey_obj = Hotkey.from_json_serializable(sound["hotkeys"])
                    self.hotkey_registry[hotkey_obj] = lambda sid=sound["id"]: self.app.play_sound(sid)
                    logging.debug(f"Registered hotkey for sound {sound['id']}: {hotkey_obj}")
                except Exception as e:
                    logging.error(f"Failed to register hotkey for sound {sound['id']}: {e}", exc_info=True)

        for action, hotkeys in self.app.sound_manager.global_hotkeys.items():
            if hotkeys:
                try:
                    hotkey_obj = Hotkey.from_json_serializable(hotkeys)
                    if action == "stop_all":
                        self.hotkey_registry[hotkey_obj] = self.app.stop_all_sounds
                    elif action == "toggle_mic_to_mixer":
                        self.hotkey_registry[hotkey_obj] = self.app.toggle_mic_to_mixer_from_hotkey
                    logging.debug(f"Registered global hotkey for {action}: {hotkey_obj}")
                except Exception as e:
                    logging.error(f"Failed to register global hotkey for {action}: {e}", exc_info=True)

        if self.hotkey_registry:
            self.start()
            logging.info(f"Registered {len(self.hotkey_registry)} hotkeys.")
        else:
            logging.info("No active hotkeys to listen for. Not starting pynput listeners.")

    def start(self):
        self.stop()
        self.active_keys = set()

        def on_press(key):
            try:
                key_str = get_pynput_key_string(key)
                self.active_keys.add(key_str)
                logging.debug(f"Key pressed: {key_str}, Active keys: {self.active_keys}")
                self.check_hotkeys()
            except Exception as e:
                logging.error(f"Error processing key press: {e}", exc_info=True)

        def on_release(key):
            try:
                key_str = get_pynput_key_string(key)
                self.active_keys.discard(key_str)
                logging.debug(f"Key released: {key_str}, Active keys: {self.active_keys}")
            except Exception as e:
                logging.error(f"Error processing key release: {e}", exc_info=True)

        def on_click(x, y, button, pressed):
            try:
                if not pressed:
                    return
                if button in [mouse.Button.left, mouse.Button.right]:
                    return
                button_str = get_pynput_key_string(button)
                self.active_keys.add(button_str)
                logging.debug(f"Mouse button pressed: {button_str}, Active keys: {self.active_keys}")
                self.check_hotkeys()
                self.active_keys.discard(button_str)
            except Exception as e:
                logging.error(f"Error processing mouse click: {e}", exc_info=True)

        self.keyboard_listener = keyboard.Listener(on_press=on_press, on_release=on_release)
        self.mouse_listener = mouse.Listener(on_click=on_click)
        self.keyboard_listener.start()
        self.mouse_listener.start()
        logging.info("Started pynput listeners for keyboard and mouse.")

    def check_hotkeys(self):
        for hotkey_obj, callback in self.hotkey_registry.items():
            hotkey_strings = set(hotkey_obj.to_json_serializable())
            if hotkey_strings == self.active_keys:
                logging.debug(f"Hotkey triggered: {hotkey_obj}")
                callback()
                non_modifiers = [k for k in hotkey_strings if k not in ['ctrl', 'shift', 'alt', 'cmd']]
                for k in non_modifiers:
                    self.active_keys.discard(k)

    def stop(self):
        if self.keyboard_listener and self.keyboard_listener.is_alive():
            self.keyboard_listener.stop()
            self.keyboard_listener = None
        if self.mouse_listener and self.mouse_listener.is_alive():
            self.mouse_listener.stop()
            self.mouse_listener = None
        self.active_keys.clear()
        logging.info("Stopped pynput listeners.")

    def set_hotkey_assignment_target(self, target_id):
        self.current_hotkey_assignment_target_id = target_id
        logging.debug(f"Set hotkey assignment target to: {target_id}")

# --- AppSettingsManager Class ---
class AppSettingsManager:
    def __init__(self):
        self.settings = {}
        self.load_settings()

    def load_settings(self):
        if os.path.exists(APP_SETTINGS_FILE):
            try:
                with open(APP_SETTINGS_FILE, 'r') as f:
                    self.settings = json.load(f)
                logging.info("App settings loaded successfully")
            except Exception as e:
                logging.error(f"Error loading app settings: {e}")
                self.settings = {}
                messagebox.showwarning("Settings Error", f"Failed to load app settings: {e}. Using defaults.")
        else:
            logging.info("No app settings file found, using defaults")
            self.settings = {}

    def save_settings(self, settings):
        self.settings.update(settings)
        try:
            with open(APP_SETTINGS_FILE, 'w') as f:
                json.dump(self.settings, f, indent=4)
            logging.info("App settings saved successfully")
        except Exception as e:
            logging.error(f"Error saving app settings: {e}")
            messagebox.showerror("Settings Error", f"Failed to save app settings: {e}")

    def get_settings(self):
        return self.settings

# --- SoundboardApp Class ---
class SoundboardApp(ttk.Window):
    def __init__(self):
        super().__init__(title="WarpBoard Soundboard", themename="litera")
        self.geometry("800x600")
        ensure_folders()
        self.sound_manager = SoundManager()
        self.audio_manager = AudioOutputManager(self)
        self.app_settings_manager = AppSettingsManager()
        self.keybind_manager = KeybindManager(self)
        
        self.selected_sound_ids = set()
        self._interactive_widgets = []
        
        self.include_mic_in_mix_var = tk.BooleanVar(value=False)
        self.soundboard_monitor_enabled_var = tk.BooleanVar(value=True)
        self.mic_monitor_enabled_var = tk.BooleanVar(value=False)
        self.master_volume_var = tk.DoubleVar(value=100.0)
        self.soundboard_monitor_volume_var = tk.DoubleVar(value=100.0)
        self.mic_monitor_volume_var = tk.DoubleVar(value=100.0)
        self.auto_start_include_mic_in_mix_var = tk.BooleanVar(value=False)
        self.single_sound_mode_var = tk.BooleanVar(value=False)
        self.current_theme_var = tk.StringVar(value=self.style.theme.name)
        self.stop_all_hotkey_var = tk.StringVar(value="Not Assigned")
        self.toggle_mic_to_mixer_hotkey_var = tk.StringVar(value="Not Assigned")
        self.status_message_var = tk.StringVar(value="Ready.")
        self.currently_playing_info = tk.StringVar(value="Now Playing: None")
        self.mic_status_info = tk.StringVar(value="Mic: Off")
        self.soundboard_output_status_info = tk.StringVar(value="Soundboard Output: Off")
        self.assign_button_style_var = tk.StringVar(value="primary")
        self.key_assignment_in_progress_var = tk.BooleanVar(value=False)
        
        self.protocol("WM_DELETE_WINDOW", self._on_app_closure)
        self._create_ui()
        self._apply_settings_to_ui()
        
        self.now_playing_updater_thread = threading.Thread(target=self._update_now_playing_periodically, daemon=True)
        self.now_playing_updater_thread.start()
        
        self.populate_sound_grid()
        self.keybind_manager.update_hotkeys()

    def _create_ui(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=BOTH, expand=True, padx=10, pady=10)
        self._interactive_widgets.append(self.notebook)
        
        self.soundboard_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.soundboard_tab, text="Soundboard")

        self.settings_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_tab, text="Settings")

        self.settings_notebook = ttk.Notebook(self.settings_tab)
        self.settings_notebook.pack(fill=BOTH, expand=True, padx=5, pady=5)

        
        self.audio_settings_tab = ttk.Frame(self.settings_notebook)
        self.hotkeys_settings_tab = ttk.Frame(self.settings_notebook)
        self.general_settings_tab = ttk.Frame(self.settings_notebook)
        self.about_tab = ttk.Frame(self.settings_notebook)
        
        self.settings_notebook.add(self.audio_settings_tab, text="Audio Settings")
        self.settings_notebook.add(self.hotkeys_settings_tab, text="Global Hotkeys")
        self.settings_notebook.add(self.general_settings_tab, text="General")
        self.settings_notebook.add(self.about_tab, text="About")
        
        self._interactive_widgets.extend([self.audio_settings_tab, self.hotkeys_settings_tab, self.general_settings_tab, self.about_tab])
        
        self._create_soundboard_widgets(self.soundboard_tab)
        self._create_audio_settings_widgets(self.audio_settings_tab)
        self._create_hotkeys_settings_widgets(self.hotkeys_settings_tab)
        self._create_general_settings_widgets(self.general_settings_tab)
        self._create_about_widgets(self.about_tab)
        
        self.status_frame = ttk.Frame(self)
        self.status_frame.pack(fill=X, padx=5, pady=5)
        
        self.status_label = ttk.Label(self.status_frame, textvariable=self.status_message_var, bootstyle="info")
        self.status_label.pack(side=LEFT, padx=5)
        
        self.currently_playing_label = ttk.Label(self.status_frame, textvariable=self.currently_playing_info, bootstyle="info")
        self.currently_playing_label.pack(side=LEFT, padx=20)
        
        self.mic_status_label = ttk.Label(self.status_frame, textvariable=self.mic_status_info, bootstyle="danger")
        self.mic_status_label.pack(side=LEFT, padx=20)
        
        self.soundboard_output_status_label = ttk.Label(self.status_frame, textvariable=self.soundboard_output_status_info, bootstyle="danger")
        self.soundboard_output_status_label.pack(side=LEFT, padx=20)

    def _create_soundboard_widgets(self, parent_frame):
        button_frame = ttk.Frame(parent_frame)
        button_frame.pack(fill=X, padx=5, pady=5)
        
        add_button = ttk.Button(button_frame, text="Add Sound", command=self.add_sound, bootstyle="primary")
        add_button.pack(side=LEFT, padx=5)
        self._interactive_widgets.append(add_button)
        
        remove_button = ttk.Button(button_frame, text="Remove Selected", command=self.remove_selected_sounds, bootstyle="danger")
        remove_button.pack(side=LEFT, padx=5)
        self._interactive_widgets.append(remove_button)
        
        play_selected_button = ttk.Button(button_frame, text="Play Selected", command=self.play_selected_sound, bootstyle="success")
        play_selected_button.pack(side=LEFT, padx=5)
        self._interactive_widgets.append(play_selected_button)
        
        stop_all_button = ttk.Button(button_frame, text="Stop All", command=self.stop_all_sounds, bootstyle="danger")
        stop_all_button.pack(side=LEFT, padx=5)
        self._interactive_widgets.append(stop_all_button)
        
        self.sound_grid_frame = ttk.Frame(parent_frame)
        self.sound_grid_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
        self._interactive_widgets.append(self.sound_grid_frame)
        
        self.sound_grid_canvas = tk.Canvas(self.sound_grid_frame)
        self.sound_grid_scrollbar = ttk.Scrollbar(self.sound_grid_frame, orient=VERTICAL, command=self.sound_grid_canvas.yview)
        self.sound_grid_canvas.configure(yscrollcommand=self.sound_grid_scrollbar.set)
        
        self.sound_grid_scrollbar.pack(side=RIGHT, fill=Y)
        self.sound_grid_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        self._interactive_widgets.append(self.sound_grid_scrollbar)
        
        self.inner_sound_grid_frame = ttk.Frame(self.sound_grid_canvas)
        self.sound_grid_canvas.create_window((0, 0), window=self.inner_sound_grid_frame, anchor="nw")
        
        self.inner_sound_grid_frame.bind("<Configure>", lambda e: self.sound_grid_canvas.configure(scrollregion=self.sound_grid_canvas.bbox("all")))
        
        self.sound_grid_canvas.bind_all("<MouseWheel>", self._on_mouse_wheel)
        self.sound_grid_canvas.bind_all("<Button-4>", self._on_mouse_wheel)
        self.sound_grid_canvas.bind_all("<Button-5>", self._on_mouse_wheel)

    def _on_mouse_wheel(self, event):
        if self.sound_grid_canvas.winfo_ismapped():
            if event.num == 4 or event.delta > 0:
                self.sound_grid_canvas.yview_scroll(-1, "units")
            elif event.num == 5 or event.delta < 0:
                self.sound_grid_canvas.yview_scroll(1, "units")
            return "break"

    def _create_audio_settings_widgets(self, parent_frame):
        output_frame = ttk.LabelFrame(parent_frame, text="Audio Output Settings", padding=10)
        output_frame.pack(fill=X, pady=10)
        self._interactive_widgets.append(output_frame)
        
        output_label = ttk.Label(output_frame, text="Main Output Device (Virtual Mic):")
        output_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.output_device_combobox = ttk.Combobox(output_frame, state="readonly")
        self.output_device_combobox.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.output_device_combobox.bind("<<ComboboxSelected>>", self._on_output_device_selected)
        ToolTip(self.output_device_combobox, "Select the output device for the mixed audio (soundboard + mic).")
        self._interactive_widgets.append(self.output_device_combobox)
        
        input_label = ttk.Label(output_frame, text="Physical Mic Input Device:")
        input_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.input_device_combobox = ttk.Combobox(output_frame, state="readonly")
        self.input_device_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.input_device_combobox.bind("<<ComboboxSelected>>", self._on_input_device_selected)
        ToolTip(self.input_device_combobox, "Select your physical microphone for input.")
        self._interactive_widgets.append(self.input_device_combobox)
        
        soundboard_monitor_label = ttk.Label(output_frame, text="Soundboard Monitor Device (Local):")
        soundboard_monitor_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.soundboard_monitor_device_combobox = ttk.Combobox(output_frame, state="readonly")
        self.soundboard_monitor_device_combobox.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.soundboard_monitor_device_combobox.bind("<<ComboboxSelected>>", self._on_soundboard_monitor_device_selected)
        ToolTip(self.soundboard_monitor_device_combobox, "Select the device to monitor only the soundboard audio locally (e.g., speakers).")
        self._interactive_widgets.append(self.soundboard_monitor_device_combobox)
        
        mic_monitor_label = ttk.Label(output_frame, text="Mic Playback Monitor Device (Local):")
        mic_monitor_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.mic_monitor_device_combobox = ttk.Combobox(output_frame, state="readonly")
        self.mic_monitor_device_combobox.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        self.mic_monitor_device_combobox.bind("<<ComboboxSelected>>", self._on_mic_monitor_device_selected)
        ToolTip(self.mic_monitor_device_combobox, "Select the device to monitor only your microphone audio locally (e.g., headphones).")
        self._interactive_widgets.append(self.mic_monitor_device_combobox)
        
        output_frame.columnconfigure(1, weight=1)
        
        controls_frame = ttk.LabelFrame(parent_frame, text="Audio Controls", padding=10)
        controls_frame.pack(fill=X, pady=10)
        self._interactive_widgets.append(controls_frame)
        
        include_mic_check = ttk.Checkbutton(controls_frame, text="Include Mic in Audio Mix", variable=self.include_mic_in_mix_var, command=self.toggle_include_mic_in_mix)
        include_mic_check.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        ToolTip(include_mic_check, "Enable to include your microphone audio in the virtual microphone output.")
        self._interactive_widgets.append(include_mic_check)
        
        soundboard_monitor_check = ttk.Checkbutton(controls_frame, text="Enable Local Soundboard Monitoring", variable=self.soundboard_monitor_enabled_var, command=self.toggle_soundboard_monitor)
        soundboard_monitor_check.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        ToolTip(soundboard_monitor_check, "Enable to hear the soundboard audio through your local speakers/headphones.")
        self._interactive_widgets.append(soundboard_monitor_check)
        
        mic_monitor_check = ttk.Checkbutton(controls_frame, text="Enable Local Mic Playback Monitoring", variable=self.mic_monitor_enabled_var, command=self.toggle_mic_monitor)
        mic_monitor_check.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        ToolTip(mic_monitor_check, "Enable to hear your microphone audio through your local speakers/headphones.")
        self._interactive_widgets.append(mic_monitor_check)
        
        controls_frame.columnconfigure(1, weight=1)
        
        volume_frame = ttk.LabelFrame(parent_frame, text="Volume Controls", padding=10)
        volume_frame.pack(fill=X, pady=10)
        self._interactive_widgets.append(volume_frame)
        
        master_volume_label = ttk.Label(volume_frame, text="Master Output Volume:")
        master_volume_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.master_volume_slider = ttk.Scale(volume_frame, from_=0, to=100, variable=self.master_volume_var, command=self._on_master_volume_changed)
        self.master_volume_slider.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self._interactive_widgets.append(self.master_volume_slider)
        self.master_volume_value_label = ttk.Label(volume_frame, text=f"{int(self.master_volume_var.get())}%")
        self.master_volume_var.trace_add("write", lambda name, index, mode: self.master_volume_value_label.config(text=f"{int(self.master_volume_var.get())}%"))
        self.master_volume_value_label.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.master_volume_slider.bind("<ButtonRelease-1>", self._on_master_volume_released)
        ToolTip(self.master_volume_slider, "Controls the overall volume of the mixed audio sent to the virtual microphone.")
        
        soundboard_monitor_volume_label = ttk.Label(volume_frame, text="Local Soundboard Monitor Volume:")
        soundboard_monitor_volume_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.soundboard_monitor_volume_slider = ttk.Scale(volume_frame, from_=0, to=100, variable=self.soundboard_monitor_volume_var, command=self._on_soundboard_monitor_volume_changed)
        self.soundboard_monitor_volume_slider.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self._interactive_widgets.append(self.soundboard_monitor_volume_slider)
        self.soundboard_monitor_volume_value_label = ttk.Label(volume_frame, text=f"{int(self.soundboard_monitor_volume_var.get())}%")
        self.soundboard_monitor_volume_var.trace_add("write", lambda name, index, mode: self.soundboard_monitor_volume_value_label.config(text=f"{int(self.soundboard_monitor_volume_var.get())}%"))
        self.soundboard_monitor_volume_value_label.grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.soundboard_monitor_volume_slider.bind("<ButtonRelease-1>", lambda e: self._on_soundboard_monitor_volume_released())
        ToolTip(self.soundboard_monitor_volume_slider, "Controls the volume of only the soundboard audio played back to your local speakers/headphones.")
        
        mic_monitor_volume_label = ttk.Label(volume_frame, text="Local Mic Playback Monitor Volume:")
        mic_monitor_volume_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.mic_monitor_volume_slider = ttk.Scale(volume_frame, from_=0, to=100, variable=self.mic_monitor_volume_var, command=self._on_mic_monitor_volume_changed)
        self.mic_monitor_volume_slider.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self._interactive_widgets.append(self.mic_monitor_volume_slider)
        self.mic_monitor_volume_value_label = ttk.Label(volume_frame, text=f"{int(self.mic_monitor_volume_var.get())}%")
        self.mic_monitor_volume_var.trace_add("write", lambda name, index, mode: self.mic_monitor_volume_value_label.config(text=f"{int(self.mic_monitor_volume_var.get())}%"))
        self.mic_monitor_volume_value_label.grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.mic_monitor_volume_slider.bind("<ButtonRelease-1>", self._on_mic_monitor_volume_released)
        ToolTip(self.mic_monitor_volume_slider, "Controls the volume of ONLY your microphone's audio played back to your local speakers/headphones.")
        
        volume_frame.columnconfigure(1, weight=1)
        
        self._populate_device_dropdowns()

    def _create_hotkeys_settings_widgets(self, parent_frame):
        hotkey_frame = ttk.LabelFrame(parent_frame, text="Global Hotkey Assignments", padding=10)
        hotkey_frame.pack(fill=X, pady=10)
        self._interactive_widgets.append(hotkey_frame)
        
        stop_all_label = ttk.Label(hotkey_frame, text="Stop All Sounds:")
        stop_all_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.stop_all_hotkey_display = ttk.Label(hotkey_frame, textvariable=self.stop_all_hotkey_var, bootstyle="info")
        self.stop_all_hotkey_display.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.assign_stop_all_button = ttk.Button(hotkey_frame, text="Assign", command=lambda: self._assign_global_hotkey("stop_all"), bootstyle=self.assign_button_style_var.get())
        self.assign_stop_all_button.grid(row=0, column=2, padx=5, pady=5)
        self._interactive_widgets.append(self.assign_stop_all_button)
        self.clear_stop_all_button = ttk.Button(hotkey_frame, text="Clear", command=lambda: self._clear_global_hotkey("stop_all"), bootstyle="secondary")
        self.clear_stop_all_button.grid(row=0, column=3, padx=5, pady=5)
        self._interactive_widgets.append(self.clear_stop_all_button)
        ToolTip(self.assign_stop_all_button, "Click to assign a global hotkey to stop all playing sounds.")
        ToolTip(self.clear_stop_all_button, "Clear the global hotkey for stopping all sounds.")
        
        toggle_mic_label = ttk.Label(hotkey_frame, text="Toggle Mic to Mixer:")
        toggle_mic_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.toggle_mic_hotkey_display = ttk.Label(hotkey_frame, textvariable=self.toggle_mic_to_mixer_hotkey_var, bootstyle="info")
        self.toggle_mic_hotkey_display.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.assign_toggle_mic_button = ttk.Button(hotkey_frame, text="Assign", command=lambda: self._assign_global_hotkey("toggle_mic_to_mixer"), bootstyle=self.assign_button_style_var.get())
        self.assign_toggle_mic_button.grid(row=1, column=2, padx=5, pady=5)
        self._interactive_widgets.append(self.assign_toggle_mic_button)
        self.clear_toggle_mic_button = ttk.Button(hotkey_frame, text="Clear", command=lambda: self._clear_global_hotkey("toggle_mic_to_mixer"), bootstyle="secondary")
        self.clear_toggle_mic_button.grid(row=1, column=3, padx=5, pady=5)
        self._interactive_widgets.append(self.clear_toggle_mic_button)
        ToolTip(self.assign_toggle_mic_button, "Click to assign a global hotkey to toggle including your microphone in the mix.")
        ToolTip(self.clear_toggle_mic_button, "Clear the global hotkey for toggling mic in mix.")
        
        hotkey_frame.columnconfigure(1, weight=1)

    def _create_general_settings_widgets(self, parent_frame):
        general_frame = ttk.LabelFrame(parent_frame, text="General Settings", padding=10)
        general_frame.pack(fill=X, pady=10)
        self._interactive_widgets.append(general_frame)
        
        theme_label = ttk.Label(general_frame, text="Theme:")
        theme_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.theme_combobox = ttk.Combobox(general_frame, textvariable=self.current_theme_var, values=self.style.theme_names(), state="readonly")
        self.theme_combobox.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.theme_combobox.bind("<<ComboboxSelected>>", self._on_theme_changed)
        ToolTip(self.theme_combobox, "Select the UI theme for the application.")
        self._interactive_widgets.append(self.theme_combobox)
        
        single_sound_check = ttk.Checkbutton(general_frame, text="Single Sound Mode", variable=self.single_sound_mode_var, command=self._on_single_sound_mode_changed)
        single_sound_check.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        ToolTip(single_sound_check, "When enabled, only one sound can play at a time, stopping any currently playing sounds.")
        self._interactive_widgets.append(single_sound_check)
        
        auto_start_mic_check = ttk.Checkbutton(general_frame, text="Auto-Start Mic in Mix", variable=self.auto_start_include_mic_in_mix_var, command=self._on_auto_start_mic_changed)
        auto_start_mic_check.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        ToolTip(auto_start_mic_check, "When enabled, the microphone will automatically be included in the mix on app startup.")
        self._interactive_widgets.append(auto_start_mic_check)
        
        general_frame.columnconfigure(1, weight=1)

    def _create_about_widgets(self, parent_frame):
        about_frame = ttk.LabelFrame(parent_frame, text="About WarpBoard", padding=10)
        about_frame.pack(fill=BOTH, expand=True, pady=10)
        self._interactive_widgets.append(about_frame)
        
        ttk.Label(about_frame, text="WarpBoard Soundboard", font=("Helvetica", 14, "bold")).pack(pady=5)
        ttk.Label(about_frame, text="Version: 1.0.0").pack(pady=5)
        ttk.Label(about_frame, text="A free and open-source soundboard application for Windows.").pack(pady=5)
        ttk.Label(about_frame, text="Developed by: Your Name").pack(pady=5)
        ttk.Label(about_frame, text="GitHub: github.com/yourusername/warpboard").pack(pady=5)
        
        vb_cable_button = ttk.Button(about_frame, text="Download VB-CABLE (Virtual Mic)", command=lambda: self._open_url(VB_CABLE_URL))
        vb_cable_button.pack(pady=10)
        ToolTip(vb_cable_button, "Download VB-CABLE, required for sending mixed audio to applications like Discord.")
        self._interactive_widgets.append(vb_cable_button)

    def _open_url(self, url):
        import webbrowser
        try:
            webbrowser.open(url)
            logging.info(f"Opened URL: {url}")
        except Exception as e:
            logging.error(f"Failed to open URL {url}: {e}")
            messagebox.showerror("Error", f"Failed to open URL: {e}")

    def _create_sound_card(self, sound):
        sound_id = sound["id"]
        sound_name = sound["name"]
        sound_duration = sound.get("duration", 0)
        sound_hotkey = sound.get("hotkeys", [])
        
        card = ttk.LabelFrame(self.inner_sound_grid_frame, text=sound_name[:MAX_DISPLAY_NAME_LENGTH], padding=5)
        card.pack(fill=X, padx=5, pady=5)
        self._interactive_widgets.append(card)
        
        select_var = tk.BooleanVar(value=sound_id in self.selected_sound_ids)
        select_check = ttk.Checkbutton(card, variable=select_var, command=lambda: self._toggle_sound_selection(sound_id, select_var))
        select_check.grid(row=0, column=0, rowspan=2, padx=5, pady=5, sticky="w")
        ToolTip(select_check, "Select this sound for batch operations like removal or playback.")
        self._interactive_widgets.append(select_check)
        
        play_button = ttk.Button(card, text="Play", command=lambda: self.play_sound(sound_id), bootstyle="success")
        play_button.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ToolTip(play_button, "Play this sound.")
        self._interactive_widgets.append(play_button)
        
        stop_button = ttk.Button(card, text="Stop", command=lambda: self.stop_sound(sound_id), bootstyle="danger")
        stop_button.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        ToolTip(stop_button, "Stop this sound.")
        self._interactive_widgets.append(stop_button)
        
        loop_var = tk.BooleanVar(value=sound.get("loop", False))
        sound["_loop_var_ui"] = loop_var
        loop_check = ttk.Checkbutton(card, text="Loop", variable=loop_var, command=lambda: self._update_sound_loop(sound_id, loop_var.get()))
        loop_check.grid(row=0, column=2, padx=5, pady=2, sticky="w")
        ToolTip(loop_check, "Enable to loop this sound during playback.")
        self._interactive_widgets.append(loop_check)
        
        volume_var = tk.DoubleVar(value=sound.get("volume", 1.0) * 100)
        sound["_volume_var_ui"] = volume_var
        volume_label = ttk.Label(card, text="Volume:")
        volume_label.grid(row=0, column=3, padx=5, pady=2, sticky="w")
        volume_scale = ttk.Scale(card, from_=0, to=100, variable=volume_var, command=lambda v: self._update_sound_volume(sound_id, float(v) / 100))
        volume_scale.grid(row=0, column=4, padx=5, pady=2, sticky="ew")
        self._interactive_widgets.append(volume_scale)
        volume_value_label = ttk.Label(card, text=f"{int(volume_var.get())}%")
        volume_var.trace_add("write", lambda name, index, mode: volume_value_label.config(text=f"{int(volume_var.get())}%"))
        volume_value_label.grid(row=0, column=5, padx=5, pady=2, sticky="w")
        ToolTip(volume_scale, "Adjust the volume for this sound.")
        
        hotkey_label_var = tk.StringVar(value=get_hotkey_display_string(sound_hotkey))
        sound["_hotkey_label_var_ui"] = hotkey_label_var
        hotkey_label = ttk.Label(card, text="Hotkey:")
        hotkey_label.grid(row=1, column=2, padx=5, pady=2, sticky="w")
        hotkey_display = ttk.Label(card, textvariable=hotkey_label_var, bootstyle="info")
        hotkey_display.grid(row=1, column=3, padx=5, pady=2, sticky="ew")
        ToolTip(hotkey_display, "The hotkey assigned to play this sound.")
        
        assign_hotkey_button = ttk.Button(card, text="Assign Hotkey", command=lambda: self._assign_sound_hotkey(sound_id, sound_name, hotkey_label_var), bootstyle=self.assign_button_style_var.get())
        assign_hotkey_button.grid(row=1, column=4, padx=5, pady=2, sticky="ew")
        self._interactive_widgets.append(assign_hotkey_button)
        ToolTip(assign_hotkey_button, "Click to assign a hotkey to play this sound.")
        
        clear_hotkey_button = ttk.Button(card, text="Clear", command=lambda: self._clear_sound_hotkey(sound_id, hotkey_label_var), bootstyle="secondary")
        clear_hotkey_button.grid(row=1, column=5, padx=5, pady=2, sticky="ew")
        self._interactive_widgets.append(clear_hotkey_button)
        ToolTip(clear_hotkey_button, "Clear the hotkey for this sound.")
        
        rename_button = ttk.Button(card, text="Rename", command=lambda: self._rename_sound(sound_id), bootstyle="info")
        rename_button.grid(row=0, column=6, padx=5, pady=2, sticky="ew")
        self._interactive_widgets.append(rename_button)
        ToolTip(rename_button, "Rename this sound.")
        
        card.columnconfigure(3, weight=1)
        card.columnconfigure(4, weight=1)

    def _toggle_sound_selection(self, sound_id, select_var):
        if select_var.get():
            self.selected_sound_ids.add(sound_id)
        else:
            self.selected_sound_ids.discard(sound_id)
        logging.debug(f"Sound {sound_id} {'selected' if select_var.get() else 'deselected'}")

    def _update_sound_volume(self, sound_id, volume):
        sound = next((s for s in self.sound_manager.sounds if s["id"] == sound_id), None)
        if sound:
            sound["volume"] = volume
            self.sound_manager.save_config()
            logging.debug(f"Updated volume for sound {sound_id} to {volume}")

    def _update_sound_loop(self, sound_id, loop_status):
        self.sound_manager.set_sound_loop(sound_id, loop_status)

    def _rename_sound(self, sound_id):
        sound = next((s for s in self.sound_manager.sounds if s["id"] == sound_id), None)
        if not sound:
            self.show_status_message("Sound not found for renaming.", bootstyle="danger")
            return
        
        new_name = simpledialog.askstring("Rename Sound", "Enter new sound name:", initialvalue=sound["name"], parent=self)
        if new_name:
            try:
                self.sound_manager.rename_sound(sound_id, new_name)
                self.populate_sound_grid()
                self.show_status_message(f"Renamed sound to '{new_name}'", bootstyle="success")
            except Exception as e:
                self.show_status_message(f"Failed to rename sound: {e}", bootstyle="danger")

    def _clear_sound_hotkey(self, sound_id, hotkey_label_var):
        try:
            self.sound_manager.update_sound_hotkeys(sound_id, [])
            hotkey_label_var.set("Not Assigned")
            self.keybind_manager.update_hotkeys()
            sound = next((s for s in self.sound_manager.sounds if s["id"] == sound_id), None)
            if sound:
                self.show_status_message(f"Cleared hotkey for '{sound['name']}'", bootstyle="success")
            else:
                self.show_status_message("Sound not found.", bootstyle="danger")
        except Exception as e:
            logging.error(f"Failed to clear hotkey for sound {sound_id}: {e}", exc_info=True)
            messagebox.showerror("Error", f"Failed to clear hotkey: {e}")

    def _clear_global_hotkey(self, action):
        try:
            self.sound_manager.set_global_hotkey(action, [])
            hotkey_var = self.stop_all_hotkey_var if action == "stop_all" else self.toggle_mic_to_mixer_hotkey_var
            hotkey_var.set("Not Assigned")
            self.keybind_manager.update_hotkeys()
            self.show_status_message(f"Cleared global hotkey for '{action}'", bootstyle="success")
        except Exception as e:
            logging.error(f"Failed to clear global hotkey for {action}: {e}", exc_info=True)
            messagebox.showerror("Error", f"Failed to clear hotkey: {e}")

    def _assign_sound_hotkey(self, sound_id, sound_name, hotkey_label_var):
        def on_hotkey_recorded(hotkey_list_strings):
            try:
                if not hotkey_list_strings:
                    self.show_status_message(f"Hotkey assignment for '{sound_name}' cancelled.", bootstyle="info")
                    hotkey_label_var.set("Not Assigned")
                    self.keybind_manager.update_hotkeys()
                    return

                all_hotkeys = self.sound_manager.get_all_assigned_hotkeys()
                if tuple(sorted(hotkey_list_strings)) in all_hotkeys:
                    self.show_status_message(f"Hotkey '{get_hotkey_display_string(hotkey_list_strings)}' is already assigned.", bootstyle="warning")
                    hotkey_label_var.set(get_hotkey_display_string(self.sound_manager.sounds[sound_id]["hotkeys"]))
                    return

                hotkey_obj = Hotkey.from_json_serializable(hotkey_list_strings)
                self.sound_manager.update_sound_hotkeys(sound_id, hotkey_obj.to_json_serializable())
                hotkey_label_var.set(str(hotkey_obj))
                self.keybind_manager.update_hotkeys()
                self.show_status_message(f"Hotkey assigned to '{sound_name}': {str(hotkey_obj)}", bootstyle="success")
            except Exception as e:
                logging.error(f"Failed to assign hotkey for sound {sound_name}: {e}", exc_info=True)
                messagebox.showerror("Error", f"Failed to assign hotkey: {e}")

        self.key_assignment_in_progress_var.set(True)
        self.keybind_manager.set_hotkey_assignment_target(sound_id)
        self._update_assign_button_styles()
        self.disable_ui_for_hotkey_assignment()
        HotkeyRecorder(self, sound_name, hotkey_label_var, lambda combo: [on_hotkey_recorded(combo), self.enable_ui_after_hotkey_assignment(), self.key_assignment_in_progress_var.set(False), self.keybind_manager.set_hotkey_assignment_target(None), self._update_assign_button_styles()])

    def _assign_global_hotkey(self, action):
        def on_hotkey_recorded(hotkey_list_strings):
            try:
                if not hotkey_list_strings:
                    self.show_status_message(f"Hotkey assignment for '{action}' cancelled.", bootstyle="info")
                    hotkey_var = self.stop_all_hotkey_var if action == "stop_all" else self.toggle_mic_to_mixer_hotkey_var
                    hotkey_var.set("Not Assigned")
                    self.keybind_manager.update_hotkeys()
                    return

                all_hotkeys = self.sound_manager.get_all_assigned_hotkeys()
                if tuple(sorted(hotkey_list_strings)) in all_hotkeys:
                    self.show_status_message(f"Hotkey '{get_hotkey_display_string(hotkey_list_strings)}' is already assigned.", bootstyle="warning")
                    hotkey_var = self.stop_all_hotkey_var if action == "stop_all" else self.toggle_mic_to_mixer_hotkey_var
                    hotkey_var.set(get_hotkey_display_string(self.sound_manager.global_hotkeys.get(action, [])))
                    return

                hotkey_obj = Hotkey.from_json_serializable(hotkey_list_strings)
                self.sound_manager.set_global_hotkey(action, hotkey_obj.to_json_serializable())
                hotkey_var = self.stop_all_hotkey_var if action == "stop_all" else self.toggle_mic_to_mixer_hotkey_var
                hotkey_var.set(str(hotkey_obj))
                self.keybind_manager.update_hotkeys()
                self.show_status_message(f"Global hotkey assigned for '{action}': {str(hotkey_obj)}", bootstyle="success")
            except Exception as e:
                logging.error(f"Failed to assign global hotkey for {action}: {e}", exc_info=True)
                messagebox.showerror("Error", f"Failed to assign hotkey: {e}")

        self.key_assignment_in_progress_var.set(True)
        self.keybind_manager.set_hotkey_assignment_target(action)
        self._update_assign_button_styles()
        self.disable_ui_for_hotkey_assignment()
        hotkey_var = self.stop_all_hotkey_var if action == "stop_all" else self.toggle_mic_to_mixer_hotkey_var
        HotkeyRecorder(self, action, hotkey_var, lambda combo: [on_hotkey_recorded(combo), self.enable_ui_after_hotkey_assignment(), self.key_assignment_in_progress_var.set(False), self.keybind_manager.set_hotkey_assignment_target(None), self._update_assign_button_styles()])

    def _update_assign_button_styles(self):
        if self.key_assignment_in_progress_var.get():
            self.assign_button_style_var.set("warning")
            for widget in self._interactive_widgets:
                if isinstance(widget, ttk.Button) and widget.cget("text") in ["Assign", "Assign Hotkey"]:
                    widget.configure(text="Press Keys...", bootstyle="warning")
        else:
            self.assign_button_style_var.set("primary")
            for widget in self._interactive_widgets:
                if isinstance(widget, ttk.Button) and widget.cget("text") == "Press Keys...":
                    original_text = "Assign" if widget in [self.assign_stop_all_button, self.assign_toggle_mic_button] else "Assign Hotkey"
                    widget.configure(text=original_text, bootstyle="primary")

    def disable_ui_for_hotkey_assignment(self):
        for widget in self._interactive_widgets:
            if isinstance(widget, (ttk.Button, ttk.Checkbutton, ttk.Combobox, ttk.Scale)):
                widget.configure(state="disabled")

    def enable_ui_after_hotkey_assignment(self):
        for widget in self._interactive_widgets:
            if isinstance(widget, (ttk.Button, ttk.Checkbutton, ttk.Combobox, ttk.Scale)):
                widget.configure(state="normal")

    def _populate_device_dropdowns(self):
        if not self.audio_manager.output_devices:
            self.show_status_message("No output devices detected.", bootstyle="danger")
        else:
            output_names = [dev["name"][:MAX_DEVICE_NAME_LENGTH] for dev in self.audio_manager.output_devices]
            self.output_device_combobox["values"] = output_names
            current_output_name = self.audio_manager.get_pyaudio_device_name_by_index(self.audio_manager.current_output_device_id_pyaudio)
            if current_output_name in output_names:
                self.output_device_combobox.set(current_output_name[:MAX_DEVICE_NAME_LENGTH])
            else:
                self.output_device_combobox.set(output_names[0] if output_names else "None")

        if not self.audio_manager.input_devices:
            self.show_status_message("No input devices detected.", bootstyle="danger")
        else:
            input_names = [dev["name"][:MAX_DEVICE_NAME_LENGTH] for dev in self.audio_manager.input_devices]
            self.input_device_combobox["values"] = input_names
            current_input_name = self.audio_manager.get_pyaudio_device_name_by_index(self.audio_manager.current_input_device_id_pyaudio)
            if current_input_name in input_names:
                self.input_device_combobox.set(current_input_name[:MAX_DEVICE_NAME_LENGTH])
            else:
                self.input_device_combobox.set(input_names[0] if input_names else "None")

        monitor_names = [dev["name"][:MAX_DEVICE_NAME_LENGTH] for dev in self.audio_manager.output_devices]
        self.soundboard_monitor_device_combobox["values"] = monitor_names
        self.mic_monitor_device_combobox["values"] = monitor_names
        soundboard_monitor_name = self.audio_manager.get_pyaudio_device_name_by_index(self.audio_manager.current_soundboard_monitor_device_id_pyaudio)
        mic_monitor_name = self.audio_manager.get_pyaudio_device_name_by_index(self.audio_manager.current_mic_monitor_device_id_pyaudio)
        if soundboard_monitor_name in monitor_names:
            self.soundboard_monitor_device_combobox.set(soundboard_monitor_name[:MAX_DEVICE_NAME_LENGTH])
        else:
            self.soundboard_monitor_device_combobox.set(monitor_names[0] if monitor_names else "None")
        if mic_monitor_name in monitor_names:
            self.mic_monitor_device_combobox.set(mic_monitor_name[:MAX_DEVICE_NAME_LENGTH])
        else:
            self.mic_monitor_device_combobox.set(monitor_names[0] if monitor_names else "None")

    def _on_output_device_selected(self, event):
        selected_name = self.output_device_combobox.get()
        for dev in self.audio_manager.output_devices:
            if dev["name"].startswith(selected_name):
                self.audio_manager.current_output_device_id_pyaudio = dev["index"]
                self.audio_manager.restart_audio_streams()
                self.show_status_message(f"Main output device set to: {selected_name}", bootstyle="success")
                self._save_settings()
                break

    def _on_input_device_selected(self, event):
        selected_name = self.input_device_combobox.get()
        for dev in self.audio_manager.input_devices:
            if dev["name"].startswith(selected_name):
                self.audio_manager.current_input_device_id_pyaudio = dev["index"]
                self.audio_manager.restart_audio_streams()
                self.show_status_message(f"Input device set to: {selected_name}", bootstyle="success")
                self._save_settings()
                break

    def _on_soundboard_monitor_device_selected(self, event):
        selected_name = self.soundboard_monitor_device_combobox.get()
        for dev in self.audio_manager.output_devices:
            if dev["name"].startswith(selected_name):
                self.audio_manager.current_soundboard_monitor_device_id_pyaudio = dev["index"]
                self.audio_manager.restart_audio_streams()
                self.show_status_message(f"Soundboard monitor device set to: {selected_name}", bootstyle="success")
                self._save_settings()
                break

    def _on_mic_monitor_device_selected(self, event):
        selected_name = self.mic_monitor_device_combobox.get()
        for dev in self.audio_manager.output_devices:
            if dev["name"].startswith(selected_name):
                self.audio_manager.current_mic_monitor_device_id_pyaudio = dev["index"]
                self.audio_manager.restart_audio_streams()
                self.show_status_message(f"Mic monitor device set to: {selected_name}", bootstyle="success")
                self._save_settings()
                break

    def _on_master_volume_changed(self, value):
        self.audio_manager.master_volume = float(value) / 100.0
        self._save_settings()

    def _on_master_volume_released(self, event):
        self.show_status_message(f"Master volume set to {int(self.master_volume_var.get())}%", bootstyle="info")

    def _on_soundboard_monitor_volume_changed(self, value):
        self.audio_manager.soundboard_monitor_volume = float(value) / 100.0
        self._save_settings()

    def _on_soundboard_monitor_volume_released(self):
        self.show_status_message(f"Soundboard monitor volume set to {int(self.soundboard_monitor_volume_var.get())}%", bootstyle="info")

    def _on_mic_monitor_volume_changed(self, value):
        self.audio_manager.mic_monitor_volume = float(value) / 100.0
        self._save_settings()

    def _on_mic_monitor_volume_released(self, event):
        self.show_status_message(f"Mic monitor volume set to {int(self.mic_monitor_volume_var.get())}%", bootstyle="info")

    def _on_single_sound_mode_changed(self):
        self.audio_manager.mixer.set_single_sound_mode(self.single_sound_mode_var.get())
        self._save_settings()
        self.show_status_message(f"Single sound mode {'enabled' if self.single_sound_mode_var.get() else 'disabled'}", bootstyle="info")

    def _on_auto_start_mic_changed(self):
        self._save_settings()
        self.show_status_message(f"Auto-start mic in mix {'enabled' if self.auto_start_include_mic_in_mix_var.get() else 'disabled'}", bootstyle="info")

    def _on_theme_changed(self, event):
        try:
            self.style.theme_use(self.current_theme_var.get())
            self._save_settings()
            self.show_status_message(f"Theme changed to {self.current_theme_var.get()}", bootstyle="success")
        except Exception as e:
            logging.error(f"Failed to change theme: {e}")
            self.show_status_message(f"Failed to change theme: {e}", bootstyle="danger")

    def _save_settings(self):
        settings = {
            "theme": self.current_theme_var.get(),
            "single_sound_mode": self.single_sound_mode_var.get(),
            "auto_start_include_mic_in_mix": self.auto_start_include_mic_in_mix_var.get(),
            "master_volume": self.master_volume_var.get(),
            "soundboard_monitor_volume": self.soundboard_monitor_volume_var.get(),
            "mic_monitor_volume": self.mic_monitor_volume_var.get(),
            "output_device_id": self.audio_manager.current_output_device_id_pyaudio,
            "input_device_id": self.audio_manager.current_input_device_id_pyaudio,
            "soundboard_monitor_device_id": self.audio_manager.current_soundboard_monitor_device_id_pyaudio,
            "mic_monitor_device_id": self.audio_manager.current_mic_monitor_device_id_pyaudio,
            "include_mic_in_mix": self.include_mic_in_mix_var.get(),
            "soundboard_monitor_enabled": self.soundboard_monitor_enabled_var.get(),
            "mic_monitor_enabled": self.mic_monitor_enabled_var.get(),
        }
        self.app_settings_manager.save_settings(settings)

    def _apply_settings_to_ui(self):
        settings = self.app_settings_manager.get_settings()
        
        self.current_theme_var.set(settings.get("theme", "litera"))
        try:
            self.style.theme_use(self.current_theme_var.get())
        except Exception as e:
            logging.warning(f"Failed to apply theme {self.current_theme_var.get()}: {e}")
        
        self.single_sound_mode_var.set(settings.get("single_sound_mode", False))
        self.audio_manager.mixer.set_single_sound_mode(self.single_sound_mode_var.get())
        
        self.auto_start_include_mic_in_mix_var.set(settings.get("auto_start_include_mic_in_mix", False))
        self.include_mic_in_mix_var.set(settings.get("include_mic_in_mix", False))
        self.soundboard_monitor_enabled_var.set(settings.get("soundboard_monitor_enabled", True))
        self.mic_monitor_enabled_var.set(settings.get("mic_monitor_enabled", False))
        
        self.master_volume_var.set(settings.get("master_volume", 100.0))
        self.soundboard_monitor_volume_var.set(settings.get("soundboard_monitor_volume", 100.0))
        self.mic_monitor_volume_var.set(settings.get("mic_monitor_volume", 100.0))
        
        self.audio_manager.master_volume = self.master_volume_var.get() / 100.0
        self.audio_manager.soundboard_monitor_volume = self.soundboard_monitor_volume_var.get() / 100.0
        self.audio_manager.mic_monitor_volume = self.mic_monitor_volume_var.get() / 100.0
        
        output_device_id = settings.get("output_device_id")
        if output_device_id is not None:
            self.audio_manager.current_output_device_id_pyaudio = output_device_id
        input_device_id = settings.get("input_device_id")
        if input_device_id is not None:
            self.audio_manager.current_input_device_id_pyaudio = input_device_id
        soundboard_monitor_device_id = settings.get("soundboard_monitor_device_id")
        if soundboard_monitor_device_id is not None:
            self.audio_manager.current_soundboard_monitor_device_id_pyaudio = soundboard_monitor_device_id
        mic_monitor_device_id = settings.get("mic_monitor_device_id")
        if mic_monitor_device_id is not None:
            self.audio_manager.current_mic_monitor_device_id_pyaudio = mic_monitor_device_id
        
        self.audio_manager.start_main_audio_output_stream()
        if self.auto_start_include_mic_in_mix_var.get():
            self.include_mic_in_mix_var.set(True)
            self.audio_manager.start_mic_input_stream()
        if self.soundboard_monitor_enabled_var.get():
            self.audio_manager.start_soundboard_monitor_stream()
        if self.mic_monitor_enabled_var.get():
            self.audio_manager.start_mic_monitor_stream()
        
        self._populate_device_dropdowns()
        
        if self.audio_manager.virtual_mic_device_id is None:
            self.show_status_message("VB-CABLE not detected. Install it for proper virtual mic functionality.", bootstyle="warning")
        
        self.stop_all_hotkey_var.set(get_hotkey_display_string(self.sound_manager.global_hotkeys.get("stop_all", [])))
        self.toggle_mic_to_mixer_hotkey_var.set(get_hotkey_display_string(self.sound_manager.global_hotkeys.get("toggle_mic_to_mixer", [])))

    def populate_sound_grid(self):
        for widget in self.inner_sound_grid_frame.winfo_children():
            widget.destroy()
        self._interactive_widgets = [w for w in self._interactive_widgets if not isinstance(w, ttk.LabelFrame) or w.winfo_parent() != str(self.inner_sound_grid_frame)]
        
        for sound in self.sound_manager.sounds:
            self._create_sound_card(sound)
        
        self.sound_grid_canvas.configure(scrollregion=self.sound_grid_canvas.bbox("all"))

    def add_sound(self):
        file_paths = filedialog.askopenfilenames(filetypes=SUPPORTED_FORMATS)
        if not file_paths:
            return
        
        for file_path in file_paths:
            custom_name = simpledialog.askstring("Sound Name", f"Enter name for sound '{os.path.basename(file_path)}':", parent=self)
            try:
                sound = self.sound_manager.add_sound(file_path, custom_name)
                self._create_sound_card(sound)
                self.show_status_message(f"Added sound: {sound['name']}", bootstyle="success")
            except Exception as e:
                self.show_status_message(f"Failed to add sound: {e}", bootstyle="danger")
        
        self.sound_grid_canvas.configure(scrollregion=self.sound_grid_canvas.bbox("all"))

    def remove_selected_sounds(self):
        if not self.selected_sound_ids:
            self.show_status_message("No sounds selected for removal.", bootstyle="warning")
            return
        
        if messagebox.askyesno("Confirm Removal", "Are you sure you want to remove the selected sounds?"):
            try:
                self.sound_manager.remove_sounds(self.selected_sound_ids)
                self.selected_sound_ids.clear()
                self.populate_sound_grid()
                self.keybind_manager.update_hotkeys()
                self.show_status_message("Selected sounds removed.", bootstyle="success")
            except Exception as e:
                self.show_status_message(f"Failed to remove sounds: {e}", bootstyle="danger")

    def play_sound(self, sound_id):
        sound = next((s for s in self.sound_manager.sounds if s["id"] == sound_id), None)
        if not sound:
            self.show_status_message("Sound not found.", bootstyle="danger")
            return
        
        try:
            data, _ = sf.read(sound["path"], dtype='float32')
            if len(data.shape) == 1:
                data = np.column_stack((data, data))
            
            self.audio_manager.mixer.add_sound(data, sound["volume"], sound["loop"], sound["id"], sound["name"])
            self.show_status_message(f"Playing sound: {sound['name']}", bootstyle="success")
        except Exception as e:
            logging.error(f"Failed to play sound {sound['name']}: {e}", exc_info=True)
            self.show_status_message(f"Failed to play sound: {e}", bootstyle="danger")

    def play_selected_sound(self):
        if len(self.selected_sound_ids) != 1:
            self.show_status_message("Select exactly one sound to play.", bootstyle="warning")
            return
        sound_id = next(iter(self.selected_sound_ids))
        self.play_sound(sound_id)

    def stop_sound(self, sound_id):
        if self.audio_manager.mixer.remove_sound_by_id(sound_id):
            sound = next((s for s in self.sound_manager.sounds if s["id"] == sound_id), None)
            if sound:
                self.show_status_message(f"Stopped sound: {sound['name']}", bootstyle="success")
            else:
                self.show_status_message("Sound not found.", bootstyle="danger")
        else:
            self.show_status_message("Sound not currently playing.", bootstyle="warning")

    def stop_all_sounds(self):
        self.audio_manager.mixer.clear_sounds()
        self.show_status_message("All sounds stopped.", bootstyle="success")

    def toggle_include_mic_in_mix(self):
        self.audio_manager.include_mic_in_mix = self.include_mic_in_mix_var.get()
        if self.audio_manager.include_mic_in_mix:
            self.audio_manager.start_mic_input_stream()
            self.mic_status_info.set("Mic: On")
            self.show_status_message("Microphone included in mix.", bootstyle="success")
        else:
            self.audio_manager.stop_mic_input_stream()
            self.mic_status_info.set("Mic: Off")
            self.show_status_message("Microphone removed from mix.", bootstyle="success")
        self._save_settings()

    def toggle_mic_to_mixer_from_hotkey(self):
        self.include_mic_in_mix_var.set(not self.include_mic_in_mix_var.get())
        self.toggle_include_mic_in_mix()

    def toggle_soundboard_monitor(self):
        self.audio_manager.soundboard_monitor_enabled = self.soundboard_monitor_enabled_var.get()
        if self.audio_manager.soundboard_monitor_enabled:
            self.audio_manager.start_soundboard_monitor_stream()
            self.show_status_message("Soundboard monitor enabled.", bootstyle="success")
        else:
            self.audio_manager.stop_soundboard_monitor_stream()
            self.show_status_message("Soundboard monitor disabled.", bootstyle="success")
        self._save_settings()

    def toggle_mic_monitor(self):
        self.audio_manager.mic_monitor_enabled = self.mic_monitor_enabled_var.get()
        if self.audio_manager.mic_monitor_enabled:
            self.audio_manager.start_mic_monitor_stream()
            self.show_status_message("Mic monitor enabled.", bootstyle="success")
        else:
            self.audio_manager.stop_mic_monitor_stream()
            self.show_status_message("Mic monitor disabled.", bootstyle="success")
        self._save_settings()

    def show_status_message(self, message, bootstyle="info"):
        self.status_message_var.set(message)
        self.status_label.configure(bootstyle=bootstyle)
        self.after(5000, lambda: self.status_label.configure(style="info.TLabel"))
        logging.info(f"Status message: {message}")

    def _update_now_playing_periodically(self):
        while True:
            try:
                with current_playing_sound_details_lock:
                    names = current_playing_sound_details.get("names", [])
                    is_active = current_playing_sound_details.get("active", False)
                
                if is_active:
                    if names:
                        display_names = ", ".join(names[:3])
                        if len(names) > 3:
                            display_names += f" (+{len(names) - 3} more)"
                        self.currently_playing_info.set(f"Now Playing: {display_names}")
                        self.soundboard_output_status_info.set("Soundboard Output: On")
                    else:
                        self.currently_playing_info.set("Now Playing: Mic Only")
                        self.soundboard_output_status_info.set("Soundboard Output: On")
                else:
                    self.currently_playing_info.set("Now Playing: None")
                    self.soundboard_output_status_info.set("Soundboard Output: Off")
                
                self.update()
                time.sleep(0.5)
            except Exception as e:
                logging.error(f"Error in now playing updater thread: {e}", exc_info=True)
                time.sleep(1)

    def _on_app_closure(self):
        self.audio_manager.close()
        self.keybind_manager.stop()
        self.destroy()
        logging.info("Application closed.")

    def get_pynput_key_string(self, key):
        return get_pynput_key_string(key)

if __name__ == "__main__":
    app = SoundboardApp()
    app.mainloop()