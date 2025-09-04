import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, font
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import json
import os
import threading
import soundfile as sf
import numpy as np
import pyaudio
import time
import re
import sys
import uuid
import pydub
import importlib.metadata
from collections import deque
from threading import Lock, Event
from pynput import keyboard, mouse
# Removed unused import
import enum
import logging
import webbrowser
import platform
import subprocess

# --- Configuration and Constants ---
def get_app_data_dir():
    """Returns the platform-specific application data directory."""
    if platform.system() == 'Windows':
        return os.path.join(os.getenv('APPDATA'), 'WarpBoard')
    elif platform.system() == 'Darwin': # macOS
        return os.path.join(os.path.expanduser('~/Library/Application Support'), 'WarpBoard')
    else: # Linux
        return os.path.join(os.path.expanduser('~/.config'), 'WarpBoard')

APP_DATA_DIR = get_app_data_dir()
SOUNDS_DIR = os.path.join(APP_DATA_DIR, "sounds")
CONFIG_DIR = os.path.join(APP_DATA_DIR, "config")
CONFIG_FILE = os.path.join(CONFIG_DIR, "soundboard_config.json")
APP_SETTINGS_FILE = os.path.join(CONFIG_DIR, "app_settings.json")
LOG_FILE = os.path.join(APP_DATA_DIR, "warpboard.log")

# --- THIS IS THE CORRECTED BLOCK ---
# Determine the root directory for bundled assets, which works for both
# development (running .py) and for an installed application (.exe).
if getattr(sys, 'frozen', False):
    # If the application is run as a bundled executable (e.g., from an installer),
    # the root directory is the directory of the executable itself.
    ROOT_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
else:
    # Otherwise (running as a normal .py script), it's the script's directory.
    ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

VB_CABLE_URL = "https://download.vb-audio.com/Download_CABLE/VBCABLE_Driver_Pack43.zip"
PROFILE_URL = "https://github.com/Imrankhadeer"
APP_VERSION = "2.4.0" # Stability and Bugfix Release

DOCS_URL = "https://github.com/Imrankhadeer/Warpboard/wiki"
ISSUES_URL = "https://github.com/Imrankhadeer/Warpboard/issues"


# --- Audio Settings ---
SUPPORTED_FORMATS = [("Audio Files", "*.wav *.mp3 *.ogg *.flac")]
VIRTUAL_MIC_NAME_PARTIAL = "CABLE Input"
SAMPLE_RATE = 44100
CHANNELS = 2
FRAME_SIZE = 1024
# NOTE: The rest of the file remains the same...

# --- UI and Validation ---
INVALID_FILENAME_CHARS = r'[<>:"/\\|?*]'
MAX_DISPLAY_NAME_LENGTH = 30
MAX_DEVICE_NAME_LENGTH = 45

# --- Global State ---
current_playing_sound_details = {}
current_playing_sound_details_lock = threading.Lock()


# --- Hotkey System ---
class KeyModifier(enum.Flag):
    NONE = 0
    SHIFT = enum.auto()
    CTRL = enum.auto()
    ALT = enum.auto()
    SUPER = enum.auto()

class KeyCode(enum.Enum):
    UNKNOWN = 'unknown'

class Hotkey:
    def __init__(self, key_code: KeyCode, modifiers: KeyModifier = KeyModifier.NONE, raw_key: str = None):
        self.key_code, self.modifiers, self.raw_key = key_code, modifiers, raw_key

    def __hash__(self):
        return hash((self.key_code, self.modifiers, self.raw_key))

    def __eq__(self, other):
        return isinstance(other, Hotkey) and (self.key_code, self.modifiers, self.raw_key) == (other.key_code, other.modifiers, other.raw_key)

    def __repr__(self):
        parts = []
        if KeyModifier.CTRL in self.modifiers: parts.append('Ctrl')
        if KeyModifier.SHIFT in self.modifiers: parts.append('Shift')
        if KeyModifier.ALT in self.modifiers: parts.append('Alt')
        if KeyModifier.SUPER in self.modifiers: parts.append('Super')
        key_name = self.raw_key.replace('_', ' ').title() if self.raw_key else self.key_code.value
        if key_name.startswith('<vk-'): key_name = f"VK-{key_name[4:-1]}"
        parts.append(key_name)
        return '+'.join(parts)

    def to_json_serializable(self):
        parts = []
        if KeyModifier.CTRL in self.modifiers: parts.append('ctrl')
        if KeyModifier.SHIFT in self.modifiers: parts.append('shift')
        if KeyModifier.ALT in self.modifiers: parts.append('alt')
        if KeyModifier.SUPER in self.modifiers: parts.append('cmd')
        parts.append(self.raw_key if self.raw_key else self.key_code.value)
        return sorted(parts)

    @classmethod
    def from_json_serializable(cls, hotkey_list: list[str]):
        modifiers, key_str = KeyModifier.NONE, None
        for s in hotkey_list:
            if s == 'ctrl': modifiers |= KeyModifier.CTRL
            elif s == 'shift': modifiers |= KeyModifier.SHIFT
            elif s == 'alt': modifiers |= KeyModifier.ALT
            elif s == 'cmd': modifiers |= KeyModifier.SUPER
            else: key_str = s
        if not key_str: raise ValueError("Hotkey must contain a non-modifier key.")
        return cls(KeyCode.UNKNOWN, modifiers, raw_key=key_str)


# --- Utility Functions ---
def run_powershell_script(script_name, *args):
    """Runs a PowerShell script from the 'scripts' directory and returns its output."""
    try:
        # This now correctly uses the fixed ROOT_DIR
        script_path = os.path.join(ROOT_DIR, 'scripts', script_name)
        logging.info(f"Attempting to run PowerShell script from path: {script_path}")

        if not os.path.exists(script_path):
            logging.error(f"Script file not found at resolved path: {script_path}")
            raise FileNotFoundError(f"Script not found: {script_path}")

        command = ['powershell.exe', '-ExecutionPolicy', 'Bypass', '-File', script_path] + list(args)
        
        result = subprocess.run(command, capture_output=True, text=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
        return result.stdout.strip()
    except subprocess.CalledProcessError as e:
        logging.error(f"Error running PowerShell script '{script_name}': {e.stderr}")
        raise
    except FileNotFoundError:
        logging.error(f"powershell.exe not found. Please ensure it is in your system's PATH.")
        raise
    except Exception as e:
        logging.error(f"An unexpected error occurred while running PowerShell script '{script_name}': {e}")
        raise

def get_executable_path(name):
    """Gets the path to an executable, now simplified to use ROOT_DIR."""
    # The logic correctly uses ROOT_DIR which handles both dev and frozen states.
    return os.path.join(ROOT_DIR, name)

def get_hotkey_display_string(hotkey_list):
    if not hotkey_list: return "Not Assigned"
    try:
        return str(Hotkey.from_json_serializable(hotkey_list))
    except (ValueError, TypeError) as e:
        logging.warning(f"Failed to create Hotkey object for display {hotkey_list}: {e}")
        return "+".join(s.title() for s in hotkey_list)

def get_pynput_key_string(key):
    if isinstance(key, keyboard.Key):
        key_name = key.name
        if key_name in ('ctrl_l', 'ctrl_r'): return 'ctrl'
        if key_name in ('alt_l', 'alt_gr', 'alt_r'): return 'alt'
        if key_name in ('shift_l', 'shift_r'): return 'shift'
        if key_name in ('cmd_l', 'cmd_r'): return 'cmd'
        return key_name
    elif isinstance(key, keyboard.KeyCode):
        return key.char.lower() if key.char and key.char.isalnum() else f"<vk_{key.vk}>"
    elif isinstance(key, mouse.Button):
        return key.name
    return None

def center_window(win):
    """Centers a tkinter window on the screen."""
    win.update_idletasks()
    width = win.winfo_width()
    height = win.winfo_height()
    x = (win.winfo_screenwidth() // 2) - (width // 2)
    y = (win.winfo_screenheight() // 2) - (height // 2)
    win.geometry(f'{width}x{height}+{x}+{y}')

# --- UI Components ---
class ToolTip(Toplevel):
    def __init__(self, widget, text_func):
        super().__init__(widget)
        self.widget = widget
        self.text_func = text_func
        self.withdraw()
        self.overrideredirect(True)
        
        self.label = ttk.Label(self, text="", justify='left', bootstyle="inverse-light", padding=5)
        self.label.pack()
        
        self._scheduled_hide = None
        self._scheduled_show = None

        self.widget.bind("<Enter>", self.on_enter)
        self.widget.bind("<Leave>", self.on_leave)
        self.label.bind("<Enter>", self.on_tooltip_enter)
        self.label.bind("<Leave>", self.on_leave)

    def on_enter(self, _=None):
        self._cancel_hide()
        self._scheduled_show = self.after(500, self.show)

    def on_leave(self, _=None):
        self._cancel_show()
        self._scheduled_hide = self.after(100, self.hide)
    
    def on_tooltip_enter(self, _=None):
        self._cancel_hide()

    def _cancel_show(self):
        if self._scheduled_show:
            self.after_cancel(self._scheduled_show)
            self._scheduled_show = None

    def _cancel_hide(self):
        if self._scheduled_hide:
            self.after_cancel(self._scheduled_hide)
            self._scheduled_hide = None

    def show(self):
        if self.state() == 'normal':
            return
        self.label.config(text=self.text_func())
        x = self.widget.winfo_rootx()
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.geometry(f"+{x}+{y}")
        self.deiconify()

    def hide(self):
        self.withdraw()

class HotkeyRecorder(Toplevel):
    def __init__(self, parent, target_name, on_complete_callback):
        super().__init__(parent)
        self.on_complete_callback = on_complete_callback
        
        self.transient(parent)
        self.title("Assign Hotkey")
        self.geometry("350x150")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        
        self.recorded_keys = set()
        self.hotkey_display_var = tk.StringVar(value="Press keys...")
        self.timeout_id = None

        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=BOTH, expand=True)

        ttk.Label(main_frame, text=f"Assigning Hotkey for:", anchor="center").pack()
        ttk.Label(main_frame, text=f"'{target_name}'", anchor="center", font="-weight bold").pack()
        ttk.Label(main_frame, textvariable=self.hotkey_display_var, bootstyle="info", font="-size 12 -weight bold", anchor="center").pack(pady=10)
        ttk.Label(main_frame, text="Press ESC to cancel. Hotkey is saved after a brief pause.", font="-size 8", bootstyle="secondary", anchor="center").pack()

        parent.center_toplevel(self)
        self.grab_set()

        self.keyboard_listener = keyboard.Listener(on_press=self._on_key_press, suppress=True)
        self.mouse_listener = mouse.Listener(on_click=self._on_mouse_click, suppress=True)
        self.keyboard_listener.start()
        self.mouse_listener.start()
        self._reset_timeout()

    def _on_key_press(self, key):
        if key == keyboard.Key.esc:
            self._on_cancel(); return False 
        key_str = get_pynput_key_string(key)
        if key_str: self.recorded_keys.add(key_str); self._update_display()
        self._reset_timeout(); return True

    def _on_mouse_click(self, _, __, button, pressed):
        if pressed and button not in [mouse.Button.left, mouse.Button.right]:
            key_str = get_pynput_key_string(button)
            if key_str: self.recorded_keys.add(key_str); self._update_display()
            self._reset_timeout()
        return True

    def _update_display(self):
        if self.recorded_keys: self.hotkey_display_var.set(get_hotkey_display_string(list(self.recorded_keys)))

    def _reset_timeout(self):
        if self.timeout_id: self.after_cancel(self.timeout_id)
        self.timeout_id = self.after(1200, self._finalize_hotkey)

    def _stop_listeners(self):
        if self.keyboard_listener and self.keyboard_listener.is_alive(): self.keyboard_listener.stop()
        if self.mouse_listener and self.mouse_listener.is_alive(): self.mouse_listener.stop()

    def _finalize_hotkey(self):
        self._stop_listeners()
        non_modifiers = [k for k in self.recorded_keys if k not in ['ctrl', 'shift', 'alt', 'cmd']]
        self.on_complete_callback(sorted(list(self.recorded_keys)) if non_modifiers else None)
        self.destroy()

    def _on_cancel(self):
        if self.timeout_id: self.after_cancel(self.timeout_id)
        self._stop_listeners(); self.on_complete_callback(None); self.destroy()

# --- Core Logic Classes ---
class DependencyChecker:
    @staticmethod
    def run_checks():
        try:
            importlib.metadata.version("pydub")
            ffmpeg_path = get_executable_path('ffmpeg.exe')
            creation_flags = subprocess.CREATE_NO_WINDOW if platform.system() == 'Windows' else 0
            subprocess.run([ffmpeg_path, '-version'], capture_output=True, check=True, creationflags=creation_flags)
        except (Exception) as e:
            logging.critical(f"Dependency check failed for ffmpeg: {e}")
            messagebox.showerror("Dependency Error", "FFmpeg is missing or not configured correctly. Please ensure ffmpeg.exe is in the application directory.")
            sys.exit(1)

class AppSettingsManager:
    def __init__(self): self.settings = self.load_settings()
    def load_settings(self):
        if os.path.exists(APP_SETTINGS_FILE):
            try:
                with open(APP_SETTINGS_FILE, 'r') as f: return json.load(f)
            except json.JSONDecodeError as e: logging.error(f"Failed to load app settings file: {e}")
        return {}
    def save_settings(self, settings_dict):
        self.settings.update(settings_dict)
        try:
            with open(APP_SETTINGS_FILE, 'w') as f: json.dump(self.settings, f, indent=4)
        except IOError as e: logging.error(f"Failed to save app settings: {e}")
    def get_setting(self, key, default=None): return self.settings.get(key, default)

class MixingBuffer:
    def __init__(self): self.sounds, self.lock, self.single_sound_mode = deque(), Lock(), False
    def set_single_sound_mode(self, enabled):
        with self.lock: self.single_sound_mode = enabled
    def add_sound(self, data, volume, loop, sound_id, sound_name):
        with self.lock:
            if self.single_sound_mode:
                self.sounds.clear()
            elif not loop:
                self.sounds = deque(s for s in self.sounds if s["id"] != sound_id)

            self.sounds.append({"id": sound_id, "data": data, "volume": volume, "loop": loop, "index": 0, "name": sound_name})
    def mix_audio(self, frames):
        with self.lock:
            mixed = np.zeros((frames, CHANNELS), dtype=np.float32)
            sounds_to_remove, playing_names = [], []
            for sound in list(self.sounds):
                start, end, data_len = sound["index"], sound["index"] + frames, len(sound["data"])
                chunk = sound["data"][start:end]
                actual_chunk_len = len(chunk)
                if actual_chunk_len < frames:
                    if sound["loop"]:
                        remaining_frames = frames - actual_chunk_len
                        sound["index"] = remaining_frames % data_len
                        while remaining_frames > 0:
                            frames_to_take = min(remaining_frames, data_len)
                            chunk = np.concatenate((chunk, sound["data"][0:frames_to_take]))
                            remaining_frames -= frames_to_take
                    else:
                        sounds_to_remove.append(sound)
                        chunk = np.pad(chunk, ((0, frames - actual_chunk_len), (0, 0)), 'constant')
                else: sound["index"] = end % data_len if sound["loop"] else end
                mixed += chunk * sound["volume"]
                if sound not in sounds_to_remove: playing_names.append(sound["name"])
            for sound in list(sounds_to_remove):
                if sound in self.sounds: self.sounds.remove(sound)
            np.clip(mixed, -1.0, 1.0, out=mixed)
            return mixed, playing_names
    def clear_sounds(self):
        with self.lock: self.sounds.clear()
    def remove_sound_by_id(self, sound_id):
        with self.lock:
            initial_len = len(self.sounds)
            self.sounds = deque(s for s in self.sounds if s["id"] != sound_id)
            return len(self.sounds) < initial_len

class SoundManager:
    def __init__(self):
        self.sounds, self.global_hotkeys, self.sound_data_cache = [], {}, {}
        self.load_config()
    def add_sound(self, file_path, custom_name=None):
        try:
            sound_name = custom_name or os.path.splitext(os.path.basename(file_path))[0]
            sound_name = re.sub(INVALID_FILENAME_CHARS, '_', sound_name)
            output_path = os.path.join(SOUNDS_DIR, f"{sound_name}.wav")
            counter = 1
            while os.path.exists(output_path):
                output_path = os.path.join(SOUNDS_DIR, f"{sound_name}_{counter}.wav")
                counter += 1
            audio = pydub.AudioSegment.from_file(file_path).set_frame_rate(SAMPLE_RATE).set_channels(CHANNELS)
            audio.export(output_path, format="wav")
            new_sound = {"id": str(uuid.uuid4()), "name": sound_name, "path": output_path, "volume": 1.0, "hotkeys": [], "loop": False, "enabled": True, "duration": sf.info(output_path).duration}
            self.sounds.append(new_sound)
            self.save_config()
            return new_sound
        except Exception as e: logging.error(f"Failed to add sound {file_path}: {e}"); raise
    def rename_sound(self, sound_id, new_name):
        sound = self.get_sound_by_id(sound_id)
        if not sound: return None
        new_name_clean = re.sub(INVALID_FILENAME_CHARS, '_', new_name.strip())
        if not new_name_clean or any(s['name'] == new_name_clean for s in self.sounds if s['id'] != sound_id):
            raise ValueError("New name is invalid or already exists.")
        old_path, new_path = sound['path'], os.path.join(SOUNDS_DIR, f"{new_name_clean}.wav")
        if os.path.exists(new_path) and old_path.lower() != new_path.lower():
            raise ValueError("A file with the new name already exists.")
        try:
            if old_path.lower() != new_path.lower():
                os.rename(old_path, new_path)
            sound['name'], sound['path'] = new_name_clean, new_path
            self.save_config()
            return sound
        except OSError as e:
            logging.error(f"Failed to rename file from {old_path} to {new_path}: {e}")
            raise ValueError(f"Could not rename the sound file. Is it in use?")

    def remove_sounds(self, sound_ids):
        for sound_id in list(sound_ids):
            sound = self.get_sound_by_id(sound_id)
            if sound:
                try:
                    if os.path.exists(sound["path"]): os.remove(sound["path"])
                    if sound_id in self.sound_data_cache: del self.sound_data_cache[sound_id]
                    self.sounds.remove(sound)
                except Exception as e: logging.error(f"Error removing sound {sound['name']}: {e}")
        self.save_config()
    def get_sound_by_id(self, sound_id): return next((s for s in self.sounds if s["id"] == sound_id), None)
    def update_sound_property(self, sound_id, key, value):
        sound = self.get_sound_by_id(sound_id)
        if sound: sound[key] = value
    def set_global_hotkey(self, action, hotkey_list):
        self.global_hotkeys[action] = hotkey_list; self.save_config()
    def get_all_assigned_hotkeys(self):
        hotkeys = set()
        for sound in self.sounds:
            if sound.get("hotkeys"): hotkeys.add(tuple(sorted(sound["hotkeys"])))
        for hotkey_list in self.global_hotkeys.values():
            if hotkey_list: hotkeys.add(tuple(sorted(hotkey_list)))
        return hotkeys
    def save_config(self):
        try:
            with open(CONFIG_FILE, 'w') as f: json.dump({"sounds": self.sounds, "global_hotkeys": self.global_hotkeys}, f, indent=4)
        except IOError as e: logging.error(f"Error saving soundboard config: {e}")
    def load_config(self):
        if not os.path.exists(CONFIG_FILE): return
        try:
            with open(CONFIG_FILE, 'r') as f: data = json.load(f)
            self.sounds = [s for s in data.get("sounds", []) if s.get("path") and os.path.exists(s.get("path"))]
            self.global_hotkeys = data.get("global_hotkeys", {})
            for sound in self.sounds:
                if 'enabled' not in sound: sound['enabled'] = True
        except (json.JSONDecodeError, KeyError) as e: logging.error(f"Error loading config: {e}")
    def preload_sound_data(self, sound):
        try:
            data, _ = sf.read(sound["path"], dtype='float32')
            if len(data.shape) == 1: data = np.column_stack((data, data))
            self.sound_data_cache[sound["id"]] = data
        except Exception as e:
            logging.error(f"Failed to pre-load audio for '{sound['name']}': {e}")
            if sound["id"] in self.sound_data_cache: del self.sound_data_cache[sound["id"]]

class AudioOutputManager:
    def __init__(self, app):
        self.app, self.p, self.mixer = app, pyaudio.PyAudio(), MixingBuffer()
        self.output_devices, self.input_devices = self._enumerate_devices()
        self.virtual_mic_device_id = self._find_virtual_mic()
        self.main_stream, self.mic_stream, self.soundboard_monitor_stream, self.mic_monitor_stream = None, None, None, None
        self._mic_buffer, self._mic_buffer_lock = deque(maxlen=10), Lock()
        self._soundboard_monitor_buffer, self._soundboard_monitor_buffer_lock = deque(maxlen=10), Lock()
        self._mic_reader_thread, self._mic_reader_stop_event = None, Event()
        self.mic_inclusion_event = Event()
        self.master_volume = 1.0
        self.sb_monitor_volume = 0.75
        self.mic_monitor_volume = 0.75

    def set_master_volume(self, volume): self.master_volume = volume
    def set_sb_monitor_volume(self, volume): self.sb_monitor_volume = volume
    def set_mic_monitor_volume(self, volume): self.mic_monitor_volume = volume

    def _enumerate_devices(self):
            output, input_devs = [], []
            seen_output_names = set()
            seen_input_names = set()
            
            FILTERED_SUBSTRINGS = ['microsoft sound mapper', 'primary sound']

            for i in range(self.p.get_device_count()):
                dev = self.p.get_device_info_by_index(i)
                device_name = dev['name']

                if any(sub in device_name.lower() for sub in FILTERED_SUBSTRINGS):
                    continue

                try:
                    if dev.get('maxOutputChannels', 0) >= CHANNELS:
                        if device_name not in seen_output_names:
                            if self.p.is_format_supported(SAMPLE_RATE,
                                                        output_device=i,
                                                        output_channels=CHANNELS,
                                                        output_format=pyaudio.paFloat32):
                                output.append({"name": device_name, "index": i})
                                seen_output_names.add(device_name)

                    if dev.get('maxInputChannels', 0) >= CHANNELS:
                        if device_name not in seen_input_names:
                            if self.p.is_format_supported(SAMPLE_RATE,
                                                        input_device=i,
                                                        input_channels=CHANNELS,
                                                        input_format=pyaudio.paFloat32):
                                input_devs.append({"name": device_name, "index": i})
                                seen_input_names.add(device_name)
                except ValueError:
                    logging.warning(f"Device check failed for {dev.get('name')}. It may not be a standard audio device. Skipping.")
                    continue
            return output, input_devs
        
    def _find_virtual_mic(self):
        for dev in self.output_devices:
            if VIRTUAL_MIC_NAME_PARTIAL.lower() in dev['name'].lower():
                logging.info(f"Virtual Mic found: {dev['name']} at index {dev['index']}")
                return dev['index']
        logging.warning("VB-CABLE input not found."); return None
    def get_device_name_by_index(self, index):
        if index is None: return "None"
        try: return self.p.get_device_info_by_index(index)['name']
        except (OSError, IndexError): return "Invalid Device"

    def _stream_callback(self, _, frame_count, __, ___):
        mixed_audio, playing_names = self.mixer.mix_audio(frame_count)
        is_mic_on = self.mic_inclusion_event.is_set() 

        with current_playing_sound_details_lock:
            current_playing_sound_details["names"] = playing_names
            current_playing_sound_details["active"] = bool(playing_names) or is_mic_on

        if self.app.soundboard_monitor_enabled_var.get():
            with self._soundboard_monitor_buffer_lock:
                self._soundboard_monitor_buffer.append(mixed_audio.copy())

        if is_mic_on:
            mixed_audio += self._get_mic_data_from_buffer(frame_count)
            np.clip(mixed_audio, -1.0, 1.0, out=mixed_audio)

        mixed_audio *= self.master_volume
        return (mixed_audio.astype(np.float32).tobytes(), pyaudio.paContinue)

    def _soundboard_monitor_callback(self, _, frame_count, __, ___):
        data = np.zeros((frame_count, CHANNELS), dtype=np.float32)
        with self._soundboard_monitor_buffer_lock:
            if self._soundboard_monitor_buffer: data = self._soundboard_monitor_buffer.popleft()
        
        return ((data * self.sb_monitor_volume).astype(np.float32).tobytes(), pyaudio.paContinue)
    def _mic_monitor_callback(self, _, frame_count, __, ___):
        data = self._get_mic_data_from_buffer(frame_count)
        return ((data * self.mic_monitor_volume).astype(np.float32).tobytes(), pyaudio.paContinue)

    def _start_stream(self, stream_attr, device_id, is_input, callback):
        self._stop_stream(stream_attr)
        if device_id is None: return
        try:
            stream = self.p.open(format=pyaudio.paFloat32, channels=CHANNELS, rate=SAMPLE_RATE, output=not is_input, input=is_input, frames_per_buffer=FRAME_SIZE, output_device_index=None if is_input else device_id, input_device_index=device_id if is_input else None, stream_callback=callback)
            setattr(self, stream_attr, stream)
            logging.info(f"{stream_attr} started on device index {device_id}")
        except Exception as e:
            logging.error(f"Failed to start {stream_attr} on device {device_id}: {e}")
            error_msg = f"Failed to start audio on '{self.get_device_name_by_index(device_id)}'."
            if isinstance(e, OSError):
                if e.errno == -9997: error_msg += "\nError: Invalid sample rate. Please ensure the device supports 44100 Hz."
                elif e.errno == -9999: error_msg += "\nError: Device may be in use by another application or disconnected."
                else: error_msg += f"\nOS Error: {e.strerror}"
            else: error_msg += f"\nDetails: {e}"
            messagebox.showerror("Audio Error", error_msg, parent=self.app)
            self.app.show_status_message(f"Failed to start audio on device {device_id}", "danger")
    
    def _stop_stream(self, stream_attr):
        stream = getattr(self, stream_attr)
        if stream:
            try:
                if stream.is_active(): stream.stop_stream()
                stream.close()
            except OSError as e:
                logging.warning(f"OSError stopping stream {stream_attr}: {e}")
            finally:
                setattr(self, stream_attr, None)
                logging.info(f"{stream_attr} stopped.")

    def start_main_stream(self): self._start_stream('main_stream', self.app.app_settings.get_setting("output_device_id", self.virtual_mic_device_id), False, self._stream_callback)
    def stop_main_stream(self): self._stop_stream('main_stream')
    def start_soundboard_monitor_stream(self): self._start_stream('soundboard_monitor_stream', self.app.app_settings.get_setting("soundboard_monitor_device_id"), False, self._soundboard_monitor_callback)
    def stop_soundboard_monitor_stream(self): self._stop_stream('soundboard_monitor_stream')
    def start_mic_monitor_stream(self): self._start_stream('mic_monitor_stream', self.app.app_settings.get_setting("mic_monitor_device_id"), False, self._mic_monitor_callback)
    def stop_mic_monitor_stream(self): self._stop_stream('mic_monitor_stream')

    def _mic_reader_thread_func(self):
        while not self._mic_reader_stop_event.is_set():
            try:
                if self.mic_stream and self.mic_stream.is_active():
                    raw_data = self.mic_stream.read(FRAME_SIZE, exception_on_overflow=False)
                    with self._mic_buffer_lock:
                        self._mic_buffer.append(np.frombuffer(raw_data, dtype=np.float32).reshape(-1, CHANNELS))
                else: time.sleep(0.01)
            except Exception as e: logging.error(f"Error in mic reader thread: {e}"); break
        logging.info("Mic reader thread stopped.")
    def start_mic_input(self):
        self.stop_mic_input()
        device_id = self.app.app_settings.get_setting("input_device_id")
        if device_id is None: return
        try:
            self.mic_stream = self.p.open(format=pyaudio.paFloat32, channels=CHANNELS, rate=SAMPLE_RATE, input=True, frames_per_buffer=FRAME_SIZE, input_device_index=device_id)
            self._mic_reader_stop_event.clear()
            self._mic_reader_thread = threading.Thread(target=self._mic_reader_thread_func, daemon=True)
            self._mic_reader_thread.start()
        except Exception as e: logging.error(f"Failed to start mic input: {e}")
    def stop_mic_input(self):
        self._mic_reader_stop_event.set()
        if self._mic_reader_thread: self._mic_reader_thread.join(timeout=0.5)
        self._stop_stream('mic_stream')
        with self._mic_buffer_lock: self._mic_buffer.clear()
    def _get_mic_data_from_buffer(self, frame_count):
        with self._mic_buffer_lock:
            if self._mic_buffer: return self._mic_buffer.popleft()
        return np.zeros((frame_count, CHANNELS), dtype=np.float32)
    def close(self):
        self.stop_mic_input(); self.stop_main_stream(); self.stop_soundboard_monitor_stream(); self.stop_mic_monitor_stream()
        self.p.terminate(); logging.info("PyAudio terminated.")

class KeybindManager:
    def __init__(self, app):
        self.app, self.hotkey_registry, self.active_keys = app, {}, set()
        self.listener, self.mouse_listener = None, None
    def update_hotkeys(self):
        self.stop(); self.hotkey_registry.clear()
        for sound in self.app.sound_manager.sounds:
            if sound.get("hotkeys") and sound.get("enabled", True):
                self.hotkey_registry[tuple(sorted(sound["hotkeys"]))] = lambda s_id=sound["id"]: self.app.play_sound(s_id)
        global_hotkeys = self.app.sound_manager.global_hotkeys
        if global_hotkeys.get("stop_all"): self.hotkey_registry[tuple(sorted(global_hotkeys["stop_all"]))] = self.app.stop_all_sounds
        if global_hotkeys.get("toggle_mic_to_mixer"): self.hotkey_registry[tuple(sorted(global_hotkeys["toggle_mic_to_mixer"]))] = self.app.toggle_mic_to_mixer_from_hotkey
        if self.hotkey_registry: self.start(); logging.info(f"KeybindManager started with {len(self.hotkey_registry)} hotkeys.")
    def _on_press(self, key):
        key_str = get_pynput_key_string(key)
        if key_str: self.active_keys.add(key_str); self.check_hotkeys()
    def _on_release(self, key):
        key_str = get_pynput_key_string(key)
        if key_str in self.active_keys: self.active_keys.remove(key_str)
    def _on_click(self, _, __, button, pressed):
        if pressed and button not in [mouse.Button.left, mouse.Button.right]:
            key_str = get_pynput_key_string(button)
            if key_str:
                hotkey_tuple = tuple(sorted(self.active_keys | {key_str}))
                if hotkey_tuple in self.hotkey_registry: self.hotkey_registry[hotkey_tuple]()
    def check_hotkeys(self):
        current_keys_tuple = tuple(sorted(self.active_keys))
        if current_keys_tuple in self.hotkey_registry:
            self.hotkey_registry[current_keys_tuple]()
    def start(self):
        if not (self.listener and self.listener.is_alive()):
            self.listener = keyboard.Listener(on_press=self._on_press, on_release=self._on_release)
            self.listener.start()
        if not (self.mouse_listener and self.mouse_listener.is_alive()):
            self.mouse_listener = mouse.Listener(on_click=self._on_click)
            self.mouse_listener.start()
    def stop(self):
        if self.listener: self.listener.stop()
        if self.mouse_listener: self.mouse_listener.stop()
        self.active_keys.clear()


# --- Main Application ---
class SoundboardApp(ttk.Window):
    """The main application class for the soundboard."""
    def __init__(self):
        self.app_settings = AppSettingsManager()
        self._set_initial_devices_if_needed()
        themename = self.app_settings.get_setting("theme", "vapor")
        super().__init__(title="WarpBoard", themename=themename, minsize=(700, 500))
        self.geometry("950x700")
        center_window(self)

        # --- THIS IS THE SECOND CORRECTED BLOCK ---
        # --- Set custom taskbar/window icon ---
        # The corrected ROOT_DIR variable now handles finding the icon correctly
        # for both development and installed (.exe) versions.
        icon_path = os.path.join(ROOT_DIR, "icon.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception as e:
                logging.warning(f"Could not set window icon: {e}")
        else:
            logging.warning(f"Icon file not found: {icon_path}")

        # --- Existing initialization ---
        self.sound_manager = SoundManager()
        self.audio_manager = AudioOutputManager(self)
        self.keybind_manager = KeybindManager(self)
        
        self.sound_card_widgets, self.ordered_sound_ids, self.selected_sound_ids, self.last_selected_id = {}, [], set(), None
        self.unfiltered_output_devices, self.unfiltered_input_devices, self.filtered_input_devices, self.filtered_monitor_devices = [], [], [], []
        self.grid_columns = 4

        self._init_tk_vars()
        self._load_settings()
        self._create_styles()
        self._create_ui()
        
        self.populate_device_dropdowns()
        self._apply_settings_to_ui()
        self.populate_sound_list()

        self.after(100, self._first_run_check) 
        
        self.keybind_manager.update_hotkeys()
        self.audio_manager.start_main_stream()

        self.update_now_playing_status()
        self.protocol("WM_DELETE_WINDOW", self._on_app_closure)
        self.update_idletasks()
        self._on_frame_configure()

    def center_toplevel(self, toplevel):
        toplevel.update_idletasks()
        
        main_x = self.winfo_x()
        main_y = self.winfo_y()
        main_width = self.winfo_width()
        main_height = self.winfo_height()

        top_width = toplevel.winfo_width()
        top_height = toplevel.winfo_height()

        x = main_x + (main_width - top_width) // 2
        y = main_y + (main_height - top_height) // 2
        
        toplevel.geometry(f"+{x}+{y}")

    def _init_tk_vars(self):
        self.include_mic_in_mix_var, self.soundboard_monitor_enabled_var, self.mic_monitor_enabled_var = tk.BooleanVar(), tk.BooleanVar(), tk.BooleanVar()
        self.master_volume_var, self.soundboard_monitor_volume_var, self.mic_monitor_volume_var = tk.DoubleVar(), tk.DoubleVar(), tk.DoubleVar()
        self.current_theme_var, self.single_sound_mode_var, self.auto_start_mic_var = tk.StringVar(), tk.BooleanVar(), tk.BooleanVar()
        self.stop_all_hotkey_var, self.toggle_mic_hotkey_var = tk.StringVar(value="Not Assigned"), tk.StringVar(value="Not Assigned")
        self.search_var = tk.StringVar()

    def _load_settings(self):
        self.current_theme_var.set(self.style.theme.name)
        settings_defaults = {"auto_start_mic": False, "soundboard_monitor_enabled": True, "mic_monitor_enabled": False, "master_volume": 100.0, "soundboard_monitor_volume": 75.0, "mic_monitor_volume": 75.0, "single_sound_mode": True}
        for key, default in settings_defaults.items():
            if hasattr(self, f"{key}_var"): getattr(self, f"{key}_var").set(self.app_settings.get_setting(key, default))
        self.include_mic_in_mix_var.set(self.auto_start_mic_var.get())
        self.stop_all_hotkey_var.set(get_hotkey_display_string(self.sound_manager.global_hotkeys.get("stop_all", [])))
        self.toggle_mic_hotkey_var.set(get_hotkey_display_string(self.sound_manager.global_hotkeys.get("toggle_mic_to_mixer", [])))
    
    def _create_styles(self):
        border_color, selected_bg = self.style.colors.primary, self.style.colors.selectbg
        self.style.configure('Card.TFrame', borderwidth=2, relief='solid', bordercolor=self.style.colors.bg)
        self.style.map('Card.TFrame', bordercolor=[('selected', border_color)], background=[('selected', selected_bg)])
        self.style.configure('Placeholder.TEntry', foreground=self.style.colors.secondary)
        self.style.configure('TButton', wraplength=150, justify='center')

    def _create_ui(self):
        notebook = ttk.Notebook(self, padding=(10, 10, 10, 0))
        notebook.pack(fill=BOTH, expand=True)
        soundboard_tab, settings_tab = ttk.Frame(notebook), ttk.Frame(notebook)
        notebook.add(soundboard_tab, text="Soundboard"); notebook.add(settings_tab, text="Settings")
        self._create_soundboard_widgets(soundboard_tab)
        self._create_settings_widgets(settings_tab)
        status_frame = ttk.Frame(self, padding=(10, 5)); status_frame.pack(fill=X, side=BOTTOM)
        self.status_message_var = tk.StringVar(value="Ready.")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_message_var)
        self.status_label.pack(side=LEFT)
        self.now_playing_var = tk.StringVar(value="Now Playing: None")
        ttk.Label(status_frame, textvariable=self.now_playing_var, anchor=E).pack(side=RIGHT)

    def _create_soundboard_widgets(self, parent):
        parent.rowconfigure(2, weight=1); parent.columnconfigure(0, weight=1)
        
        search_frame = ttk.Frame(parent)
        search_frame.grid(row=0, column=0, sticky=EW, pady=(10, 10), padx=10)
        search_frame.columnconfigure(0, weight=1)
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, font=("-size", 12))
        self.search_entry.grid(row=0, column=0, sticky=EW, ipady=4)
        self.clear_search_btn = ttk.Button(search_frame, text="✖", command=lambda: self.search_var.set(""), bootstyle="light", width=3)
        self._add_placeholder(self.search_entry, "Search sounds...")

        button_frame = ttk.Frame(parent)
        button_frame.grid(row=1, column=0, sticky=EW, pady=(0, 10), padx=10)
        ttk.Button(button_frame, text="✚ Add Sound", command=self.add_sound, bootstyle="primary").pack(side=LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="✖ Remove Selected", command=self.remove_selected_sounds, bootstyle="danger").pack(side=LEFT)
        self.stop_all_button = ttk.Button(button_frame, text="Stop All Sounds", command=self.stop_all_sounds, bootstyle="danger-outline")
        self.stop_all_button.pack(side=RIGHT)

        list_container = ttk.Frame(parent)
        list_container.grid(row=2, column=0, sticky='nsew', padx=10)
        self.sound_canvas = tk.Canvas(list_container, highlightthickness=0, background=self.style.colors.bg)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self.sound_canvas.yview)
        self.sound_list_frame = ttk.Frame(self.sound_canvas, padding=5)
        self.sound_canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y); self.sound_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        self.canvas_window = self.sound_canvas.create_window((0, 0), window=self.sound_list_frame, anchor="nw")
        
        self.sound_list_frame.bind("<Configure>", lambda _: self.sound_canvas.configure(scrollregion=self.sound_canvas.bbox("all")))
        self.sound_canvas.bind("<Configure>", self._on_frame_configure)
        self.sound_canvas.bind("<Enter>", self._bind_mousewheel)
        self.sound_canvas.bind("<Leave>", self._unbind_mousewheel)
        
        self.search_var.trace_add("write", lambda *_: self._filter_sounds())
    
    def _create_settings_widgets(self, parent):
        notebook = ttk.Notebook(parent, padding=(0, 10, 0, 0))
        notebook.pack(fill=BOTH, expand=True)
        audio_tab, hotkey_tab, general_tab, audio_setup_tab, about_tab = ttk.Frame(notebook), ttk.Frame(notebook), ttk.Frame(notebook), ttk.Frame(notebook), ttk.Frame(notebook)
        notebook.add(audio_tab, text="Audio"); notebook.add(hotkey_tab, text="Hotkeys"); notebook.add(general_tab, text="General"); notebook.add(audio_setup_tab, text="Audio Setup"); notebook.add(about_tab, text="About")
        self._populate_audio_tab(audio_tab); self._populate_hotkey_tab(hotkey_tab); self._populate_general_tab(general_tab); self._populate_audio_setup_tab(audio_setup_tab); self._populate_about_tab(about_tab)
    
    def _populate_audio_tab(self, parent):
        device_frame = ttk.Labelframe(parent, text="Audio Devices", padding=10)
        device_frame.pack(fill=X, pady=10, padx=10)
        self.output_device_combo = self._create_device_combo(device_frame, "App Output", 0, self._on_output_device_selected, "The virtual audio device that other apps (like Discord, OBS) will listen to. Select your VB-CABLE here.")
        app_output_note = ttk.Label(device_frame, text='(Shows as “CABLE Input (VB-Audio Virtual Cable)” in other apps)', bootstyle="secondary", font="-size 8")
        app_output_note.grid(row=0, column=2, sticky=W, padx=5)
        self.input_device_combo = self._create_device_combo(device_frame, "Your Microphone", 1, self._on_input_device_selected, "Select your physical microphone here.")
        self.sb_monitor_combo = self._create_device_combo(device_frame, "Listen to Soundboard", 2, self._on_sb_monitor_device_selected, "To hear the soundboard through your headphones/speakers, select them here and check 'Enable' below.")
        self.mic_monitor_combo = self._create_device_combo(device_frame, "Hear Your Voice", 3, self._on_mic_monitor_device_selected, "To hear your own microphone (sidetone), select your headphones/speakers here and check 'Enable' below.")
        
        vol_frame = ttk.Labelframe(parent, text="Volume & Monitoring", padding=10)
        vol_frame.pack(fill=X, pady=10, padx=10)
        self._create_volume_slider(vol_frame, "Master Output", 0, self.master_volume_var, self.audio_manager.set_master_volume)
        self._create_volume_slider(vol_frame, "Soundboard Monitor", 1, self.soundboard_monitor_volume_var, self.audio_manager.set_sb_monitor_volume)
        self._create_volume_slider(vol_frame, "Voice Monitor", 2, self.mic_monitor_volume_var, self.audio_manager.set_mic_monitor_volume)
        ttk.Separator(vol_frame).grid(row=3, column=0, columnspan=3, sticky=EW, pady=10)
        sb_monitor_check = ttk.Checkbutton(vol_frame, text="Enable 'Listen to Soundboard'", variable=self.soundboard_monitor_enabled_var, command=self.toggle_soundboard_monitor, bootstyle="round-toggle")
        sb_monitor_check.grid(row=4, column=0, columnspan=3, sticky=W)
        ToolTip(sb_monitor_check, lambda: "Hear the soundboard audio through your selected 'Listen to Soundboard' device.")

        mic_monitor_check = ttk.Checkbutton(vol_frame, text="Enable 'Hear Your Voice'", variable=self.mic_monitor_enabled_var, command=self.toggle_mic_monitor, bootstyle="round-toggle")
        mic_monitor_check.grid(row=5, column=0, columnspan=3, sticky=W, pady=(5,0))
        ToolTip(mic_monitor_check, lambda: "Hear your own microphone (sidetone) through your selected 'Hear Your Voice' device.")

        include_mic_check = ttk.Checkbutton(vol_frame, text="Include Mic in App Output", variable=self.include_mic_in_mix_var, command=self.toggle_include_mic_in_mix, bootstyle="round-toggle")
        include_mic_check.grid(row=6, column=0, columnspan=3, sticky=W, pady=(5,0))
        ToolTip(include_mic_check, lambda: "Mix your actual microphone into the main 'App Output' so others can hear you and the sounds.")

    def _create_device_combo(self, parent, label, row, command, tooltip_text):
        label_widget = ttk.Label(parent, text=label)
        label_widget.grid(row=row, column=0, sticky=W, padx=5, pady=2)
        ToolTip(label_widget, lambda: tooltip_text)
        combo = ttk.Combobox(parent, state="readonly", width=35)
        combo.grid(row=row, column=1, sticky=EW, padx=5)
        combo.bind("<<ComboboxSelected>>", command)
        parent.columnconfigure(1, weight=1); return combo

    def _create_volume_slider(self, parent, label, row, var, setter_func):
        if label: ttk.Label(parent, text=label).grid(row=row, column=0, sticky=W, padx=5, pady=2)
        scale = ttk.Scale(parent, from_=0, to=125, variable=var)
        scale.grid(row=row, column=1, sticky=EW, padx=5)
        label_widget = ttk.Label(parent, text=f"{int(var.get())}%", width=4)
        label_widget.grid(row=row, column=2, padx=5)
        
        def update_volume(*_):
            volume_percent = var.get()
            label_widget.config(text=f"{int(volume_percent)}%")
            if setter_func:
                setter_func(volume_percent / 100.0)

        var.trace_add("write", update_volume)
        scale.bind("<ButtonRelease-1>", lambda _: self._save_app_settings())
        parent.columnconfigure(1, weight=1)
    
    def _populate_hotkey_tab(self, parent):
        frame = ttk.Labelframe(parent, text="Global Hotkeys", padding=10)
        frame.pack(fill=X, padx=10, pady=10)
        self._create_hotkey_entry(frame, "Stop All Sounds", 0, self.stop_all_hotkey_var, "stop_all")
        self._create_hotkey_entry(frame, "Toggle Mic in Mix", 1, self.toggle_mic_hotkey_var, "toggle_mic_to_mixer")

    def _create_hotkey_entry(self, parent, label_text, row, var, action):
        ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky=W, padx=5, pady=5)
        ttk.Label(parent, textvariable=var, bootstyle="info").grid(row=row, column=1, sticky=EW, padx=5)
        ttk.Button(parent, text="Assign", command=lambda: self._assign_global_hotkey(action, var)).grid(row=row, column=2, padx=5)
        ttk.Button(parent, text="Clear", command=lambda: self._clear_global_hotkey(action, var)).grid(row=row, column=3, padx=5)
        parent.columnconfigure(1, weight=1)
        
    def _populate_general_tab(self, parent):
        frame = ttk.Labelframe(parent, text="Application Settings", padding=10)
        frame.pack(fill=X, padx=10, pady=10)
        ttk.Label(frame, text="Theme:").grid(row=0, column=0, sticky=W, padx=5, pady=5)
        theme_combo = ttk.Combobox(frame, textvariable=self.current_theme_var, values=self.style.theme_names(), state="readonly")
        theme_combo.grid(row=0, column=1, sticky=EW, padx=5)
        theme_combo.bind("<<ComboboxSelected>>", self._on_theme_changed)
        single_sound_check = ttk.Checkbutton(frame, text="Single Sound Mode (stop others on play)", variable=self.single_sound_mode_var, command=self._on_single_sound_mode_changed, bootstyle="round-toggle")
        single_sound_check.grid(row=1, column=0, columnspan=2, sticky=W, pady=5)
        ToolTip(single_sound_check, lambda: "When enabled, playing a new sound will automatically stop any sound that is already playing.")

        auto_start_mic_check = ttk.Checkbutton(frame, text="Automatically include Mic on startup", variable=self.auto_start_mic_var, command=lambda: self._save_app_settings(), bootstyle="round-toggle")
        auto_start_mic_check.grid(row=2, column=0, columnspan=2, sticky=W)
        ToolTip(auto_start_mic_check, lambda: "If checked, your microphone will automatically be included in the 'App Output' every time you start WarpBoard.")
        frame.columnconfigure(1, weight=1)

    def _populate_audio_setup_tab(self, parent):
        setup_frame = ttk.Labelframe(parent, text="VB-CABLE Virtual Mic Setup", padding=15)
        setup_frame.pack(fill=X, padx=10, pady=10)

        instructions = (
            "WarpBoard uses a free 'virtual audio cable' to send audio to other applications (like Discord, OBS, or games). "
            "If it's not installed, the app can install it for you on first launch.\n\n"
            "If you have issues, you can use the tools below to manage the audio driver."
        )
        label = ttk.Label(setup_frame, text=instructions, justify=LEFT)
        label.pack(anchor=W, pady=5, fill=X)
        label.winfo_toplevel().update_idletasks()
        label.configure(wraplength=setup_frame.winfo_width() - 30)

        ttk.Button(setup_frame, text="Download VB-CABLE manually", command=lambda: webbrowser.open(VB_CABLE_URL)).pack(pady=10)

        actions_frame = ttk.Labelframe(parent, text="Maintenance Actions", padding=15)
        actions_frame.pack(fill=X, padx=10, pady=10)
        
        ttk.Button(actions_frame, text="Fix Audio Setup", command=self._fix_audio_setup).pack(side=LEFT, padx=5)
        ttk.Button(actions_frame, text="Fix Duplicate Devices", command=self._fix_duplicate_devices).pack(side=LEFT, padx=5)
        
        reinstall_btn = ttk.Button(actions_frame, text="⚠️ Force Reinstall VB-Cable", command=self._force_reinstall_vb_cable, bootstyle="danger")
        reinstall_btn.pack(side=LEFT, padx=5)

        log_frame = ttk.Labelframe(parent, text="Troubleshooting", padding=15)
        log_frame.pack(fill=X, padx=10, pady=10)
        ttk.Label(log_frame, text="If you encounter issues, the log file can help identify the problem.").pack(anchor=W, pady=5)
        ttk.Button(log_frame, text="Open Log File", command=self._open_log_file, bootstyle="info-outline").pack(pady=10)

    def _first_run_check(self, force_install=False):
        is_installed = any(VIRTUAL_MIC_NAME_PARTIAL.lower() in dev['name'].lower() for dev in self.audio_manager.output_devices)
        
        if force_install or (not is_installed and not self.app_settings.get_setting("vb_cable_installed", False)):
            if not messagebox.askyesno("VB-CABLE Driver Required", 
                                     "WarpBoard requires the free VB-CABLE virtual audio driver to function. May we install it for you now? \n\n(Your default audio devices will be temporarily changed and then restored.)", 
                                     parent=self):
                return

            popup = Toplevel(self)
            popup.title("Installing...")
            ttk.Label(popup, text="Installing WarpBoard Virtual Mic…\nThis may take a few seconds.\nYour regular headphones and mic will remain unchanged.", padding=20).pack()
            self.center_toplevel(popup)
            self.update_idletasks()

            try:
                logging.info("Getting current default audio devices.")
                defaults_json = run_powershell_script('get_default_devices.ps1')
                defaults = json.loads(defaults_json)
                logging.info(f"Saved defaults: {defaults}")

                logging.info("Running VB-CABLE installer...")
                installer_path = os.path.join(ROOT_DIR, 'VBCABLE_Setup_x64.exe')
                if not os.path.exists(installer_path):
                    raise FileNotFoundError("VB-CABLE installer not found in application directory.")
                
                subprocess.run(f'start /wait "" "{installer_path}" /S', shell=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
                logging.info("VB-CABLE installer finished.")
                
                time.sleep(5)

                logging.info("Restoring default audio devices.")
                run_powershell_script('set_default_devices.ps1',
                                      f"-DefaultPlayback \"{defaults['DefaultPlayback']}\"",
                                      f"-DefaultCommunicationsPlayback \"{defaults['DefaultCommunicationsPlayback']}\"",
                                      f"-DefaultRecording \"{defaults['DefaultRecording']}\"",
                                      f"-DefaultCommunicationsRecording \"{defaults['DefaultCommunicationsRecording']}\"")
                logging.info("Default devices restored.")

                self.app_settings.save_settings({"vb_cable_installed": True})
                messagebox.showinfo("Installation Complete", "VB-CABLE has been installed successfully! Please restart WarpBoard for the changes to take full effect.", parent=self)

            except Exception as e:
                logging.error(f"VB-CABLE automatic installation failed: {e}")
                messagebox.showerror("Installation Failed", f"An error occurred during the automatic installation of VB-CABLE. You may need to install it manually.\n\nError: {e}", parent=self)
            finally:
                popup.destroy()
                self.audio_manager.output_devices, self.audio_manager.input_devices = self.audio_manager._enumerate_devices()
                self.populate_device_dropdowns()

    def _fix_audio_setup(self):
        try:
            defaults = self.app_settings.get_setting("original_system_defaults")
            if not defaults:
                messagebox.showwarning("Not Found", "Original default device settings not found. Cannot restore.", parent=self)
                return

            messagebox.showinfo("Fixing Audio", "Attempting to restore your original default audio devices...", parent=self)
            logging.info("Running 'Fix Audio Setup' to restore original devices.")
            run_powershell_script('set_default_devices.ps1',
                                  f"-DefaultPlayback \"{defaults['DefaultPlaybackId']}\"",
                                  f"-DefaultCommunicationsPlayback \"{defaults['DefaultCommunicationsPlaybackId']}\"",
                                  f"-DefaultRecording \"{defaults['DefaultRecordingId']}\"",
                                  f"-DefaultCommunicationsRecording \"{defaults['DefaultCommunicationsRecordingId']}\"")
            messagebox.showinfo("Complete", "Default audio devices have been restored.", parent=self)
        except Exception as e:
            logging.error(f"Failed to run 'Fix Audio Setup' script: {e}")
            messagebox.showerror("Error", f"Failed to run fix script: {e}", parent=self)

    def _force_reinstall_vb_cable(self):
        if messagebox.askyesno("Confirm Re-installation", 
                               "This will attempt to force a re-installation of the VB-CABLE driver. This can be useful if the driver is corrupted. Are you sure you want to continue?", 
                               parent=self):
            self._first_run_check(force_install=True)

    def _fix_duplicate_devices(self):
        try:
            messagebox.showinfo("Fixing Duplicates", "Searching for and disabling duplicate audio devices. This may take a moment.", parent=self)
            logging.info("Running duplicate device cleanup script...")
            output = run_powershell_script('remove_duplicate_devices.ps1')
            logging.info(f"Cleanup script output: {output}")
            messagebox.showinfo("Complete", "Duplicate device cleanup finished. It's recommended to restart the application.", parent=self)
            self.audio_manager.output_devices, self.audio_manager.input_devices = self.audio_manager._enumerate_devices()
            self.populate_device_dropdowns()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to run cleanup script: {e}", parent=self)

    def _set_initial_devices_if_needed(self):
        if self.app_settings.get_setting("input_device_id") is None:
            logging.info("First run or settings reset: detecting system default audio devices.")
            try:
                defaults_json = run_powershell_script('get_default_devices.ps1')
                defaults = json.loads(defaults_json)
                
                default_mic_id = defaults.get("DefaultCommunicationsRecordingId") or defaults.get("DefaultRecordingId")
                default_speaker_id = defaults.get("DefaultCommunicationsPlaybackId") or defaults.get("DefaultPlaybackId")

                if default_mic_id:
                    self.app_settings.save_settings({"input_device_id": default_mic_id})
                if default_speaker_id:
                    self.app_settings.save_settings({
                        "soundboard_monitor_device_id": default_speaker_id,
                        "mic_monitor_device_id": default_speaker_id
                    })
                
                self.app_settings.save_settings({"original_system_defaults": defaults})
                logging.info(f"Saved system defaults: {defaults}")
            except Exception as e:
                logging.error(f"Failed to get system default devices on first run: {e}")

    def _open_log_file(self):
        try:
            if not os.path.exists(LOG_FILE):
                messagebox.showinfo("Info", "Log file has not been created yet.", parent=self)
                return
            if platform.system() == "Windows":
                os.startfile(LOG_FILE)
            elif platform.system() == "Darwin":
                subprocess.run(["open", LOG_FILE], check=True)
            else:
                subprocess.run(["xdg-open", LOG_FILE], check=True)
        except Exception as e:
            logging.error(f"Failed to open log file: {e}")
            messagebox.showerror("Error", f"Could not open log file.\nLocation: {LOG_FILE}", parent=self)
        
    def _populate_about_tab(self, parent):
        frame = ttk.Labelframe(parent, text="About WarpBoard", padding=15)
        frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

        ttk.Label(frame, text="WarpBoard", font="-size 16 -weight bold").pack(pady=5)
        ttk.Label(frame, text=f"Version: {APP_VERSION}").pack()
        
        about_text = ("WarpBoard is a free, open-source soundboard built for the gaming and streaming community. "
                      "It's designed to be simple to use, lightweight on your system, and powerful enough to bring your audio clips to life. "
                      "Whether you're creating content, spicing up your voice chats, or just having fun, WarpBoard is here to help you make some noise.")
        about_label = ttk.Label(frame, text=about_text, justify=LEFT, wraplength=400)
        about_label.pack(pady=(10, 20))

        default_font = font.nametofont("TkDefaultFont")
        link_font = default_font.copy()
        link_font.configure(underline=True)

        profile_label = ttk.Label(
            frame,
            text="Created by: Asta main",
            foreground=self.style.colors.primary,
            font=link_font,
            cursor="hand2"
        )
        profile_label.pack(pady=(10,0))
        profile_label.bind("<Button-1>", lambda _: webbrowser.open(PROFILE_URL))
        ToolTip(profile_label, lambda: f"Open profile: {PROFILE_URL}")

        docs_label = ttk.Label(
            frame,
            text="Documentation",
            foreground=self.style.colors.primary,
            font=link_font,
            cursor="hand2"
        )
        docs_label.pack()
        docs_label.bind("<Button-1>", lambda _: webbrowser.open(DOCS_URL))
        ToolTip(docs_label, lambda: f"Open documentation: {DOCS_URL}")

        github_label = ttk.Label(
            frame,
            text="Report an issue",
            foreground=self.style.colors.primary,
            font=link_font,
            cursor="hand2"
        )
        github_label.pack()
        github_label.bind("<Button-1>", lambda _: webbrowser.open(ISSUES_URL))
        ToolTip(github_label, lambda: f"Report bugs at: {ISSUES_URL}")

        ttk.Label(frame, text="License: MIT", foreground="gray").pack(pady=(10, 5))
    
    def _on_frame_configure(self, _=None):
        canvas_width = self.sound_canvas.winfo_width()
        if canvas_width > 1:
            self.sound_canvas.itemconfig(self.canvas_window, width=canvas_width)
            new_cols = max(1, canvas_width // 200)
            if hasattr(self, 'grid_columns') and new_cols != self.grid_columns:
                self.grid_columns = new_cols
                self._filter_sounds()
    
    def _bind_mousewheel(self, _): self.bind_all("<MouseWheel>", self._on_mousewheel)
    def _unbind_mousewheel(self, _): self.unbind_all("<MouseWheel>")
    def _on_mousewheel(self, event):
        try:
            widget = self.winfo_containing(event.x_root, event.y_root)
            if widget and str(widget).startswith(str(self.sound_canvas)):
                if event.delta > 0: self.sound_canvas.yview_scroll(-1, "units")
                else: self.sound_canvas.yview_scroll(1, "units")
        except KeyError:
            pass

    def _add_placeholder(self, entry, placeholder):
        entry.insert(0, placeholder)
        entry.configure(style='Placeholder.TEntry')
        def on_focus_in(_):
            if entry.get() == placeholder:
                entry.delete(0, "end")
                entry.configure(style='TEntry')
        def on_focus_out(_):
            if not entry.get():
                entry.insert(0, placeholder)
                entry.configure(style='Placeholder.TEntry')
        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)

    def _filter_sounds(self, *_):
        query = self.search_var.get().lower()
        if query == "search sounds...": query = ""

        if query:
            self.clear_search_btn.grid(row=0, column=1, sticky=E, padx=(4,0))
        else:
            self.clear_search_btn.grid_forget()

        for i in range(self.sound_list_frame.grid_size()[0]):
            self.sound_list_frame.columnconfigure(i, weight=0)
        for i in range(self.grid_columns):
            self.sound_list_frame.columnconfigure(i, weight=1)

        row, col = 0, 0
        for sound_id in self.ordered_sound_ids:
            sound = self.sound_manager.get_sound_by_id(sound_id)
            if sound:
                card_frame = self.sound_card_widgets[sound_id]['frame']
                if query in sound['name'].lower():
                    card_frame.grid(row=row, column=col, padx=5, pady=5, sticky='nsew')
                    col += 1
                    if col >= self.grid_columns:
                        col = 0
                        row += 1
                else:
                    card_frame.grid_forget()
        
        self.sound_canvas.yview_moveto(0)
        self.sound_canvas.update_idletasks()
        self.sound_canvas.configure(scrollregion=self.sound_canvas.bbox("all"))

    def populate_sound_list(self):
        for widget in self.sound_list_frame.winfo_children(): widget.destroy()
        self.sound_card_widgets.clear()
        self.ordered_sound_ids = sorted([s for s in self.sound_manager.sounds], key=lambda x: x['name'].lower())
        self.ordered_sound_ids = [s['id'] for s in self.ordered_sound_ids]
        
        for sound_id in self.ordered_sound_ids:
            sound = self.sound_manager.get_sound_by_id(sound_id)
            if sound: self._add_sound_card_to_ui(sound)
        self._filter_sounds()

    def _add_sound_card_to_ui(self, sound):
        sound_id = sound["id"]
        card_frame = ttk.Frame(self.sound_list_frame, style='Card.TFrame', padding=10)
        card_frame.columnconfigure(0, weight=1)

        top_button_frame = ttk.Frame(card_frame)
        top_button_frame.grid(row=0, column=0, sticky=EW, pady=5)
        top_button_frame.columnconfigure(0, weight=1)
        
        play_btn = ttk.Button(top_button_frame, text=sound['name'][:MAX_DISPLAY_NAME_LENGTH], command=lambda: self.play_sound(sound_id), bootstyle="success")
        play_btn.grid(row=0, column=0, sticky=EW, ipady=5)
        
        stop_btn = ttk.Button(top_button_frame, text="■", command=lambda: self.stop_sound(sound_id), bootstyle="danger", width=3)
        stop_btn.grid(row=0, column=1, sticky="ns", padx=(5,0))


        hotkey_var = tk.StringVar(value=get_hotkey_display_string(sound['hotkeys']))
        hotkey_label = ttk.Label(card_frame, textvariable=hotkey_var, bootstyle="secondary", anchor="center")
        
        bottom_frame = ttk.Frame(card_frame)
        loop_var = tk.BooleanVar(value=sound.get("loop", False))
        loop_check = ttk.Checkbutton(bottom_frame, text="Loop", variable=loop_var, bootstyle="round-toggle", command=lambda s_id=sound_id, v=loop_var: self._update_sound_property_and_save(s_id, "loop", v.get()))
        edit_btn = ttk.Button(bottom_frame, text="⚙", command=lambda s_id=sound_id: self._open_edit_sound_menu(s_id), bootstyle="light-outline", width=3)

        hotkey_label.grid(row=1, column=0, sticky=EW)
        bottom_frame.grid(row=2, column=0, sticky=EW, pady=(10, 0))
        bottom_frame.columnconfigure(0, weight=1)
        bottom_frame.columnconfigure(1, weight=1)
        loop_check.grid(row=0, column=0, sticky=W)
        edit_btn.grid(row=0, column=1, sticky=E)

        if not sound.get("enabled", True):
            play_btn.configure(state="disabled")
        
        self.sound_card_widgets[sound_id] = {"frame": card_frame, "hotkey_var": hotkey_var, "loop_var": loop_var, "play_btn": play_btn}
        
        tooltip_text_func = lambda s=sound: f"Name: {s['name']}\nDuration: {s.get('duration', 0):.2f} seconds"
        ToolTip(card_frame, tooltip_text_func)

        for widget in [card_frame, hotkey_label, bottom_frame, top_button_frame]:
            widget.bind("<Button-1>", lambda e, s_id=sound_id: self._on_card_click(e, s_id))

    def _on_card_click(self, event, sound_id):
        ctrl_pressed, shift_pressed = (event.state & 4) != 0, (event.state & 1) != 0
        if shift_pressed and self.last_selected_id:
            try:
                visible_ids = []
                for sid in self.ordered_sound_ids:
                    try:
                        if self.sound_card_widgets[sid]['frame'].winfo_ismapped():
                            visible_ids.append(sid)
                    except (KeyError, tk.TclError):
                        continue
                start = visible_ids.index(self.last_selected_id)
                end = visible_ids.index(sound_id)
                if start > end: start, end = end, start
                if not ctrl_pressed: self.selected_sound_ids.clear()
                for i in range(start, end + 1): self.selected_sound_ids.add(visible_ids[i])
            except (ValueError, KeyError, tk.TclError): 
                self.selected_sound_ids = {sound_id}
        elif ctrl_pressed:
            if sound_id in self.selected_sound_ids: self.selected_sound_ids.remove(sound_id)
            else: self.selected_sound_ids.add(sound_id)
        else: self.selected_sound_ids = {sound_id}
        self.last_selected_id = sound_id
        self._update_card_styles()

    def _update_card_styles(self):
        for sound_id, widgets in self.sound_card_widgets.items():
            widgets['frame'].state(['selected'] if sound_id in self.selected_sound_ids else ['!selected'])

    def _remove_sound_card_from_ui(self, sound_id):
        if sound_id in self.sound_card_widgets:
            self.sound_card_widgets[sound_id]['frame'].destroy()
            del self.sound_card_widgets[sound_id]
            if sound_id in self.ordered_sound_ids: self.ordered_sound_ids.remove(sound_id)
    
    def _open_edit_sound_menu(self, sound_id):
        sound = self.sound_manager.get_sound_by_id(sound_id);
        if not sound: return
        
        edit_window = Toplevel(self); edit_window.title(f"Edit '{sound['name']}'"); edit_window.transient(self)
        main_frame = ttk.Frame(edit_window, padding=15); main_frame.pack(fill=BOTH, expand=True)

        enabled_frame = ttk.Labelframe(main_frame, text="Status", padding=5)
        enabled_frame.pack(fill=X, pady=5)
        enabled_var = tk.BooleanVar(value=sound.get("enabled", True))
        ttk.Checkbutton(enabled_frame, text="Sound Enabled", variable=enabled_var, bootstyle="round-toggle").pack(padx=5, pady=5, anchor=W)

        name_frame = ttk.Labelframe(main_frame, text="Sound Name", padding=5); name_frame.pack(fill=X, pady=5)
        name_var = tk.StringVar(value=sound['name']); ttk.Entry(name_frame, textvariable=name_var).pack(fill=X, padx=5, pady=5)
        
        hotkey_frame = ttk.Labelframe(main_frame, text="Hotkey", padding=5); hotkey_frame.pack(fill=X, pady=5)
        hotkey_var = tk.StringVar(value=get_hotkey_display_string(sound['hotkeys']))
        ttk.Label(hotkey_frame, textvariable=hotkey_var, bootstyle="info").grid(row=0, column=0, sticky=EW, padx=5, pady=5)
        ttk.Button(hotkey_frame, text="Assign", command=lambda: self._assign_sound_hotkey(sound_id, hotkey_var)).grid(row=0, column=1, padx=5)
        ttk.Button(hotkey_frame, text="Clear", command=lambda: self._clear_sound_hotkey(sound_id, hotkey_var)).grid(row=0, column=2, padx=5)
        hotkey_frame.columnconfigure(0, weight=1)
        
        volume_frame = ttk.Labelframe(main_frame, text="Volume", padding=5); volume_frame.pack(fill=X, pady=5)
        volume_var = tk.DoubleVar(value=sound.get("volume", 1.0) * 100)
        self._create_volume_slider(volume_frame, "", 0, volume_var, None)
        
        def _save_changes():
            try:
                self.sound_manager.update_sound_property(sound_id, "volume", volume_var.get() / 100.0)
                self.sound_manager.update_sound_property(sound_id, "enabled", enabled_var.get())
                if name_var.get() != sound['name']:
                    self.sound_manager.rename_sound(sound_id, name_var.get())
                self.sound_manager.save_config()
                self.keybind_manager.update_hotkeys()
                self.populate_sound_list() 
                edit_window.destroy()
            except ValueError as e: messagebox.showerror("Rename Error", str(e), parent=edit_window)

        ttk.Button(main_frame, text="Save and Close", command=_save_changes, bootstyle="primary").pack(pady=15)
        self.center_toplevel(edit_window)
        edit_window.grab_set()

    def _update_sound_property_and_save(self, sound_id, key, value):
        self.sound_manager.update_sound_property(sound_id, key, value)
        self.sound_manager.save_config()
    def add_sound(self):
        file_paths = filedialog.askopenfilenames(filetypes=SUPPORTED_FORMATS)
        if not file_paths: return
        for path in file_paths:
            try:
                self.sound_manager.add_sound(path)
                self.show_status_message(f"Added: {os.path.basename(path)}", "success")
            except Exception as e:
                messagebox.showerror("Add Sound Error", f"Failed to add sound from {os.path.basename(path)}.\nError: {e}", parent=self)
                self.show_status_message(f"Failed to add sound.", "danger")
        self.keybind_manager.update_hotkeys()
        self.populate_sound_list()
        
    def remove_selected_sounds(self):
        if not self.selected_sound_ids: self.show_status_message("No sounds selected.", "warning"); return
        if messagebox.askyesno("Confirm Removal", f"Are you sure you want to permanently remove {len(self.selected_sound_ids)} sound(s)?", parent=self):
            ids_to_remove = self.selected_sound_ids.copy()
            self.sound_manager.remove_sounds(ids_to_remove)
            for sound_id in ids_to_remove: self._remove_sound_card_from_ui(sound_id)
            self.selected_sound_ids.clear(); self.last_selected_id = None
            self.keybind_manager.update_hotkeys()
            self.show_status_message(f"Removed {len(ids_to_remove)} sound(s).", "success")
            self.populate_sound_list()

    def play_sound(self, sound_id):
        sound = self.sound_manager.get_sound_by_id(sound_id)
        if not sound or not sound.get("enabled", True): return

        audio_data = self.sound_manager.sound_data_cache.get(sound_id)
        if audio_data is None:
            try:
                self.sound_manager.preload_sound_data(sound)
                audio_data = self.sound_manager.sound_data_cache.get(sound_id)
            except Exception as e:
                logging.error(f"Failed to load audio on demand for '{sound['name']}': {e}")
                self.show_status_message(f"Error playing {sound['name']}", "danger")
                return

        if audio_data is not None:
            self.audio_manager.mixer.add_sound(audio_data, sound["volume"], sound["loop"], sound_id, sound["name"])

    def stop_sound(self, sound_id):
        if self.audio_manager.mixer.remove_sound_by_id(sound_id):
            self.show_status_message(f"Stopped: {self.sound_manager.get_sound_by_id(sound_id)['name']}", "info")
    def stop_all_sounds(self):
        self.audio_manager.mixer.clear_sounds()
        self.after(0, self.show_status_message, "All sounds stopped.", "success")

    def toggle_include_mic_in_mix(self):
        is_enabled = self.include_mic_in_mix_var.get()
        if is_enabled:
            self.audio_manager.mic_inclusion_event.set()
            self.audio_manager.start_mic_input()
        else:
            self.audio_manager.mic_inclusion_event.clear()
            self.audio_manager.stop_mic_input()
        self._save_app_settings()

    def toggle_soundboard_monitor(self):
        is_enabled = self.soundboard_monitor_enabled_var.get()
        if is_enabled: self.audio_manager.start_soundboard_monitor_stream()
        else: self.audio_manager.stop_soundboard_monitor_stream()
        self._save_app_settings()
        
    def toggle_mic_monitor(self):
        is_enabled = self.mic_monitor_enabled_var.get()
        if is_enabled: self.audio_manager.start_mic_monitor_stream()
        else: self.audio_manager.stop_mic_monitor_stream()
        self._save_app_settings()
        
    def toggle_mic_to_mixer_from_hotkey(self):
        self.after(0, self._toggle_mic_from_hotkey)

    def _assign_sound_hotkey(self, sound_id, hotkey_var):
        sound = self.sound_manager.get_sound_by_id(sound_id);
        if not sound: return
        def on_complete(hotkey_list):
            if hotkey_list is None: return
            if tuple(sorted(hotkey_list)) in self.sound_manager.get_all_assigned_hotkeys():
                messagebox.showwarning("Hotkey In Use", "This hotkey is already assigned.", parent=self)
                return
            self.sound_manager.update_sound_property(sound_id, "hotkeys", hotkey_list)
            display_str = get_hotkey_display_string(hotkey_list)
            hotkey_var.set(display_str)
            self.sound_card_widgets[sound_id]['hotkey_var'].set(display_str)
            self.keybind_manager.update_hotkeys()
        HotkeyRecorder(self, sound['name'], on_complete)
        
    def _clear_sound_hotkey(self, sound_id, hotkey_var):
        self.sound_manager.update_sound_property(sound_id, "hotkeys", [])
        hotkey_var.set("Not Assigned")
        self.sound_card_widgets[sound_id]['hotkey_var'].set("Not Assigned")
        self.keybind_manager.update_hotkeys()
        
    def _assign_global_hotkey(self, action, var):
        def on_complete(hotkey_list):
            if hotkey_list is None: return
            self.sound_manager.set_global_hotkey(action, hotkey_list)
            var.set(get_hotkey_display_string(hotkey_list)); self.keybind_manager.update_hotkeys()
        HotkeyRecorder(self, action.replace("_", " ").title(), on_complete)
        
    def _clear_global_hotkey(self, action, var):
        self.sound_manager.set_global_hotkey(action, []); var.set("Not Assigned"); self.keybind_manager.update_hotkeys()
        
    def show_status_message(self, message, style="info"):
        self.status_message_var.set(message); self.status_label.configure(bootstyle=style)
        self.after(5000, lambda: self.status_label.configure(bootstyle="secondary"))
        
    def update_now_playing_status(self):
        with current_playing_sound_details_lock:
            names, is_active = current_playing_sound_details.get("names", []), current_playing_sound_details.get("active", False)
        status_text = "Now Playing: "
        if is_active:
            self.stop_all_button.configure(bootstyle="danger")
            if names:
                status_text += ", ".join(names[:2]);
                if len(names) > 2: status_text += f" & {len(names) - 2} more"
            else: status_text += "Microphone"
        else:
            self.stop_all_button.configure(bootstyle="danger-outline"); status_text += "None"
        self.now_playing_var.set(status_text)
        self.after(250, self.update_now_playing_status)
        
    def _save_app_settings(self):
        settings = {"theme": self.current_theme_var.get(), "master_volume": self.master_volume_var.get(), "soundboard_monitor_volume": self.soundboard_monitor_volume_var.get(), "mic_monitor_volume": self.mic_monitor_volume_var.get(), "soundboard_monitor_enabled": self.soundboard_monitor_enabled_var.get(), "mic_monitor_enabled": self.mic_monitor_enabled_var.get(), "auto_start_mic": self.auto_start_mic_var.get(), "single_sound_mode": self.single_sound_mode_var.get()}
        self.app_settings.save_settings(settings); logging.info("Application settings saved.")
        
    def _on_app_closure(self):
        self._save_app_settings(); self.sound_manager.save_config()
        self.audio_manager.close(); self.keybind_manager.stop(); self.destroy()
        logging.info("--- WarpBoard Closed ---")
        
    def populate_device_dropdowns(self):
            self.unfiltered_output_devices = self.audio_manager.output_devices
            self.unfiltered_input_devices = self.audio_manager.input_devices
            
            output_display_names = []
            for dev in self.unfiltered_output_devices:
                display_name = "🎙 WarpBoard Virtual Mic" if VIRTUAL_MIC_NAME_PARTIAL.lower() in dev['name'].lower() else dev['name']
                
                if display_name not in output_display_names:
                    output_display_names.append(display_name)
            
            self.output_device_combo['values'] = output_display_names
            
            is_virtual_device = lambda name: VIRTUAL_MIC_NAME_PARTIAL.lower() in name.lower() or 'vb-audio' in name.lower()
            self.filtered_input_devices = [dev for dev in self.unfiltered_input_devices if not is_virtual_device(dev['name'])]
            self.filtered_monitor_devices = [dev for dev in self.unfiltered_output_devices if not is_virtual_device(dev['name'])]

            self.input_device_combo['values'] = [dev['name'] for dev in self.filtered_input_devices]
            self.sb_monitor_combo['values'] = [dev['name'] for dev in self.filtered_monitor_devices]
            self.mic_monitor_combo['values'] = [dev['name'] for dev in self.filtered_monitor_devices]

            self._set_combo_from_setting(self.output_device_combo, "output_device_id", self.unfiltered_output_devices, "🎙 WarpBoard Virtual Mic")
            self._set_combo_from_setting(self.input_device_combo, "input_device_id", self.filtered_input_devices)
            self._set_combo_from_setting(self.sb_monitor_combo, "soundboard_monitor_device_id", self.filtered_monitor_devices)
            self._set_combo_from_setting(self.mic_monitor_combo, "mic_monitor_device_id", self.filtered_monitor_devices)
            
    def _set_combo_from_setting(self, combo, setting_key, device_list, preferred_device_name=None):
        device_id = self.app_settings.get_setting(setting_key)
        if device_id is None:
            if combo['values']:
                pass
            return

        try:
            device_name = self.audio_manager.get_device_name_by_index(device_id)
            if device_name in combo['values']:
                combo.set(device_name)
        except Exception as e:
            logging.warning(f"Failed to set device for {setting_key} from saved ID {device_id}: {e}")

    def _on_output_device_selected(self, event):
        selected_index_in_list = event.widget.current()
        device_id = self.unfiltered_output_devices[selected_index_in_list]['index']
        self.app_settings.save_settings({"output_device_id": device_id})
        self.audio_manager.start_main_stream()

    def _on_input_device_selected(self, event):
        selected_index_in_list = event.widget.current()
        device_id = self.filtered_input_devices[selected_index_in_list]['index']
        self.app_settings.save_settings({"input_device_id": device_id})
        if self.include_mic_in_mix_var.get(): self.audio_manager.start_mic_input()

    def _on_sb_monitor_device_selected(self, event):
        selected_index_in_list = event.widget.current()
        device_id = self.filtered_monitor_devices[selected_index_in_list]['index']
        self.app_settings.save_settings({"soundboard_monitor_device_id": device_id})
        if self.soundboard_monitor_enabled_var.get(): self.audio_manager.start_soundboard_monitor_stream()

    def _on_mic_monitor_device_selected(self, event):
        selected_index_in_list = event.widget.current()
        device_id = self.filtered_monitor_devices[selected_index_in_list]['index']
        self.app_settings.save_settings({"mic_monitor_device_id": device_id})
        if self.mic_monitor_enabled_var.get(): self.audio_manager.start_mic_monitor_stream()
        
    def _on_theme_changed(self, _):
        messagebox.showinfo("Theme Change", "Theme will be applied on next restart.", parent=self)
        self._save_app_settings()
        
    def _on_single_sound_mode_changed(self):
        self.audio_manager.mixer.set_single_sound_mode(self.single_sound_mode_var.get())
        self._save_app_settings()
        
    def _apply_settings_to_ui(self):
        self.audio_manager.mixer.set_single_sound_mode(self.single_sound_mode_var.get())
        
        self.audio_manager.set_master_volume(self.master_volume_var.get() / 100.0)
        self.audio_manager.set_sb_monitor_volume(self.soundboard_monitor_volume_var.get() / 100.0)
        self.audio_manager.set_mic_monitor_volume(self.mic_monitor_volume_var.get() / 100.0)

        self.toggle_include_mic_in_mix()
        self.toggle_soundboard_monitor()
        self.toggle_mic_monitor()

def ensure_folders():
    """Creates the necessary application data folders if they don't exist."""
    for folder in [APP_DATA_DIR, SOUNDS_DIR, CONFIG_DIR]:
        os.makedirs(folder, exist_ok=True)

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        pydub.AudioSegment.ffmpeg = get_executable_path('ffmpeg.exe')
        pydub.AudioSegment.ffprobe = get_executable_path('ffprobe.exe')

    ensure_folders()
    
    logging.basicConfig(
        filename=LOG_FILE,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(module)s - %(message)s'
    )

    DependencyChecker.run_checks()
    app = SoundboardApp()
    app.mainloop()