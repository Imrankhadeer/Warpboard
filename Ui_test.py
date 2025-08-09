
import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import sounddevice as sd
import numpy as np
import threading
import pygame
import os
import json
import time
import queue
from pydub import AudioSegment
from pydub.playback import play
from ttkbootstrap.tooltip import ToolTip
import sys
import uuid
from tkinter import filedialog, messagebox, simpledialog
from pynput import keyboard

# Global set to track currently pressed keys for hotkey assignment
current_keys_for_assignment = set()

# Global variable to store the currently playing sound's details for the "Now Playing" display
current_playing_sound_details = {}
current_playing_sound_details_lock = threading.Lock()

# --- Utility Functions ---
def get_pynput_key_string(key):
    """Converts a pynput key object to a human-readable string."""
    try:
        if isinstance(key, keyboard.Key):
            return str(key).split('.')[-1].replace('_', ' ').title()
        elif isinstance(key, keyboard.KeyCode):
            if key.char:
                return key.char.upper() if key.char.isalpha() else key.char 
            else:
                return f"VK_{key.vk}" # Fallback for special chars using virtual key code
        else:
            return str(key).replace('Key.', '').replace('KeyCode.', '') # Catch-all
    except Exception:
        return str(key) # Fallback

def get_hotkey_display_string(hotkey_combination):
    """Converts a list of pynput key objects to a displayable string."""
    if not hotkey_combination:
        return "None"
    
    # Sort keys for consistent display (e.g., Ctrl+A always, not A+Ctrl)
    # Sort order: Modifier keys first (Ctrl, Alt, Shift), then other keys by string representation
    modifier_keys = {keyboard.Key.ctrl, keyboard.Key.alt, keyboard.Key.shift, 
                     keyboard.Key.ctrl_l, keyboard.Key.ctrl_r, 
                     keyboard.Key.alt_l, keyboard.Key.alt_r, 
                     keyboard.Key.shift_l, keyboard.Key.shift_r}
    
    sorted_combination = []
    
    # Add modifiers first
    for key in hotkey_combination:
        if key in modifier_keys:
            sorted_combination.append(key)
    
    # Add non-modifiers, sorted alphabetically by their string representation
    non_modifiers = sorted([key for key in hotkey_combination if key not in modifier_keys], key=get_pynput_key_string)
    sorted_combination.extend(non_modifiers)

    return " + ".join([get_pynput_key_string(key) for key in sorted_combination])

def get_audio_device_names(kind='output'):
    """Fetches a list of available audio device names, including 'Default'."""
    device_names = ["Default"] # Always include Default
    try:
        devices = sd.query_devices()
        for i, d in enumerate(devices):
            # Use .get() with a default value to safely access dictionary keys
            # to prevent KeyError if 'max_outputs' or 'max_inputs' is missing.
            max_channels = d.get(f'max_{kind}_channels', 0)
            if max_channels > 0:
                device_names.append(d['name'])
    except sd.PortAudioError as e:
        print(f"PortAudio error querying audio devices for {kind}: {e}")
        print("Please ensure your audio drivers are correctly installed and working.")
    except Exception as e:
        print(f"Unexpected error querying audio devices for {kind}: {e}")
    
    # Remove duplicates while preserving order (approximate order of discovery + Default first)
    return list(dict.fromkeys(device_names))

# --- AudioOutputManager Class ---
class AudioOutputManager:
    def __init__(self, app_instance, output_device_name=None, input_device_name=None, monitor_device_name=None):
        self.app_instance = app_instance
        self.output_device_name = output_device_name
        self.input_device_name = input_device_name
        self.monitor_device_name = monitor_device_name

        self.loopback_stream = None
        self.mic_passthrough_stream = None
        self.monitor_stream = None

        # Pygame mixer will always initialize to the system's default device.
        # sounddevice will manage actual device routing if a virtual cable is selected.
        self.initialize_pygame_mixer_default()

    def initialize_pygame_mixer_default(self):
        try:
            if pygame.mixer.get_init():
                pygame.mixer.quit()
            pygame.mixer.init() # Initialize with default device
            print("Pygame mixer initialized with default device.")
        except Exception as e:
            print(f"Error initializing Pygame mixer: {e}")
            # Do not re-raise or quit, just log. Playback might be affected.

    def update_output_device(self, device_name):
        self.output_device_name = device_name
        # Pygame is no longer explicitly set to a device, it always uses default.
        # So, no pygame update needed here.
        self.restart_streams() # Restart relevant streams to apply new device

    def update_input_device(self, device_name):
        self.input_device_name = device_name
        self.restart_streams()

    def update_monitor_device(self, device_name):
        self.monitor_device_name = device_name
        self.restart_streams()

    def restart_streams(self):
        # Stop all streams
        self.stop_loopback_stream()
        self.stop_mic_passthrough_stream()
        self.stop_monitor_stream()

        # Re-start based on current UI state after a short delay to ensure devices are free
        self.app_instance.after(100, self._delayed_stream_restart)

    def _delayed_stream_restart(self):
        if self.app_instance.loopback_var.get():
            self.start_loopback_stream()
        if self.app_instance.mic_passthrough_var.get():
            self.start_mic_passthrough_stream()
        if self.app_instance.monitor_output_var.get():
            self.start_monitor_stream()

    def _loopback_callback(self, outdata, frames, time_info, status):
        if status:
            print(f"Loopback stream status: {status}")
        outdata.fill(0) # Fill with zeros for silence if no actual loopback logic needed

    def start_loopback_stream(self):
        if self.loopback_stream is None or self.loopback_stream.stopped:
            try:
                output_device_id = None
                if self.output_device_name and self.output_device_name != "Default":
                    devices = sd.query_devices()
                    for i, d in enumerate(devices):
                        if d['name'] == self.output_device_name and d['max_output_channels'] > 0:
                            output_device_id = i
                            break
                
                self.loopback_stream = sd.OutputStream(
                    device=output_device_id,
                    channels=2, # Stereo output
                    callback=self._loopback_callback,
                    samplerate=44100,
                    dtype='float32'
                )
                self.loopback_stream.start()
                print(f"Loopback stream started on device: {self.output_device_name}")
            except Exception as e:
                print(f"Error starting loopback stream: {e}")
                self.app_instance.after(0, lambda current_e=e: messagebox.showerror("Audio Error", f"Could not start loopback stream: {current_e}"))
                self.app_instance.after(0, self.app_instance.loopback_var.set(False))

    def stop_loopback_stream(self):
        if self.loopback_stream and self.loopback_stream.active:
            self.loopback_stream.stop()
        if self.loopback_stream: # Also close if not active but exists
            self.loopback_stream.close()
            self.loopback_stream = None
            print("Loopback stream stopped.")

    def _mic_passthrough_callback(self, indata, outdata, frames, time_info, status):
        if status:
            print(f"Mic passthrough stream status: {status}")
        outdata[:] = indata # Route input directly to main output

        # If monitoring is enabled and monitor stream is active, send mic data to it too
        if self.app_instance.monitor_output_var.get() and self.monitor_stream and self.monitor_stream.active:
            try:
                indata_channels = indata.shape[1]
                monitor_stream_channels = self.monitor_stream.channels

                data_to_write = indata.copy()
                if indata_channels == 1 and monitor_stream_channels == 2:
                    # Duplicate mono to stereo
                    data_to_write = np.repeat(indata, 2, axis=1)
                elif indata_channels != monitor_stream_channels:
                    print(f"Warning: Channel mismatch in monitor stream. Input: {indata_channels}, Monitor Output: {monitor_stream_channels}. Writing directly, may cause issues.")
                
                self.monitor_stream.write(data_to_write)
            except sd.CallbackStop: # If the monitor stream itself got stopped
                raise # Propagate stop signal
            except Exception as e:
                print(f"Error writing to monitor stream in mic passthrough callback: {e}")

    def start_mic_passthrough_stream(self):
        if self.mic_passthrough_stream is None or self.mic_passthrough_stream.stopped:
            try:
                input_device_id = None
                output_device_id = None
                
                if self.input_device_name and self.input_device_name != "Default":
                    devices = sd.query_devices()
                    for i, d in enumerate(devices):
                        if d['name'] == self.input_device_name and d['max_input_channels'] > 0:
                            input_device_id = i
                            break
                
                if self.output_device_name and self.output_device_name != "Default":
                    devices = sd.query_devices()
                    for i, d in enumerate(devices):
                        if d['name'] == self.output_device_name and d['max_output_channels'] > 0:
                            output_device_id = i
                            break

                self.mic_passthrough_stream = sd.Stream(
                    samplerate=44100,
                    channels=[1, 2], # Input mono, Output stereo
                    dtype='float32',
                    callback=self._mic_passthrough_callback,
                    device=(input_device_id, output_device_id)
                )
                self.mic_passthrough_stream.start()
                print(f"Mic passthrough started (Input: {self.input_device_name}, Output: {self.output_device_name})")
            except Exception as e:
                print(f"Error starting mic passthrough stream: {e}")
                self.app_instance.after(0, lambda current_e=e: messagebox.showerror("Audio Error", f"Could not start mic passthrough stream: {current_e}"))
                self.app_instance.after(0, self.app_instance.mic_passthrough_var.set(False))

    def stop_mic_passthrough_stream(self):
        if self.mic_passthrough_stream and self.mic_passthrough_stream.active:
            self.mic_passthrough_stream.stop()
        if self.mic_passthrough_stream:
            self.mic_passthrough_stream.close()
            self.mic_passthrough_stream = None
            print("Mic passthrough stopped.")

    def start_monitor_stream(self):
        if self.monitor_stream is None or self.monitor_stream.stopped:
            try:
                monitor_device_id = None
                if self.monitor_device_name and self.monitor_device_name != "Default":
                    devices = sd.query_devices()
                    for i, d in enumerate(devices):
                        if d['name'] == self.monitor_device_name and d['max_output_channels'] > 0:
                            monitor_device_id = i
                            break
                
                # If monitor_device_id is still None, it means either "Default" was chosen or the named device wasn't found.
                # sd.OutputStream handles None device gracefully by using default.

                self.monitor_stream = sd.OutputStream(
                    device=monitor_device_id,
                    channels=2, # Stereo output
                    samplerate=44100,
                    dtype='float32'
                )
                self.monitor_stream.start()
                print(f"Monitor stream started on device: {self.monitor_device_name}")
            except Exception as e:
                print(f"Error starting monitor stream: {e}")
                self.app_instance.after(0, lambda current_e=e: messagebox.showerror("Audio Error", f"Could not start monitor stream: {current_e}"))
                self.app_instance.after(0, self.app_instance.monitor_output_var.set(False))

    def stop_monitor_stream(self):
        if self.monitor_stream and self.monitor_stream.active:
            self.monitor_stream.stop()
        if self.monitor_stream:
            self.monitor_stream.close()
            self.monitor_stream = None
            print("Monitor stream stopped.")

# --- KeyBindManager Class ---
class KeyBindManager:
    def __init__(self, app_instance):
        self.app_instance = app_instance
        self.sound_hotkeys = {} # {hotkey_combination_tuple: sound_button_instance}
        self.global_stop_hotkey_combination = None
        self.global_stop_hotkey_callback = None
        self.listener = None
        self.key_assignment_callback = None
        self.is_assigning_key = False
        self.current_assigned_key_combination = set() # For assignment mode
        self.pressed_keys = set() # Track currently pressed keys for hotkey detection during normal operation
        self._global_hotkey_triggered = False # Flag to prevent repeated global hotkey triggers
        self._sound_hotkey_triggered_per_press = {} # Tracks if a sound hotkey was triggered for the current press cycle

        self.listener_thread = None

        self.start_listener()

    def start_listener(self):
        if self.listener is None or not self.listener.running:
            self.listener = keyboard.Listener(
                on_press=self._on_press,
                on_release=self._on_release
            )
            self.listener_thread = threading.Thread(target=self.listener.start, daemon=True)
            self.listener_thread.start()
            print("Keyboard listener started.")

    def stop_listener(self):
        if self.listener and self.listener.running:
            self.listener.stop()
            if self.listener_thread and self.listener_thread.is_alive():
                self.listener_thread.join(timeout=1)
            print("Keyboard listener stopped.")

    def _on_press(self, key):
        if self.is_assigning_key:
            if key not in current_keys_for_assignment:
                current_keys_for_assignment.add(key)
                self.app_instance.after(0, self._update_key_assignment_display)
        else:
            if key not in self.pressed_keys:
                self.pressed_keys.add(key)
            
            # Reset triggered flags when new keys are pressed if needed, or manage on release
            # For now, keep the _sound_hotkey_triggered_per_press state per hotkey combination.

            # Check for global stop hotkey FIRST
            if self.global_stop_hotkey_combination and set(self.global_stop_hotkey_combination).issubset(self.pressed_keys):
                if not hasattr(self, '_global_hotkey_triggered') or not self._global_hotkey_triggered:
                    self.app_instance.after(0, self.global_stop_hotkey_callback)
                    self._global_hotkey_triggered = True 
                return

            # Check for sound hotkeys
            for hotkey_tuple, sound_button in self.sound_hotkeys.items():
                if set(hotkey_tuple).issubset(self.pressed_keys):
                    # Only trigger if not already triggered in this press cycle for this specific hotkey combo
                    if hotkey_tuple not in self._sound_hotkey_triggered_per_press:
                        self.app_instance.after(0, sound_button.play_sound_from_hotkey)
                        self._sound_hotkey_triggered_per_press[hotkey_tuple] = True
                    # Do not break here if multiple hotkeys might share common keys, allowing all to fire if subsets match.
                    # If you only want the "most specific" or one hotkey to fire, you need more complex logic (e.g., sort by length).
                    # For now, it will trigger all matching subsets on the first press.

    def _on_release(self, key):
        if self.is_assigning_key:
            if key in current_keys_for_assignment:
                current_keys_for_assignment.discard(key)
                self.app_instance.after(0, self._update_key_assignment_display)
        else:
            if key in self.pressed_keys:
                self.pressed_keys.discard(key)
                
            # Reset global hotkey triggered flag if all relevant keys for it are released
            if self.global_stop_hotkey_combination and not set(self.global_stop_hotkey_combination).issubset(self.pressed_keys):
                self._global_hotkey_triggered = False

            # Reset sound hotkey triggered flags if their relevant keys are released
            # This is to allow re-triggering if the hotkey combination is pressed again.
            keys_to_remove_from_triggered = []
            for hotkey_tuple, _ in self._sound_hotkey_triggered_per_press.items():
                if not set(hotkey_tuple).issubset(self.pressed_keys):
                    keys_to_remove_from_triggered.append(hotkey_tuple)
            for hotkey_tuple in keys_to_remove_from_triggered:
                del self._sound_hotkey_triggered_per_press[hotkey_tuple]


    def start_key_assignment_mode(self, callback):
        self.is_assigning_key = True
        self.key_assignment_callback = callback
        current_keys_for_assignment.clear()
        self.current_assigned_key_combination.clear() # Clear accumulated keys for new assignment
        self.app_instance.after(0, self._update_key_assignment_display)

    def _update_key_assignment_display(self):
        if self.is_assigning_key:
            display_str = get_hotkey_display_string(list(current_keys_for_assignment))
            if hasattr(self.app_instance, 'current_hotkey_label'):
                self.app_instance.current_hotkey_label.config(text=f"Press hotkey... Current: {display_str}")
            
            # Update the set of keys that were part of this assignment attempt
            if current_keys_for_assignment:
                self.current_assigned_key_combination.update(current_keys_for_assignment)
            
            # Check if all keys have been released and a combination was captured
            elif not current_keys_for_assignment and self.current_assigned_key_combination:
                self.is_assigning_key = False
                final_assigned_keys = tuple(sorted(list(self.current_assigned_key_combination), key=str))
                self.app_instance.after(0, lambda assigned_keys=final_assigned_keys: self.key_assignment_callback(assigned_keys))
                self.app_instance.after(0, lambda: self.app_instance.current_hotkey_label.config(text=""))
                self.current_assigned_key_combination.clear()
                
    def assign_hotkey_to_sound(self, sound_button_instance, hotkey_combination_tuple):
        # Remove any existing assignment for this sound_button_instance
        self.remove_hotkey_from_sound(sound_button_instance) # This clears old hotkey from SoundButton itself

        # Check if the hotkey combination is already assigned to another sound
        if hotkey_combination_tuple in self.sound_hotkeys:
            existing_sound_button = self.sound_hotkeys[hotkey_combination_tuple]
            # Unassign from the old sound button, but only if it still exists
            if existing_sound_button and hasattr(existing_sound_button, 'winfo_exists') and existing_sound_button.winfo_exists():
                existing_sound_button.hotkey_combination = None
                existing_sound_button.update_hotkey_display()
                print(f"Hotkey {get_hotkey_display_string(hotkey_combination_tuple)} unassigned from {os.path.basename(existing_sound_button.sound_file_path)}")
            else:
                # If widget doesn't exist, remove it from sound_hotkeys explicitly
                del self.sound_hotkeys[hotkey_combination_tuple]
                print(f"Warning: Hotkey {get_hotkey_display_string(hotkey_combination_tuple)} was assigned to a destroyed sound button and removed.")

        self.sound_hotkeys[hotkey_combination_tuple] = sound_button_instance
        sound_button_instance.hotkey_combination = hotkey_combination_tuple
        sound_button_instance.update_hotkey_display()
        print(f"Assigned hotkey {get_hotkey_display_string(hotkey_combination_tuple)} to {os.path.basename(sound_button_instance.sound_file_path)}")


    def remove_hotkey_from_sound(self, sound_button_instance):
        hotkey_to_remove = None
        for hotkey_tuple, sb_instance in self.sound_hotkeys.items():
            if sb_instance == sound_button_instance:
                hotkey_to_remove = hotkey_tuple
                break
        if hotkey_to_remove:
            del self.sound_hotkeys[hotkey_to_remove]
        sound_button_instance.hotkey_combination = None
        sound_button_instance.update_hotkey_display()
        print(f"Removed hotkey from {os.path.basename(sound_button_instance.sound_file_path)}")

    def set_global_hotkey(self, hotkey_combination, hotkey_var, callback):
        self.global_stop_hotkey_combination = tuple(sorted(hotkey_combination, key=str))
        self.global_stop_hotkey_callback = callback
        hotkey_var.set(get_hotkey_display_string(hotkey_combination))
        print(f"Global stop hotkey set to: {get_hotkey_display_string(hotkey_combination)}")

# --- SoundButton Class ---
class SoundButton(ttk.Frame):
    def __init__(self, parent, app_instance, sound_file_path, hotkey_combination=None):
        super().__init__(parent, relief=ttk.RAISED, borderwidth=2)
        self.app_instance = app_instance
        self.sound_file_path = sound_file_path
        self.hotkey_combination = hotkey_combination # Stored as tuple of pynput key objects
        self.sound = None # Will be loaded on demand

        self.create_widgets()
        self.update_hotkey_display()

    def create_widgets(self):
        file_name = os.path.basename(self.sound_file_path)

        # File name label
        self.name_label = ttk.Label(self, text=file_name, anchor=W, bootstyle="inverse-primary")
        self.name_label.pack(side=TOP, fill=X, expand=True, padx=5, pady=2)

        # Hotkey display Entry
        self.hotkey_entry_var = tk.StringVar(value=get_hotkey_display_string(self.hotkey_combination))
        self.hotkey_entry = ttk.Entry(self, textvariable=self.hotkey_entry_var, state="readonly", bootstyle="info")
        self.hotkey_entry.pack(side=TOP, fill=X, expand=True, padx=5, pady=1)
        self.hotkey_entry.bind("<Button-1>", self.assign_hotkey_click) # Bind click to start assignment
        ToolTip(self.hotkey_entry, text="Click to assign/reassign hotkey")

        # Buttons frame for Play, Clear Hotkey, Delete
        buttons_frame = ttk.Frame(self)
        buttons_frame.pack(side=BOTTOM, fill=X, padx=5, pady=5)

        # Play button
        self.play_button = ttk.Button(buttons_frame, text="Play", command=self.play_sound, bootstyle="success")
        self.play_button.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 5))

        # Clear Hotkey button
        self.clear_hotkey_button = ttk.Button(buttons_frame, text="Clear Hotkey", command=self.clear_hotkey, bootstyle="secondary-outline")
        self.clear_hotkey_button.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 5))

        # Delete button
        self.delete_button = ttk.Button(buttons_frame, text="Delete", command=self.delete_sound, bootstyle="danger")
        self.delete_button.pack(side=RIGHT, fill=BOTH, expand=True, padx=(5, 0)) # Pad to the right

    def update_hotkey_display(self):
        if not self.winfo_exists(): # Check if the SoundButton frame itself exists
            return
        if not hasattr(self, 'hotkey_entry') or not self.hotkey_entry.winfo_exists(): # Check if the Entry widget exists
            return
            
        display_text = get_hotkey_display_string(self.hotkey_combination)
        # Temporarily make entry writable to update text, then set back to readonly
        self.hotkey_entry.config(state="normal")
        self.hotkey_entry_var.set(display_text)
        self.hotkey_entry.config(state="readonly")

    def load_sound(self):
        if self.sound is None:
            try:
                self.sound = pygame.mixer.Sound(self.sound_file_path)
                print(f"Loaded sound: {os.path.basename(self.sound_file_path)}")
            except pygame.error as e:
                messagebox.showerror("Sound Load Error", f"Could not load {os.path.basename(self.sound_file_path)}: {e}", parent=self.app_instance)
                self.sound = None

    def play_sound(self):
        self.load_sound()
        if self.sound:
            self.app_instance.sound_queue.put((self.sound, self.sound_file_path))
            print(f"Queued sound: {os.path.basename(self.sound_file_path)}")

    def play_sound_from_hotkey(self):
        self.play_sound()

    def assign_hotkey_click(self, event=None):
        # Clear the hotkey entry and show prompt
        self.hotkey_entry.config(state="normal")
        self.hotkey_entry_var.set("Press hotkey...")
        self.hotkey_entry.config(state="readonly")
        
        self.app_instance.key_bind_manager.start_key_assignment_mode(self._on_hotkey_assigned)

    def _on_hotkey_assigned(self, hotkey_combination_tuple):
        # This is called from the key_bind_manager listener thread via app_instance.after(0, ...)
        # So it runs on the main Tkinter thread.
        self.app_instance.key_bind_manager.assign_hotkey_to_sound(self, hotkey_combination_tuple)
        self.app_instance.save_settings()

    def clear_hotkey(self):
        if messagebox.askyesno("Clear Hotkey", f"Are you sure you want to clear the hotkey for '{os.path.basename(self.sound_file_path)}'?", parent=self.app_instance):
            self.app_instance.key_bind_manager.remove_hotkey_from_sound(self)
            self.app_instance.save_settings()

    def delete_sound(self):
        if messagebox.askyesno("Delete Sound", f"Are you sure you want to delete '{os.path.basename(self.sound_file_path)}'?", parent=self.app_instance):
            self.app_instance.remove_sound_button(self)
            self.app_instance.key_bind_manager.remove_hotkey_from_sound(self) # Ensure hotkey is removed from manager
            self.destroy() # Destroy the tkinter widget
            self.app_instance.save_settings()

# --- SoundboardApp Class ---
class SoundboardApp(ttk.Window):
    def __init__(self):
        super().__init__(themename="superhero")
        self.title("WarpBoard")
        self.geometry("800x600")

        self.sound_buttons = []
        self.sound_queue = queue.Queue()
        self.playback_thread = threading.Thread(target=self._playback_worker, daemon=True)
        self.playback_thread.start()

        self.key_bind_manager = KeyBindManager(self)
        
        # Load settings first to get device names
        self.load_settings()

        # Initialize AudioOutputManager with loaded settings
        self.audio_output_manager = AudioOutputManager(
            self,
            output_device_name=self.settings.get("output_device", "Default"),
            input_device_name=self.settings.get("input_device", "Default"),
            monitor_device_name=self.settings.get("monitor_device", "Default")
        )

        self.create_widgets()
        self.update_now_playing_display()

        # Ensure correct device is set on startup for AudioOutputManager
        # These are now handled by _delayed_stream_restart
        # But ensure initial state is set
        if self.loopback_var.get():
            self.audio_output_manager.start_loopback_stream()
        if self.mic_passthrough_var.get():
            self.audio_output_manager.start_mic_passthrough_stream()
        if self.monitor_output_var.get():
            self.audio_output_manager.start_monitor_stream()

        # Bind closing event
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        # Top-level "Now Playing" frame remains outside tabs
        self.now_playing_frame = ttk.LabelFrame(self, text="Now Playing", bootstyle="info")
        self.now_playing_frame.pack(side=TOP, fill=X, padx=10, pady=5)
        self.create_now_playing_widgets(self.now_playing_frame)

        # Notebook for tabs (Sounds and Settings)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(side=TOP, fill=BOTH, expand=True, padx=10, pady=5)

        # --- Sounds Tab ---
        self.sounds_tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.sounds_tab_frame, text="Sounds", sticky=NSEW)
        self.sounds_tab_frame.columnconfigure(0, weight=1) # Allow sound buttons to expand

        # Controls for sounds tab (Add, Stop All, Current Hotkey display)
        self.sounds_controls_frame = ttk.Frame(self.sounds_tab_frame)
        self.sounds_controls_frame.pack(side=TOP, fill=X, padx=10, pady=5)

        self.add_sound_button = ttk.Button(self.sounds_controls_frame, text="Add Sound", command=self.add_sound, bootstyle="primary")
        self.add_sound_button.pack(side=LEFT, padx=5)

        self.stop_all_button = ttk.Button(self.sounds_controls_frame, text="Stop All Sounds", command=self.stop_all_sounds, bootstyle="warning")
        self.stop_all_button.pack(side=LEFT, padx=5)

        self.current_hotkey_label = ttk.Label(self.sounds_controls_frame, text="", bootstyle="warning")
        self.current_hotkey_label.pack(side=LEFT, padx=5)

        # Canvas for scrollable sound buttons (reparented to sounds_tab_frame)
        self.canvas = tk.Canvas(self.sounds_tab_frame, highlightthickness=0) # Remove default canvas border
        self.scrollbar = ttk.Scrollbar(self.sounds_tab_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas) # This frame holds the sound buttons

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        self.scrollbar.pack(side="right", fill="y")

        # --- Settings Tab ---
        self.settings_tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_tab_frame, text="Settings", sticky=NSEW)
        # Configure grid for settings frame to allow columns to expand
        self.settings_tab_frame.columnconfigure(1, weight=1)
        self.settings_tab_frame.columnconfigure(2, weight=1) # For the virtual mic tip
        self.create_settings_widgets(self.settings_tab_frame) # Pass the new settings frame

        self.repopulate_sound_buttons() # Add existing sounds from settings

    def create_now_playing_widgets(self, parent):
        self.now_playing_sound_name = tk.StringVar(value="No sound playing")
        self.now_playing_current_time = tk.StringVar(value="00:00")
        self.now_playing_total_duration = tk.StringVar(value="00:00")

        ttk.Label(parent, textvariable=self.now_playing_sound_name, font=("TkDefaultFont", 12, "bold")).pack(pady=2)
        ttk.Label(parent, text="—").pack(pady=0)
        
        time_frame = ttk.Frame(parent)
        time_frame.pack(pady=2)
        ttk.Label(time_frame, textvariable=self.now_playing_current_time).pack(side=LEFT)
        ttk.Label(time_frame, text=" / ").pack(side=LEFT)
        ttk.Label(time_frame, textvariable=self.now_playing_total_duration).pack(side=LEFT)

        self.now_playing_progress_bar = ttk.Progressbar(parent, orient="horizontal", mode="determinate", length=400)
        self.now_playing_progress_bar.pack(pady=5)

        self.now_playing_stop_button = ttk.Button(parent, text="Stop Current", command=self.stop_current_sound, bootstyle="danger")
        self.now_playing_stop_button.pack(pady=2)
    
    def create_settings_widgets(self, parent):
        # Audio Output Device Selection
        output_devices = get_audio_device_names('output')
        self.selected_output_device = tk.StringVar(value=self.settings.get("output_device", "Default"))
        self.output_device_menu = ttk.OptionMenu(parent, self.selected_output_device, self.selected_output_device.get(), *output_devices, command=self.on_output_device_change)
        ttk.Label(parent, text="Main Output Device:").grid(row=0, column=0, padx=5, pady=5, sticky=W)
        self.output_device_menu.grid(row=0, column=1, padx=5, pady=5, sticky=EW)
        
        # New: Virtual Mic explanation
        virtual_mic_label = ttk.Label(parent, text="Tip: Select a Virtual Mic (e.g., VB-Cable) here if you want other apps to hear the output.", bootstyle="info")
        virtual_mic_label.grid(row=0, column=2, padx=10, pady=5, sticky=W)


        # Audio Input Device Selection (for mic passthrough)
        input_devices = get_audio_device_names('input')
        self.selected_input_device = tk.StringVar(value=self.settings.get("input_device", "Default"))
        self.input_device_menu = ttk.OptionMenu(parent, self.selected_input_device, self.selected_input_device.get(), *input_devices, command=self.on_input_device_change)
        ttk.Label(parent, text="Input Device (Mic):").grid(row=1, column=0, padx=5, pady=5, sticky=W)
        self.input_device_menu.grid(row=1, column=1, padx=5, pady=5, sticky=EW)

        # Loopback checkbox
        self.loopback_var = tk.BooleanVar(value=self.settings.get("loopback_enabled", False))
        self.loopback_checkbox = ttk.Checkbutton(parent, text="Enable Loopback (Listen to your Mic)", variable=self.loopback_var, bootstyle="round-toggle")
        self.loopback_checkbox.grid(row=2, column=0, padx=10, pady=5, sticky=W)
        self.loopback_var.trace_add("write", self.toggle_loopback)
        ToolTip(self.loopback_checkbox, text="Routes your microphone input directly to your selected MAIN OUTPUT device. Use for basic mic monitoring.")

        # Mic Passthrough checkbox
        self.mic_passthrough_var = tk.BooleanVar(value=self.settings.get("mic_passthrough_enabled", False))
        self.mic_passthrough_checkbox = ttk.Checkbutton(parent, text="Enable Mic Passthrough (Send Mic to Output)", variable=self.mic_passthrough_var, bootstyle="round-toggle")
        self.mic_passthrough_checkbox.grid(row=2, column=1, padx=10, pady=5, sticky=W)
        self.mic_passthrough_var.trace_add("write", self.toggle_mic_passthrough)
        ToolTip(self.mic_passthrough_checkbox, text="Routes microphone input directly to the selected MAIN OUTPUT device. Others will hear your mic.")

        # New: Monitor Output Option (Hear what others hear)
        monitor_devices = get_audio_device_names('output')
        self.selected_monitor_device = tk.StringVar(value=self.settings.get("monitor_device", "Default"))
        self.monitor_output_var = tk.BooleanVar(value=self.settings.get("monitor_output_enabled", False))
        
        ttk.Label(parent, text="Monitor Output On:").grid(row=3, column=0, padx=5, pady=5, sticky=W)
        self.monitor_output_menu = ttk.OptionMenu(parent, self.selected_monitor_device, self.selected_monitor_device.get(), *monitor_devices, command=self.on_monitor_device_change)
        self.monitor_output_menu.grid(row=3, column=1, padx=5, pady=5, sticky=EW)

        self.monitor_output_checkbox = ttk.Checkbutton(parent, text="Enable Monitor (Hear Mic Passthrough Output)", variable=self.monitor_output_var, bootstyle="round-toggle")
        self.monitor_output_checkbox.grid(row=3, column=2, padx=10, pady=5, sticky=W)
        self.monitor_output_var.trace_add("write", self.toggle_monitor_output)
        ToolTip(self.monitor_output_checkbox, text="Routes microphone audio from 'Mic Passthrough' to your selected 'Monitor Output On' device.\n\nTo hear *all* output (mic + soundboard sounds):\n1. Set 'Main Output Device' (above) to a Virtual Cable (e.g., VB-Cable).\n2. In your OS sound settings, 'Listen to this device' for that Virtual Cable, routing its output to your headphones.")

        # Global Stop Hotkey UI
        self.set_stop_all_hotkey_button = ttk.Button(parent, text="Set Stop All Hotkey", command=self.set_stop_all_hotkey, bootstyle="info")
        self.set_stop_all_hotkey_button.grid(row=4, column=0, padx=5, pady=5, sticky=W)
        
        self.stop_all_hotkey_var = tk.StringVar(value=get_hotkey_display_string(self.key_bind_manager.global_stop_hotkey_combination) if self.key_bind_manager.global_stop_hotkey_combination else "None")
        self.stop_all_hotkey_label = ttk.Label(parent, textvariable=self.stop_all_hotkey_var, bootstyle="info")
        self.stop_all_hotkey_label.grid(row=4, column=1, padx=5, pady=5, sticky=EW)

    def on_output_device_change(self, device_name):
        self.audio_output_manager.update_output_device(device_name)
        self.save_settings()

    def on_input_device_change(self, device_name):
        self.audio_output_manager.update_input_device(device_name)
        self.save_settings()

    def on_monitor_device_change(self, device_name):
        self.audio_output_manager.update_monitor_device(device_name)
        self.save_settings()
        # If monitoring is active, restart stream to apply new device
        if self.monitor_output_var.get():
            self.audio_output_manager.stop_monitor_stream()
            self.audio_output_manager.start_monitor_stream()

    def toggle_loopback(self, *args):
        if self.loopback_var.get():
            self.audio_output_manager.start_loopback_stream()
        else:
            self.audio_output_manager.stop_loopback_stream()
        self.save_settings()

    def toggle_mic_passthrough(self, *args):
        if self.mic_passthrough_var.get():
            self.audio_output_manager.start_mic_passthrough_stream()
        else:
            self.audio_output_manager.stop_mic_passthrough_stream()
        self.save_settings()

    def toggle_monitor_output(self, *args):
        if self.monitor_output_var.get():
            self.audio_output_manager.start_monitor_stream()
        else:
            self.audio_output_manager.stop_monitor_stream()
        self.save_settings()

    def _playback_worker(self):
        while True:
            try:
                sound, file_path = self.sound_queue.get()
                if sound is None: # Sentinel value to stop the thread
                    break

                if not pygame.mixer.get_init():
                    print("Pygame mixer not initialized in playback worker. Attempting to re-init.")
                    # Re-initialize pygame mixer if it's not initialized (e.g., after suspend/resume)
                    self.audio_output_manager.initialize_pygame_mixer_default()

                with current_playing_sound_details_lock:
                    current_playing_sound_details["name"] = os.path.basename(file_path)
                    current_playing_sound_details["start_time"] = time.time()
                    current_playing_sound_details["duration"] = sound.get_length()

                self.after(0, self.update_now_playing_display)

                channel = sound.play()
                print(f"Playing {os.path.basename(file_path)}")

                while channel.get_busy():
                    try:
                        current_pos_ms = channel.get_pos()
                        current_pos_s = current_pos_ms / 1000.0
                        total_duration_s = sound.get_length()

                        self.after(0, lambda current_s=current_pos_s, total_s=total_duration_s: self._update_progress_bar_and_time(current_s, total_s))
                        time.sleep(0.1)
                    except AttributeError as ae: # Catch specifically for get_pos
                        print(f"AttributeError in playback worker (get_pos): {ae}. Channel may have become invalid, stopping progress update.")
                        break # Exit loop if get_pos is problematic
                    except Exception as e:
                        print(f"Error during playback progress update: {e}")
                        break # Exit loop on other errors

                print(f"Finished playing {os.path.basename(file_path)}")

            except queue.Empty:
                time.sleep(0.1)
            except Exception as e:
                print(f"Error in playback worker: {e}")
            finally:
                with current_playing_sound_details_lock:
                    current_playing_sound_details.clear()
                self.after(0, self.update_now_playing_display)

    def _update_progress_bar_and_time(self, current_pos_s, total_duration_s):
        minutes_current = int(current_pos_s // 60)
        seconds_current = int(current_pos_s % 60)
        minutes_total = int(total_duration_s // 60)
        seconds_total = int(total_duration_s % 60)

        self.now_playing_current_time.set(f"{minutes_current:02}:{seconds_current:02}")
        self.now_playing_total_duration.set(f"{minutes_total:02}:{seconds_total:02}")

        if total_duration_s > 0:
            progress_value = (current_pos_s / total_duration_s) * 100
            self.now_playing_progress_bar.config(value=progress_value)
        else:
            self.now_playing_progress_bar.config(value=0)

    def update_now_playing_display(self):
        with current_playing_sound_details_lock:
            if current_playing_sound_details:
                name = current_playing_sound_details.get("name", "Unknown Sound")
                duration = current_playing_sound_details.get("duration", 0)

                minutes_total = int(duration // 60)
                seconds_total = int(duration % 60)

                self.now_playing_sound_name.set(name)
                self.now_playing_total_duration.set(f"{minutes_total:02}:{seconds_total:02}")
                self.now_playing_current_time.set("00:00")
                self.now_playing_progress_bar.config(value=0)
            else:
                self.now_playing_sound_name.set("No sound playing")
                self.now_playing_current_time.set("00:00")
                self.now_playing_total_duration.set("00:00")
                self.now_playing_progress_bar.config(value=0)
    
    def stop_current_sound(self):
        if pygame.mixer.get_busy():
            pygame.mixer.stop()
            print("Stopped current sound(s).")
        with current_playing_sound_details_lock:
            current_playing_sound_details.clear()
        self.update_now_playing_display()

    def stop_all_sounds(self):
        if pygame.mixer.get_busy():
            pygame.mixer.stop()
            print("Stopped all sounds via hotkey.")
        with current_playing_sound_details_lock:
            current_playing_sound_details.clear()
        self.update_now_playing_display()

    def add_sound(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Audio Files", "*.mp3 *.wav *.ogg")]
        )
        if file_path:
            for btn in self.sound_buttons:
                if btn.sound_file_path == file_path:
                    messagebox.showwarning("Duplicate Sound", "This sound is already added.", parent=self)
                    return

            sound_button = SoundButton(self.scrollable_frame, self, file_path)
            self.sound_buttons.append(sound_button)
            sound_button.pack(fill=X, padx=5, pady=5)
            self.canvas.update_idletasks()
            self.canvas.config(scrollregion=self.canvas.bbox("all"))
            self.save_settings()

    def remove_sound_button(self, button_to_remove):
        self.sound_buttons.remove(button_to_remove)
        self.canvas.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def repopulate_sound_buttons(self):
        # Clear existing buttons from UI first, but keep hotkeys in manager
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.sound_buttons.clear()
        self.key_bind_manager.sound_hotkeys.clear() # Clear manager's hotkeys for fresh repopulation

        for sound_data in self.settings.get("sounds", []):
            file_path = sound_data.get("file_path")
            hotkey_combo_serializable = sound_data.get("hotkey_combination")
            
            hotkey_combination_pynput = []
            if hotkey_combo_serializable:
                for key_str in hotkey_combo_serializable:
                    if key_str.startswith('Key.'):
                        try:
                            hotkey_combination_pynput.append(getattr(keyboard.Key, key_str.split('.')[-1].lower())) # lower() for consistency
                        except AttributeError:
                            print(f"Warning: Unknown pynput.Key '{key_str}' for hotkey.")
                            pass
                    elif key_str.startswith('VK_'): # Custom format for VK codes
                        try:
                            vk_code = int(key_str[3:])
                            hotkey_combination_pynput.append(keyboard.KeyCode.from_vk(vk_code))
                        except ValueError:
                            print(f"Warning: Invalid VK code format '{key_str}' for hotkey.")
                            pass
                    elif len(key_str) == 1: # Single character
                        hotkey_combination_pynput.append(keyboard.KeyCode(char=key_str))
                    else:
                        print(f"Warning: Unknown hotkey key format '{key_str}'.")


            if os.path.exists(file_path):
                # Ensure hotkey_combination_pynput is a tuple for hashing if it's not empty
                hotkey_tuple_for_manager = tuple(sorted(hotkey_combination_pynput, key=str)) if hotkey_combination_pynput else None
                
                sound_button = SoundButton(self.scrollable_frame, self, file_path, hotkey_tuple_for_manager)
                self.sound_buttons.append(sound_button)
                sound_button.pack(fill=X, padx=5, pady=5)
                
                if hotkey_tuple_for_manager:
                    # Directly add to manager, handling potential existing assignments
                    self.key_bind_manager.assign_hotkey_to_sound(sound_button, hotkey_tuple_for_manager)
            else:
                print(f"Warning: Sound file not found and skipped: {file_path}")
        
        self.canvas.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

        # Set global stop hotkey if exists
        global_hotkey_data = self.settings.get("global_stop_hotkey")
        if global_hotkey_data:
            global_hotkey_combo_pynput = []
            for key_str in global_hotkey_data:
                if key_str.startswith('Key.'):
                    try:
                        global_hotkey_combo_pynput.append(getattr(keyboard.Key, key_str.split('.')[-1].lower()))
                    except AttributeError:
                        print(f"Warning: Unknown pynput.Key '{key_str}' for global hotkey.")
                        pass
                elif key_str.startswith('VK_'):
                    try:
                        vk_code = int(key_str[3:])
                        global_hotkey_combo_pynput.append(keyboard.KeyCode.from_vk(vk_code))
                    except ValueError:
                        print(f"Warning: Invalid VK code format '{key_str}' for global hotkey.")
                        pass
                elif len(key_str) == 1:
                    global_hotkey_combo_pynput.append(keyboard.KeyCode(char=key_str))
                else:
                    print(f"Warning: Unknown global hotkey key format '{key_str}'.")
            
            if global_hotkey_combo_pynput:
                self.key_bind_manager.set_global_hotkey(global_hotkey_combo_pynput, self.stop_all_hotkey_var, self.stop_all_sounds)
        
        self.stop_all_hotkey_var.set(get_hotkey_display_string(self.key_bind_manager.global_stop_hotkey_combination) if self.key_bind_manager.global_stop_hotkey_combination else "None")


    def save_settings(self):
        settings_data = {
            "output_device": self.selected_output_device.get(),
            "input_device": self.selected_input_device.get(),
            "loopback_enabled": self.loopback_var.get(),
            "mic_passthrough_enabled": self.mic_passthrough_var.get(),
            "monitor_device": self.selected_monitor_device.get(),
            "monitor_output_enabled": self.monitor_output_var.get(),
            "sounds": [],
            "global_stop_hotkey": []
        }
        for sound_button in self.sound_buttons:
            hotkey_combo_serializable = []
            if sound_button.hotkey_combination:
                for key in sound_button.hotkey_combination:
                    if isinstance(key, keyboard.Key):
                        hotkey_combo_serializable.append(f"Key.{str(key).split('.')[-1]}")
                    elif isinstance(key, keyboard.KeyCode):
                        if key.char:
                            hotkey_combo_serializable.append(key.char)
                        elif key.vk is not None:
                            hotkey_combo_serializable.append(f"VK_{key.vk}") # Custom format for VK codes
            
            settings_data["sounds"].append({
                "file_path": sound_button.sound_file_path,
                "hotkey_combination": hotkey_combo_serializable
            })
        
        if self.key_bind_manager.global_stop_hotkey_combination:
            global_hotkey_serializable = []
            for key in self.key_bind_manager.global_stop_hotkey_combination:
                if isinstance(key, keyboard.Key):
                    global_hotkey_serializable.append(f"Key.{str(key).split('.')[-1]}")
                elif isinstance(key, keyboard.KeyCode):
                    if key.char:
                        global_hotkey_serializable.append(key.char)
                    elif key.vk is not None:
                        global_hotkey_serializable.append(f"VK_{key.vk}")
            settings_data["global_stop_hotkey"] = global_hotkey_serializable

        with open("settings.json", "w") as f:
            json.dump(settings_data, f, indent=4)
        print("Settings saved.")

    def load_settings(self):
        try:
            with open("settings.json", "r") as f:
                self.settings = json.load(f)
        except FileNotFoundError:
            self.settings = {}
        except json.JSONDecodeError:
            print("Error decoding settings.json, starting with default settings.")
            self.settings = {}

    def set_stop_all_hotkey(self):
        # No more messagebox.showinfo here
        self.key_bind_manager.start_key_assignment_mode(self._on_global_hotkey_assigned)

    def _on_global_hotkey_assigned(self, hotkey_combination_tuple):
        self.after(0, lambda assigned_keys=hotkey_combination_tuple: self.key_bind_manager.set_global_hotkey(assigned_keys, self.stop_all_hotkey_var, self.stop_all_sounds))
    
    def on_closing(self):
        print("Closing application...")
        self.audio_output_manager.stop_loopback_stream()
        self.audio_output_manager.stop_mic_passthrough_stream()
        self.audio_output_manager.stop_monitor_stream()
        
        self.sound_queue.put((None, None)) # Signal playback thread to stop
        self.playback_thread.join(timeout=2)

        self.key_bind_manager.stop_listener()

        if pygame.mixer.get_init():
            pygame.mixer.quit()
        print("Pygame mixer quit.")

        self.save_settings()
        self.destroy() # Close the Tkinter window


# Main application entry point
if __name__ == "__main__":
    try:
        pygame.mixer.init()
        print(f"Pygame {pygame.version.ver} (SDL {pygame.get_sdl_version()[0]}.{pygame.get_sdl_version()[1]}.{pygame.get_sdl_version()[2]}, Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro})")
        print("Hello from the pygame community. https://www.pygame.org/contribute.html")
    except pygame.error as e:
        messagebox.showerror("Pygame Initialization Error", f"Could not initialize Pygame mixer: {e}\nAudio playback may not work.")
        sys.exit(1)

    app = SoundboardApp()
    app.mainloop()