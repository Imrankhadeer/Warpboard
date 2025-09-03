import unittest
import os
import shutil
import uuid
import wave
import numpy as np
from Warpboard import SoundManager, SOUNDS_DIR, CONFIG_FILE, SAMPLE_RATE, CHANNELS

class TestSoundManager(unittest.TestCase):
    def setUp(self):
        self.sound_manager = SoundManager()
        os.makedirs(SOUNDS_DIR, exist_ok=True)
        # Create a dummy WAV file
        self.sound_file = os.path.join(SOUNDS_DIR, "test_sound.wav")
        with wave.open(self.sound_file, "w") as f:
            f.setnchannels(CHANNELS)
            f.setsampwidth(2)
            f.setframerate(SAMPLE_RATE)
            f.writeframes(np.zeros(SAMPLE_RATE, dtype=np.int16).tobytes())

    def tearDown(self):
        if os.path.exists(SOUNDS_DIR):
            shutil.rmtree(SOUNDS_DIR)
        if os.path.exists(CONFIG_FILE):
            os.remove(CONFIG_FILE)

    def test_rename_sound_case_insensitive(self):
        sound = self.sound_manager.add_sound(self.sound_file, custom_name="test_sound")
        sound_id = sound["id"]

        # Rename with different capitalization
        new_name = "Test_Sound"
        self.sound_manager.rename_sound(sound_id, new_name)

        # Verify that the sound was renamed
        renamed_sound = self.sound_manager.get_sound_by_id(sound_id)
        self.assertEqual(renamed_sound["name"], new_name)
        self.assertTrue(os.path.exists(renamed_sound["path"]))

if __name__ == "__main__":
    unittest.main()
