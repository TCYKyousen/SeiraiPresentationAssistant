from PyQt6.QtMultimedia import QMediaPlayer, QAudioOutput
from PyQt6.QtCore import QUrl
import os
import win32com.client

class SoundManager:
    def __init__(self, resource_dir):
        self.resource_dir = resource_dir
        self.players = {}
        self.outputs = {}
        try:
            self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        except:
            self.speaker = None

    def play(self, sound_name, loop=False):
        """
        Play a sound effect.
        :param sound_name: The name of the sound file (e.g., 'Switch' or 'Switch.ogg')
        :param loop: Whether to loop the sound
        """
        filename = sound_name if sound_name.endswith('.ogg') else f"{sound_name}.ogg"
        
        if filename not in self.players:
            path = os.path.join(self.resource_dir, filename)
            if os.path.exists(path):
                player = QMediaPlayer()
                audio_output = QAudioOutput()
                player.setAudioOutput(audio_output)
                player.setSource(QUrl.fromLocalFile(path))
                self.players[filename] = player
                self.outputs[filename] = audio_output
            else:
                print(f"Sound file not found: {path}")
                return

        player = self.players[filename]
        if player.playbackState() == QMediaPlayer.PlaybackState.PlayingState:
            player.stop()
            
        if loop:
            player.setLoops(QMediaPlayer.Loops.Infinite)
        else:
            player.setLoops(1)
            
        player.play()

    def stop(self, sound_name):
        filename = sound_name if sound_name.endswith('.ogg') else f"{sound_name}.ogg"
        if filename in self.players:
            self.players[filename].stop()

    def speak(self, text):
        if self.speaker:
            try:
                # 1 = SVSFlagsAsync
                self.speaker.Speak(text, 1)
            except Exception as e:
                print(f"TTS Error: {e}")
