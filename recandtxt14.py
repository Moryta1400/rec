
# admin_check.py
import os
import sys
import ctypes

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    # Re-run the program with admin rights
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    sys.exit()

# recandtxt11.py
import os
import sys
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import pyaudio
import wave
import numpy as np
from datetime import datetime
from pydub import AudioSegment
import speech_recognition as sr
import ctypes
import subprocess

def get_internal_path():
    """Get path to _internal directory"""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    internal_path = os.path.join(base_path, '_internal')
    if not os.path.exists(internal_path):
        try:
            os.makedirs(internal_path)
        except PermissionError:
            messagebox.showerror(
                "Error",
                "Permission denied. Please run the program as administrator."
            )
            sys.exit(1)
    return internal_path

# تنظیم مسیر برای فایل‌های خارجی
INTERNAL_PATH = get_internal_path()
os.environ['PATH'] = f"{INTERNAL_PATH};{os.environ['PATH']}"

# تنظیم مسیر ffmpeg
ffmpeg_path = os.path.join(INTERNAL_PATH, "ffmpeg.exe")
AudioSegment.converter = ffmpeg_path

# Check for admin rights
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    if sys.platform.startswith('win'):
        try:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            sys.exit()
        except:
            messagebox.showerror("Error", "This program requires administrator privileges.")
            sys.exit(1)
class AudioTranscriber:
    def __init__(self, root, auto_file=None):
        self.root = tk.Tk() if root is None else root
        self.root.title("Transcriber")
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
        
        # Icon setup if available
        try:
            icon_path = os.path.join(INTERNAL_PATH, "app.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except:
            pass
            
        self.root.protocol("WM_DELETE_WINDOW", self.exit_application)
        
        self.center_container = ttk.Frame(self.root)
        self.center_container.place(relx=0.5, rely=0.5, anchor='center')
        
        self.setup_ui()
        self.running = False
        self.current_temp_file = None
        
        if auto_file:
            self.file_path = auto_file
            self.file_label.config(text=os.path.basename(auto_file))
            self.language_var.set("fa-IR")
            self.start_button.config(state=tk.NORMAL)
            self.root.after(100, self.start_transcription)

    def setup_ui(self):
        style = ttk.Style()
        style.configure('Custom.TRadiobutton', font=('Tahoma', 10))
        
        main_frame = ttk.Frame(self.center_container)
        main_frame.pack(expand=True)
        
        title_label = tk.Label(
            main_frame,
            text="Audio Transcriber",
            font=('Tahoma', 16, 'bold')
        )
        title_label.pack(pady=(0, 20))
        
        self.file_button = ttk.Button(
            main_frame, 
            text="Select Audio File",
            command=self.select_file,
            width=25
        )
        self.file_button.pack(pady=(0, 10))
        
        self.file_label = tk.Label(
            main_frame,
            text="No file selected",
            font=('Tahoma', 10)
        )
        self.file_label.pack(pady=(0, 20))
        
        language_frame = ttk.Frame(main_frame)
        language_frame.pack(pady=(0, 20))
        
        lang_label = tk.Label(
            language_frame, 
            text="Select Language:",
            font=('Tahoma', 10)
        )
        lang_label.pack(side=tk.LEFT, padx=5)
        
        self.language_var = tk.StringVar(value="fa-IR")
        ttk.Radiobutton(
            language_frame,
            text="Persian",
            variable=self.language_var,
            value="fa-IR",
            style='Custom.TRadiobutton'
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Radiobutton(
            language_frame,
            text="English",
            variable=self.language_var,
            value="en-US",
            style='Custom.TRadiobutton'
        ).pack(side=tk.LEFT, padx=5)
        
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(pady=(0, 10))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            length=500,
            mode='determinate',
            variable=self.progress_var
        )
        self.progress_bar.pack(pady=(0, 10))
        
        self.status_label = tk.Label(
            progress_frame,
            text="",
            font=('Tahoma', 10)
        )
        self.status_label.pack(pady=(0, 5))
        
        self.eta_label = tk.Label(
            progress_frame,
            text="",
            font=('Tahoma', 10)
        )
        self.eta_label.pack(pady=(0, 20))
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(0, 20))
        
        # Open folder button
        self.open_folder_btn = ttk.Button(
            button_frame,
            text="Open Location",
            command=self.open_file_location,
            width=20
        )
        self.open_folder_btn.pack(side=tk.LEFT, padx=5)
        
        self.exit_button = ttk.Button(
            button_frame,
            text="Exit",
            command=self.exit_application,
            width=20
        )
        self.exit_button.pack(side=tk.LEFT, padx=5)
        
        self.start_button = ttk.Button(
            button_frame,
            text="Start Transcription",
            command=self.start_transcription,
            state=tk.DISABLED,
            width=20
        )
        self.start_button.pack(side=tk.LEFT, padx=5)

    def open_file_location(self):
        """Open file location in explorer"""
        if hasattr(self, 'file_path') and os.path.exists(os.path.dirname(self.file_path)):
            os.startfile(os.path.dirname(self.file_path))

    def select_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Audio Files", "*.mp3 *.wav *.m4a *.ogg")]
        )
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.start_button.config(state=tk.NORMAL)

    def clean_temp_files(self):
        try:
            temp_dir = os.path.join(INTERNAL_PATH, 'temp')
            if os.path.exists(temp_dir):
                for file in os.listdir(temp_dir):
                    if file.startswith("temp_chunk_") and file.endswith(".wav"):
                        file_path = os.path.join(temp_dir, file)
                        for attempt in range(3):
                            try:
                                if os.path.exists(file_path):
                                    os.remove(file_path)
                                    break
                            except:
                                time.sleep(0.2)
        except Exception as e:
            print(f"Error in clean_temp_files: {str(e)}")

    def exit_application(self):
        try:
            self.running = False
            time.sleep(0.2)
            self.clean_temp_files()
            self.root.quit()
            self.root.destroy()
        except Exception as e:
            print(f"Error during exit: {str(e)}")
            self.root.destroy()
    def update_progress(self, current, total):
        if not self.running:
            return
            
        progress = (current / total) * 100
        self.progress_var.set(progress)
        
        elapsed = datetime.now() - self.start_time
        elapsed_seconds = elapsed.total_seconds()
        if current > 0:
            eta_seconds = elapsed_seconds * (total - current) / current
            self.eta_label.config(text=f"Estimated time remaining: {int(eta_seconds)} seconds")
        
        self.status_label.config(text=f"Processing: {progress:.1f}%")
        self.root.update_idletasks()

    def transcribe(self):
        try:
            # Create temp directory if it doesn't exist
            temp_dir = os.path.join(INTERNAL_PATH, 'temp')
            if not os.path.exists(temp_dir):
                os.makedirs(temp_dir)

            self.start_time = datetime.now()
            self.status_label.config(text="Loading audio file...")
            audio = AudioSegment.from_file(self.file_path)
            
            self.status_label.config(text="Preparing audio...")
            if audio.channels > 1:
                audio = audio.set_channels(1)
            audio = audio.set_frame_rate(16000).normalize()

            recognizer = sr.Recognizer()
            recognizer.energy_threshold = 300
            
            chunks = [audio[i:i + 30000] for i in range(0, len(audio), 30000)]
            full_text = []
            error_count = 0
            
            for i, chunk in enumerate(chunks):
                if not self.running:
                    break
                    
                temp_wav = os.path.join(temp_dir, f"temp_chunk_{i}.wav")
                self.current_temp_file = temp_wav
                
                try:
                    chunk.export(temp_wav, format="wav")
                    
                    with sr.AudioFile(temp_wav) as source:
                        audio_data = recognizer.record(source)
                    
                    time.sleep(0.1)
                    
                    text = recognizer.recognize_google(
                        audio_data,
                        language=self.language_var.get()
                    )
                    full_text.append(text)
                    
                except Exception as e:
                    error_count += 1
                    full_text.append("")
                    print(f"Error in chunk {i}: {str(e)}")
                
                finally:
                    try:
                        if os.path.exists(temp_wav):
                            os.remove(temp_wav)
                    except:
                        pass
                    
                    self.current_temp_file = None
                
                self.update_progress(i + 1, len(chunks))

            if self.running:
                output_path = os.path.splitext(self.file_path)[0] + ".txt"
                with open(output_path, "w", encoding="utf-8") as f:
                    f.write(' '.join(full_text))
                
                self.status_label.config(text=f"Transcription completed. Errors: {error_count}")
                messagebox.showinfo("Complete", f"Text file saved to:\n{output_path}")
                self.root.after(500, self.exit_application)
            
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}")
            messagebox.showerror("Error", str(e))
            self.root.after(500, self.exit_application)
        
        finally:
            time.sleep(0.2)
            self.clean_temp_files()
            self.stop_transcription()

    def start_transcription(self):
        if not hasattr(self, 'file_path'):
            messagebox.showwarning("Warning", "Please select an audio file first.")
            return
            
        self.running = True
        self.start_button.config(state=tk.DISABLED, text="Processing...")
        self.file_button.config(state=tk.DISABLED)
        self.progress_var.set(0)
        self.eta_label.config(text="Calculating remaining time...")
        
        threading.Thread(target=self.transcribe, daemon=True).start()

    def stop_transcription(self):
        self.running = False
        time.sleep(0.2)
        self.clean_temp_files()
        self.start_button.config(state=tk.NORMAL, text="Start Transcription")
        self.file_button.config(state=tk.NORMAL)
        self.status_label.config(text="Operation stopped")

class AudioRecorderApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Audio Recorder & Transcriber")
        
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        window_width = 500
        window_height = 350
        x = screen_width - window_width - 20
        y = screen_height - window_height - 60
        
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.attributes('-topmost', True)
        
        # Icon setup if available
        try:
            icon_path = os.path.join(INTERNAL_PATH, "app.ico")
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except:
            pass
            
        self.recording = False
        self.paused = False
        self.frames = []
        self.temp_frames = []
        
        self.save_interval = 5 * 60
        self.last_save_time = time.time()
        self.current_audio_file = None
        
        self.save_directory = os.path.expanduser("~/Documents")
        
        self.CHUNK = 1024 * 4
        self.FORMAT = pyaudio.paInt16
        self.CHANNELS = 1
        self.RATE = 44100
        self.loopback_device_index = 1
        
        self.protocol("WM_DELETE_WINDOW", self.exit_program)
        
        self.setup_gui()
    def setup_gui(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        creator_frame = ttk.Frame(main_frame)
        creator_frame.pack(fill=tk.X, pady=(0, 10))
        
        creator_label = ttk.Label(creator_frame, 
                                text="Created By M.Taghizadeh",
                                font=('Tahoma', 10, 'bold'))
        creator_label.pack(anchor='center')
        
        email_label = ttk.Label(creator_frame,
                             text="M.Taghizadeh@ymail.com",
                             font=('Tahoma', 9))
        email_label.pack(anchor='center')

        # Save Location Frame with Open button
        dir_frame = ttk.LabelFrame(main_frame, text="Save Location", padding="5")
        dir_frame.pack(fill=tk.X, pady=(0, 10))

        self.dir_label = ttk.Label(dir_frame, text=self.save_directory, 
                                 font=('Tahoma', 8))
        self.dir_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        dir_buttons_frame = ttk.Frame(dir_frame)
        dir_buttons_frame.pack(side=tk.RIGHT)

        self.open_btn = ttk.Button(dir_buttons_frame, text="Open", 
                                 command=self.open_save_location,
                                 width=8)
        self.open_btn.pack(side=tk.LEFT, padx=(0, 5))

        dir_button = ttk.Button(dir_buttons_frame, text="Browse", 
                              command=self.select_directory,
                              width=8)
        dir_button.pack(side=tk.LEFT)
        
        self.status_var = tk.StringVar(value="Ready to record")
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var)
        self.status_label.pack(pady=(0, 10))
        
        # Top button frame (for Record and Pause)
        top_button_frame = ttk.Frame(main_frame)
        top_button_frame.pack(fill=tk.X, pady=5)
        
        # Bottom button frame (for End and Convert)
        bottom_button_frame = ttk.Frame(main_frame)
        bottom_button_frame.pack(fill=tk.X, pady=5)
        
        # Button width
        button_width = 15
        
        # Record and Pause buttons in top frame
        self.record_btn = ttk.Button(top_button_frame, text="Start Recording", 
                                   command=self.start_recording,
                                   width=button_width)
        self.record_btn.pack(side=tk.LEFT, padx=5, expand=True)
        
        self.pause_btn = ttk.Button(top_button_frame, text="Pause", 
                                  command=self.pause_recording,
                                  state='disabled',
                                  width=button_width)
        self.pause_btn.pack(side=tk.LEFT, padx=5, expand=True)
        
        # End and Convert buttons in bottom frame
        self.stop_btn = ttk.Button(bottom_button_frame, text="End Recording", 
                                 command=self.stop_recording,
                                 state='disabled',
                                 width=button_width)
        self.stop_btn.pack(side=tk.LEFT, padx=5, expand=True)
        
        self.transcribe_btn = ttk.Button(bottom_button_frame, text="Convert to Text", 
                                       command=self.transcribe_audio,
                                       state='disabled',
                                       width=button_width)
        self.transcribe_btn.pack(side=tk.LEFT, padx=5, expand=True)
        
        # Exit button and version frame
        exit_frame = ttk.Frame(main_frame)
        exit_frame.pack(pady=(10, 0), fill=tk.X)
        
        self.exit_btn = ttk.Button(exit_frame, text="Exit", 
                                 command=self.exit_program,
                                 width=button_width)
        self.exit_btn.pack(fill=tk.X)
        
        # Version label
        version_label = ttk.Label(exit_frame, 
                                text="Ver 1.4",
                                font=('Tahoma', 8))
        version_label.pack(pady=(2, 0))

    def select_directory(self):
        """Select directory for saving files"""
        directory = filedialog.askdirectory(initialdir=self.save_directory)
        if directory:
            self.save_directory = directory
            self.dir_label.config(text=self.save_directory)
            # Reset current audio file to force new file creation in new location
            self.current_audio_file = None

    def open_save_location(self):
        """Open save location in file manager"""
        try:
            if self.current_audio_file and os.path.exists(os.path.dirname(self.current_audio_file)):
                os.startfile(os.path.dirname(self.current_audio_file))
            else:
                os.startfile(self.save_directory)
        except Exception as e:
            print(f"[ERROR] Failed to open directory: {str(e)}")
            messagebox.showerror("Error", "Failed to open directory")

    def pause_recording(self):
        if self.recording:
            self.paused = not self.paused
            if self.paused:
                self.pause_btn.config(text="Resume")
                self.status_var.set("Recording paused")
            else:
                self.pause_btn.config(text="Pause")
                self.status_var.set("Recording resumed")

    def save_audio_chunk(self):
        if not self.temp_frames:
            return

        try:
            print("[INFO] Saving audio chunk...")
            temp_wav = os.path.join(INTERNAL_PATH, 'temp', 'temp_chunk.wav')
            
            os.makedirs(os.path.dirname(temp_wav), exist_ok=True)
            
            with wave.open(temp_wav, 'wb') as wf:
                wf.setnchannels(self.CHANNELS)
                wf.setsampwidth(2)
                wf.setframerate(self.RATE)
                wf.writeframes(b''.join(self.temp_frames))
            
            if os.path.exists(self.current_audio_file):
                existing_audio = AudioSegment.from_mp3(self.current_audio_file)
                new_chunk = AudioSegment.from_wav(temp_wav)
                combined = existing_audio + new_chunk
                combined.export(self.current_audio_file, format='mp3', bitrate='128k')
            else:
                audio = AudioSegment.from_wav(temp_wav)
                audio.export(self.current_audio_file, format='mp3', bitrate='128k')
            
            if os.path.exists(temp_wav):
                os.remove(temp_wav)
            
            print(f"[SUCCESS] Saved audio chunk to {self.current_audio_file}")
            self.status_var.set(f"Auto-saved to: {os.path.basename(self.current_audio_file)}")
            
        except Exception as e:
            print(f"[ERROR] Failed to save audio chunk: {str(e)}")

    def start_recording(self):
        if self.recording:
            return

        try:
            print("[INFO] Starting recording...")
            self.p = pyaudio.PyAudio()

            # Create new audio file for each recording
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            self.current_audio_file = os.path.join(self.save_directory, 
                                                 f'recording_{timestamp}.mp3')
            
            def audio_callback(in_data, frame_count, time_info, status):
                if self.recording and not self.paused:
                    audio_data = np.frombuffer(in_data, dtype=np.int16)
                    self.frames.append(audio_data.tobytes())
                    self.temp_frames.append(audio_data.tobytes())
                    
                    current_time = time.time()
                    if current_time - self.last_save_time >= self.save_interval:
                        self.save_audio_chunk()
                        self.temp_frames = []
                        self.last_save_time = current_time
                        
                return (in_data, pyaudio.paContinue)

            self.stream = self.p.open(
                format=self.FORMAT,
                channels=self.CHANNELS,
                rate=self.RATE,
                input=True,
                input_device_index=self.loopback_device_index,
                frames_per_buffer=self.CHUNK,
                stream_callback=audio_callback
            )

            self.recording = True
            self.paused = False
            self.stream.start_stream()
            self.last_save_time = time.time()
            
            # Update UI
            self.record_btn['state'] = 'disabled'
            self.pause_btn['state'] = 'normal'
            self.stop_btn['state'] = 'normal'
            self.transcribe_btn['state'] = 'disabled'
            self.status_var.set("Recording in progress...")
            
            print("[SUCCESS] Recording started")

        except Exception as e:
            print(f"[ERROR] Failed to start recording: {str(e)}")
            self.status_var.set("Recording failed to start!")
            messagebox.showerror("Error", f"Failed to start recording: {str(e)}")

    def stop_recording(self):
        if not self.recording:
            return

        try:
            print("[INFO] Stopping recording...")
            self.recording = False
            self.paused = False
            
            self.stream.stop_stream()
            self.stream.close()
            self.p.terminate()
            
            if self.temp_frames:
                self.save_audio_chunk()
            
            self.frames = []
            self.temp_frames = []
            
            # Update UI
            self.record_btn['state'] = 'normal'
            self.pause_btn['state'] = 'disabled'
            self.stop_btn['state'] = 'disabled'
            self.transcribe_btn['state'] = 'normal'
            self.pause_btn.config(text="Pause")
            self.status_var.set(f"Recording saved: {os.path.basename(self.current_audio_file)}")
            
            print(f"[SUCCESS] Recording saved: {self.current_audio_file}")
            
        except Exception as e:
            print(f"[ERROR] Failed to stop recording: {str(e)}")
            self.status_var.set("Failed to save recording!")

    def transcribe_audio(self):
        if self.current_audio_file and os.path.exists(self.current_audio_file):
            try:
                print("[INFO] Starting transcription...")
                self.transcribe_btn['state'] = 'disabled'
                transcriber = AudioTranscriber(None, auto_file=self.current_audio_file)
                transcriber.root.mainloop()
                self.transcribe_btn['state'] = 'normal'
            except Exception as e:
                print(f"[ERROR] Failed to start transcription: {str(e)}")
                self.status_var.set("Transcription failed!")
                self.transcribe_btn['state'] = 'normal'
                messagebox.showerror("Error", f"Transcription failed: {str(e)}")
        else:
            print("[ERROR] No recording file found")
            self.status_var.set("No recording file found!")
            messagebox.showerror("Error", "Audio file not found!")

    def exit_program(self):
        try:
            if self.recording:
                print("[INFO] Saving final recording before exit...")
                self.status_var.set("Saving recording before exit...")
                self.update_idletasks()
                self.stop_recording()
                time.sleep(0.5)
            
            temp_dir = os.path.join(INTERNAL_PATH, 'temp')
            if os.path.exists(temp_dir):
                for file in os.listdir(temp_dir):
                    try:
                        os.remove(os.path.join(temp_dir, file))
                    except:
                        pass
                try:
                    os.rmdir(temp_dir)
                except:
                    pass
            
            self.quit()
            self.destroy()
            
        except Exception as e:
            print(f"[ERROR] Error during cleanup: {str(e)}")
            try:
                self.destroy()
            except:
                pass
        finally:
            sys.exit(0)

def main():
    try:
        if not is_admin():
            print("[WARNING] Running without admin privileges")
            if sys.platform.startswith('win'):
                try:
                    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
                    sys.exit()
                except:
                    messagebox.showwarning(
                        "Warning",
                        "Running without administrator privileges.\nSome features might not work correctly."
                    )
        
        if not os.path.exists(INTERNAL_PATH):
            os.makedirs(INTERNAL_PATH)
            
        temp_dir = os.path.join(INTERNAL_PATH, 'temp')
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
            
        if not os.path.exists(ffmpeg_path):
            messagebox.showerror(
                "Error", 
                "ffmpeg.exe not found in _internal directory.\n"
                "Please make sure ffmpeg.exe is placed in the _internal folder."
            )
            return
            
        app = AudioRecorderApp()
        app.mainloop()
        
    except Exception as e:
        print(f"[ERROR] Application error: {str(e)}")
        messagebox.showerror(
            "Error", 
            f"Application error: {str(e)}\n"
            "Please check the error log for details."
        )
    finally:
        try:
            temp_dir = os.path.join(INTERNAL_PATH, 'temp')
            if os.path.exists(temp_dir):
                for file in os.listdir(temp_dir):
                    try:
                        os.remove(os.path.join(temp_dir, file))
                    except:
                        pass
                try:
                    os.rmdir(temp_dir)
                except:
                    pass
        except:
            pass

if __name__ == "__main__":
    main()
