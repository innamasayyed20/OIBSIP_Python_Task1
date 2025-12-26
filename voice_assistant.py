"""
Voice Assistant with GUI - Using Windows SAPI for Speech
Reliable speech output on Windows
"""

import speech_recognition as sr
import datetime
import webbrowser
import threading
import tkinter as tk
from tkinter import scrolledtext
import time
import os

# Try to import win32com for Windows speech, fallback to pyttsx3
try:
    import win32com.client
    USE_WIN32 = True
    print("Using Windows SAPI for speech")
except ImportError:
    import pyttsx3
    USE_WIN32 = False
    print("Using pyttsx3 for speech")

class VoiceAssistantGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("NOVA Voice Assistant")
        self.root.geometry("950x750")
        self.root.configure(bg="#fdf2f8")
        
        # Center window on screen
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
        # Initialize speech components
        self.recognizer = sr.Recognizer()
        self.is_listening = False
        self.use_win32 = USE_WIN32
        
        self.setup_ui()
        
        # Welcome message
        self.root.after(500, lambda: self.display_and_speak("Hello! I'm your voice assistant. Click 'Start Listening' to begin!", "assistant"))
    
    def speak(self, text):
        """Text-to-speech using Windows SAPI or pyttsx3"""
        def speak_thread():
            try:
                print(f"🔊 Speaking: {text[:50]}...")
                if self.use_win32:
                    # Initialize COM for this thread
                    import pythoncom
                    pythoncom.CoInitialize()
                    try:
                        # Use Windows SAPI - much more reliable
                        speaker = win32com.client.Dispatch("SAPI.SpVoice")
                        speaker.Rate = 1  # Speed: -10 to 10
                        speaker.Volume = 100  # Volume: 0 to 100
                        speaker.Speak(text)
                    finally:
                        pythoncom.CoUninitialize()
                else:
                    # Fallback to pyttsx3
                    engine = pyttsx3.init()
                    engine.setProperty('rate', 170)
                    engine.setProperty('volume', 1.0)
                    engine.say(text)
                    engine.runAndWait()
                    engine.stop()
                print(f"✅ Finished speaking")
            except Exception as e:
                print(f"❌ Speech error: {e}")
        
        thread = threading.Thread(target=speak_thread, daemon=True)
        thread.start()
    
    def setup_ui(self):
        """Create the GUI interface with modern gradient design"""

        # Header with gradient effect
        header_frame = tk.Frame(self.root, bg="#fce7f3", height=140)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)

        # Title with glow effect
        title = tk.Label(
            header_frame,
            text="NOVA VOICE ASSISTANT",
            font=("Arial", 32, "bold"),
            fg="#ec4899",
            bg="#fce7f3"
        )
        title.pack(pady=(25, 5))

        subtitle = tk.Label(
            header_frame,
            text="Intelligent • Fast • Natural Communication",
            font=("Arial", 11),
            fg="#f472b6",
            bg="#fce7f3"
        )
        subtitle.pack()

        # Status with animated dot
        status_container = tk.Frame(self.root, bg="#fdf2f8")
        status_container.pack(pady=15)

        self.status_label = tk.Label(
            status_container,
            text="● Ready to Assist",
            font=("Arial", 14, "bold"),
            fg="#ec4899",
            bg="#fdf2f8"
        )
        self.status_label.pack()

        # Chat container with card design
        chat_container = tk.Frame(self.root, bg="#fdf2f8")
        chat_container.pack(padx=30, pady=10, fill=tk.BOTH, expand=True)

        # Chat card with shadow effect
        chat_card = tk.Frame(chat_container, bg="#ffffff", highlightthickness=2, highlightbackground="#fbcfe8")
        chat_card.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.chat_display = scrolledtext.ScrolledText(
            chat_card,
            wrap=tk.WORD,
            font=("Consolas", 11),
            bg="#ffffff",
            fg="#4a4a4a",
            insertbackground="#ec4899",
            relief=tk.FLAT,
            padx=20,
            pady=20,
            height=15,
            selectbackground="#fce7f3",
            selectforeground="#ec4899"
        )
        self.chat_display.pack(fill=tk.BOTH, expand=True, padx=3, pady=3)
        self.chat_display.config(state=tk.DISABLED)

        # Enhanced text styles with emojis
        self.chat_display.tag_config(
            "user",
            foreground="#ec4899",
            font=("Consolas", 11, "bold")
        )
        self.chat_display.tag_config(
            "assistant",
            foreground="#f472b6",
            font=("Consolas", 11)
        )
        self.chat_display.tag_config(
            "timestamp",
            foreground="#d1a3c4",
            font=("Consolas", 9)
        )

        # Button panel with modern styling
        btn_panel = tk.Frame(self.root, bg="#fdf2f8")
        btn_panel.pack(pady=25)

        # Main action button with gradient simulation
        self.listen_btn = tk.Button(
            btn_panel,
            text="🎤 Start Listening",
            command=self.toggle_listening,
            font=("Arial", 14, "bold"),
            bg="#ec4899",
            fg="white",
            activebackground="#db2777",
            activeforeground="white",
            relief=tk.FLAT,
            padx=35,
            pady=16,
            cursor="hand2",
            borderwidth=0
        )
        self.listen_btn.pack(side=tk.LEFT, padx=10)

        # Secondary buttons with vibrant colors
        help_btn = tk.Button(
            btn_panel,
            text="💡 Help",
            command=self.show_help,
            font=("Arial", 12, "bold"),
            bg="#d946ef",
            fg="white",
            activebackground="#c026d3",
            activeforeground="white",
            relief=tk.FLAT,
            padx=25,
            pady=16,
            cursor="hand2",
            borderwidth=0
        )
        help_btn.pack(side=tk.LEFT, padx=10)

        clear_btn = tk.Button(
            btn_panel,
            text="🗑️ Clear",
            command=self.clear_chat,
            font=("Arial", 12, "bold"),
            bg="#fb7185",
            fg="white",
            activebackground="#f43f5e",
            activeforeground="white",
            relief=tk.FLAT,
            padx=25,
            pady=16,
            cursor="hand2",
            borderwidth=0
        )
        clear_btn.pack(side=tk.LEFT, padx=10)

        # Footer with tips
        footer_frame = tk.Frame(self.root, bg="#fdf2f8")
        footer_frame.pack(pady=(10, 20))

        footer = tk.Label(
            footer_frame,
            text="💡 Quick Tips: Say 'What time is it?' • 'Search for Python' • 'What's the date?' • 'Help'",
            font=("Arial", 9),
            fg="#d1a3c4",
            bg="#fdf2f8"
        )
        footer.pack()

    
    def display_message(self, message, sender="assistant"):
        """Display message in chat window"""
        self.chat_display.config(state=tk.NORMAL)
        
        timestamp = datetime.datetime.now().strftime("%I:%M %p")
        
        if sender == "user":
            self.chat_display.insert(tk.END, f"\n[{timestamp}] ", "timestamp")
            self.chat_display.insert(tk.END, "👤 You: ", "user")
            self.chat_display.insert(tk.END, f"{message}\n")
        else:
            self.chat_display.insert(tk.END, f"\n[{timestamp}] ", "timestamp")
            self.chat_display.insert(tk.END, "🤖 Assistant: ", "assistant")
            self.chat_display.insert(tk.END, f"{message}\n")
        
        self.chat_display.see(tk.END)
        self.chat_display.config(state=tk.DISABLED)
    
    def display_and_speak(self, message, sender="assistant"):
        """Display message and speak it"""
        self.display_message(message, sender)
        if sender == "assistant":
            self.speak(message)
    
    def update_status(self, status, color):
        """Update status indicator"""
        self.status_label.config(text=f"● {status}", fg=color)
    
    def toggle_listening(self):
        """Start or stop listening"""
        if not self.is_listening:
            self.is_listening = True
            self.listen_btn.config(text="⏹️ Stop Listening", bg="#fb7185", activebackground="#f43f5e")
            self.update_status("Listening...", "#f472b6")
            threading.Thread(target=self.listen_continuous, daemon=True).start()
        else:
            self.is_listening = False
            self.listen_btn.config(text="🎤 Start Listening", bg="#ec4899", activebackground="#db2777")
            self.update_status("Ready to Assist", "#ec4899")
    
    def listen_continuous(self):
        while self.is_listening:
            command = self.listen()
            if command:
                self.process_command(command)
            else:
                time.sleep(0.5)  # wait quietly before listening again

    
    def listen(self):
        """Listen to voice input"""
        try:
            with sr.Microphone() as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=0.3)
                audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=10)
                
                self.update_status("Processing...", "#d946ef")
                command = self.recognizer.recognize_google(audio).lower()
                self.display_message(command, "user")
                return command
                
        except sr.WaitTimeoutError:
            return None
        except sr.UnknownValueError:
            # User didn't say anything clearly → stay silent and keep waiting
            self.update_status("Listening...", "#f472b6")
            return None

        except sr.RequestError:
            self.display_and_speak("Speech recognition service is unavailable.")
            self.is_listening = False
            self.listen_btn.config(text="🎤 Start Listening", bg="#ec4899", activebackground="#db2777")
            self.update_status("Error", "#fb7185")
            return None
        except Exception as e:
            self.display_and_speak(f"An error occurred: {str(e)}")
            return None
    
    def process_command(self, command):
        """Process voice commands"""
        # Greeting
        if any(word in command for word in ["hello", "hi", "hey"]):
            self.display_and_speak("Hello! How can I help you today?")
        
        # Exit
        elif any(word in command for word in ["exit", "quit", "bye", "goodbye", "stop"]):
            self.display_and_speak("Goodbye! Have a great day!")
            self.is_listening = False
            self.listen_btn.config(text="🎤 Start Listening", bg="#ec4899", activebackground="#db2777")
            self.update_status("Ready to Assist", "#ec4899")
        
        # Time
        elif "time" in command:
            current_time = datetime.datetime.now().strftime("%I:%M %p")
            response = f"The current time is {current_time}"
            self.display_and_speak(response)
        
        # Date
        elif "date" in command or "today" in command:
            now = datetime.datetime.now()
            current_date = now.strftime("%B %d, %Y")
            day = now.strftime("%A")
            response = f"Today is {day}, {current_date}"
            self.display_and_speak(response)
        
        # Search
        elif "search" in command:
            query = command.replace("search for", "").replace("search", "").strip()
            url = f"https://www.google.com/search?q={query}"
            webbrowser.open(url)
            self.display_and_speak(f"Searching for {query} on Google")
        
        # Name
        elif "your name" in command or "who are you" in command:
            self.display_and_speak("I am your AI voice assistant, here to help you!")
        
        # Help
        elif "help" in command or "what can you do" in command:
            self.show_help()
        
        # Unknown
        else:
            self.display_and_speak("I'm not sure how to help with that. Say 'help' to see what I can do.")
        
        self.update_status("Listening...", "#f472b6")
    
    def show_help(self):
        """Show help information"""
        help_text = """I can help you with: Tell the time, Tell the date, Search the web, Greet you, and Exit. Just say what you need!"""
        self.display_and_speak(help_text)
    
    def clear_chat(self):
        """Clear chat display"""
        self.chat_display.config(state=tk.NORMAL)
        self.chat_display.delete(1.0, tk.END)
        self.chat_display.config(state=tk.DISABLED)
        self.display_and_speak("Chat cleared!")

def main():
    root = tk.Tk()
    app = VoiceAssistantGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()