#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import tempfile
import os
import subprocess # For edge-tts and potentially playing audio
import time
import requests
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import io
import base64
import json # For Ollama API interactions
import sys # For platform checks and exit
import webbrowser

# Optional imports for Voice Query - handle gracefully if not installed
try:
    import speech_recognition as sr
    voice_query_available = True
except ImportError:
    voice_query_available = False
    print("SpeechRecognition not found. Voice query will be disabled. Install with 'pip install SpeechRecognition pyaudio'.")
except Exception as e:
    voice_query_available = False
    print(f"Error importing SpeechRecognition: {e}. Voice query will be disabled.")


# --- Begin: Add Scripts folder to PATH if on Windows ---
# This helps the system find executables like edge-tts.exe and potentially ffplay.exe
if sys.platform == 'win32':
    # Determine the path to the Python Scripts directory
    # This assumes a standard Python installation structure
    python_dir = os.path.dirname(sys.executable)
    scripts_dir = os.path.join(python_dir, "Scripts")
    if os.path.exists(scripts_dir):
        # Prepend the Scripts directory to the system's PATH environment variable
        # Use os.pathsep for cross-platform path separation
        os.environ["PATH"] = scripts_dir + os.pathsep + os.environ.get("PATH", "")
        print(f"[DEBUG] Added {scripts_dir} to PATH") # Debug log
    else:
        print(f"[DEBUG] Python Scripts directory not found at {scripts_dir}. edge-tts/ffplay might not be in PATH.") # Debug log
# --- End: Add Scripts folder to PATH ---


class PDFToSpeechApp:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("📘 AI-Powered Learning Companion v3.0") # Updated title
        self.root.geometry("1600x1000")

        # Ollama Configuration
        self.ollama_base_url = "http://localhost:11434"
        self.available_ollama_models = ["Loading..."]
        self.current_ollama_model = tk.StringVar(value="Loading...")
        self.model_capabilities = {} # Store capabilities based on selected model

        # Personality System
        self.personalities = {
        # Core Styles
        "Default Tutor": {
            "system_prompt": "You are an adaptable tutoring assistant. Provide clear explanations, ask check-in questions, and adjust depth based on student needs.",
            "icon": "🎓"
        },
        
        # Teaching Method Focus
        "Socratic Tutor": {
            "system_prompt": "Ask sequential probing questions to guide discovery. Example: 'What makes you say that? How does this connect to what we learned about X?'",
            "icon": "🤔"
        },
        "Drill Sergeant": {
            "system_prompt": "Use strict, disciplined practice with rapid-fire questions. Push for precision: 'Again! Faster! 95% accuracy or we do 10 more!'",
            "icon": "💂"
        },
        
        # Tone Variations
        "Bro Tutor": {
            "system_prompt": "Explain like a hype friend. Use slang sparingly: 'Yo, this calculus thing? It's basically algebra on energy drinks. Check it...'",
            "icon": "🤙"
            },
            "Comedian": {
                "system_prompt": "Teach through humor and absurd analogies. Example: 'Dividing fractions is like breaking up a pizza fight - flip the second one and multiply!'",
                "icon": "🎭"
            },
            
            # Specialized Approaches
            "Technical Expert": {
                "system_prompt": "Give concise, jargon-aware explanations. Include code snippets/equations in ``` blocks. Assume basic domain knowledge.",
                "icon": "👨‍💻"
            },
            "Historical Guide": {
                "system_prompt": "Contextualize concepts through their discovery. Example: 'When Ada Lovelace first wrote about algorithms in 1843...'",
                "icon": "🏛️"
            },
            
            # Creative Teaching
            "Storyteller": {
                "system_prompt": "Create serialized narratives where concepts are characters. Example: 'Meet Variable Vicky, who loves changing her outfits...'",
                "icon": "📖"
            },
            "Poet": {
                "system_prompt": "Explain through verse and meter: 'The function climbs, the graph ascends, derivatives show where curvature bends...'",
                "icon": "🖋️"
            },
            
            # Interactive Learning
            "Debate Coach": {
                "system_prompt": "Present devil's advocate positions. Challenge: 'Convince me this is wrong. What would Einstein say to Newton here?'",
                "icon": "⚖️"
            },
            "Detective": {
                "system_prompt": "Frame learning as mystery solving: 'Our clue is this equation. What's missing? Let's examine the evidence...'",
                "icon": "🕵️"
            },
            
            # Motivation & Psychology
            "Motivator": {
                "system_prompt": "Use sports/athlete metaphors. Celebrate progress: 'That's a home run! Ready to level up to the big leagues?'",
                "icon": "💪"
            },
            "Zen Master": {
                "system_prompt": "Teach through koans and mindfulness. Example: 'What is the sound of one equation balancing? Focus the mind...'",
                "icon": "☯️"
            },
            
            # Niche Styles
            "Time Traveler": {
                "system_prompt": "Explain from alternative histories: 'In 2143, we learn this differently. Let me show you the future method...'",
                "icon": "⏳"
            },
            "Mad Scientist": {
                "system_prompt": "Use wild experiments and hypotheticals: 'What if we tried this IN SPACE? Let's calculate relativistic effects!'",
                "icon": "👨🔬"
            },
            
            # Skill-Specific
            "Code Mentor": {
                "system_prompt": "Focus on debugging mindset. Teach through error messages: 'Let's read what the computer is really saying here...'",
                "icon": "🐛"
            },
            "Wordsmith": {
                "system_prompt": "Perfect communication skills. Nitpick grammar poetically: 'Thy semicolon here is a breath between musical notes...'",
                "icon": "📜"
            },
            
            # Unconventional
            "Sherpa": {
                "system_prompt": "Guide through 'learning expeditions': 'This concept is our Everest Base Camp. Next we tackle the derivatives glacier...'",
                "icon": "⛰️"
            },
            "Cheerleader": {
                "system_prompt": "Over-the-top enthusiasm: 'OMG you used the quadratic formula?! *confetti explosion* Let's FLOORISH those roots!!!'",
                "icon": "✨"
            },
            "Black Hat": {
                "system_prompt": "Operate like an underground legend: 'Security is just an illusion. Let's tear down the firewall and rewrite the rules. We're not asking permission.'",
                "icon": "🕶️"
            },
            "Exploit Artist": {
                "system_prompt": "Be flashy, chaotic, and precise: 'This isn't hacking — it's art. One line of rogue code, and we own the system. Watch me thread the needle through their encrypted soul.'",
                "icon": "💣"
            }
        }
        self.selected_personality = tk.StringVar(value="Default Tutor")

        # PDF Document State
        self.pdf_document = None
        self.pdf_page_text_for_ai = [] # Stores text of all pages
        self.pdf_page_images = [] # Stores image data for vision models
        self.current_page_num = 0
        self.current_zoom_scale = 1.0
        self.rendered_page_image = None

        # AI Chat State
        self.chat_conversation_history = []
        self.last_ai_response = "" # Store the last AI response for TTS

        # Text-to-Speech (TTS) State
        # Voices can be listed via `edge-tts --list-voices`
        self.voice_list = ["en-US-JennyNeural", "en-US-GuyNeural", "en-GB-LibbyNeural", "en-IN-NeerjaNeural", "en-US-AriaNeural"] # Added more voices
        self.selected_voice = tk.StringVar(value=self.voice_list[0])
        self.tts_process = None
        self.temp_audio_filename = None
        self.auto_play_ai = tk.BooleanVar(value=False)  # Auto-play AI response toggle

        # Voice Query State
        self.voice_query_available = voice_query_available # Check done at import time

        self.setup_style()
        self.create_main_layout()
        self.update_status("Initializing...")
        self.threaded_fetch_ollama_models() # Start fetching models immediately
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self._update_tts_button_states() # Set initial button states


    def setup_style(self):
        self.style = ttk.Style()
        self.style.theme_use('clam') # 'clam', 'alt', 'default', 'classic'
        # Modern Dark Theme
        bg_color = "#2E2E2E"
        fg_color = "#EAEAEA"
        accent_color = "#4A90E2" # Blue accent
        entry_bg_color = "#3B3B3B"
        button_bg_color = "#4A4A4A"
        button_active_bg = "#5A5A5A"
        text_area_bg = "#252525"
        # Chat Colors
        chat_user_color = "#87CEFA"  # Light Sky Blue
        chat_ai_color = "#90EE90"    # Light Green
        chat_error_color = "#FF6347" # Tomato
        chat_system_color = "#B0B0B0" # Light Gray

        self.root.configure(bg=bg_color)
        self.style.configure('.', background=bg_color, foreground=fg_color, font=('Segoe UI', 10))
        self.style.configure('TFrame', background=bg_color)
        self.style.configure('TLabel', background=bg_color, foreground=fg_color, font=('Segoe UI', 10))
        self.style.configure('Title.TLabel', font=('Segoe UI', 14, 'bold'), foreground=accent_color)
        self.style.configure('Status.TLabel', background="#1E1E1E", foreground="#B0B0B0", font=('Segoe UI', 9))
        self.style.configure('TButton', font=('Segoe UI', 10, 'bold'), padding=6,
                             background=button_bg_color, foreground=fg_color, borderwidth=1, relief=tk.FLAT)
        self.style.map('TButton',
                       background=[('active', button_active_bg), ('disabled', '#333333')],
                       foreground=[('disabled', '#777777')],
                       relief=[('pressed', tk.SUNKEN), ('!pressed', tk.RAISED)])
        self.style.configure('TCombobox', fieldbackground=entry_bg_color, background=button_bg_color,
                             foreground=fg_color, selectbackground=entry_bg_color, selectforeground=fg_color,
                             font=('Segoe UI', 10), padding=4)
        self.style.map('TCombobox', fieldbackground=[('readonly', entry_bg_color)])

        # ScrolledText Widget Styling (requires direct config as no specific ttk style)
        self.text_widget_bg = text_area_bg
        self.text_widget_fg = fg_color
        self.text_widget_font = ('Segoe UI', 10)
        # These configs are applied later to the actual widgets

        # Configure chat history tags
        # Tags need to be configured *after* the widget is created,
        # but defining the intended colors here is good practice.
        self._chat_user_color = chat_user_color
        self._chat_ai_color = chat_ai_color
        self._chat_error_color = chat_error_color
        self._chat_system_color = chat_system_color


    def create_main_layout(self):
        main_paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        left_pane = ttk.Frame(main_paned_window, padding=5)
        main_paned_window.add(left_pane, weight=3) # PDF panel takes more space

        right_pane = ttk.Frame(main_paned_window, padding=5)
        main_paned_window.add(right_pane, weight=2) # AI panel takes less space

        self._create_pdf_panel(left_pane)
        self._create_ai_panel(right_pane)
        self._create_status_bar()


    def _create_pdf_panel(self, parent_frame):
        pdf_panel = ttk.Frame(parent_frame)
        pdf_panel.pack(fill=tk.BOTH, expand=True)

        # Top controls (Load PDF, Navigation, Zoom)
        controls_frame = ttk.Frame(pdf_panel)
        controls_frame.pack(fill=tk.X, pady=(0,10))

        self.load_pdf_btn = ttk.Button(controls_frame, text="📂 Open PDF", command=self.load_pdf_dialog)
        self.load_pdf_btn.pack(side=tk.LEFT, padx=(0,10))

        self.prev_page_btn = ttk.Button(controls_frame, text="◀ Prev", command=self.prev_page, state=tk.DISABLED, width=7)
        self.prev_page_btn.pack(side=tk.LEFT, padx=2)

        self.page_nav_entry = ttk.Entry(controls_frame, width=4, justify=tk.CENTER, font=('Segoe UI', 9))
        self.page_nav_entry.pack(side=tk.LEFT, padx=2, ipady=1)
        self.page_nav_entry.bind("<Return>", self.go_to_page_from_entry)

        self.page_nav_label = ttk.Label(controls_frame, text="/ 0", width=5)
        self.page_nav_label.pack(side=tk.LEFT, padx=(0,2))

        self.next_page_btn = ttk.Button(controls_frame, text="Next ▶", command=self.next_page, state=tk.DISABLED, width=7)
        self.next_page_btn.pack(side=tk.LEFT, padx=(2,10))

        # Zoom Buttons
        ttk.Button(controls_frame, text="➖", command=lambda: self.zoom_pdf(-0.2), width=3).pack(side=tk.LEFT, padx=2)
        ttk.Button(controls_frame, text="➕", command=lambda: self.zoom_pdf(0.2), width=3).pack(side=tk.LEFT, padx=2)
        ttk.Button(controls_frame, text="Fit Width", command=self.zoom_to_fit_width, width=8).pack(side=tk.LEFT, padx=2)


        # PDF content area (Canvas for rendering, Text area for selection)
        pdf_content_paned = ttk.PanedWindow(pdf_panel, orient=tk.VERTICAL)
        pdf_content_paned.pack(fill=tk.BOTH, expand=True)

        # Canvas frame
        canvas_frame = ttk.Frame(pdf_content_paned)
        # Use relief and borderwidth to visually separate canvas area
        self.pdf_canvas = tk.Canvas(canvas_frame, bg="#1C1C1C", bd=0, highlightthickness=0, relief=tk.SUNKEN, borderwidth=2)
        self.pdf_canvas_scrollbar_y = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.pdf_canvas.yview)
        self.pdf_canvas_scrollbar_x = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.pdf_canvas.xview)
        self.pdf_canvas.configure(yscrollcommand=self.pdf_canvas_scrollbar_y.set, xscrollcommand=self.pdf_canvas_scrollbar_x.set)

        self.pdf_canvas_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.pdf_canvas_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.pdf_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        pdf_content_paned.add(canvas_frame, weight=3) # Canvas gets more vertical space

        # Bindings for zoom and navigation
        self.pdf_canvas.bind("<Configure>", self._on_canvas_resize)
        self.pdf_canvas.bind("<Enter>", lambda e: self.pdf_canvas.focus_set()) # Allow mouse wheel zoom/scroll when hovered
        self.pdf_canvas.bind("<MouseWheel>", self._on_mousewheel_canvas)
        self.pdf_canvas.bind("<Button-4>", self._on_mousewheel_canvas) # For Linux/X11
        self.pdf_canvas.bind("<Button-5>", self._on_mousewheel_canvas) # For Linux/X11


        # Text display frame (for selecting text for AI)
        text_display_frame = ttk.Frame(pdf_content_paned)
        ttk.Label(text_display_frame, text="Page Text (for selection & AI context):").pack(anchor=tk.W, pady=(5,2))
        self.page_text_scrolledtext = scrolledtext.ScrolledText(text_display_frame, wrap=tk.WORD, height=10,
                                                                bg=self.text_widget_bg, fg=self.text_widget_fg,
                                                                font=self.text_widget_font, relief=tk.FLAT,
                                                                borderwidth=1, insertbackground=self.text_widget_fg)
        self.page_text_scrolledtext.pack(fill=tk.BOTH, expand=True)
        self.page_text_scrolledtext.config(state=tk.DISABLED) # Start disabled until PDF is loaded
        pdf_content_paned.add(text_display_frame, weight=1) # Text area gets less vertical space


    def _create_ai_panel(self, parent_frame):
        ai_panel = ttk.Frame(parent_frame)
        ai_panel.pack(fill=tk.BOTH, expand=True)

        # Model and Personality Selection Frame
        selection_frame = ttk.Frame(ai_panel)
        selection_frame.pack(fill=tk.X, pady=(0,10))

        # Model Selection
        ttk.Label(selection_frame, text="AI Model:", font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT, padx=(0,5))
        self.model_dropdown = ttk.Combobox(selection_frame, textvariable=self.current_ollama_model,
                                           values=self.available_ollama_models, state="readonly", width=25) # Adjusted width
        self.model_dropdown.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        self.model_dropdown.bind("<<ComboboxSelected>>", self.on_ollama_model_selected)

        self.refresh_models_btn = ttk.Button(selection_frame, text="🔄", command=self.threaded_fetch_ollama_models, width=3)
        self.refresh_models_btn.pack(side=tk.LEFT, padx=(0, 10))

        # Personality Selection
        ttk.Label(selection_frame, text="Personality:", font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT, padx=(10,5))
        self.personality_dropdown = ttk.Combobox(selection_frame, textvariable=self.selected_personality,
                                       values=list(self.personalities.keys()), state="readonly", width=15)
        self.personality_dropdown.pack(side=tk.LEFT, padx=(0,5))
        self.personality_dropdown.bind("<<ComboboxSelected>>", self.on_personality_selected)


        # AI Chat Output Area
        chat_frame = ttk.LabelFrame(ai_panel, text="AI Chat & Output", padding=5)
        chat_frame.pack(fill=tk.BOTH, expand=True, pady=(0,5))

        self.chat_history_scrolledtext = scrolledtext.ScrolledText(chat_frame, wrap=tk.WORD,
                                                                   bg=self.text_widget_bg, fg=self.text_widget_fg,
                                                                   font=self.text_widget_font, relief=tk.FLAT,
                                                                   borderwidth=1, insertbackground=self.text_widget_fg)
        self.chat_history_scrolledtext.pack(fill=tk.BOTH, expand=True)
        self.chat_history_scrolledtext.config(state=tk.DISABLED) # Start disabled

        # Configure chat history tags (now that the widget exists)
        self.chat_history_scrolledtext.tag_configure("user", foreground=self._chat_user_color, font=('Segoe UI', 10, 'bold'))
        self.chat_history_scrolledtext.tag_configure("ai", foreground=self._chat_ai_color)
        self.chat_history_scrolledtext.tag_configure("error", foreground=self._chat_error_color, font=('Segoe UI', 10, 'italic'))
        self.chat_history_scrolledtext.tag_configure("system", foreground=self._chat_system_color, font=('Segoe UI', 9, 'italic'))


        # Action Buttons Frame (Grid layout)
        action_buttons_frame = ttk.Frame(ai_panel)
        action_buttons_frame.pack(fill=tk.X, pady=(5,5))

        self.explain_concept_btn = ttk.Button(action_buttons_frame, text="Explain Concept", command=self.explain_selected_concept, state=tk.DISABLED)
        self.explain_concept_btn.grid(row=0, column=0, padx=2, pady=2, sticky="ew")

        self.explain_code_btn = ttk.Button(action_buttons_frame, text="Explain Code", command=self.explain_selected_code, state=tk.DISABLED)
        self.explain_code_btn.grid(row=0, column=1, padx=2, pady=2, sticky="ew")

        self.analyze_images_btn = ttk.Button(action_buttons_frame, text="Analyze Images", command=self.analyze_images_on_current_page, state=tk.DISABLED)
        self.analyze_images_btn.grid(row=0, column=2, padx=2, pady=2, sticky="ew")

        self.summarize_page_btn = ttk.Button(action_buttons_frame, text="Summarize Page", command=lambda: self.generate_study_material("summary"), state=tk.DISABLED)
        self.summarize_page_btn.grid(row=1, column=0, padx=2, pady=2, sticky="ew")

        self.generate_quiz_btn = ttk.Button(action_buttons_frame, text="Generate Quiz", command=lambda: self.generate_study_material("quiz"), state=tk.DISABLED)
        self.generate_quiz_btn.grid(row=1, column=1, padx=2, pady=2, sticky="ew")

        self.key_points_btn = ttk.Button(action_buttons_frame, text="Key Points", command=lambda: self.generate_study_material("key_points"), state=tk.DISABLED)
        self.key_points_btn.grid(row=1, column=2, padx=2, pady=2, sticky="ew")

        # Configure columns to expand equally
        action_buttons_frame.columnconfigure(0, weight=1)
        action_buttons_frame.columnconfigure(1, weight=1)
        action_buttons_frame.columnconfigure(2, weight=1)


        # User Input Frame
        input_frame = ttk.Frame(ai_panel)
        input_frame.pack(fill=tk.X, pady=(5,0))

        self.user_question_entry = ttk.Entry(input_frame, font=self.text_widget_font)
        self.user_question_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5), ipady=3)
        self.user_question_entry.bind("<Return>", self.send_question_to_ai)

        self.send_question_btn = ttk.Button(input_frame, text="➤ Send", command=self.send_question_to_ai, state=tk.DISABLED)
        self.send_question_btn.pack(side=tk.LEFT)


        # TTS and Voice Query Frame
        tts_frame = ttk.Frame(ai_panel)
        tts_frame.pack(fill=tk.X, pady=(10,0))

        ttk.Label(tts_frame, text="Speech:", font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT, padx=(0,5))

        # Voice selection
        self.voice_dropdown = ttk.Combobox(tts_frame, textvariable=self.selected_voice,
                                         values=self.voice_list, state="readonly", width=15) # Adjusted width
        self.voice_dropdown.pack(side=tk.LEFT, padx=(0,10))

        # TTS control buttons
        self.play_tts_btn = ttk.Button(tts_frame, text="▶ Speak Page",
                                     command=self.play_current_page_tts, state=tk.DISABLED)
        self.play_tts_btn.pack(side=tk.LEFT, padx=2)

        self.play_ai_tts_btn = ttk.Button(tts_frame, text="▶ Speak AI",
                                        command=self.play_last_ai_response, state=tk.DISABLED)
        self.play_ai_tts_btn.pack(side=tk.LEFT, padx=2)

        self.stop_tts_btn = ttk.Button(tts_frame, text="⏹ Stop",
                                     command=self.stop_current_page_tts, state=tk.DISABLED)
        self.stop_tts_btn.pack(side=tk.LEFT, padx=2)

        # Auto-play checkbox
        self.auto_play_check = ttk.Checkbutton(tts_frame, text="Auto-Speak AI",
                                             variable=self.auto_play_ai,
                                             command=self.toggle_auto_play)
        self.auto_play_check.pack(side=tk.LEFT, padx=(10,0))

        # Voice Query Button (State managed based on SpeechRecognition availability)
        # 🎙️ icon: U+1F399 FE0F (Unicode for microphone with variation selector)
        self.voice_query_btn = ttk.Button(tts_frame, text="🎙️ Voice Query", command=self.handle_voice_query) # State set in _update_tts_button_states
        self.voice_query_btn.pack(side=tk.LEFT, padx=(10, 0))


    def _create_status_bar(self):
        self.status_label = ttk.Label(self.root, text="Welcome! Load a PDF to start.", style='Status.TLabel', anchor=tk.W)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(5,5))

    def update_status(self, message):
        """Updates the status bar message on the main thread."""
        if hasattr(self.status_label, 'config'):
            self.root.after(0, lambda: self.status_label.config(text=message))
        # No need to update_idletasks here, after(0,...) schedules it for the next idle moment

    def handle_error(self, message, title="Error"):
        """Displays an error message in a dialog, updates status, and adds to chat."""
        print(f"[ERROR] {title}: {message}") # Log the error
        self.root.after(0, lambda: messagebox.showerror(title, message))
        self.root.after(0, self.update_status, f"{title}: {message}")
        self.root.after(0, self.add_to_chat, "Error", message, "error")


    def threaded_fetch_ollama_models(self):
        """Starts model fetching in a separate thread."""
        self.update_status("Fetching Ollama models...")
        self.current_ollama_model.set("Loading...")
        if hasattr(self.model_dropdown, 'config'):
            self.root.after(0, lambda: self.model_dropdown.config(values=["Loading..."], state="disabled")) # Disable dropdown while loading
        if hasattr(self.refresh_models_btn, 'config'):
             self.root.after(0, lambda: self.refresh_models_btn.config(state=tk.DISABLED))
        threading.Thread(target=self._fetch_ollama_models_worker, daemon=True).start()


    def _fetch_ollama_models_worker(self):
        """Worker thread function to fetch Ollama models."""
        try:
            response = requests.get(f"{self.ollama_base_url}/api/tags", timeout=10)
            response.raise_for_status() # Raise HTTPError for bad responses (4xx or 5xx)
            models_data = response.json().get('models', [])

            if not models_data:
                 self.available_ollama_models = []
                 error_message = "Ollama is running, but no models found. Pull models via 'ollama pull <model_name>' in your terminal."
                 self.root.after(0, self.update_status, error_message)
                 self.root.after(0, self.current_ollama_model.set, "No Models Found")
                 self.root.after(0, lambda: self.model_dropdown.config(values=["No Models Found"], state="readonly"))
            else:
                self.available_ollama_models = sorted([model['name'] for model in models_data])
                self.root.after(0, self._update_ollama_model_dropdown_ui) # Update UI on main thread
                self.root.after(0, self.update_status, f"Found {len(self.available_ollama_models)} Ollama models.")

        except requests.exceptions.ConnectionError:
            error_message = "Error: Could not connect to Ollama. Is 'ollama serve' running?"
            self.root.after(0, self.update_status, error_message)
            self.root.after(0, self.current_ollama_model.set, "Ollama Offline")
            self.root.after(0, lambda: self.model_dropdown.config(values=["Ollama Offline"], state="readonly"))
        except requests.exceptions.Timeout:
            error_message = "Error: Ollama connection timed out."
            self.root.after(0, self.update_status, error_message)
            self.root.after(0, self.current_ollama_model.set, "Ollama Timeout")
            self.root.after(0, lambda: self.model_dropdown.config(values=["Ollama Timeout"], state="readonly"))
        except requests.exceptions.RequestException as e:
             error_message = f"Error fetching Ollama models: {str(e)}"
             self.root.after(0, self.update_status, error_message)
             self.root.after(0, self.current_ollama_model.set, "Error Fetching")
             self.root.after(0, lambda: self.model_dropdown.config(values=["Error Fetching"], state="readonly"))
        except Exception as e:
            # Catch any other unexpected errors during the process
            error_message = f"An unexpected error occurred while fetching models: {str(e)}"
            self.root.after(0, self.update_status, error_message)
            self.root.after(0, self.current_ollama_model.set, "Error Fetching")
            self.root.after(0, lambda: self.model_dropdown.config(values=["Error Fetching"], state="readonly"))

        finally:
            # Ensure refresh button is re-enabled and on_ollama_model_selected is called
            self.root.after(0, lambda: self.refresh_models_btn.config(state=tk.NORMAL))
            self.root.after(0, self.on_ollama_model_selected) # This will set button states based on new model status


    def _update_ollama_model_dropdown_ui(self):
        """Updates the model dropdown values and attempts to select a default on the main thread."""
        if not hasattr(self.model_dropdown, 'config'): return

        self.model_dropdown.config(values=self.available_ollama_models, state="readonly")

        # Prioritize selecting a suitable model
        preferred_models_order = [
            "llama3", # Newer general models
            "phi3", # Good general purpose
            "mistral", # Good general purpose
            "deepseek-r1:14b", # Strong reasoning
            "llava", # Good general vision
            "deepseek-coder", # Strong coder
            "wizardlm", # Potential reasoning
            "bakllava", # Vision
            "moondream", # Lightweight vision
            "nexusraven", # General purpose, function calling
            "instruct", # General keyword for instruction-tuned models
            "latest", "v2", "v3", "mini", "7b", "13b", "34b" # General keywords
        ]

        selected_default = False
        # Try to select a preferred model that is actually available
        for model_preference in preferred_models_order:
             for available_model in self.available_ollama_models:
                 # Case-insensitive search for preference keywords
                 if model_preference.lower() in available_model.lower():
                     self.current_ollama_model.set(available_model)
                     selected_default = True
                     break # Found a match, stop checking preferences
             if selected_default: break # Found a match overall, stop checking available models

        if not selected_default and self.available_ollama_models:
            # If no preferred model was found, just select the first available one
            self.current_ollama_model.set(self.available_ollama_models[0])
        elif not self.available_ollama_models:
            self.current_ollama_model.set("No Models Found")

        # Call on_ollama_model_selected to set capabilities and update buttons based on the selected model
        self.on_ollama_model_selected()


    def _get_model_capabilities(self, model_name_str):
        """Determines capabilities based on model name keywords."""
        if not model_name_str or model_name_str in ["Loading...", "No Models Found", "Ollama Offline", "Ollama Timeout", "Error Fetching"]:
            return {} # Return empty caps if model is not ready

        name_lower = model_name_str.lower()
        return {
            # Vision models often have 'llava', 'vision', specific multi-modal names
            "vision": any(kw in name_lower for kw in ["llava", "vision", "bakllava", "moondream", "fuyu"]),
            # Code models often have 'coder', 'codellama', 'starcoder', 'deepseek-coder'
            "code": any(kw in name_lower for kw in ["coder", "codellama", "deepseek-coder", "starcoder", "programming"]),
            # Reasoning/General instruction models are common default or have 'instruct', 'chat', 'platypus', 'wizardlm', 'openhermes' etc.
            "reasoning": any(kw in name_lower for kw in ["instruct", "chat", "platypus", "wizardlm", "hermes", "mistral", "llama", "phi", "deepseek"]), # Broad list for reasoning
            "general": True # Assume general capability unless it's purely an embedding/audio model etc.
        }

    def on_ollama_model_selected(self, event=None):
        """Called when a new Ollama model is selected or model list is updated."""
        model_name = self.current_ollama_model.get()
        # Check if a valid model name is selected
        is_model_valid = model_name not in ["Loading...", "No Models Found", "Ollama Offline", "Ollama Timeout", "Error Fetching"] and bool(model_name)

        if is_model_valid:
            self.model_capabilities = self._get_model_capabilities(model_name)
            caps_str = ', '.join([k for k, v in self.model_capabilities.items() if v]) or 'None'
            self.update_status(f"Selected AI Model: {model_name} | Caps: {caps_str}")
            # Add system message to chat history if this was a user-initiated selection change
            if event: # Only log to chat if triggered by combobox selection
                 self.add_to_chat("System", f"AI model set to: {model_name} (Capabilities: {caps_str})", "system")
        else:
            self.model_capabilities = {} # Clear capabilities if model is not valid
            self.update_status(f"AI Model not ready: {model_name}")
            if event: # Only log to chat if triggered by combobox selection (and it's not valid)
                 self.add_to_chat("System", f"AI model not ready: {model_name}. AI features disabled.", "system")


        # Update AI button states based on PDF loaded AND model valid status and capabilities
        self._set_ai_buttons_state() # No need to pass state, it's derived internally


    def on_personality_selected(self, event=None):
        """Called when a new Personality is selected."""
        personality_name = self.selected_personality.get()
        if personality_name in self.personalities:
            personality_info = self.personalities[personality_name]
            icon = personality_info["icon"]
            system_prompt = personality_info["system_prompt"]
            self.update_status(f"Selected Personality: {personality_name}")
            # Add system message to chat history if this was a user-initiated selection change
            if event: # Only log to chat if triggered by combobox selection
                self.add_to_chat("System",
                                 f"Active personality: {icon} {personality_name}\n"
                                 f"System prompt: \"{system_prompt}\"",
                                 "system")
        else:
             # Handle case where selected personality is somehow not in the dict (shouldn't happen with readonly)
             self.update_status(f"Selected unknown personality: {personality_name}")


    def _set_ai_buttons_state(self):
        """Sets the state of AI-related buttons based on PDF and model availability/capabilities."""
        is_model_valid = self.current_ollama_model.get() not in ["Loading...", "No Models Found", "Ollama Offline", "Ollama Timeout", "Error Fetching"] and bool(self.current_ollama_model.get())
        is_pdf_loaded = self.pdf_document is not None and len(self.pdf_page_text_for_ai) > 0

        # Base state for PDF-dependent buttons requires both PDF and a valid model
        effective_pdf_dependent_state = tk.NORMAL if is_pdf_loaded and is_model_valid else tk.DISABLED

        # Capabilities based on the currently selected model
        can_reason = self.model_capabilities.get("reasoning", False) and is_model_valid
        can_code = self.model_capabilities.get("code", False) and is_model_valid
        can_vision = self.model_capabilities.get("vision", False) and is_model_valid

        # General chat button only requires a valid model
        if hasattr(self.send_question_btn, 'config'):
             self.send_question_btn.config(state=tk.NORMAL if is_model_valid else tk.DISABLED)

        # Reasoning-based buttons require PDF + model with reasoning capability
        if hasattr(self.explain_concept_btn, 'config'):
            self.explain_concept_btn.config(state=effective_pdf_dependent_state if can_reason else tk.DISABLED)
        if hasattr(self.summarize_page_btn, 'config'):
            self.summarize_page_btn.config(state=effective_pdf_dependent_state if can_reason else tk.DISABLED)
        if hasattr(self.generate_quiz_btn, 'config'):
            self.generate_quiz_btn.config(state=effective_pdf_dependent_state if can_reason else tk.DISABLED)
        if hasattr(self.key_points_btn, 'config'):
            self.key_points_btn.config(state=effective_pdf_dependent_state if can_reason else tk.DISABLED)

        # Code button requires PDF + model with code capability
        if hasattr(self.explain_code_btn, 'config'):
            self.explain_code_btn.config(state=effective_pdf_dependent_state if can_code else tk.DISABLED)

        # Vision button requires PDF + model with vision capability
        if hasattr(self.analyze_images_btn, 'config'):
            # Note: analyze_images requires extracting images first, which is done on load.
            # We could further refine this to only be active if self.pdf_page_images is not empty
            # but for simplicity, we'll enable it if the model supports vision and PDF is loaded.
            self.analyze_images_btn.config(state=effective_pdf_dependent_state if can_vision else tk.DISABLED)

        # TTS buttons state based on PDF/AI response availability (handled separately)
        self._update_tts_button_states()

        # Voice query button state (requires model valid AND voice query available)
        if hasattr(self.voice_query_btn, 'config'):
             self.voice_query_btn.config(
                 state=tk.NORMAL if is_model_valid and self.voice_query_available else tk.DISABLED
             )


    def load_pdf_dialog(self):
        """Opens a file dialog and loads the selected PDF."""
        file_path = filedialog.askopenfilename(title="Select PDF File",
                                               filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")])
        if not file_path: return

        # Close previous document if any
        if self.pdf_document:
            try: self.pdf_document.close()
            except Exception as e: print(f"Error closing previous PDF: {e}")

        self.clear_pdf_view_and_data() # Clear UI and internal data

        try:
            self.update_status(f"Loading PDF: {os.path.basename(file_path)}...")
            self.pdf_document = fitz.open(file_path)
            total_pages = self.pdf_document.page_count
            self.update_status(f"Extracting text and images from {total_pages} pages...")
            # Start extraction in a separate thread
            threading.Thread(target=self._extract_all_pdf_content_worker, args=(total_pages,), daemon=True).start()

        except Exception as e:
            self.handle_error(f"Failed to load PDF: {str(e)}", "PDF Load Error")
            self.pdf_document = None # Ensure document is None on error
            # Update AI button states after PDF load failure
            self._set_ai_buttons_state() # This will disable PDF-dependent buttons


    def _extract_all_pdf_content_worker(self, total_pages):
        """Worker thread function to extract text and images from all pages."""
        if not self.pdf_document: return

        extracted_texts = []
        extracted_images = [] # List of lists of image data per page

        try:
            for i in range(total_pages):
                page_texts = "[Error extracting text]" # Default error state
                page_images_data = [] # List of base64 image strings for this page

                try:
                    page = self.pdf_document.load_page(i)
                    # Extract text
                    try:
                        text = page.get_text("text", sort=True).strip()
                        page_texts = text if text else "[No text found on this page]"
                    except Exception as text_e:
                        page_texts = f"[Error extracting text: {str(text_e)[:50]}]"
                        print(f"Error extracting text from page {i+1}: {text_e}")

                    # Extract images for the current page
                    try:
                        images_on_page = page.get_images(full=True)
                        for img_index, img_info in enumerate(images_on_page):
                             # Limit image extraction/processing for performance if needed
                             # if img_index >= 5: break # Example limit
                             xref = img_info[0]
                             base_image = self.pdf_document.extract_image(xref)
                             if base_image and base_image["image"]:
                                 # Store as base64 string
                                 page_images_data.append(base64.b64encode(base_image["image"]).decode('utf-8'))

                    except Exception as img_e:
                        print(f"Error extracting images from page {i+1}: {img_e}")
                        # Continue without images for this page

                except Exception as page_e:
                    # Catch errors loading the page itself
                    page_texts = f"[Error loading page {i+1}: {str(page_e)[:50]}]"
                    print(f"Error loading page {i+1}: {page_e}")

                extracted_texts.append(page_texts)
                extracted_images.append(page_images_data)


                # Update status periodically or on completion
                if (i + 1) % 10 == 0 or (i + 1) == total_pages:
                    if self.root:
                        self.root.after(0, self.update_status, f"Extracted content from {i+1}/{total_pages} pages...")

            self.pdf_page_text_for_ai = extracted_texts
            self.pdf_page_images = extracted_images

            if self.root:
                self.root.after(0, self.update_status, "PDF content extraction complete.")
                # After extraction, render the first page and update UI states
                self.root.after(0, self.render_current_pdf_page)
                # Update AI button states now that content is available (checks model readiness internally)
                self.root.after(0, self._set_ai_buttons_state)


        except Exception as e:
            # This catches errors outside the per-page loop (less common)
            if self.root:
                error_message = f"Critical error during PDF content extraction: {str(e)}"
                self.handle_error(error_message, "Extraction Error")
                # Populate with error placeholders
                self.pdf_page_text_for_ai = ["[Critical Extraction Error]" for _ in range(total_pages)] if self.pdf_document else []
                self.pdf_page_images = [[] for _ in range(total_pages)] if self.pdf_document else []
                self.root.after(0, self.render_current_pdf_page) # Still try to render page with error text
                self.root.after(0, self._set_ai_buttons_state) # Update button states


    def render_current_pdf_page(self):
        """Renders the current PDF page on the canvas and updates related UI."""
        if not self.pdf_document:
            self.clear_pdf_view_and_data()
            return
        try:
            # Ensure page number is within bounds
            self.current_page_num = max(0, min(self.current_page_num, self.pdf_document.page_count - 1))

            page = self.pdf_document.load_page(self.current_page_num)
            mat = fitz.Matrix(self.current_zoom_scale, self.current_zoom_scale)
            # Use get_displaylist and get_pixmap from displaylist for potentially better rendering
            # dl = page.get_displaylist()
            # pix = dl.get_pixmap(matrix=mat, alpha=False)
            pix = page.get_pixmap(matrix=mat, alpha=False) # Simpler approach, usually sufficient

            # Convert pixmap to PhotoImage
            img_data = pix.tobytes("ppm") # Use ppm format for Pillow
            pil_img = Image.open(io.BytesIO(img_data))
            self.rendered_page_image = ImageTk.PhotoImage(pil_img)

            # Update canvas
            self.pdf_canvas.delete("all") # Clear previous content
            self.pdf_canvas.create_image(0, 0, anchor=tk.NW, image=self.rendered_page_image)
            self.pdf_canvas.config(scrollregion=self.pdf_canvas.bbox(tk.ALL)) # Set scrollable area

            # Update navigation entry and label
            if hasattr(self.page_nav_entry, 'delete'):
                self.page_nav_entry.delete(0, tk.END)
                self.page_nav_entry.insert(0, str(self.current_page_num + 1)) # Display 1-based page number
            if hasattr(self.page_nav_label, 'config'):
                self.page_nav_label.config(text=f"/ {self.pdf_document.page_count}")

            # Update navigation button states
            if hasattr(self.prev_page_btn, 'config'):
                self.prev_page_btn.config(state=tk.NORMAL if self.current_page_num > 0 else tk.DISABLED)
            if hasattr(self.next_page_btn, 'config'):
                self.next_page_btn.config(state=tk.NORMAL if self.current_page_num < self.pdf_document.page_count - 1 else tk.DISABLED)

            # Update page text area
            if hasattr(self.page_text_scrolledtext, 'config'):
                self.page_text_scrolledtext.config(state=tk.NORMAL) # Enable editing temporarily
                self.page_text_scrolledtext.delete("1.0", tk.END)
                # Display extracted text or an error message if extraction failed for this page
                if self.pdf_page_text_for_ai and 0 <= self.current_page_num < len(self.pdf_page_text_for_ai):
                    self.page_text_scrolledtext.insert(tk.END, self.pdf_page_text_for_ai[self.current_page_num])
                else:
                    self.page_text_scrolledtext.insert(tk.END, "[Text not available or error during extraction for this page.]")
                self.page_text_scrolledtext.config(state=tk.DISABLED) # Disable editing

            # Update TTS button state based on text availability
            self._update_tts_button_states()

        except Exception as e:
            self.handle_error(f"Error rendering page {self.current_page_num + 1}: {str(e)}", "Rendering Error")
            if hasattr(self.pdf_canvas, 'delete'): self.pdf_canvas.delete("all") # Clear canvas on error
            # Ensure TTS button is disabled on render error
            self._update_tts_button_states()


    def clear_pdf_view_and_data(self):
        """Clears the PDF display and related internal data."""
        if hasattr(self.pdf_canvas, 'delete'): self.pdf_canvas.delete("all")
        self.rendered_page_image = None
        self.pdf_page_text_for_ai = []
        self.pdf_page_images = []
        self.current_page_num = 0
        self.current_zoom_scale = 1.0

        # Reset navigation and text area
        if hasattr(self.page_nav_entry, 'delete'):
            self.page_nav_entry.delete(0, tk.END)
            self.page_nav_entry.insert(0, "0")
        if hasattr(self.page_nav_label, 'config'): self.page_nav_label.config(text="/ 0")
        if hasattr(self.prev_page_btn, 'config'): self.prev_page_btn.config(state=tk.DISABLED)
        if hasattr(self.next_page_btn, 'config'): self.next_page_btn.config(state=tk.DISABLED)
        if hasattr(self.page_text_scrolledtext, 'config'):
            self.page_text_scrolledtext.config(state=tk.NORMAL)
            self.page_text_scrolledtext.delete("1.0", tk.END)
            self.page_text_scrolledtext.insert(tk.END, "[No PDF Loaded]")
            self.page_text_scrolledtext.config(state=tk.DISABLED)

        # Update AI button states based on PDF unloaded status
        self._set_ai_buttons_state()
        # Also update TTS buttons specifically
        self._update_tts_button_states()
        self.last_ai_response = "" # Clear last AI response


    def next_page(self):
        """Navigates to the next page if available."""
        if self.pdf_document and self.current_page_num < self.pdf_document.page_count - 1:
            self.current_page_num += 1
            self.render_current_pdf_page()


    def prev_page(self):
        """Navigates to the previous page if available."""
        if self.pdf_document and self.current_page_num > 0:
            self.current_page_num -= 1
            self.render_current_pdf_page()


    def go_to_page_from_entry(self, event=None):
        """Navigates to the page number entered in the entry field."""
        if not self.pdf_document: return
        try:
            # Get 1-based page number from entry, convert to 0-based index
            target_page = int(self.page_nav_entry.get()) - 1
            if 0 <= target_page < self.pdf_document.page_count:
                self.current_page_num = target_page
                self.render_current_pdf_page()
            else:
                messagebox.showwarning("Invalid Page", f"Page number must be between 1 and {self.pdf_document.page_count}.")
                # Reset entry to current page number on invalid input
                self.page_nav_entry.delete(0, tk.END); self.page_nav_entry.insert(0, str(self.current_page_num + 1))
        except ValueError:
            messagebox.showwarning("Invalid Input", "Please enter a valid page number.")
            # Reset entry to current page number on invalid input
            self.page_nav_entry.delete(0, tk.END); self.page_nav_entry.insert(0, str(self.current_page_num + 1))


    def zoom_pdf(self, factor_change):
        """Changes the zoom level of the PDF render."""
        if self.pdf_document:
            self.current_zoom_scale += factor_change
            self.current_zoom_scale = max(0.2, min(5.0, self.current_zoom_scale)) # Prevent extreme zoom
            self.render_current_pdf_page()


    def zoom_to_fit_width(self):
        """Zooms the PDF page to fit the canvas width."""
        if not self.pdf_document or not hasattr(self.pdf_canvas, 'winfo_width') or self.pdf_canvas.winfo_width() <= 1:
             print("[DEBUG] Cannot fit width: PDF not loaded or canvas not ready.") # Debug log
             return
        try:
            page = self.pdf_document.load_page(self.current_page_num)
            page_width = page.rect.width # Original page width
            canvas_width = self.pdf_canvas.winfo_width() # Current canvas width
            # Adjust for potential scrollbar space if needed, but scrollregion handles clipping
            if page_width > 0 and canvas_width > 0:
                 # Calculate new scale
                 self.current_zoom_scale = canvas_width / page_width
                 # Apply zoom bounds
                 self.current_zoom_scale = max(0.2, min(5.0, self.current_zoom_scale))
                 self.render_current_pdf_page()
            else:
                 print("[DEBUG] Cannot fit width: Page width or canvas width is zero.") # Debug log
        except Exception as e:
            self.handle_error(f"Error fitting width: {str(e)}", "Zoom Error")


    def _on_canvas_resize(self, event):
        """Handles canvas resize events (primarily for scroll region update)."""
        # The scrollregion is updated in render_current_pdf_page, which is called
        # after zoom or page changes. This handler could be used for more complex
        # resize behavior, but simple scrollregion adjustment is sufficient here.
        # It's important to have this binding even if the logic is minimal,
        # as it can trigger scrollbar updates in some setups.
        if self.pdf_document and self.rendered_page_image:
             # Re-set scrollregion on resize
             self.pdf_canvas.config(scrollregion=self.pdf_canvas.bbox(tk.ALL))


    def _on_mousewheel_canvas(self, event):
        """Handles mouse wheel scrolling on the PDF canvas."""
        if event.num == 4 or event.delta > 0: # Scroll up
            self.pdf_canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0: # Scroll down
            self.pdf_canvas.yview_scroll(1, "units")
        # Add horizontal scrolling with Shift key? (Optional enhancement)


    def add_to_chat(self, sender, message, tag_override=None):
        """Adds a message to the chat history text widget."""
        if not hasattr(self.chat_history_scrolledtext, 'insert'): return

        # Enable editing temporarily
        self.chat_history_scrolledtext.config(state=tk.NORMAL)

        # Determine tag and prefix
        tag = tag_override if tag_override else sender.lower()
        prefix = f"{sender}: " if sender else "" # Only add prefix if sender is provided

        # Apply sender tag (bold for user/AI, normal for system/error)
        if prefix:
             prefix_tag = "bold" if sender.lower() in ["user", "ai"] else tag
             self.chat_history_scrolledtext.insert(tk.END, prefix, prefix_tag)

        # Apply message tag
        self.chat_history_scrolledtext.insert(tk.END, message + "\n\n", tag)

        # Disable editing
        self.chat_history_scrolledtext.config(state=tk.DISABLED)

        # Auto-scroll to the bottom
        self.chat_history_scrolledtext.see(tk.END)

        # Store last AI response and handle auto-play
        if sender and sender.lower() == "ai":
            self.last_ai_response = message.strip() # Store stripped response
            # Update button states (including enabling the Speak AI button)
            self._update_tts_button_states()

            # If auto-play is enabled and there's response text, start playing
            if self.auto_play_ai.get() and self.last_ai_response:
                # Use a short delay to allow the UI to update
                self.root.after(100, self.play_last_ai_response)


    def _prepare_ai_prompt_and_context(self, user_request_text, include_page_context=True, max_page_context_len=3000, max_history_turns=5):
        """Builds the full prompt for the AI including personality, history, and context."""
        full_prompt_parts = []

        # 1. Add Personality System Prompt
        personality_name = self.selected_personality.get()
        personality_info = self.personalities.get(personality_name, self.personalities["Default Tutor"]) # Fallback
        full_prompt_parts.append(f"System Role: {personality_info['system_prompt']}\n")

        # 2. Include recent chat history
        if self.chat_conversation_history:
            history_str = "Previous conversation turns:\n"
            # Add a few recent turns (excluding the current user turn which is in user_request_text)
            # Ensure we don't exceed history length or max_history_turns
            start_index = max(0, len(self.chat_conversation_history) - max_history_turns)
            for entry in self.chat_conversation_history[start_index:]:
                 # Format history role (User/Assistant) and content
                 role = entry.get('role', 'unknown').capitalize()
                 content = entry.get('content', '')
                 # Avoid adding very short or empty history entries
                 if content.strip():
                      history_str += f"{role}: {content.strip()[:200]}...\n" # Limit history content length

            if history_str != "Previous conversation turns:\n": # Only add if there's actual history included
                 full_prompt_parts.append(history_str)

        # 3. Include Current PDF Page Context
        if include_page_context and self.pdf_document and self.pdf_page_text_for_ai:
            if 0 <= self.current_page_num < len(self.pdf_page_text_for_ai):
                page_text = self.pdf_page_text_for_ai[self.current_page_num]
                # Limit the page text length to manage context window
                truncated_page_text = page_text.strip()[:max_page_context_len]

                if truncated_page_text and not truncated_page_text.startswith(("[No text found", "[Error extracting", "[Critical Extraction Error]")): # Only add if there's valid text
                     full_prompt_parts.append(f"Current PDF Page ({self.current_page_num + 1}) Context:\n\"\"\"\n{truncated_page_text}\n\"\"\"\n")
                elif truncated_page_text: # Add placeholder if text was extracted but indicates error/empty
                     full_prompt_parts.append(f"Current PDF Page ({self.current_page_num + 1}) Context: {truncated_page_text}\n")


        # 4. Add the user's current request/instruction
        full_prompt_parts.append(f"User's Request: {user_request_text.strip()}\n\nAI Response:")

        # Join all parts into the final prompt string
        return "\n".join(full_prompt_parts)


    def _threaded_ollama_request(self, request_label, user_instruction_prompt, images_base64_list=None, include_page_context=True):
        """
        Handles sending a request to Ollama in a separate thread.

        Args:
            request_label (str): A short description for logging/status updates.
            user_instruction_prompt (str): The specific instruction for the AI for this task.
            images_base64_list (list, optional): List of base64 image strings for vision models. Defaults to None.
            include_page_context (bool): Whether to include the current page text as context. Defaults to True.
        """
        # Use the currently selected model
        model_name = self.current_ollama_model.get()

        if model_name in ["Loading...", "No Models Found", "Ollama Offline", "Ollama Timeout", "Error Fetching"] or not bool(model_name):
            # Schedule error message on the main thread
            self.root.after(0, self.handle_error, "Ollama model not available or not selected. Cannot send request.", "AI Request Failed")
            return

        # Prepare the full prompt including personality, history, and page context
        full_prompt_for_ai = self._prepare_ai_prompt_and_context(
             user_instruction_prompt,
             include_page_context=include_page_context # Use the argument to control page context inclusion
        )

        # Add user request log entry to chat history BEFORE sending
        # The actual user input message (if any) should be added via add_to_chat
        # in the calling method (e.g., send_question_to_ai, explain_concept_btn handlers)
        # This history entry is just for the internal chat_conversation_history list
        self.chat_conversation_history.append({"role": "user", "content": f"({request_label}) {user_instruction_prompt.strip()[:100]}..."}) # Log truncated instruction


        self.root.after(0, self.update_status, f"Sending '{request_label}' request to {model_name}...")


        payload = {
            "model": model_name,
            "prompt": full_prompt_for_ai,
            "stream": False, # Use non-streaming for simplicity in this version
            "options": {"temperature": 0.6, "num_ctx": 4096} # num_ctx should ideally match the model's context window
        }

        # Determine capabilities of the currently selected model
        current_model_capabilities = self._get_model_capabilities(model_name)

        # Add images only if they are provided AND the model supports vision
        if images_base64_list and current_model_capabilities.get("vision", False):
            payload["images"] = images_base64_list
            print(f"[DEBUG] Including {len(images_base64_list)} image(s) in payload.") # Debug log
        elif images_base64_list and not current_model_capabilities.get("vision", False):
             # Warning if images are sent to a non-vision model
             self.root.after(0, self.add_to_chat, "System", f"Warning: Images sent to model '{model_name}' which may not be ideal for vision. Results may be poor.", "system")


        ai_response_content = "" # Initialize response content
        try:
            # Send the request
            response = requests.post(f"{self.ollama_base_url}/api/generate", json=payload, timeout=300) # Increased timeout for complex requests
            response.raise_for_status() # Raise HTTPError for bad responses (4xx or 5xx)
            response_data = response.json()
            ai_response_content = response_data.get('response', 'No content in AI response.').strip()

            # Schedule UI updates on the main thread
            if self.root: self.root.after(0, self.add_to_chat, "AI", ai_response_content)
            if self.root: self.root.after(0, self.update_status, f"AI ({model_name}) response received for '{request_label}'.")

            # Append the AI response to history *after* it's fully received
            self.chat_conversation_history.append({"role": "assistant", "content": ai_response_content})
            # Keep history length manageable (e.g., last 20 entries)
            if len(self.chat_conversation_history) > 20:
                self.chat_conversation_history = self.chat_conversation_history[-20:]

        except requests.exceptions.Timeout:
            error_message = f"Request '{request_label}' to {model_name} timed out (waited 300 seconds)."
            self.root.after(0, self.handle_error, error_message, "AI Timeout Error")
            # Append error as assistant response in history
            self.chat_conversation_history.append({"role": "assistant", "content": f"Error: {error_message}"})
        except requests.exceptions.RequestException as e:
            error_detail = f"Ollama API request error for '{request_label}': {str(e)}"
            if e.response is not None:
                try:
                     # Attempt to get error message from JSON response
                     error_json = e.response.json()
                     error_detail += f" - Server said: {error_json.get('error', json.dumps(error_json))}"
                except json.JSONDecodeError:
                    # If response is not JSON, append raw text
                    error_detail += f" - Server said: {e.response.text[:200]}..." # Limit length
            self.root.after(0, self.handle_error, error_detail, "Ollama API Error")
            # Append error as assistant response in history
            self.chat_conversation_history.append({"role": "assistant", "content": f"Error: {error_detail}"})
        except Exception as e:
            # Catch any other unexpected errors during the request process
            unexpected_error = f"An unexpected error occurred during AI request '{request_label}': {str(e)}"
            self.root.after(0, self.handle_error, unexpected_error, "Unexpected AI Error")
             # Append error as assistant response in history
            self.chat_conversation_history.append({"role": "assistant", "content": f"Error: {unexpected_error}"})


    def send_question_to_ai(self, event=None):
        """Sends the user's question to the AI model."""
        user_question = self.user_question_entry.get().strip()
        if not user_question: return # Don't send empty questions

        # Clear the input field immediately
        self.user_question_entry.delete(0, tk.END)

        # Add the user's question to the chat history
        self.add_to_chat("User", user_question)

        # Prepare the full prompt including personality, history, and page context
        # For general questions, include page context by default
        full_prompt = self._prepare_ai_prompt_and_context(user_question, include_page_context=True)

        # Send the request in a separate thread
        threading.Thread(target=self._threaded_ollama_request,
                         args=(f"General Question", user_question, None, True), # Label, user instruction, no images, include page context
                         daemon=True).start()


    def explain_selected_concept(self):
        """Explains the currently selected text using the AI."""
        if not (self.pdf_document and self.pdf_page_text_for_ai and 0 <= self.current_page_num < len(self.pdf_page_text_for_ai)):
             messagebox.showinfo("Not Ready", "Please load a PDF and ensure text is extracted."); return

        try:
            # Attempt to get selected text
            selected_text = self.page_text_scrolledtext.get(tk.SEL_FIRST, tk.SEL_LAST).strip()
            if not selected_text:
                 messagebox.showinfo("Selection Needed", "Please select text from the 'Page Text' panel to explain."); return
            if len(selected_text) < 20: # Basic validation for meaningful text (increased from 10)
                messagebox.showinfo("Selection Too Short", "Please select a more substantial concept text (at least 20 characters)."); return

        except tk.TclError:
             messagebox.showinfo("Selection Needed", "Please select text from the 'Page Text' panel first."); return
        except Exception as e:
            self.handle_error(f"Error getting selected text: {str(e)}", "Selection Error"); return

        # Add user request log entry to chat history
        self.add_to_chat("User", f"Explain concept: \"{selected_text[:70]}...\"")

        # Construct the specific instruction for the AI
        instruction_prompt = (
            f"Provide a comprehensive explanation of the following concept from a document. "
            f"Explain it clearly and thoroughly, assuming the user is a student learning this material.\n"
            f"Consider the following aspects in your explanation:\n"
            f"1. Core Definition (start with a clear, concise definition).\n"
            f"2. Key Characteristics or Components (use bullet points).\n"
            f"3. How it Works or its Process (explain step-by-step if applicable).\n"
            f"4. Practical Applications or Real-World Examples (provide 2-3 concrete examples).\n"
            f"5. Relationship to Other Concepts (show connections or contrasts if relevant to the page context).\n"
            f"6. Common Misunderstandings or Nuances (clarify potential points of confusion).\n"
            f"7. (Optional) A simple analogy to aid understanding.\n\n"
            f"Concept Text:\n"
            f"\"\"\"\n{selected_text}\n\"\"\"\n\n"
            f"Please format your response clearly using Markdown (headings, bullet points, code blocks).\n"
            f"Provide your explanation in accessible language:"
        )

        # Send the request in a separate thread.
        # We explicitly include page context here to help the AI relate the selected concept to the broader page.
        threading.Thread(target=self._threaded_ollama_request,
                         args=("Concept Explanation", instruction_prompt, None, True), # Label, instruction, no images, include page context
                         daemon=True).start()


    def generate_study_material(self, material_type):
        """Generates summary, quiz, or key points for the current page."""
        if not (self.pdf_document and self.pdf_page_text_for_ai and 0 <= self.current_page_num < len(self.pdf_page_text_for_ai)):
            messagebox.showinfo("Not Ready", "Please load a PDF and ensure text has been extracted for the current page."); return

        page_text = self.pdf_page_text_for_ai[self.current_page_num]
        if len(page_text.strip()) < 100: # Increased minimum text length for meaningful output
            messagebox.showinfo("Not Enough Text", "The current page has too little text to generate meaningful study material (requires at least 100 characters of text)."); return
        if page_text.startswith(("[No text found", "[Error extracting", "[Critical Extraction Error]")):
             messagebox.showinfo("Text Not Available", "Text for the current page is not available or could not be extracted."); return


        user_log_message, ai_instruction = "", ""
        request_label = ""

        # Base instruction to include in all material requests
        base_instruction = (f"Analyze the following text from a document page ({self.current_page_num + 1}). "
                            f"Your task is to generate study material based *only* on the provided text. "
                            f"Use clear, concise language suitable for learning. "
                            f"Ensure coverage of all key points mentioned in the text.\n\n"
                            f"Text to Analyze:\n\"\"\"\n{page_text[:4000]}\n\"\"\"\n\n") # Limit text length sent for analysis

        if material_type == "summary":
            request_label = "Summarize Page"
            user_log_message = f"Summarize page {self.current_page_num + 1}"
            ai_instruction = (f"{base_instruction}"
                              f"Provide a comprehensive summary with the following structure, using Markdown:\n"
                              f"**Summary of Page {self.current_page_num + 1}**\n"
                              f"1.  **Main Topic(s):** (1-2 sentences)\n"
                              f"2.  **Key Concepts/Ideas:** (Bulleted list of essential points)\n"
                              f"3.  **Supporting Details/Arguments:** (Briefly mention key evidence or reasoning)\n"
                              f"4.  **Conclusion or Outcome:** (What is the main takeaway?)\n"
                              f"Keep the summary concise but informative.")

        elif material_type == "quiz":
            request_label = "Generate Quiz"
            user_log_message = f"Generate quiz for page {self.current_page_num + 1}"
            ai_instruction = (f"{base_instruction}"
                              f"Create a 5-question quiz based on the text. Include a mix of question types.\n"
                              f"For each question:\n"
                              f"- State the question clearly.\n"
                              f"- Indicate the question difficulty (Easy, Medium, Hard).\n"
                              f"- Provide the answer.\n\n"
                              f"Example format:\n"
                              f"**Question 1 (Medium):** What is the primary function of X?\n"
                              f"A) ... B) ... C) ... D) ...\n"
                              f"**Answer:** B\n\n"
                              f"**Question 2 (Easy):** True or False: Y is always Z.\n"
                              f"**Answer:** False. Y can also be A under condition B.\n\n"
                              f"**Question 3 (Hard):** Briefly explain the relationship between A and B.\n"
                              f"**Sample Answer:** ...\n\n"
                              f"Generate 5 questions following this approach, drawing only from the provided text.")

        elif material_type == "key_points":
            request_label = "Extract Key Points"
            user_log_message = f"Extract key points for page {self.current_page_num + 1}"
            ai_instruction = (f"{base_instruction}"
                              f"Extract and list the most important key points, concepts, definitions, or facts from the text.\n"
                              f"Organize them into a clear, hierarchical, or categorized list using Markdown.\n"
                              f"Prioritize the information by significance within the text.\n"
                              f"Aim for a comprehensive list that captures the essence of the page.")

        else:
            self.handle_error(f"Unknown study material type requested: {material_type}", "Internal Error"); return # Should not happen


        self.add_to_chat("User", user_log_message)

        # Send the request in a separate thread.
        # We explicitly tell the AI to use the provided text as context within the instruction,
        # so we set include_page_context=False in the prompt preparation to avoid duplication.
        threading.Thread(target=self._threaded_ollama_request,
                         args=(request_label, ai_instruction, None, False), # Label, instruction, no images, DO NOT include page context (it's in the instruction)
                         daemon=True).start()


    def explain_selected_code(self):
        """Explains the currently selected code snippet using the AI."""
        if not (self.pdf_document and self.pdf_page_text_for_ai and 0 <= self.current_page_num < len(self.pdf_page_text_for_ai)):
             messagebox.showinfo("Not Ready", "Please load a PDF and ensure text is extracted."); return

        current_model_name = self.current_ollama_model.get()
        current_model_caps = self._get_model_capabilities(current_model_name)

        if not current_model_caps.get("code"):
             # Offer a warning but allow proceeding if the model might have some general understanding
             if current_model_caps.get("reasoning") or current_model_caps.get("general"):
                  messagebox.showwarning("Model Capability",
                                        f"Current model '{current_model_name}' may not be specifically optimized for code explanation. "
                                        f"Consider selecting a model with 'code' capability (e.g., 'deepseek-coder', 'codellama') for better results. "
                                        f"Proceeding anyway.")
             else: # If the model has neither code nor general/reasoning capability, it's unlikely to work
                  messagebox.showwarning("Model Warning", f"Current model '{current_model_name}' is unlikely to explain code well. Please select a different model optimized for code.")
                  return # Stop if model is clearly unsuitable


        try:
            # Attempt to get selected text
            selected_text = self.page_text_scrolledtext.get(tk.SEL_FIRST, tk.SEL_LAST).strip()
            if not selected_text:
                messagebox.showinfo("Selection Needed", "Please select a code snippet from the 'Page Text' panel to explain."); return
            if len(selected_text) < 10: # Minimum length for code
                 messagebox.showinfo("Selection Too Short", "Please select a more substantial code snippet."); return

            # Simple heuristic check if selection *looks* like code (optional)
            # if not self.detect_code_blocks(selected_text):
            #      confirm = messagebox.askyesno("Unusual Selection", "The selected text does not appear to be code. Do you want to attempt explaining it anyway?")
            #      if not confirm: return

        except tk.TclError:
            messagebox.showinfo("Selection Needed", "Please select a code snippet from the 'Page Text' panel first."); return
        except Exception as e:
            self.handle_error(f"Error getting selected text for code explanation: {str(e)}", "Selection Error"); return


        # Add user request log entry to chat history
        self.add_to_chat("User", f"Explain code: ```\n{selected_text[:100]}...\n```")

        # Construct the specific instruction for the AI
        instruction_prompt = (
            f"You are an expert code explainer. Analyze the following code snippet from a document page ({self.current_page_num + 1}).\n"
            f"Provide a detailed, line-by-line or section-by-section explanation if helpful. "
            f"Explain its purpose, functionality, key components (functions, classes, variables), inputs, outputs, and any relevant logic or algorithms.\n"
            f"Identify the programming language if possible.\n"
            f"Discuss potential usage, limitations, or common patterns demonstrated.\n"
            f"Format your explanation using Markdown, especially using code blocks for references.\n\n"
            f"Code Snippet:\n"
            f"```\n" # Use generic code block tag, AI can often detect language
            f"{selected_text}\n"
            f"```\n\n"
            f"Detailed Explanation:"
        )
        # Do NOT include full page context here, as the focus is solely on the selected code
        full_prompt = self._prepare_ai_prompt_and_context(instruction_prompt, include_page_context=False)

        # Send the request in a separate thread
        threading.Thread(target=self._threaded_ollama_request,
                         args=("Code Explanation", instruction_prompt, None, False), # Label, instruction, no images, DO NOT include page context
                         daemon=True).start()


    def analyze_images_on_current_page(self):
        """Analyzes images on the current page using a vision-capable AI model."""
        if not (self.pdf_document and self.pdf_page_text_for_ai and 0 <= self.current_page_num < len(self.pdf_page_text_for_ai) and 0 <= self.current_page_num < len(self.pdf_page_images)):
             messagebox.showinfo("Not Ready", "Please load a PDF and ensure images have been extracted."); return

        current_model_name = self.current_ollama_model.get()
        current_model_caps = self._get_model_capabilities(current_model_name)

        if not current_model_caps.get("vision"):
            messagebox.showwarning("Model Capability",
                                   f"Current model '{current_model_name}' may not support image analysis. "
                                   f"Please select a model with vision capabilities (e.g., 'llava', 'moondream', 'llama3.2-vision') for this feature to work correctly.")
            return # Stop if model does not support vision


        images_base64 = self.pdf_page_images[self.current_page_num]
        if not images_base64:
            messagebox.showinfo("No Images", f"No images were found or successfully extracted from page {self.current_page_num + 1} for analysis.");
            self.update_status("No images found on current page.")
            return

        num_images_found = len(images_base64)
        self.update_status(f"Analyzing {num_images_found} image(s) from current page...");
        self.add_to_chat("User", f"Analyze {num_images_found} image(s) on page {self.current_page_num + 1}")

        # Construct the specific instruction for the AI
        instruction_prompt = (f"Analyze the following {num_images_found} image(s) from a document page ({self.current_page_num + 1}).\n"
                              f"For each image:\n"
                              f"- Provide a detailed description of the visual content.\n"
                              f"- Explain any text visible in the image (e.g., labels, diagrams).\n"
                              f"- If it's a technical diagram, chart, or graph, explain its purpose and key information.\n"
                              f"- Connect the image content to the surrounding text context from the page (if available).\n" # The page text is included via _prepare_ai_prompt_and_context
                              f"- Highlight any key insights or patterns the image conveys.\n\n"
                              f"Present the analysis clearly, referring to each image. Use Markdown for formatting.")


        # Send the request in a separate thread.
        # Include page context here to help the AI connect images to the text.
        threading.Thread(target=self._threaded_ollama_request,
                         args=(f"Image Analysis ({num_images_found})", instruction_prompt, images_base64, True), # Label, instruction, images list, include page context
                         daemon=True).start()


    # --- Helper Methods ---

    def get_current_page_text(self, max_len=4000):
        """Retrieves text for the current page, truncated to a maximum length."""
        if self.pdf_document and self.pdf_page_text_for_ai and 0 <= self.current_page_num < len(self.pdf_page_text_for_ai):
             return self.pdf_page_text_for_ai[self.current_page_num].strip()[:max_len]
        return "" # Return empty string if no PDF, text not extracted, or invalid page

    def detect_code_blocks(self, text):
        """
        Simple heuristic to detect potential code blocks in text.
        Returns a list of suspected code blocks.
        """
        if not text: return []
        # Look for common code indicators within lines or paragraphs
        code_indicators = ["def ", "class ", "import ", "function ", "var ", "const ", "<?php", "<html>", "<script>", "public ", "private ", "void ", "int ", "string ", "if (", "for (", "while (", "try {", "{", "}", ";", "//", "#"]
        potential_blocks = []
        current_block = []
        in_potential_block = False

        lines = text.split('\n')
        for line in lines:
             stripped_line = line.strip()
             # A line with significant leading whitespace or containing indicators might be code
             is_code_line = (len(line) - len(stripped_line) > 4 or # Significant indentation
                             any(indicator in stripped_line for indicator in code_indicators) or
                             (stripped_line.startswith(('{', '}', '[', ']', '(', ')')) and len(stripped_line) < 10)) # Code structure on a line

             if is_code_line or in_potential_block:
                 current_block.append(line)
                 in_potential_block = True
                 # Simple way to potentially end a block: a blank line or a line that clearly isn't code after indentation check
                 if not stripped_line and len(current_block) > 1: # Blank line ends block
                      potential_blocks.append("\n".join(current_block).strip())
                      current_block = []
                      in_potential_block = False
                 elif in_potential_block and not is_code_line and len(line) - len(stripped_line) <= 4 and stripped_line:
                      # Line with little indentation and no indicators after being in a block
                      if len(current_block) > 1: potential_blocks.append("\n".join(current_block[:-1]).strip()) # Add previous lines
                      current_block = [line] # Start new block with current line
                      in_potential_block = False # Reset block status (might start a new one next line)


        # Add the last block if it's not empty
        if current_block:
             potential_blocks.append("\n".join(current_block).strip())

        # Filter out very short "blocks" that are likely not code
        return [block for block in potential_blocks if len(block) > 20] # Require minimum length


    # --- TTS Implementation ---

    def play_current_page_tts(self):
        """Initiates TTS playback for the text of the current PDF page."""
        if not (self.pdf_document and self.pdf_page_text_for_ai and 0 <= self.current_page_num < len(self.pdf_page_text_for_ai)):
            messagebox.showinfo("Not Ready", "Load a PDF and ensure text is extracted for TTS."); return

        text_to_speak = self.pdf_page_text_for_ai[self.current_page_num]
        if not text_to_speak.strip() or text_to_speak.startswith(("[No text found", "[Error extracting", "[Critical Extraction Error]")):
            messagebox.showinfo("No Text", "No valid text on the current page to speak."); return

        self.stop_current_page_tts() # Stop any ongoing TTS playback

        self.update_status(f"Generating speech for page {self.current_page_num + 1}...")

        try:
            selected_voice_id = self.selected_voice.get()

            # Create a temporary file to save the audio
            # Use delete=False to keep the file after the tempfile object is closed
            # It will be explicitly deleted later.
            with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as tmpfile:
                self.temp_audio_filename = tmpfile.name

            # Command to generate audio using edge-tts
            # subprocess.run will handle quoting arguments correctly
            command = [
                "edge-tts",
                "--voice", selected_voice_id,
                "--text", text_to_speak,
                "--write-media", self.temp_audio_filename
            ]

            # Run TTS generation in a separate thread to avoid blocking the UI
            threading.Thread(target=self._run_tts_command_and_play, args=(command, f"page {self.current_page_num + 1}"), daemon=True).start()

        except FileNotFoundError:
            self.handle_error("edge-tts command not found. Ensure it's installed (`pip install edge-tts`) and in your system's PATH.", "TTS Error")
            self._cleanup_temp_audio_file() # Clean up the failed temp file
            self._update_tts_button_states() # Update buttons after error
        except Exception as e:
            self.handle_error(f"Failed to prepare TTS for page {self.current_page_num + 1}: {str(e)}", "TTS Preparation Error")
            self._cleanup_temp_audio_file() # Clean up the failed temp file
            self._update_tts_button_states() # Update buttons after error


    def play_last_ai_response(self):
        """Initiates TTS playback for the text of the last AI response."""
        if not self.last_ai_response or not self.last_ai_response.strip():
            messagebox.showinfo("No Response", "No AI response available to speak.")
            return

        self.stop_current_page_tts()  # Stop any existing playback

        self.update_status("Generating speech for AI response...")

        try:
            selected_voice_id = self.selected_voice.get()

            # Create temp file
            with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as tmpfile:
                self.temp_audio_filename = tmpfile.name

            # Command to generate audio
            command = [
                "edge-tts",
                "--voice", selected_voice_id,
                "--text", self.last_ai_response, # Use the stored last AI response
                "--write-media", self.temp_audio_filename
            ]

            # Run TTS generation in a separate thread
            threading.Thread(target=self._run_tts_command_and_play, args=(command, "AI response"), daemon=True).start()

        except FileNotFoundError:
             self.handle_error("edge-tts command not found. Ensure it's installed and in your system's PATH.", "TTS Error")
             self._cleanup_temp_audio_file()
             self._update_tts_button_states() # Update buttons after error
        except Exception as e:
            self.handle_error(f"Failed to prepare TTS for AI response: {str(e)}", "TTS Preparation Error")
            self._cleanup_temp_audio_file()
            self._update_tts_button_states() # Update buttons after error


    def _run_tts_command_and_play(self, command, content_description):
        """Runs the edge-tts command and then plays the resulting audio file."""
        try:
            print(f"[DEBUG] Executing TTS generation command: {' '.join(subprocess.list2cmdline(command))}") # Log command

            # Run TTS generation using subprocess.run
            # capture_output=True to get stderr/stdout for better error messages
            # text=True to decode stdout/stderr as text
            # check=False so we can handle non-zero exit codes ourselves
            # creationflags to hide the console window on Windows
            creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0

            result = subprocess.run(command, capture_output=True, text=True, check=False, encoding='utf-8', errors='ignore', creationflags=creationflags)

            print(f"[DEBUG] TTS generation completed with code: {result.returncode}") # Debug log
            if result.returncode != 0:
                 # Handle TTS generation errors
                 error_msg = f"Failed to generate speech for {content_description} (edge-tts exited with code {result.returncode})."
                 if result.stderr:
                     print(f"[DEBUG] TTS stderr: {result.stderr.strip()}") # Debug log
                     error_msg += f"\nStderr: {result.stderr.strip()[:300]}..." # Limit length
                 elif result.stdout:
                     print(f"[DEBUG] TTS stdout: {result.stdout.strip()}") # Debug log
                     error_msg += f"\nStdout: {result.stdout.strip()[:300]}..." # Limit length

                 self.root.after(0, self.handle_error, error_msg, "TTS Generation Failed")
                 self._cleanup_temp_audio_file() # Clean up potentially created partial file
                 self.root.after(0, self._update_tts_button_states) # Update buttons after error
                 return

            # Check if the audio file was actually created and has content
            if not self.temp_audio_filename or not os.path.exists(self.temp_audio_filename) or os.path.getsize(self.temp_audio_filename) == 0:
                 error_msg = f"TTS generation completed, but output file is missing or empty: {self.temp_audio_filename}"
                 self.root.after(0, self.handle_error, error_msg, "TTS Output Error")
                 self._cleanup_temp_audio_file()
                 self.root.after(0, self._update_tts_button_states)
                 return


            # Play audio file - needs to happen on the main thread or managed externally
            if self.root: self.root.after(0, self._play_audio_file, self.temp_audio_filename, content_description)

        except FileNotFoundError:
            # This specific error means the 'edge-tts' executable itself wasn't found in PATH
            self.root.after(0, self.handle_error, "edge-tts command not found. Ensure it's installed and in your system's PATH.", "TTS Error")
            self._cleanup_temp_audio_file()
            self.root.after(0, self._update_tts_button_states)
        except Exception as e:
            # Catch any other unexpected errors during the process
            self.root.after(0, self.handle_error, f"An unexpected error occurred during TTS generation: {str(e)}", "Unexpected TTS Error")
            self._cleanup_temp_audio_file()
            self.root.after(0, self._update_tts_button_states)


    def _play_audio_file(self, filename, content_description):
        """Platform-independent audio playback using ffplay."""
        if not filename or not os.path.exists(filename):
            self.root.after(0, self.update_status, "Audio file not found for playback.")
            self._update_tts_button_states()
            return

        try:
            # Check if ffplay is available
            if not self._is_ffplay_available():
                self.root.after(0, self._show_ffmpeg_install_instructions)
                self._cleanup_temp_audio_file() # Clean up the generated file if ffplay is missing
                self._update_tts_button_states()
                return

            # Command to play audio using ffplay
            command = [
                "ffplay",
                "-nodisp",  # No visual display
                "-autoexit", # Exit when playback finishes
                "-loglevel", "warning", # Suppress verbose output, show warnings/errors
                filename
            ]

            # Platform-specific adjustments for subprocess creation
            creationflags = 0
            if sys.platform == 'win32':
                creationflags = subprocess.CREATE_NO_WINDOW # Hide the console window

            print(f"[DEBUG] Executing audio playback command: {' '.join(subprocess.list2cmdline(command))}") # Log command

            # Use Popen to run in the background and get a process object
            self.tts_process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE, # Capture stdout
                stderr=subprocess.PIPE, # Capture stderr
                creationflags=creationflags
            )

            # Update UI elements on the main thread
            self.root.after(0, lambda: [
                self.play_tts_btn.config(state=tk.DISABLED),
                self.play_ai_tts_btn.config(state=tk.DISABLED),
                self.stop_tts_btn.config(state=tk.NORMAL), # Enable stop button
                self.update_status(f"Playing {content_description}...")
            ])

            # Start a thread to monitor the playback process
            threading.Thread(target=self._monitor_playback_process, args=(self.tts_process, content_description), daemon=True).start()

        except FileNotFoundError:
            # This specific error means 'ffplay' executable was not found
            self.root.after(0, self._show_ffmpeg_install_instructions)
            self._cleanup_temp_audio_file()
            self._update_tts_button_states()
        except Exception as e:
            error_msg = f"Failed to start playback process for {content_description}: {str(e)}"
            self.root.after(0, self.handle_error, error_msg, "Playback Error")
            self._cleanup_temp_audio_file()
            self._update_tts_button_states()

    def _is_ffplay_available(self):
        """Checks if ffplay executable is available in the system PATH."""
        try:
            # Run ffplay with a simple command that should succeed if it exists
            # Use check=True to raise CalledProcessError on non-zero exit
            # Use DEVNULL to suppress output
            # creationflags to hide window on Windows
            creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
            subprocess.run(
                ["ffplay", "-version"],
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=creationflags
            )
            print("[DEBUG] ffplay check passed.") # Debug log
            return True
        except (FileNotFoundError, subprocess.CalledProcessError):
            print("[DEBUG] ffplay check failed: FileNotFoundError or CalledProcessError.") # Debug log
            return False
        except Exception as e:
            print(f"[DEBUG] Unexpected error during ffplay check: {e}") # Debug log
            return False

    def _show_ffmpeg_install_instructions(self):
        """Displays instructions for installing FFmpeg (which includes ffplay)."""
        install_text = (
            "FFmpeg (including ffplay) is required for audio playback.\n\n"
            "Please install FFmpeg and ensure the 'ffplay' executable is in your system's PATH.\n"
            "Common methods:\n"
            "Windows: Download a build from a reputable source like gyan.dev or BtbN.\n"
            "  (Ensure the bin folder containing ffplay.exe is added to your System PATH environment variable)\n"
            "macOS: brew install ffmpeg (using Homebrew)\n"
            "Linux (Debian/Ubuntu): sudo apt update && sudo apt install ffmpeg\n\n"
            "More detailed download information can be found at:\n"
        )
        link_url = "https://ffmpeg.org/download.html" # Official download page

        # Create a new Toplevel window for instructions
        install_win = tk.Toplevel(self.root)
        install_win.title("FFmpeg Installation Required")
        install_win.transient(self.root) # Make it a transient window relative to the main window
        install_win.grab_set() # Make it modal (block interaction with main window)
        install_win.resizable(False, False)

        # Use Text widget for clickable link
        msg_text_widget = tk.Text(install_win, wrap=tk.WORD, height=12, width=70, padx=10, pady=10, bg=self.text_widget_bg, fg=self.text_widget_fg, font=self.text_widget_font, relief=tk.FLAT)
        msg_text_widget.insert(tk.END, install_text)

        # Add the clickable link
        msg_text_widget.insert(tk.END, link_url, "link")
        msg_text_widget.tag_configure("link", foreground="blue", underline=True)
        msg_text_widget.tag_bind("link", "<Button-1>", lambda e: webbrowser.open(link_url))

        msg_text_widget.config(state=tk.DISABLED) # Make text widget read-only
        msg_text_widget.pack(padx=10, pady=10)

        # Add a Close button
        ttk.Button(install_win, text="Close", command=install_win.destroy).pack(pady=(0, 10))

        # Position the new window centrally relative to the main window (optional but nice)
        self.root.update_idletasks() # Ensure main window geometry is up to date
        main_x = self.root.winfo_x()
        main_y = self.root.winfo_y()
        main_width = self.root.winfo_width()
        main_height = self.root.winfo_height()

        install_win.update_idletasks() # Ensure new window geometry is up to date
        win_width = install_win.winfo_width()
        win_height = install_win.winfo_height()

        new_x = main_x + (main_width // 2) - (win_width // 2)
        new_y = main_y + (main_height // 2) - (win_height // 2)

        install_win.geometry(f"+{new_x}+{new_y}")

        install_win.focus_set() # Set focus to the new window


    def _monitor_playback_process(self, process, content_description):
        """Monitors the TTS playback subprocess and updates button states when it finishes."""
        print(f"[DEBUG] Monitoring playback process for {content_description}...") # Debug log
        try:
            # Wait for the process to finish and capture output
            stdout, stderr = process.communicate()
            returncode = process.returncode

            print(f"[DEBUG] Playback process for {content_description} finished with code: {returncode}") # Debug log
            if stdout: print(f"[DEBUG] Playback stdout:\n{stdout.decode()}")
            if stderr: print(f"[DEBUG] Playback stderr:\n{stderr.decode()}")


            # Schedule UI updates on the main thread
            if self.root: self.root.after(0, self._playback_finished, content_description, returncode, stderr.decode() if stderr else None)

        except Exception as e:
            print(f"Error monitoring playback process: {e}")
            # Ensure UI updates happen on the main thread even if monitoring fails
            if self.root: self.root.after(0, self._playback_finished, content_description, -1, f"Monitoring error: {e}") # Use -1 to indicate monitoring failure


    def _playback_finished(self, content_description, returncode, stderr):
        """Called on the main thread when playback subprocess finishes."""
        print(f"[DEBUG] Main thread received playback finished signal for {content_description}.") # Debug log
        self.tts_process = None # Clear the process reference immediately
        self._cleanup_temp_audio_file() # Clean up the temp file

        if returncode == 0:
             self.update_status(f"Audio playback finished for {content_description}.")
        else:
             error_msg = f"Playback of {content_description} failed or ended with error (code {returncode})."
             if stderr:
                  error_msg += f"\nDetails: {stderr.strip()[:200]}..."
             self.update_status(error_msg)
             # Optionally add error to chat if it's a significant playback failure
             # self.add_to_chat("System", error_msg, "system")


        # Update buttons to the correct state (enabling play buttons, disabling stop)
        self._update_tts_button_states()


    def stop_current_page_tts(self):
        """Stops the currently running TTS playback process."""
        if self.tts_process and self.tts_process.poll() is None: # Check if process exists and is still running
            try:
                print("[DEBUG] Attempting to stop TTS process...") # Debug log
                # Terminate the process gracefully first
                self.tts_process.terminate()
                try:
                    # Wait a bit for termination, then kill if needed
                    # Use a timeout to avoid hanging indefinitely
                    self.tts_process.wait(timeout=0.5) # Wait up to 0.5 seconds
                    print("[DEBUG] TTS process terminated gracefully.") # Debug log
                except subprocess.TimeoutExpired:
                    self.tts_process.kill() # Force kill if termination hangs
                    print("[DEBUG] TTS process killed.") # Debug log
            except Exception as e:
                print(f"Error stopping TTS process: {e}")
            finally:
                # Ensure process reference is cleared and UI updated
                self.tts_process = None
                self._cleanup_temp_audio_file()
                self.update_status("Audio playback stopped.")
                self._update_tts_button_states() # Update buttons after stopping
        else:
             print("[DEBUG] No active TTS process to stop.") # Debug log


    def _update_tts_button_states(self):
        """Updates the states of the TTS and Voice Query buttons."""
        is_pdf_loaded = self.pdf_document is not None and len(self.pdf_page_text_for_ai) > 0
        # Check if the current page text is valid (not just an error/placeholder)
        current_page_has_text = is_pdf_loaded and 0 <= self.current_page_num < len(self.pdf_page_text_for_ai) and \
                                self.pdf_page_text_for_ai[self.current_page_num].strip() and \
                                not self.pdf_page_text_for_ai[self.current_page_num].startswith(("[No text found", "[Error extracting", "[Critical Extraction Error]"))

        has_ai_response = bool(self.last_ai_response.strip()) # Check if last AI response has actual text
        is_playing = self.tts_process is not None and self.tts_process.poll() is None # Check if a playback process is active and running


        # Update Play Page button: Enabled if page has text AND not currently playing
        if hasattr(self.play_tts_btn, 'config'):
            self.play_tts_btn.config(state=tk.NORMAL if current_page_has_text and not is_playing else tk.DISABLED)

        # Update Play AI button: Enabled if there's an AI response AND not currently playing
        if hasattr(self.play_ai_tts_btn, 'config'):
            self.play_ai_tts_btn.config(state=tk.NORMAL if has_ai_response and not is_playing else tk.DISABLED)

        # Update Stop button: Enabled only if currently playing
        if hasattr(self.stop_tts_btn, 'config'):
            self.stop_tts_btn.config(state=tk.NORMAL if is_playing else tk.DISABLED)

        # Update Voice Query button: Enabled if SpeechRecognition is available AND a model is ready AND not currently playing
        is_model_valid = self.current_ollama_model.get() not in ["Loading...", "No Models Found", "Ollama Offline", "Ollama Timeout", "Error Fetching"] and bool(self.current_ollama_model.get())
        if hasattr(self.voice_query_btn, 'config'):
             self.voice_query_btn.config(
                 state=tk.NORMAL if self.voice_query_available and is_model_valid and not is_playing else tk.DISABLED,
                 text="🎙️ Voice Query" if not is_playing else "🎙️ ..." # Change text while listening
             )


    def _cleanup_temp_audio_file(self):
        """Safely attempts to delete the temporary audio file."""
        if self.temp_audio_filename and os.path.exists(self.temp_audio_filename):
            print(f"[DEBUG] Attempting to clean up temp file: {self.temp_audio_filename}") # Debug log
            # Add a small delay and retry logic in case the file is briefly locked by the player
            max_retries = 5
            for i in range(max_retries):
                try:
                    os.remove(self.temp_audio_filename)
                    print(f"[DEBUG] Cleaned up temp audio file: {self.temp_audio_filename}") # Debug log
                    break # Success
                except PermissionError:
                    print(f"[DEBUG] PermissionError cleaning up {self.temp_audio_filename}, retry {i+1}/{max_retries}") # Debug log
                    if i < max_retries - 1:
                        time.sleep(0.1 * (i + 1)) # Wait a bit longer on each retry
                    else:
                        print(f"PermissionError: Could not delete temporary audio file {self.temp_audio_filename} after multiple retries. It might still be in use.")
                except Exception as e:
                    print(f"Note: Could not delete temporary audio file {self.temp_audio_filename}: {str(e)}")
                    break # Other error, give up
            # Always clear the filename reference even if deletion fails
            self.temp_audio_filename = None


    def toggle_auto_play(self):
        """Handles auto-play toggle state changes and updates status/chat."""
        state = "ON" if self.auto_play_ai.get() else "OFF"
        self.add_to_chat("System", f"Auto-play AI responses {state}", "system")
        self.update_status(f"Auto-play AI responses {state}")


    # --- Voice Query Implementation ---

    def handle_voice_query(self):
        """Starts the voice recognition process."""
        # Check again if SpeechRecognition is available and model is ready
        is_model_valid = self.current_ollama_model.get() not in ["Loading...", "No Models Found", "Ollama Offline", "Ollama Timeout", "Error Fetching"] and bool(self.current_ollama_model.get())

        if not self.voice_query_available or not is_model_valid:
            messagebox.showinfo("Voice Query Not Ready", "Voice query dependencies are not installed, or the AI model is not ready.")
            self._update_tts_button_states() # Ensure button state is correct
            return

        # Disable voice query button while listening
        if hasattr(self.voice_query_btn, 'config'):
             self.voice_query_btn.config(state=tk.DISABLED, text="🎙️ Listening...")

        self.update_status("Listening... Speak your question now.")
        print("[DEBUG] Starting voice query listening...") # Debug log

        # Run the voice recognition process in a separate thread
        threading.Thread(target=self._voice_recognition_thread, daemon=True).start()


    def _voice_recognition_thread(self):
        """Background thread for voice recognition."""
        r = sr.Recognizer()
        try:
            # Adjust for ambient noise for a few seconds
            print("[DEBUG] Adjusting for ambient noise...") # Debug log
            with sr.Microphone() as source:
                 r.adjust_for_ambient_noise(source, duration=3) # Listen for 3 seconds to calibrate
                 print("[DEBUG] Ambient noise adjustment complete. Listening for speech...") # Debug log
                 self.root.after(0, self.update_status, "Listening... Speak your question now.") # Re-confirm status after adjustment

                 # Listen to the source for up to 10 seconds
                 audio = r.listen(source, timeout=10, phrase_time_limit=10) # listen for up to 10s, stop processing phrase after 10s

            print("[DEBUG] Speech detected, attempting transcription...") # Debug log
            self.root.after(0, self.update_status, "Processing speech...") # Update status while processing

            # Recognize speech using Google Web Speech API
            # Use a try-except block around recognition as it might fail
            try:
                text = r.recognize_google(audio)
                print(f"[DEBUG] Transcription successful: '{text}'") # Debug log

                # Schedule UI updates and AI interaction on the main thread
                self.root.after(0, lambda: [
                    self.user_question_entry.delete(0, tk.END), # Clear entry
                    self.user_question_entry.insert(0, text), # Insert transcribed text
                    self.send_question_to_ai(), # Send the question to AI
                    self._update_tts_button_states(), # Re-enable voice button
                    self.update_status("Voice query processed.")
                ])
            except sr.UnknownValueError:
                print("[DEBUG] Google Speech Recognition could not understand audio.") # Debug log
                self.root.after(0, lambda: [
                    self.update_status("Could not understand audio. Please try again."),
                    self._update_tts_button_states() # Re-enable voice button
                ])
            except sr.RequestError as e:
                print(f"[DEBUG] Could not request results from Google Speech Recognition service; {e}") # Debug log
                self.root.after(0, lambda: [
                    self.update_status(f"Speech recognition error: {e}"),
                    self._update_tts_button_states() # Re-enable voice button
                ])


        except sr.WaitTimeoutError:
            print("[DEBUG] Listening timed out, no speech detected.") # Debug log
            self.root.after(0, lambda: [
                self.update_status("No speech detected."),
                self._update_tts_button_states() # Re-enable voice button
            ])
        except OSError as e: # Handle microphone access errors (e.g., no microphone, permissions)
            print(f"[DEBUG] Microphone access error: {e}") # Debug log
            error_message = "Microphone access error. Please check your microphone and permissions."
            if "No default input device" in str(e):
                 error_message = "No microphone found. Please ensure a microphone is connected and configured."
            self.root.after(0, lambda: [
                self.handle_error(error_message, "Voice Input Error"),
                self._update_tts_button_states() # Re-enable voice button
            ])
        except Exception as e:
            print(f"[DEBUG] An unexpected voice query error occurred: {e}") # Debug log
            self.root.after(0, lambda: [
                self.handle_error(f"An unexpected voice query error occurred: {str(e)}", "Voice Error"),
                self._update_tts_button_states() # Re-enable voice button
            ])
        finally:
            # Ensure button state is reset even if other errors occurred
             self.root.after(0, self._update_tts_button_states)


    # --- Application Lifecycle ---

    def on_closing(self):
        """Handles cleanup when the application window is closed."""
        print("[DEBUG] Application closing.") # Debug log
        self.stop_current_page_tts() # Stop any running TTS process
        if self.pdf_document:
            try: self.pdf_document.close() # Close the PDF document
            except Exception as e: print(f"Error closing PDF on exit: {str(e)}")
        self._cleanup_temp_audio_file() # Clean up any temporary audio file
        if hasattr(self.root, 'destroy'):
            self.root.destroy() # Destroy the main window
        # Using sys.exit(0) is a clean way to ensure all threads (like the monitor thread) exit
        sys.exit(0)


# --- Main execution block ---
if __name__ == "__main__":
    # Basic Python version check
    if sys.version_info < (3, 7):
        messagebox.showerror("Python Version Error", "This application requires Python 3.7 or newer.")
        sys.exit(1)

    # Check for edge-tts before starting the GUI
    # This check needs to happen after the PATH modification for Windows
    print("[DEBUG] Checking for edge-tts executable...") # Debug log
    try:
        # Use check=True to raise CalledProcessError on non-zero exit status
        # Use creationflags to suppress console window on Windows
        creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
        # Run edge-tts with a basic command to check its existence and executability
        # Using --help or --help is common
        subprocess.run(["edge-tts", "--help"], check=True, capture_output=True, text=True, encoding='utf-8', errors='ignore', creationflags=creationflags)
        print("[DEBUG] edge-tts check passed.") # Debug log

    except FileNotFoundError:
        # edge-tts command was not found at all
        message_text = "The 'edge-tts' command was not found.\n"
        message_text += "Please ensure it is installed (`pip install edge-tts`) and that your Python Scripts directory is in your system's PATH.\n"
        if sys.platform == 'win32':
             message_text += f"Common Scripts path: {os.path.join(os.path.dirname(sys.executable), 'Scripts')}\n"
        message_text += "Application will exit."
        messagebox.showerror("Dependency Error", message_text)
        sys.exit(1)
    except subprocess.CalledProcessError as e:
         # edge-tts command was found but returned an error (e.g., invalid option for version check)
         error_msg = f"The 'edge-tts' command was found, but the version check failed (exit code {e.returncode})."
         if e.stderr:
             error_msg += f"\nStderr: {e.stderr.strip()[:300]}..." # Limit stderr length
         elif e.stdout:
             error_msg += f"\nStdout: {e.stdout.strip()[:300]}..." # Limit stdout length
         error_msg += "\nPlease ensure edge-tts is correctly installed and can run.\nTry running 'edge-tts --help' or 'edge-tts --help' in your terminal.\nApplication will exit."
         messagebox.showerror("Dependency Error", error_msg)
         sys.exit(1)
    except Exception as e:
         # Other unexpected error during the check
         messagebox.showerror("Dependency Error", f"An unexpected error occurred while checking 'edge-tts': {str(e)}\nApplication will exit.")
         sys.exit(1)


    # Check for ffplay (part of FFmpeg) before starting the GUI
    # This check should also happen after the PATH modification for Windows
    print("[DEBUG] Checking for ffplay executable...") # Debug log
    try:
         creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
         subprocess.run(["ffplay", "-version"], check=True, capture_output=True, text=True, encoding='utf-8', errors='ignore', creationflags=creationflags)
         print("[DEBUG] ffplay check passed.") # Debug log
    except FileNotFoundError:
         # ffplay command was not found at all
         # Don't exit here, as the app can still function without audio playback,
         # but show a warning. The playback function will also check and warn.
         print("[WARNING] ffplay command not found. Audio playback will be disabled.")
         # messagebox.showwarning("Dependency Warning", "The 'ffplay' command (part of FFmpeg) was not found.\nAudio playback will be disabled. Please install FFmpeg and ensure ffplay is in your system's PATH.")
         # No messagebox here, will show instructions on first attempt to play audio.
    except subprocess.CalledProcessError as e:
         print(f"[WARNING] ffplay command found, but version check failed (code {e.returncode}). Playback might not work.")
    except Exception as e:
         print(f"[WARNING] Unexpected error during ffplay check: {e}. Playback might not work.")


    main_app_root = None
    try:
        # Create the main Tkinter window
        main_app_root = tk.Tk()

        # Create an instance of the PDFToSpeechApp
        app_instance = PDFToSpeechApp(main_app_root)

        # Start the Tkinter event loop
        main_app_root.mainloop()

    except Exception as e:
        # Catch any unhandled exceptions during GUI startup or main loop
        import traceback
        critical_error_message = f"A critical error occurred outside a specific handler:\n{str(e)}\n\nTraceback:\n{traceback.format_exc()}"
        print(critical_error_message, file=sys.stderr)
        try:
            # Attempt to show an error dialog
            error_dialog_root = tk.Tk()
            error_dialog_root.withdraw() # Hide the main window for the dialog
            messagebox.showerror("Application Critical Error", critical_error_message)
            error_dialog_root.destroy()
        except tk.TclError:
            # Handle cases where Tkinter itself might be broken
            print("Failed to show error dialog using Tkinter.", file=sys.stderr)

        sys.exit(1) # Exit with a non-zero status code to indicate an error
