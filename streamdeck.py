import serial
import time
from serial.tools import list_ports
import pyautogui
import keyboard
import tkinter as tk
from tkinter import ttk, messagebox
import threading
from PIL import Image, ImageTk
import os
import sys
import json
import ctypes
from ctypes import wintypes

# Add Windows API constants and functions
KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP = 0x0002

# Virtual key codes for F1-F24
VK_F1 = 0x70
VK_F13 = 0x7C

# Define ULONG_PTR based on system architecture
ULONG_PTR = ctypes.c_ulong if ctypes.sizeof(ctypes.c_void_p) == 4 else ctypes.c_ulonglong

user32 = ctypes.WinDLL('user32', use_last_error=True)
user32.keybd_event.argtypes = [ctypes.c_byte, ctypes.c_byte, wintypes.DWORD, ULONG_PTR]

def press_function_key(key_number):
    """Press a function key using Windows API."""
    if 1 <= key_number <= 24:
        # Calculate virtual key code
        vk = VK_F1 + (key_number - 1) if key_number <= 12 else VK_F13 + (key_number - 13)
        
        try:
            # Press the key
            user32.keybd_event(vk, 0, KEYEVENTF_EXTENDEDKEY, 0)
            # Small delay between press and release
            time.sleep(0.05)
            # Release the key
            user32.keybd_event(vk, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0)
            # Small delay after key press
            time.sleep(0.05)
        except Exception as e:
            print(f"Error pressing function key {key_number}: {e}")

def throttle(seconds=0):
    """A decorator that prevents a function from being called more than once every `seconds` seconds."""
    def decorator(func):
        last_called = [0]  # Using list to maintain state in closure
        def wrapper(*args, **kwargs):
            now = time.time()
            if now - last_called[0] >= seconds:
                last_called[0] = now
                return func(*args, **kwargs)
        return wrapper
    return decorator

class ModernStreamDeckApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Stream Deck Kontrol Paneli")
        self.root.geometry("800x700")
        self.root.minsize(700, 600)

        # Add recording state variables
        self.recording_hotkey = False
        self.current_recording_entry = None
        self.current_recording_button = None

        # Add throttling for intensive operations
        self.last_arduino_check = 0
        self.arduino_check_interval = 0.1  # Check every 100ms instead of continuous polling
        self.last_ui_update = 0
        self.ui_update_interval = 0.5  # Update UI every 500ms

        self.settings_file = os.path.join(os.path.expanduser("~"), "Documents", "StreamDeckSettings.json")
        self.theme_mode = tk.StringVar(value="dark")
        
        # Initialize variables and command types first
        self.command_types = ["yazı", "press", "hotkey", "volume", "media"]
        self.press_keys = [
            "press:enter", "press:esc", "press:tab", "press:space", "press:backspace",
            "press:delete", "press:up", "press:down", "press:left", "press:right",
            "press:f1", "press:f2", "press:f3", "press:f4", "press:f5", "press:f6",
            "press:f7", "press:f8", "press:f9", "press:f10", "press:f11", "press:f12",
            "press:f13", "press:f14", "press:f15", "press:f16", "press:f17", "press:f18",
            "press:f19", "press:f20", "press:f21", "press:f22", "press:f23", "press:f24"
        ]
        self.volume_keys = ["volume:up", "volume:down", "volume:mute"]
        self.media_keys = ["media:play/pause", "media:next", "media:previous", "media:stop"]

        # Initialize command variables
        self.command_vars = {}
        buttons = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
        for button in buttons:
            self.command_vars[button] = {
                'type': tk.StringVar(value="yazı"),
                'subtype': tk.StringVar(),
                'entry': tk.StringVar(value=f"{button} Butonu işlevi"),
                'message': f"{button} Butonu işlevi"
            }

        # Load settings after initializing variables
        self.load_settings()
        
        # Set theme
        self.set_theme(self.theme_mode.get())
        
        # Create main container
        self.main_container = ttk.Frame(self.root)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Create UI components
        self.create_header()
        self.create_command_interface_with_scrollbar()
        self.create_footer()
        
        # Arduino connection setup
        self.arduino = None
        self.arduino_connected = False
        self.arduino_lock = threading.Lock()
        
        # Start Arduino monitoring in a separate thread
        self.arduino_thread = threading.Thread(target=self.monitor_arduino, daemon=True)
        self.arduino_thread.start()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def save_settings(self):
        """Throttled settings save"""
        try:
            settings = {
                "theme_mode": self.theme_mode.get(),
                "buttons": {
                    button: {
                        "command_type": vars['type'].get(),
                        "command_subtype": vars['subtype'].get(),
                        "command_text": vars['entry'].get(),
                        "message": vars['message']
                    }
                    for button, vars in self.command_vars.items()
                }
            }
            
            os.makedirs(os.path.dirname(self.settings_file), exist_ok=True)
            
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
            
        except Exception as e:
            print(f"Settings save error: {e}")

    def load_settings(self):
        """Optimized settings loading"""
        try:
            if not os.path.exists(self.settings_file):
                return

            with open(self.settings_file, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            
            if "theme_mode" in settings:
                self.theme_mode.set(settings["theme_mode"])
            
            if "buttons" in settings:
                for button, data in settings["buttons"].items():
                    if button in self.command_vars:
                        vars = self.command_vars[button]
                        vars['type'].set(data.get("command_type", "yazı"))
                        vars['subtype'].set(data.get("command_subtype", ""))
                        vars['entry'].set(data.get("command_text", ""))
                        vars['message'] = data.get("message", "")
                        
        except Exception as e:
            print(f"Settings load error: {e}")

    def on_closing(self):
        try:
            self.update_message(auto_save=True)  # Önce son komutları güncelle
            self.save_settings()  # Sonra ayarları kaydet
            self.root.destroy()
        except Exception as e:
            print(f"Uygulama kapatılırken hata oluştu: {e}")
            self.root.destroy()

    def set_theme(self, mode):
        # Tema renkleri tanımla
        if mode == "light":
            self.primary_color = "#3498db"
            self.accent_color = "#2ecc71"
            self.warning_color = "#e74c3c"
            self.bg_color = "#f9f9f9"
            self.card_bg_color = "white"
            self.text_color = "#2c3e50"
            self.label_bg = self.bg_color
            self.secondary_text = "#7f8c8d"
            self.entry_bg = "white"
            self.entry_fg = "#2c3e50"
            self.combo_fieldbackground = "white"
            self.combo_background = "white"
            self.combo_foreground = "#2c3e50"
            self.combo_selectbackground = "#3498db"
            self.combo_selectforeground = "white"
        else:  # dark mode
            self.primary_color = "#3498db"
            self.accent_color = "#2ecc71"
            self.warning_color = "#e74c3c"
            self.bg_color = "#1e1e2e"
            self.card_bg_color = "#2d2d42"
            self.text_color = "#e0e0e0"
            self.label_bg = self.bg_color
            self.secondary_text = "#a0a0a0"
            self.entry_bg = "#3a3a5a"
            self.entry_fg = "#e0e0e0"
            self.combo_fieldbackground = "#3a3a5a"
            self.combo_background = "#3a3a5a"
            self.combo_foreground = "#e0e0e0"
            self.combo_selectbackground = "#3498db"
            self.combo_selectforeground = "#e0e0e0"
        
        # ttk teması oluştur
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        # Arkaplan rengi
        self.root.configure(bg=self.bg_color)
        
        # Widget stilleri
        self.style.configure("TFrame", background=self.bg_color)
        self.style.configure("TLabel", background=self.bg_color, foreground=self.text_color, font=("Segoe UI", 10))
        
        # Entry stil ayarları
        self.style.configure("TEntry",
                          fieldbackground=self.entry_bg,
                          foreground=self.entry_fg,
                          insertcolor=self.text_color)
        self.style.map("TEntry",
                    fieldbackground=[("readonly", self.entry_bg)],
                    foreground=[("readonly", self.entry_fg)])
        
        # Combobox stil ayarları
        self.style.configure("TCombobox",
                          fieldbackground=self.combo_fieldbackground,
                          background=self.combo_background,
                          foreground=self.combo_foreground,
                          arrowcolor=self.text_color,
                          selectbackground=self.combo_selectbackground,
                          selectforeground=self.combo_selectforeground)
        self.style.map("TCombobox",
                    fieldbackground=[("readonly", self.combo_fieldbackground)],
                    selectbackground=[("readonly", self.combo_selectbackground)],
                    selectforeground=[("readonly", self.combo_selectforeground)],
                    background=[("readonly", self.combo_background)],
                    foreground=[("readonly", self.combo_foreground)])
                        
        self.style.configure("TButton", font=("Segoe UI", 10))
        self.style.configure("Header.TLabel", font=("Segoe UI", 18, "bold"), foreground=self.primary_color, background=self.bg_color)
        self.style.configure("Subheader.TLabel", font=("Segoe UI", 12, "bold"), foreground=self.text_color, background=self.bg_color)
        
        # Özel buton stilleri
        self.style.configure("Accent.TButton",
                             background=self.accent_color,
                             foreground="white",
                             padding=(15, 8),
                             font=("Segoe UI", 11, "bold"))
        self.style.map("Accent.TButton",
                       background=[("active", "#27ae60"), ("pressed", "#219653")],
                       foreground=[("active", "white"), ("pressed", "white")])
        
        # Tema değiştirme butonu için stil
        self.style.configure("Theme.TButton",
                            padding=(10, 5),
                            font=("Segoe UI", 9))
        
        # LabelFrame stili
        self.style.configure("Card.TLabelframe", background=self.card_bg_color, borderwidth=0)
        self.style.configure("Card.TLabelframe.Label",
                           background=self.card_bg_color,
                           foreground=self.primary_color,
                           font=("Segoe UI", 12, "bold"))
        
        # Status stilleri
        self.style.configure("Connected.TLabel", foreground=self.accent_color, font=("Segoe UI", 10), background=self.bg_color)
        self.style.configure("Waiting.TLabel", foreground="#f39c12", font=("Segoe UI", 10), background=self.bg_color)
        self.style.configure("Disconnected.TLabel", foreground=self.warning_color, font=("Segoe UI", 10), background=self.bg_color)
        
        # Bilgi metni stili
        self.style.configure("Info.TLabel", foreground=self.secondary_text, font=("Segoe UI", 9), background=self.bg_color)
        
        # Kart içi frame'lerin arkaplan rengini ayarla
        self.style.configure("CardInner.TFrame", background=self.card_bg_color)
        self.style.configure("CardLabel.TLabel", background=self.card_bg_color, foreground=self.text_color)
        
        # Scrollbar stilleri
        self.style.configure("TScrollbar", 
                         background=self.card_bg_color, 
                         troughcolor=self.bg_color, 
                         borderwidth=0,
                         arrowcolor=self.text_color)
        self.style.map("TScrollbar",
                    background=[("active", self.primary_color)],
                    arrowcolor=[("active", "white")])

    def configure_dark_mode_combobox(self):
        """Koyu mod için combobox özel ayarları"""
        try:
            for child in self.root.winfo_children():
                self.configure_comboboxes_recursive(child)
        except Exception as e:
            print(f"Combobox yapılandırma hatası: {e}")
            
    def configure_comboboxes_recursive(self, parent):
        """Tüm alt widget'ları bularak combobox'ları yapılandır"""
        for child in parent.winfo_children():
            if isinstance(child, ttk.Combobox):
                child.configure(
                    background=self.combo_background,
                    foreground=self.combo_foreground,
                    selectbackground=self.combo_selectbackground,
                    selectforeground=self.combo_selectforeground
                )
                child.tk.eval(f"""
                    [ttk::combobox::PopdownWindow %s].f.l configure -background {self.combo_background} -foreground {self.combo_foreground}
                    [ttk::combobox::PopdownWindow %s].f.l configure -selectbackground {self.combo_selectbackground} -selectforeground {self.combo_selectforeground}
                """ % (child, child))
            
            if len(child.winfo_children()) > 0:
                self.configure_comboboxes_recursive(child)

    def toggle_theme(self):
        new_mode = "dark" if self.theme_mode.get() == "light" else "light"
        self.theme_mode.set(new_mode)
        self.set_theme(new_mode)
        
        self.main_container.destroy()
        self.main_container = ttk.Frame(self.root)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        self.create_header()
        self.create_command_interface_with_scrollbar()  # Scrollbar'lı sürümü kullan
        self.create_footer()
        
        if new_mode == "dark":
            self.configure_dark_mode_combobox()
        
        theme_text = "☀️ Açık Mod" if new_mode == "dark" else "🌙 Koyu Mod"
        self.theme_button.configure(text=theme_text)
        
        # Bağlantı durumunu yenile
        if self.arduino_connected:
            self.update_status_indicator("connected")
        else:
            self.update_status_indicator("waiting")

    def create_header(self):
        header_frame = ttk.Frame(self.main_container)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.logo_text = ttk.Label(header_frame, text="STREAM DECK", style="Header.TLabel")
        self.logo_text.pack(side=tk.LEFT)
        
        right_frame = ttk.Frame(header_frame)
        right_frame.pack(side=tk.RIGHT)
        
        theme_text = "☀️ Açık Mod" if self.theme_mode.get() == "dark" else "🌙 Koyu Mod"
        self.theme_button = ttk.Button(right_frame,
                                     text=theme_text,
                                     style="Theme.TButton",
                                     command=self.toggle_theme)
        self.theme_button.pack(side=tk.RIGHT, padx=(10, 0))
        
        self.status_frame = ttk.Frame(right_frame)
        self.status_frame.pack(side=tk.RIGHT, padx=10)
        
        self.status_indicator = tk.Frame(self.status_frame, width=12, height=12, bg="#f39c12")
        self.status_indicator.pack(side=tk.LEFT, padx=(0, 5))
        
        self.status_label = ttk.Label(self.status_frame, text="Arduino bağlantısı bekleniyor...", style="Waiting.TLabel")
        self.status_label.pack(side=tk.LEFT)

    def update_status_indicator(self, status):
        if status == "connected":
            self.status_indicator.config(bg=self.accent_color)
            self.status_label.configure(text="Arduino bağlı", style="Connected.TLabel")
            self.arduino_connected = True
        elif status == "waiting":
            self.status_indicator.config(bg="#f39c12")
            self.status_label.configure(text="Arduino bağlantısı bekleniyor...", style="Waiting.TLabel")
        elif status == "disconnected":
            self.status_indicator.config(bg=self.warning_color)
            self.status_label.configure(text="Bağlantı koptu! Yeniden deneniyor...", style="Disconnected.TLabel")
            self.arduino_connected = False

    def create_command_interface_with_scrollbar(self):
        """Optimized command interface creation with reduced redraws"""
        commands_outer_frame = ttk.Frame(self.main_container)
        commands_outer_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create canvas with scrollbar
        canvas_frame = ttk.Frame(commands_outer_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        canvas = tk.Canvas(canvas_frame, bg=self.bg_color, highlightthickness=0)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar.config(command=canvas.yview)
        canvas.config(yscrollcommand=scrollbar.set)
        
        commands_frame = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=commands_frame, anchor=tk.NW, tags="commands_frame")
        
        # Create command cards efficiently
        for button in self.command_vars.keys():
            self.create_command_card(
                commands_frame,
                f"{button} Butonu",
                self.command_vars[button]['type'],
                self.command_vars[button]['subtype'],
                self.command_vars[button]['entry'],
                self.get_button_color(button)
            )
        
        # Optimize canvas resizing
        def update_scroll_region(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        def configure_canvas(event):
            canvas.itemconfig(canvas_window, width=canvas.winfo_width())
        
        # Efficient event binding
        canvas.bind("<Configure>", configure_canvas)
        commands_frame.bind("<Configure>", update_scroll_region)
        
        # Optimize mousewheel scrolling
        def on_mousewheel(event):
            if event.num == 5 or event.delta < 0:
                canvas.yview_scroll(1, "units")
            elif event.num == 4 or event.delta > 0:
                canvas.yview_scroll(-1, "units")
        
        # Bind mousewheel events based on platform
        if sys.platform.startswith('win'):
            canvas.bind_all("<MouseWheel>", on_mousewheel)
        else:
            canvas.bind_all("<Button-4>", on_mousewheel)
            canvas.bind_all("<Button-5>", on_mousewheel)

    def create_command_card(self, parent, title, command_type_var, command_subtype_var, entry_var, color_accent):
        card = ttk.LabelFrame(parent, text=title, style="Card.TLabelframe")
        card.pack(fill=tk.X, pady=10, ipady=10, padx=5)
        
        inner_frame = ttk.Frame(card, style="CardInner.TFrame")
        inner_frame.pack(fill=tk.X, padx=15, pady=10)
        
        inner_frame.columnconfigure(0, weight=1)
        inner_frame.columnconfigure(1, weight=3)
        
        color_indicator = tk.Frame(inner_frame, width=4, height=40, bg=color_accent)
        color_indicator.grid(row=0, column=0, rowspan=3, sticky="ns", padx=(0, 10))
        
        type_frame = ttk.Frame(inner_frame, style="CardInner.TFrame")
        type_frame.grid(row=0, column=1, sticky="ew", pady=(0, 8))
        
        ttk.Label(type_frame, text="Komut Tipi:", width=12, style="CardLabel.TLabel").pack(side=tk.LEFT)
        
        type_combobox = ttk.Combobox(
            type_frame,
            values=self.command_types,
            width=15,
            textvariable=command_type_var,
            state="readonly"
        )
        type_combobox.pack(side=tk.LEFT, padx=5)
        
        subtype_frame = ttk.Frame(inner_frame, style="CardInner.TFrame")
        subtype_frame.grid(row=1, column=1, sticky="ew", pady=(0, 8))
        
        ttk.Label(subtype_frame, text="Alt Tip:", width=12, style="CardLabel.TLabel").pack(side=tk.LEFT)
        
        subtype_combobox = ttk.Combobox(
            subtype_frame,
            width=30,
            textvariable=command_subtype_var,
            state="readonly"
        )
        subtype_combobox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        entry_frame = ttk.Frame(inner_frame, style="CardInner.TFrame")
        entry_frame.grid(row=2, column=1, sticky="ew")
        
        ttk.Label(entry_frame, text="Özel Komut:", width=12, style="CardLabel.TLabel").pack(side=tk.LEFT)
        
        entry = ttk.Entry(entry_frame, textvariable=entry_var)
        entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # Create record button frame
        record_frame = ttk.Frame(inner_frame, style="CardInner.TFrame")
        record_frame.grid(row=3, column=1, sticky="ew", pady=(8, 0))
        record_frame.grid_remove()  # Initially hidden

        # Create record button
        self.record_button = ttk.Button(
            record_frame,
            text="🎙️ Tuş Kombinasyonu Kaydet",
            command=lambda e=entry: self.start_recording_hotkey(e, self.record_button)
        )
        self.record_button.pack(fill=tk.X)
        
        type_combobox.bind("<<ComboboxSelected>>",
                         lambda e, st=subtype_combobox, sf=subtype_frame, ef=entry_frame, rf=record_frame:
                         self.update_subtype_options(command_type_var.get(), st, sf, ef, rf))
        
        self.update_subtype_options(command_type_var.get(), subtype_combobox, subtype_frame, entry_frame, record_frame)

    def update_subtype_options(self, command_type, command_subtype_combobox, subtype_frame, entry_frame, record_frame=None):
        # Mevcut alt tip değerini koru
        current_subtype = command_subtype_combobox.get()
        
        if command_type == "press":
            command_subtype_combobox['values'] = self.press_keys
            if current_subtype in self.press_keys:
                command_subtype_combobox.set(current_subtype)
            else:
                command_subtype_combobox.set(self.press_keys[0])
            subtype_frame.grid(row=1, column=1, sticky="ew", pady=(0, 8))
            entry_frame.grid_remove()
            if record_frame:
                record_frame.grid_remove()
        elif command_type == "volume":
            command_subtype_combobox['values'] = self.volume_keys
            if current_subtype in self.volume_keys:
                command_subtype_combobox.set(current_subtype)
            else:
                command_subtype_combobox.set(self.volume_keys[0])
            subtype_frame.grid(row=1, column=1, sticky="ew", pady=(0, 8))
            entry_frame.grid_remove()
            if record_frame:
                record_frame.grid_remove()
        elif command_type == "media":
            command_subtype_combobox['values'] = self.media_keys
            if current_subtype in self.media_keys:
                command_subtype_combobox.set(current_subtype)
            else:
                command_subtype_combobox.set(self.media_keys[0])
            subtype_frame.grid(row=1, column=1, sticky="ew", pady=(0, 8))
            entry_frame.grid_remove()
            if record_frame:
                record_frame.grid_remove()
        elif command_type == "hotkey":
            command_subtype_combobox.set("")
            command_subtype_combobox['values'] = []
            subtype_frame.grid_remove()
            entry_frame.grid(row=1, column=1, sticky="ew")
            if record_frame:
                record_frame.grid(row=3, column=1, sticky="ew", pady=(8, 0))
        elif command_type == "yazı":
            subtype_frame.grid_remove()
            entry_frame.grid(row=1, column=1, sticky="ew")
            if record_frame:
                record_frame.grid_remove()

    def start_recording_hotkey(self, entry_widget, button):
        if self.recording_hotkey:
            self.stop_recording_hotkey()
            return

        self.recording_hotkey = True
        self.current_recording_entry = entry_widget
        self.current_recording_button = button
        button.configure(text="⏺️ Kaydı Durdur (tuşları bırakın)")
        
        # Clear the entry
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, "Tuş kombinasyonunu girin...")
        
        # Block all existing hotkeys
        keyboard.unhook_all()
        
        # Start a thread to record the hotkey
        self.record_thread = threading.Thread(target=self.record_hotkey, daemon=True)
        self.record_thread.start()

    def stop_recording_hotkey(self):
        self.recording_hotkey = False
        if self.current_recording_button:
            self.current_recording_button.configure(text="🎙️ Tuş Kombinasyonu Kaydet")
        self.current_recording_entry = None
        self.current_recording_button = None

    def record_hotkey(self):
        pressed_keys = set()
        
        try:
            # Block all default hotkeys
            keyboard.unhook_all()
            
            while self.recording_hotkey:
                event = keyboard.read_event(suppress=True)  # Suppress all key events
                
                if event.event_type == keyboard.KEY_DOWN:
                    key_name = event.name
                    if key_name not in pressed_keys:
                        # Special key name mappings
                        if key_name == 'alt gr':
                            key_name = 'alt'
                        elif key_name == 'right shift':
                            key_name = 'shift'
                        elif key_name == 'right alt':
                            key_name = 'alt'
                        elif key_name == 'right ctrl':
                            key_name = 'ctrl'
                            
                        pressed_keys.add(key_name)
                        # Update the entry with current combination
                        if self.current_recording_entry:
                            # Sort keys to maintain consistent order (modifiers first)
                            sorted_keys = sorted(pressed_keys, key=lambda x: (
                                0 if x in ['ctrl', 'alt', 'shift', 'windows'] else 1,
                                x.lower()
                            ))
                            combination = "+".join(sorted_keys)
                            self.root.after(0, lambda: self.current_recording_entry.delete(0, tk.END))
                            self.root.after(0, lambda: self.current_recording_entry.insert(0, combination))
                
                elif event.event_type == keyboard.KEY_UP:
                    key_name = event.name
                    # Apply the same key name mappings
                    if key_name == 'alt gr':
                        key_name = 'alt'
                    elif key_name == 'right shift':
                        key_name = 'shift'
                    elif key_name == 'right alt':
                        key_name = 'alt'
                    elif key_name == 'right ctrl':
                        key_name = 'ctrl'
                        
                    if key_name in pressed_keys:
                        pressed_keys.remove(key_name)
                    if not pressed_keys:
                        self.root.after(0, self.stop_recording_hotkey)
                        break
        
        except Exception as e:
            print(f"Error during hotkey recording: {e}")
        finally:
            # Ensure we stop recording if there's an error
            if self.recording_hotkey:
                self.root.after(0, self.stop_recording_hotkey)

    def create_footer(self):
        footer_frame = ttk.Frame(self.main_container)
        footer_frame.pack(fill=tk.BOTH, pady=20)

        self.update_button = ttk.Button(
            footer_frame,
            text="Komutları Güncelle",
            command=self.update_message,
            style="Accent.TButton"
        )
        self.update_button.pack(pady=5, fill=tk.X)

        info_text = ttk.Label(
            footer_frame,
            text="Tüm değişiklikler Arduino'ya otomatik olarak gönderilecektir.",
            style="Info.TLabel"
        )
        info_text.pack(pady=(5, 0))

    @throttle(seconds=0.5)
    def update_message(self, auto_save=False):
        """Throttled message update"""
        try:
            for button, vars in self.command_vars.items():
                if vars['type'].get() == "yazı":
                    vars['message'] = vars['entry'].get()
                elif vars['type'].get() == "hotkey":
                    vars['message'] = f"hotkey:{vars['entry'].get()}"
                else:
                    vars['message'] = vars['subtype'].get()

            if not auto_save:
                self.update_button_status()
                self.save_settings()

        except Exception as e:
            messagebox.showerror("Hata", f"Komutlar güncellenirken bir hata oluştu: {e}")

    @throttle(seconds=1)
    def update_button_status(self):
        """Throttled button status update"""
        original_text = self.update_button["text"]
        self.update_button["text"] = "✓ Komutlar Güncellendi!"
        self.root.after(2000, lambda: self.update_button.configure(text=original_text))

    def monitor_arduino(self):
        """Optimized Arduino monitoring with continuous communication"""
        reconnect_delay = 2  # Initial delay between reconnection attempts
        max_reconnect_delay = 30  # Maximum delay between attempts
        
        while True:
            current_time = time.time()
            
            if current_time - self.last_arduino_check < self.arduino_check_interval:
                time.sleep(0.01)
                continue
                
            self.last_arduino_check = current_time
            
            with self.arduino_lock:
                if not self.arduino:
                    print(f"\nAttempting to reconnect (delay: {reconnect_delay}s)...")
                    self.try_connect_arduino()
                    time.sleep(reconnect_delay)
                    # Increase reconnect delay (with maximum limit)
                    reconnect_delay = min(reconnect_delay * 1.5, max_reconnect_delay)
                else:
                    try:
                        if not self.arduino.is_open:
                            print("\nArduino connection lost (port closed)")
                            self.arduino = None
                            self.update_status_indicator("disconnected")
                            reconnect_delay = 2  # Reset delay on disconnect
                            continue

                        # Send keepalive and check for response
                        self.arduino.write(b"PING\n")
                        time.sleep(0.1)
                        
                        if not self.check_arduino_data():
                            print("No response to keepalive, checking connection...")
                            self.arduino.write(b"TEST\n")
                            time.sleep(0.1)
                            if not self.check_arduino_data():
                                print("Connection appears to be lost")
                                self.arduino.close()
                                self.arduino = None
                                self.update_status_indicator("disconnected")
                                reconnect_delay = 2
                        else:
                            # Reset reconnect delay on successful communication
                            reconnect_delay = 2
                            
                    except Exception as e:
                        print(f"\nError in Arduino monitoring: {str(e)}")
                        try:
                            self.arduino.close()
                        except:
                            pass
                        self.arduino = None
                        self.update_status_indicator("disconnected")
                        reconnect_delay = 2
            
            time.sleep(self.arduino_check_interval)

    def try_connect_arduino(self):
        """Attempt to connect to Arduino with improved error handling and debugging"""
        try:
            print("\n=== Arduino Connection Attempt ===")
            available_ports = list(list_ports.comports())
            if not available_ports:
                print("No COM ports found!")
                self.update_status_indicator("waiting")
                return

            print(f"Available ports: {[port.device for port in available_ports]}")
            
            # Try COM3 first if available
            com3_port = next((port for port in available_ports if port.device == "COM3"), None)
            if com3_port:
                print("COM3 found - attempting connection first...")
                if self.try_connect_to_port(com3_port):
                    return
                print("COM3 connection failed, trying other ports...")

            # Try other ports
            for port in available_ports:
                if port.device != "COM3":  # Skip COM3 as we already tried it
                    if self.try_connect_to_port(port):
                        return

            print("No Arduino found on any port")
            self.update_status_indicator("waiting")
            
        except Exception as e:
            print(f"Connection error: {str(e)}")
            self.update_status_indicator("disconnected")

    def try_connect_to_port(self, port):
        """Try to connect to a specific port with continuous communication"""
        try:
            print(f"\nTrying to connect to {port.device}...")
            print(f"Port details: {port.description}")
            
            ser = serial.Serial(
                port=port.device,
                baudrate=9600,
                timeout=1,
                write_timeout=1,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE
            )
            
            # Give Arduino time to reset
            time.sleep(2)
            print(f"Port opened: {ser.is_open}")
            
            # Clear any existing data
            ser.reset_input_buffer()
            ser.reset_output_buffer()
            
            # Keep trying to establish initial communication
            attempts = 0
            max_attempts = 5
            while attempts < max_attempts:
                print(f"Connection attempt {attempts + 1}/{max_attempts}")
                try:
                    ser.write(b"TEST\n")
                    time.sleep(0.5)
                    
                    if ser.in_waiting > 0:
                        response = ser.readline().decode().strip()
                        print(f"Received response: '{response}'")
                        
                        if response == "DECK":
                            print(f"Arduino successfully connected on {port.device}")
                            self.arduino = ser
                            self.update_status_indicator("connected")
                            return True
                except Exception as e:
                    print(f"Attempt failed: {str(e)}")
                
                attempts += 1
                time.sleep(0.5)
            
            print(f"No valid response from {port.device} after {max_attempts} attempts")
            ser.close()
            return False
            
        except Exception as e:
            print(f"Error on port {port.device}: {str(e)}")
            try:
                ser.close()
            except:
                pass
            return False

    def check_arduino_data(self):
        """Check for Arduino data with improved continuous communication"""
        try:
            if not self.arduino:
                return False

            data_received = False
            start_time = time.time()
            timeout = 1.0  # 1 second timeout
            
            while self.arduino.in_waiting > 0 and (time.time() - start_time) < timeout:
                data = self.arduino.readline().decode().strip()
                if not data:
                    continue
                    
                print(f"Raw data received: '{data}'")
                data_received = True
                
                if data == "PONG":
                    print("Received keepalive response")
                elif data == "DECK":
                    print("Received identification response")
                elif data in self.command_vars:
                    print(f"Valid button press received: {data}")
                    self.root.after(0, lambda d=data: self.handle_command(d))
                else:
                    print(f"Unknown command received: {data}")
            
            return data_received
                        
        except Exception as e:
            print(f"Arduino read error: {str(e)}")
            self.arduino = None
            self.update_status_indicator("disconnected")
            return False

    @throttle(seconds=0.1)
    def execute_action(self, action):
        """Throttled action execution"""
        try:
            action = action.strip()
            if action.startswith("hotkey:"):
                keys = action.replace("hotkey:", "").split("+")
                keyboard.press_and_release("+".join(keys))
            elif action.startswith("press:"):
                key = action.replace("press:", "")
                if key.startswith("f") and key[1:].isdigit():
                    # Get the function key number
                    f_num = int(key[1:])
                    if 1 <= f_num <= 24:
                        press_function_key(f_num)
                    else:
                        keyboard.press_and_release(key)
                else:
                    keyboard.press_and_release(key)
            elif action.startswith("volume:"):
                key = action.split(":")[1]
                if key == "up":
                    keyboard.press_and_release("volume up")
                elif key == "down":
                    keyboard.press_and_release("volume down")
                elif key == "mute":
                    keyboard.press_and_release("volume mute")
            elif action.startswith("media:"):
                action_type = action.split(":")[1]
                if action_type == "play/pause":
                    keyboard.press_and_release("play/pause media")
                elif action_type == "next":
                    keyboard.press_and_release("next track")
                elif action_type == "previous":
                    keyboard.press_and_release("previous track")
                elif action_type == "stop":
                    keyboard.press_and_release("stop media")
            else:
                keyboard.write(action)
        except Exception as e:
            print(f"Action execution error: {e}")
            messagebox.showerror("Hata", f"Komut yürütülürken hata oluştu: {e}")

    def handle_command(self, button):
        """Handle button commands"""
        if button in self.command_vars:
            vars = self.command_vars[button]
            command_type = vars['type'].get()
            
            if command_type == "yazı":
                self.execute_action(vars['entry'].get())
            elif command_type == "hotkey":
                self.execute_action(f"hotkey:{vars['entry'].get()}")
            else:
                self.execute_action(vars['subtype'].get())

    def get_button_color(self, button):
        """Get color for button"""
        colors = {
            'A': "#3498db",
            'B': "#9b59b6",
            'C': "#e67e22",
            'D': "#2ecc71",
            'E': "#1abc9c",
            'F': "#3498db",
            'G': "#e74c3c",
            'H': "#f39c12"
        }
        return colors.get(button, "#3498db")

if __name__ == "__main__":
    root = tk.Tk()
    app = ModernStreamDeckApp(root)
    root.mainloop()