# -*- coding: utf-8 -*-
import customtkinter
from tkinter import messagebox
from PIL import Image, ImageGrab, ImageEnhance

import pytesseract
import os
import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import csv
from datetime import datetime
import threading
import json
# Removed http.server and webbrowser as they were for Picker

# --- Configuration ---
TESSERACT_CMD_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
if os.name != 'nt':
    TESSERACT_CMD_PATH = 'tesseract'
else:
    if not os.path.exists(TESSERACT_CMD_PATH):
        prog_files = os.environ.get("ProgramFiles", "C:\\Program Files")
        alt_path = os.path.join(prog_files, "Tesseract-OCR", "tesseract.exe")
        if os.path.exists(alt_path):
            TESSERACT_CMD_PATH = alt_path
        else:
            TESSERACT_CMD_PATH = 'tesseract'

# IMPORTANT: Scope reverted to access specific sheets by ID/URL
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
TOKEN_JSON_PATH = 'token.json'
CLIENT_SECRETS_FILE = 'client_secret_desktop.json' # For Python backend GSheet access

CSV_FILENAME = 'gtl_listings.csv'
APP_SETTINGS_FILE = 'app_settings.json'
KNOWN_INPUT_TYPES = ["GTL", "GM"]

# --- Tesseract Configuration ---
try:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD_PATH
    version = pytesseract.get_tesseract_version()
    print(f"Tesseract version {version} found.")
except Exception as e:
    print(f"Tesseract Error: {e}")

# --- Google Authentication and Sheets Setup (OAuth User Flow) ---
def get_user_credentials():
    creds = None
    if os.path.exists(TOKEN_JSON_PATH):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_JSON_PATH, SCOPES)
        except Exception as e:
            print(f"Error loading {TOKEN_JSON_PATH}: {e}. It might be corrupted or for different scopes.")
            creds = None # Force re-auth if token is bad

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                print("Refreshing access token for GSheet operations...")
                creds.refresh(Request())
            except Exception as e:
                print(f"Token refresh failed: {e}. You may need to re-authenticate.")
                creds = None
                if os.path.exists(TOKEN_JSON_PATH):
                    try:
                        os.remove(TOKEN_JSON_PATH) # Remove bad token to force re-auth
                        print(f"Removed invalid token file: {TOKEN_JSON_PATH}")
                    except Exception as ex_remove:
                        print(f"Error removing token file: {ex_remove}")
        else: # New authentication needed
            if not os.path.exists(CLIENT_SECRETS_FILE):
                messagebox.showerror("OAuth Error", f"Desktop OAuth credentials ('{CLIENT_SECRETS_FILE}') not found.")
                return None
            
            # --- Pre-Authentication Onboarding/Explanation ---
            explanation_title = "Google Account Permission"
            explanation_message = (
                "To save your listings to a Google Sheet, GTLHelper needs to connect to your Google Account.\n\n"
                "Google will ask for permission to 'See, edit, create, and delete all your Google Sheets spreadsheets.' "
                "This is the standard permission Google provides for apps to interact with spreadsheets YOU specify.\n\n"
                "IMPORTANT:\n"
                "GTLHelper will ONLY access the single spreadsheet you provide by its URL or ID in the settings. "
                "It will NOT access, read, or modify any other files or spreadsheets in your Google Drive.\n\n"
                "The connection token created by Google is stored securely on your computer and is only used for this purpose.\n\n"
                "Click OK to proceed to Google Authentication."
            )
            messagebox.showinfo(explanation_title, explanation_message)
            # --- End of Pre-Authentication Explanation ---

            try:
                flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS_FILE, SCOPES)
                print("Starting OAuth flow for GSheet operations. Please follow browser instructions.")
                creds = flow.run_local_server(port=0) # User authenticates in browser
                print("OAuth flow for GSheet operations completed.")
            except Exception as e:
                messagebox.showerror("OAuth Error", f"Desktop OAuth flow failed: {e}")
                print(f"Desktop OAuth err: {e}")
                return None
        
        if creds: # If new creds obtained or refreshed successfully
            try:
                with open(TOKEN_JSON_PATH, 'w') as token_file:
                    token_file.write(creds.to_json())
                print(f"GSheet operation credentials saved to {TOKEN_JSON_PATH}")
            except Exception as e:
                print(f"Error saving {TOKEN_JSON_PATH}: {e}")
    return creds

def get_gspread_client(credentials):
    if not credentials:
        return None
    try:
        client = gspread.authorize(credentials)
        print("gspread client authorized for GSheet operations.")
        return client
    except Exception as e:
        print(f"gspread auth error: {e}")
        messagebox.showerror("gspread Error", f"Auth fail for gspread: {e}")
        return None

def load_app_settings():
    defaults = {
        "spreadsheet_id": None, 
        "worksheet_name": "Sheet1"
        # Removed picker keys
    }
    if os.path.exists(APP_SETTINGS_FILE):
        try:
            with open(APP_SETTINGS_FILE, 'r') as f:
                settings = json.load(f)
            # Ensure all default keys exist, remove obsolete picker keys if present
            final_settings = {k: settings.get(k, defaults.get(k)) for k in defaults}
            return final_settings
        except Exception as e:
            print(f"Error loading {APP_SETTINGS_FILE}: {e}")
    return defaults

def save_app_settings(spreadsheet_id, worksheet_name):
    try:
        settings_data = {
            "spreadsheet_id": spreadsheet_id,
            "worksheet_name": worksheet_name
            # Removed picker keys
        }
        with open(APP_SETTINGS_FILE, 'w') as f:
            json.dump(settings_data, f, indent=4)
        print(f"App settings saved to {APP_SETTINGS_FILE}")
    except Exception as e:
        print(f"Error saving app settings: {e}")

# --- CORE OCR AND PARSING FUNCTIONS (Unchanged - Code omitted for brevity) ---
def preprocess_image(image_obj):
    print("Preprocessing image...")
    scale_factor = 1.5
    w, h = image_obj.size
    nw, nh = int(w * scale_factor), int(h * scale_factor)
    img = image_obj.resize((nw, nh), Image.Resampling.LANCZOS).convert('L')
    image_obj_binarized = img.point(lambda x: 0 if x < 128 else 255, '1')
    return image_obj_binarized

def perform_ocr_single_line(image_obj):
    print("Performing OCR (single-line)...")
    try:
        text = pytesseract.image_to_string(image_obj, config=r'--oem 3 --psm 7')
        print(f"Raw OCR (single): '{text.strip()}'")
        return text
    except Exception as e:
        print(f"OCR Error (single): {e}")
        return ""

def perform_ocr_for_region(image_obj):
    print("Performing OCR (region)...")
    try:
        text = pytesseract.image_to_string(image_obj, config=r'--oem 3 --psm 6')
        print(f"Raw OCR (region):\n'''{text.strip()}'''")
        return text
    except Exception as e:
        print(f"OCR Error (region): {e}")
        return ""

def parse_raw_ocr_to_list_of_parts(ocr_text):
    print("Parsing OCR to list of parts...")
    lines = ocr_text.strip().split('\n')
    all_listings_parts = []
    for line_num, line_content in enumerate(lines):
        line = line_content.strip()
        if not line:
            continue
        if "Sent Received Type Date" in line and line_num < 3:
            print(f"  Skip header: {line}")
            continue
        if len(line.split()) < 3:
            print(f"  Skip short line: {line}")
            continue
        current_line_parts = line.split()
        if current_line_parts:
            print(f"  Parts: {current_line_parts}")
            all_listings_parts.append(current_line_parts)
    if not all_listings_parts:
        print("  No significant lines parsed.")
    return all_listings_parts

MONTH_MAP = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
def structure_listing_data(raw_parts):
    if not raw_parts or len(raw_parts) < 4:
        print(f"StructError: Min 4 parts required. Got:{raw_parts}")
        return None
    price_str = raw_parts[0]
    if not (price_str.startswith('$') and len(price_str) > 1 and price_str[1:].replace(',', '').isdigit()):
        print(f"StructError: Invalid price format for '{price_str}'.")
        return None

    formatted_date = None
    potential_date_parts = raw_parts[-3:]
    try:
        m_str, d_str, y_str = potential_date_parts[0], potential_date_parts[1], potential_date_parts[2]
        m = MONTH_MAP.get(m_str.strip('.').capitalize())
        d = int(d_str.replace(',', ''))
        y = int(y_str)
        if m and isinstance(d, int) and isinstance(y, int) and 1990 < y < 2100:
            formatted_date = datetime(y, m, d).strftime('%d/%m/%Y')
    except (ValueError, IndexError, TypeError) as e:
        print(f"  Date parse err for '{potential_date_parts}': {e}")

    if not formatted_date:
        print(f"StructError: Could not parse valid date from {potential_date_parts}.")
        return None

    name_coll = raw_parts[1:len(raw_parts) - 3]
    if name_coll and name_coll[-1].upper() in KNOWN_INPUT_TYPES:
        name_coll = name_coll[:-1]
    name = " ".join(name_coll).strip() if name_coll else "Unknown Item"

    try:
        price = float(price_str.replace('$', '').replace(',', ''))
    except ValueError:
        print(f"StructError: PriceNum convert error for '{price_str}'.")
        return None
    return [name, price, formatted_date]

def append_to_csv(data_row):
    if not data_row:
        return False
    try:
        with open(CSV_FILENAME, 'a', newline='', encoding='utf-8') as f:
            csv.writer(f).writerow(data_row)
        return True
    except Exception as e:
        print(f"CSV Error: {e}")
        return False

def append_to_google_sheet_batch(worksheet, list_of_data_rows):
    if not worksheet or not list_of_data_rows:
        return False
    try:
        worksheet.append_rows(list_of_data_rows, value_input_option='USER_ENTERED')
        return True
    except Exception as e:
        print(f"Error batch appending to Sheets: {e}")
        return False

# --- Screen Region Selector Class (Unchanged - Code omitted for brevity) ---
class RegionSelector:
    def __init__(self, parent_for_after, on_capture_callback):
        self.parent_for_after = parent_for_after
        self.on_capture_callback = on_capture_callback
        self.overlay = customtkinter.CTkToplevel()
        self.overlay.attributes('-fullscreen', True)
        self.overlay.attributes('-alpha', 0.25)
        self.overlay.attributes('-topmost', True)
        self.overlay.overrideredirect(True)
        self.canvas = customtkinter.CTkCanvas(self.overlay, cursor="cross", bg="#404040", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.start_x, self.start_y, self.rect = None, None, None
        self.canvas.bind("<ButtonPress-1>", self.on_mouse_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_release)
        self.overlay.bind("<Escape>", self.cancel_capture)

    def on_mouse_press(self, event):
        if not self.canvas:
            return
        try:
            self.start_x = self.canvas.canvasx(event.x)
            self.start_y = self.canvas.canvasy(event.y)
            if self.rect:
                try:
                    self.canvas.delete(self.rect)
                except customtkinter.tkinter.TclError:
                    print("Debug (on_mouse_press): TclError deleting old rect.")
            self.rect = None
        except customtkinter.tkinter.TclError:
            print("Debug (on_mouse_press): TclError. Aborting capture.")
            self.cancel_capture()

    def on_mouse_drag(self, event):
        if not self.canvas or self.start_x is None:
            return
        if self.rect:
            try:
                self.canvas.delete(self.rect)
            except customtkinter.tkinter.TclError:
                print("Debug (on_mouse_drag): TclError deleting rect.")
                self.rect = None
                return
        try:
            cur_x = self.canvas.canvasx(event.x)
            cur_y = self.canvas.canvasy(event.y)
            self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, cur_x, cur_y, outline='#00AFFF', width=2)
        except customtkinter.tkinter.TclError:
            print("Debug (on_mouse_drag): TclError creating new rect.")
            self.rect = None
            return

    def on_mouse_release(self, event):
        if not self.canvas or not self.overlay or self.start_x is None:
            if self.on_capture_callback:
                self.on_capture_callback(None)
            if self.overlay:
                try:
                    self.overlay.destroy()
                except:
                    pass
            self.overlay = None
            self.canvas = None
            return

        final_canvas_x, final_canvas_y = 0, 0
        try:
            final_canvas_x = self.canvas.canvasx(event.x)
            final_canvas_y = self.canvas.canvasy(event.y)
        except customtkinter.tkinter.TclError:
            if self.on_capture_callback:
                self.on_capture_callback(None)
            if self.overlay:
                try:
                    self.overlay.destroy()
                except:
                    pass
            self.overlay = None
            self.canvas = None
            return

        try:
            self.overlay.destroy()
        except:
            pass
        self.overlay = None
        self.canvas = None

        x1, y1 = min(self.start_x, final_canvas_x), min(self.start_y, final_canvas_y)
        x2, y2 = max(self.start_x, final_canvas_x), max(self.start_y, final_canvas_y)

        if x2 - x1 > 5 and y2 - y1 > 5:
            self.parent_for_after.after(100, lambda: self.grab_screen_region(int(x1), int(y1), int(x2), int(y2)))
        else:
            if self.on_capture_callback:
                self.on_capture_callback(None)

    def grab_screen_region(self, x1, y1, x2, y2):
        try:
            img = ImageGrab.grab(bbox=(x1, y1, x2, y2), all_screens=True)
            if self.on_capture_callback:
                self.on_capture_callback(img)
        except Exception as e:
            messagebox.showerror("Capture Error", f"Capture Fail: {e}")
            if self.on_capture_callback:
                self.on_capture_callback(None)

    def cancel_capture(self, event=None):
        if self.overlay:
            try:
                self.overlay.destroy()
            except:
                pass
        self.overlay = None
        self.canvas = None
        if self.on_capture_callback:
            self.on_capture_callback(None)

# --- GUI Application ---
class GTLHelperApp:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("GTL OCR Helper")
        self.root.attributes('-topmost', True)
        customtkinter.set_appearance_mode("Dark")
        customtkinter.set_default_color_theme("blue")

        self.gspread_client = None
        self.worksheet = None
        self.app_settings = load_app_settings()
        self.current_spreadsheet_id = self.app_settings.get("spreadsheet_id")
        self.current_worksheet_name = self.app_settings.get("worksheet_name", "Sheet1")

        self.layout_is_mini = False
        self.normal_geom = "580x295" # Adjusted height slightly for explanation label
        self.mini_geom = "170x165" 
        self.icon_size_normal = (18, 18)
        self.icon_size_mini_action = (26, 26)
        self.icon_size_mini_layout_toggle = (20, 20)
        self.save_to_sheets_var = customtkinter.BooleanVar(value=True)
        self.save_to_csv_var = customtkinter.BooleanVar(value=True)

        self.icons = {}
        self.icons_mini = {}
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.icon_text_fallbacks = {
            "capture": "📷", "clipboard": "📋", "save": "💾",
            "layout": "⇄", "settings": "⚙️", "link": "🔗"
            # Removed "picker"
        }
        icon_map = {
            "capture": "camera_icon.png", "clipboard": "clipboard_icon.png",
            "save": "save_icon.png", "layout": "layout_icon.png",
            "settings": "settings_icon.png", "link": "link_icon.png"
            # Removed "picker"
        }

        def load_ctk_icon(fn, sz, fb):
            try:
                path = os.path.join(script_dir, "icons", fn)
                return customtkinter.CTkImage(Image.open(path), size=sz) if os.path.exists(path) else fb
            except Exception as e:
                print(f"IconErr {fn}: {e}")
                return fb

        for n, fn in icon_map.items():
            self.icons[n] = load_ctk_icon(fn, self.icon_size_normal, self.icon_text_fallbacks[n])
            msz = self.icon_size_mini_action if n in ["capture", "clipboard", "save"] else self.icon_size_mini_layout_toggle
            self.icons_mini[n] = load_ctk_icon(fn, msz, self.icon_text_fallbacks[n])

        self.content_frame = customtkinter.CTkFrame(self.root, corner_radius=0, fg_color="transparent")
        self.content_frame.pack(fill="both", expand=True, padx=3, pady=3)
        self.status_label = customtkinter.CTkLabel(self.content_frame, text="Authenticating...", anchor="w", font=customtkinter.CTkFont(size=10))
        self.top_controls_frame = customtkinter.CTkFrame(self.content_frame)
        self.capture_region_btn = customtkinter.CTkButton(self.top_controls_frame, command=self.start_region_capture)
        self.preview_clipboard_btn = customtkinter.CTkButton(self.top_controls_frame, command=self.preview_listing_from_clipboard)
        self.save_btn = customtkinter.CTkButton(self.top_controls_frame, command=self.save_listing_action, fg_color="#28A745", hover_color="#218838", state="disabled")
        self.preview_frame = customtkinter.CTkFrame(self.content_frame, fg_color="transparent")
        self.preview_text_var = customtkinter.StringVar(value="Preview")
        self.preview_entry = customtkinter.CTkEntry(self.preview_frame, textvariable=self.preview_text_var, state="readonly", justify="center", font=customtkinter.CTkFont(size=11))
        self.settings_control_frame = customtkinter.CTkFrame(self.content_frame, fg_color="transparent")
        self.settings_toggle_btn = customtkinter.CTkButton(self.settings_control_frame, command=self.toggle_settings_visibility)
        self.layout_toggle_btn_for_mini = customtkinter.CTkButton(self.settings_control_frame, command=self.toggle_app_layout)
        
        self.actual_settings_options_frame = customtkinter.CTkFrame(self.content_frame)
        self.sheet_selection_frame = customtkinter.CTkFrame(self.actual_settings_options_frame)
        
        self.sheet_id_label = customtkinter.CTkLabel(self.sheet_selection_frame, text="Sheet URL/ID:")
        self.sheet_id_var = customtkinter.StringVar(value=self.current_spreadsheet_id or "")
        self.sheet_id_entry = customtkinter.CTkEntry(self.sheet_selection_frame, textvariable=self.sheet_id_var, width=250) # Adjusted width

        link_icon = self.icons.get("link")
        set_sheet_txt = "Set" if link_icon and isinstance(link_icon, customtkinter.CTkImage) else self.icon_text_fallbacks["link"]
        set_sheet_compound = "left" if link_icon and isinstance(link_icon, customtkinter.CTkImage) else "top"
        self.set_sheet_btn = customtkinter.CTkButton( # Renamed from set_sheet_id_btn for clarity
            self.sheet_selection_frame, text=set_sheet_txt,
            image=link_icon if isinstance(link_icon, customtkinter.CTkImage) else None,
            compound=set_sheet_compound, command=self.set_target_sheet, height=28) # Was set_target_sheet_from_input

        # --- Reinforcement Label for Sheet ID input ---
        self.sheet_id_explanation_label = customtkinter.CTkLabel(
            self.actual_settings_options_frame, # Placed in the main settings area below sheet input
            text="Note: GTLHelper will only access and append data to this specific sheet.",
            font=customtkinter.CTkFont(size=9),
            text_color="gray",
            wraplength=350 # Adjust as needed
        )
        # --- End of Reinforcement Label ---

        self.sheets_check = customtkinter.CTkCheckBox(self.actual_settings_options_frame, text="Sheets", variable=self.save_to_sheets_var)
        self.csv_check = customtkinter.CTkCheckBox(self.actual_settings_options_frame, text="CSV", variable=self.save_to_csv_var)
        self.layout_toggle_btn_in_settings = customtkinter.CTkButton(self.actual_settings_options_frame, command=self.toggle_app_layout)
        self.current_structured_preview_data = None
        self.show_settings_expanded = customtkinter.BooleanVar(value=False)
        
        self.root.after(50, self.initialize_google_auth_and_ui)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing) # Keep if any cleanup, otherwise remove

    def on_closing(self):
        # Add any other cleanup logic if needed in the future
        print("Closing GTLHelper App.")
        self.root.destroy()

    def initialize_google_auth_and_ui(self):
        """Handles Google Auth for GSpread and then builds the UI."""
        user_creds = get_user_credentials() # For gspread
        if user_creds:
            self.gspread_client = get_gspread_client(user_creds)

        if self.gspread_client:
            self.update_status("GSheet Auth OK. Set target sheet if needed.", error=False)
            if self.current_spreadsheet_id:
                self.load_worksheet()
            else:
                self.update_status("Please set Target Google Sheet in Settings.", error=False)
        else:
            self.update_status("Google GSheet Auth Failed! Sheets disabled.", error=True)
            self.save_to_sheets_var.set(False)
            if hasattr(self, 'sheets_check') and self.sheets_check:
                self.sheets_check.configure(state="disabled")

        if hasattr(self, 'sheets_check') and self.sheets_check:
            self.sheets_check.configure(state="normal" if self.gspread_client and self.worksheet else "disabled")
            if not self.gspread_client or not self.worksheet:
                self.save_to_sheets_var.set(False)
        self.rebuild_ui_for_mode()

    def set_target_sheet(self): # Renamed from set_target_sheet_from_input
        input_val = self.sheet_id_var.get().strip()
        if not input_val:
            messagebox.showwarning("Set Sheet", "Sheet ID/URL field is empty. Please paste a valid Google Sheet URL or ID.")
            return

        sheet_id = input_val
        if "spreadsheets/d/" in input_val: # Basic check for URL
            try:
                sheet_id = input_val.split("/d/")[1].split("/")[0]
            except IndexError:
                messagebox.showerror("Set Sheet", "Invalid Google Sheet URL format.")
                return
        
        self.current_spreadsheet_id = sheet_id # Update internal ID
        if self.load_worksheet():
            save_app_settings(self.current_spreadsheet_id, self.current_worksheet_name)
            self.update_status(f"Sheet set: '{self.worksheet.title if self.worksheet else 'FAIL'}'", error=(not self.worksheet))
            if hasattr(self, 'sheets_check'):
                self.sheets_check.configure(state="normal" if self.worksheet else "disabled")
            self.save_to_sheets_var.set(bool(self.worksheet))
        else:
            self.update_status(f"Failed to load sheet: {self.current_spreadsheet_id}. Check ID/permissions.", error=True)
            if hasattr(self, 'sheets_check'): # Ensure checkbox is disabled
                self.sheets_check.configure(state="disabled")
            self.save_to_sheets_var.set(False)

    def load_worksheet(self):
        if not self.gspread_client:
            self.update_status("GSheet client not authorized. Please re-authenticate if needed.", error=True)
            return False
        if not self.current_spreadsheet_id:
            self.update_status("Sheet ID not set. Please enter a Sheet ID in settings.", error=True)
            return False
            
        self.update_status(f"Loading sheet: {self.current_spreadsheet_id[:20]}...")
        try:
            spreadsheet = self.gspread_client.open_by_key(self.current_spreadsheet_id)
            if self.current_worksheet_name and self.current_worksheet_name != "Sheet1":
                 try:
                    self.worksheet = spreadsheet.worksheet(self.current_worksheet_name)
                 except gspread.exceptions.WorksheetNotFound:
                    self.update_status(f"Worksheet '{self.current_worksheet_name}' not found. Using first sheet.", error=True)
                    self.worksheet = spreadsheet.sheet1 
                    self.current_worksheet_name = self.worksheet.title 
            else: # Default to first sheet or if current_worksheet_name is "Sheet1"
                self.worksheet = spreadsheet.sheet1
                self.current_worksheet_name = self.worksheet.title

            print(f"Loaded worksheet: '{self.worksheet.title}' from spreadsheet ID: {self.current_spreadsheet_id}")
            return True
        except gspread.exceptions.APIError as e:
            print(f"gspread API Error: {e}")
            json_response = e.response.json()
            error_details = json_response.get("error", {})
            error_status = error_details.get("status")
            error_message = error_details.get("message","Unknown API error")

            if error_status == "PERMISSION_DENIED":
                 self.update_status(f"Permission Denied for sheet. Ensure the account has access and correct scopes are granted (you may need to delete token.json and re-auth).", error=True)
            elif error_status == "NOT_FOUND":
                 self.update_status(f"Sheet not found: {self.current_spreadsheet_id}. Check the ID.", error=True)
            else:
                 self.update_status(f"GSheet API Error: {error_message} ({error_status})", error=True)
            self.worksheet = None
            return False
        except Exception as e:
            print(f"Worksheet load error: {e}")
            self.update_status(f"Failed to load sheet. Error: {e}", error=True)
            self.worksheet = None
            return False

    def rebuild_ui_for_mode(self):
        for w in self.content_frame.winfo_children():
            w.pack_forget()

        sml_pad, norm_pad = 2, 5
        btn_h_norm, btn_h_mini_act, btn_h_mini_tog = 28, 36, 28
        btn_w_norm_txt = 120 
        btn_w_mini_act, btn_w_mini_tog = 36, 80

        current_save_state = "normal" if self.current_structured_preview_data else "disabled"
        common_btn_config = {
            "capture": {"command": self.start_region_capture},
            "clipboard": {"command": self.preview_listing_from_clipboard},
            "save": {"command": self.save_listing_action, "fg_color": "#28A745",
                     "hover_color": "#218838", "state": current_save_state},
        }

        self.status_label.pack(side="bottom", fill="x", padx=sml_pad, pady=(sml_pad, sml_pad))
        self.top_controls_frame.pack(side="top", fill="x", padx=sml_pad, pady=(sml_pad, 0))

        if self.layout_is_mini:
            self.root.geometry(self.mini_geom)
            cap_icon = self.icons_mini.get("capture")
            cap_txt = "" if cap_icon and isinstance(cap_icon, customtkinter.CTkImage) else self.icon_text_fallbacks["capture"]
            clip_icon = self.icons_mini.get("clipboard")
            clip_txt = "" if clip_icon and isinstance(clip_icon, customtkinter.CTkImage) else self.icon_text_fallbacks["clipboard"]
            save_icon_mini = self.icons_mini.get("save")
            save_txt_mini = "" if save_icon_mini and isinstance(save_icon_mini, customtkinter.CTkImage) else self.icon_text_fallbacks["save"]

            self.capture_region_btn.configure(
                image=cap_icon if isinstance(cap_icon, customtkinter.CTkImage) else None,
                text=cap_txt, width=btn_w_mini_act, height=btn_h_mini_act, compound="top",
                **common_btn_config["capture"]
            )
            self.preview_clipboard_btn.configure(
                image=clip_icon if isinstance(clip_icon, customtkinter.CTkImage) else None,
                text=clip_txt, width=btn_w_mini_act, height=btn_h_mini_act, compound="top",
                **common_btn_config["clipboard"]
            )
            self.save_btn.configure(
                image=save_icon_mini if isinstance(save_icon_mini, customtkinter.CTkImage) else None,
                text=save_txt_mini, width=btn_w_mini_act, height=btn_h_mini_act, compound="top",
                **common_btn_config["save"]
            )

            self.capture_region_btn.pack(side="left", padx=sml_pad, pady=sml_pad, expand=True)
            self.preview_clipboard_btn.pack(side="left", padx=sml_pad, pady=sml_pad, expand=True)
            self.save_btn.pack(side="left", padx=sml_pad, pady=sml_pad, expand=True)

            self.preview_frame.pack(side="top", fill="x", padx=sml_pad, pady=(sml_pad, 0))
            self.preview_entry.pack(fill="x", ipady=1, padx=sml_pad)
            
            self.settings_control_frame.pack(side="top", fill="x", padx=sml_pad, pady=(norm_pad, sml_pad))
            layout_icon_mini = self.icons_mini.get("layout")
            layout_txt_mini = "Full" if layout_icon_mini and isinstance(layout_icon_mini, customtkinter.CTkImage) else self.icon_text_fallbacks["layout"] + " Full"
            layout_compound_mini = "left" if layout_icon_mini and isinstance(layout_icon_mini, customtkinter.CTkImage) else "top"
            self.layout_toggle_btn_for_mini.configure(
                image=layout_icon_mini if isinstance(layout_icon_mini, customtkinter.CTkImage) else None,
                text=layout_txt_mini, compound=layout_compound_mini,
                height=btn_h_mini_tog, width=btn_w_mini_tog
            )
            self.layout_toggle_btn_for_mini.pack(side="top", padx=sml_pad, pady=sml_pad, expand=True) # Changed from left
            
            self.settings_toggle_btn.pack_forget()
            self.actual_settings_options_frame.pack_forget() 

        else:  # Normal layout
            self.root.geometry(self.normal_geom)
            cap_icon_norm = self.icons.get("capture")
            cap_txt_norm = "Capture" if cap_icon_norm and isinstance(cap_icon_norm, customtkinter.CTkImage) else self.icon_text_fallbacks["capture"]
            clip_icon_norm = self.icons.get("clipboard")
            clip_txt_norm = "Clipboard" if clip_icon_norm and isinstance(clip_icon_norm, customtkinter.CTkImage) else self.icon_text_fallbacks["clipboard"]
            save_icon_norm = self.icons.get("save")
            save_txt_norm = "Save" if save_icon_norm and isinstance(save_icon_norm, customtkinter.CTkImage) else self.icon_text_fallbacks["save"]

            self.capture_region_btn.configure(
                image=cap_icon_norm if isinstance(cap_icon_norm, customtkinter.CTkImage) else None,
                text=cap_txt_norm, compound="left", width=btn_w_norm_txt, height=btn_h_norm,
                **common_btn_config["capture"]
            )
            self.preview_clipboard_btn.configure(
                image=clip_icon_norm if isinstance(clip_icon_norm, customtkinter.CTkImage) else None,
                text=clip_txt_norm, compound="left", width=btn_w_norm_txt, height=btn_h_norm,
                **common_btn_config["clipboard"]
            )
            self.save_btn.configure(
                image=save_icon_norm if isinstance(save_icon_norm, customtkinter.CTkImage) else None,
                text=save_txt_norm, compound="left", width=btn_w_norm_txt, height=btn_h_norm,
                **common_btn_config["save"]
            )

            self.capture_region_btn.pack(side="left", padx=norm_pad, pady=norm_pad, expand=True)
            self.preview_clipboard_btn.pack(side="left", padx=norm_pad, pady=norm_pad, expand=True)
            self.save_btn.pack(side="left", padx=norm_pad, pady=norm_pad, expand=True)

            self.settings_control_frame.pack(side="top", fill="x", padx=norm_pad, pady=(norm_pad, 0))
            settings_icon_obj = self.icons.get("settings")
            arrow = "▴" if self.show_settings_expanded.get() else "▾"
            settings_text_val = f"Settings {arrow}" 
            settings_compound = "left" if settings_icon_obj and isinstance(settings_icon_obj, customtkinter.CTkImage) else "top"
            self.settings_toggle_btn.configure(
                image=settings_icon_obj if isinstance(settings_icon_obj, customtkinter.CTkImage) else None,
                text=settings_text_val if settings_icon_obj and isinstance(settings_icon_obj, customtkinter.CTkImage) else self.icon_text_fallbacks["settings"] + arrow,
                compound=settings_compound, height=btn_h_norm
            )
            self.settings_toggle_btn.pack(side="top", fill="x")
            self.layout_toggle_btn_for_mini.pack_forget()

            if self.show_settings_expanded.get():
                self.actual_settings_options_frame.pack(side="top", fill="x", pady=(sml_pad, 0), padx=norm_pad)
                self.sheet_selection_frame.pack(side="top", fill="x", pady=(0, sml_pad)) # Reduced bottom padding
                
                self.sheet_id_label.pack(side="left", padx=(0, sml_pad))
                self.sheet_id_entry.pack(side="left", expand=True, fill="x", padx=(0,sml_pad))
                self.set_sheet_btn.pack(side="left", padx=(0,0)) # Removed picker button

                # Pack the explanation label
                self.sheet_id_explanation_label.pack(side="top", fill="x", pady=(sml_pad, norm_pad), padx=sml_pad)

                self.sheets_check.pack(side="left", padx=norm_pad, pady=(0, norm_pad), expand=True)
                self.csv_check.pack(side="left", padx=norm_pad, pady=(0, norm_pad), expand=True)

                layout_icon_norm = self.icons.get("layout")
                layout_text_norm = "Mini" if layout_icon_norm and isinstance(layout_icon_norm, customtkinter.CTkImage) else self.icon_text_fallbacks["layout"] + " Mini"
                layout_compound_norm = "left" if layout_icon_norm and isinstance(layout_icon_norm, customtkinter.CTkImage) else "top"
                self.layout_toggle_btn_in_settings.configure(
                    image=layout_icon_norm if isinstance(layout_icon_norm, customtkinter.CTkImage) else None,
                    text=layout_text_norm, compound=layout_compound_norm, height=btn_h_norm
                )
                self.layout_toggle_btn_in_settings.pack(side="left", padx=norm_pad, pady=(0, norm_pad), expand=True)
            else:
                self.actual_settings_options_frame.pack_forget()

            self.preview_frame.pack(side="top", fill="both", expand=True, padx=norm_pad, pady=norm_pad)
            self.preview_entry.pack(fill="x", expand=True, ipady=5, padx=sml_pad)

        self.update_preview_display(self.current_structured_preview_data)
        # Update status based on whether GSheet client is available and using broad scope.
        status_scope_msg = "Scope: spreadsheets (full)" if self.gspread_client else "Scope: (Auth Needed)"
        self.update_status_text_only(f"Mode: {'Mini' if self.layout_is_mini else 'Normal'}. {status_scope_msg}")


    def toggle_app_layout(self):
        self.layout_is_mini = not self.layout_is_mini
        if self.layout_is_mini and self.show_settings_expanded.get():
            self.show_settings_expanded.set(False)
        self.rebuild_ui_for_mode()

    def toggle_settings_visibility(self):
        if self.layout_is_mini: 
            self.layout_is_mini = False 
        self.show_settings_expanded.set(not self.show_settings_expanded.get())
        self.rebuild_ui_for_mode()

    def update_status(self, message, error=False):
        ec = "#FF5555"
        try: 
            dc = customtkinter.ThemeManager.theme["CTkLabel"]["text_color"]
        except:
            dc = "#FFFFFF" 
        self.status_label.configure(text=message, text_color=(ec if error else dc))
        if self.root and self.root.winfo_exists(): self.root.update_idletasks()


    def update_status_text_only(self, message):
        try:
            dc = customtkinter.ThemeManager.theme["CTkLabel"]["text_color"]
        except:
            dc = "#FFFFFF"
        self.status_label.configure(text=message, text_color=dc)
        if self.root and self.root.winfo_exists(): self.root.update_idletasks()

    def update_preview_display(self, list_of_structured_data):
        self.current_structured_preview_data = list_of_structured_data
        save_btn_state = "disabled"

        if list_of_structured_data and isinstance(list_of_structured_data, list) and len(list_of_structured_data) > 0:
            save_btn_state = "normal"
            first_item = list_of_structured_data[0]
            item_count = len(list_of_structured_data)
            count_str = f" ({item_count})" if item_count > 1 else ""

            if self.layout_is_mini:
                name_prev = str(first_item[0])
                price_prev = first_item[1]
                disp_text = f"{name_prev[:6]}..|{price_prev:.0f}{count_str}" if len(name_prev) > 6 else f"{name_prev}|{price_prev:.0f}{count_str}"
            else:
                if item_count == 1:
                    disp_text = f"Name: {first_item[0]} | Price: {first_item[1]:.0f} | Date: {first_item[2]}"
                else:
                    disp_text = f"{item_count} items. 1st: {first_item[0]} | {first_item[1]:.0f} | {first_item[2]}"
            self.preview_text_var.set(disp_text)
        else:
            self.preview_text_var.set("..." if self.layout_is_mini else "No preview data.")

        if hasattr(self, 'save_btn') and self.save_btn and self.save_btn.winfo_exists():
            self.save_btn.configure(state=save_btn_state)
        if self.root and self.root.winfo_exists(): self.root.update_idletasks()


    def _process_image_for_preview(self, image_object, is_region_capture=False):
        if image_object is None:
            self.update_preview_display(None)
            self.update_status("Capture fail/cancel.", error=True)
            return

        self.update_status("Processing OCR...")
        preproc_img = preprocess_image(image_object)
        ocr_text = perform_ocr_for_region(preproc_img) if is_region_capture else perform_ocr_single_line(preproc_img)
        struct_list = []

        if ocr_text:
            list_raw_parts = parse_raw_ocr_to_list_of_parts(ocr_text)
            if list_raw_parts:
                self.update_status(f"OCR done, {len(list_raw_parts)}L. Structuring...")
                for parts in list_raw_parts:
                    item = structure_listing_data(parts)
                    if item:
                        struct_list.append(item)
                self.update_preview_display(struct_list if struct_list else None)
                if struct_list:
                    self.update_status(f"{len(struct_list)} list(s) loaded. Save?")
                else:
                    self.update_status("Structure fail. Check OCR/selection.", error=True)
            else:
                self.update_status("No parts from OCR.", error=True)
                self.update_preview_display(None)
        else:
            self.update_status("OCR no text.", error=True)
            self.update_preview_display(None)

    def preview_listing_from_clipboard(self):
        self.update_status("Preview from clipboard...")
        try:
            img = ImageGrab.grabclipboard()
            if isinstance(img, Image.Image):
                self._process_image_for_preview(img, is_region_capture=False)
            elif img is None:
                self.update_status("No image on clipboard.", error=True)
                messagebox.showwarning("Clipboard", "No image on clipboard.")
                self.update_preview_display(None)
            else: 
                self.update_status("Clipboard content is not a direct image.", error=True)
                self.update_preview_display(None)
        except Exception as e: 
            self.update_status(f"Clipboard Error: {e}", error=True)
            self.update_preview_display(None)

    def start_region_capture(self):
        self.update_status("Select region...")
        self.root.iconify()
        try:
            RegionSelector(self.root, self.handle_captured_image)
        except Exception as e:
            print(f"Error creating RegionSelector: {e}")
            # De-iconify and show error if RegionSelector fails
            self.root.deiconify()
            self.root.attributes('-topmost', True)
            self.root.focus_force()
            self.update_status(f"Capture init error: {e}", error=True)
            self.update_preview_display(None)


    def handle_captured_image(self, captured_image):
        self.root.deiconify()
        self.root.attributes('-topmost', True) # Bring app to front
        self.root.focus_force() # Ensure it has focus
        #self.root.attributes('-topmost', False) # Allow other windows to come on top again

        if captured_image:
            self._process_image_for_preview(captured_image, is_region_capture=True)
        else:
            self.update_status("Capture cancelled or failed.")
            self.update_preview_display(None)

    def save_listing_action(self):
        self.update_status("Initiating save...")
        if not self.current_structured_preview_data or not isinstance(self.current_structured_preview_data, list):
            self.update_status("No preview data to save.", error=True)
            messagebox.showwarning("Save", "No data available to save.")
            return

        data_to_save_list = list(self.current_structured_preview_data)
        if not data_to_save_list:
            self.update_status("No valid listings to save.", error=True)
            return

        print(f"\n--- Saving {len(data_to_save_list)} Listing(s) ---")
        for item in data_to_save_list:
            print(item)
        print("------\n")

        if hasattr(self, 'save_btn') and self.save_btn and self.save_btn.winfo_exists():
            self.save_btn.configure(state="disabled")

        self.preview_text_var.set("Saving..." if self.layout_is_mini else "Saving... Please wait.")
        
        threading.Thread(
            target=self._threaded_save_operation,
            args=(data_to_save_list, self.save_to_sheets_var.get(), self.save_to_csv_var.get()),
            daemon=True
        ).start()

    def _threaded_save_operation(self, data_list, save_sheets, save_csv):
        num = len(data_list)
        sheet_ok_c, csv_ok_c = 0, 0
        final_msg, final_err = "Save complete.", False

        if save_sheets:
            if self.worksheet and self.gspread_client: 
                if append_to_google_sheet_batch(self.worksheet, data_list):
                    sheet_ok_c = num
                else: 
                    self.root.after(0, self.update_status, "Sheets save FAILED (batch).", True)
            elif not self.gspread_client:
                 self.root.after(0, messagebox.showerror, "Save Error", "Google Sheets client not authorized. Please re-auth via settings if needed.")
                 self.root.after(0, self.update_status, "Sheets FAIL: Client not authorized.", True)
            elif not self.worksheet:
                self.root.after(0, messagebox.showerror, "Save Error", "Google Sheets target not set or loaded. Please set it in Settings.")
                self.root.after(0, self.update_status, "Sheets FAIL: No target sheet.", True)

        if save_csv:
            temp_csv_ok = 0
            for item in data_list:
                if append_to_csv(item):
                    temp_csv_ok += 1
            if temp_csv_ok == num:
                csv_ok_c = num
            elif temp_csv_ok > 0:
                csv_ok_c = temp_csv_ok
                self.root.after(0, self.update_status, f"CSV: Partially saved {csv_ok_c}/{num}", True)
            else:
                self.root.after(0, self.update_status, "CSV save FAILED.", True)

        if save_sheets and save_csv:
            if sheet_ok_c == num and csv_ok_c == num:
                final_msg = f"Saved {num} to Sheets & CSV."
            else:
                final_msg = f"Sheets: {sheet_ok_c}/{num}. CSV: {csv_ok_c}/{num}."
                final_err = True
        elif save_sheets:
            if sheet_ok_c == num:
                final_msg = f"Saved {num} to Sheets."
            else:
                final_msg = f"Sheets: Failed to save {num-sheet_ok_c}/{num}." 
                final_err = True
        elif save_csv:
            if csv_ok_c == num:
                final_msg = f"Saved {num} to CSV."
            else:
                final_msg = f"CSV: Failed to save {num-csv_ok_c}/{num}." 
                final_err = True
        else: # Neither save destination was enabled
            final_msg = "No save destination enabled."
            # Not necessarily an error, so final_err remains False unless other issues.

        self.root.after(0, self.update_status, final_msg, final_err)
        self.root.after(0, self.update_preview_display, None) 


if __name__ == "__main__":
    if not os.path.exists(CLIENT_SECRETS_FILE):
        messagebox.showerror("Setup Error", f"Desktop OAuth Client Secrets file ('{CLIENT_SECRETS_FILE}') not found. Please ensure it is in the same directory as the application.")
        exit()

    root = customtkinter.CTk()
    app = GTLHelperApp(root)
    root.mainloop()
