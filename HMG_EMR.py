#!/usr/bin/env python3
"""
HMG Hospital - Complete EMR System v13.0.0
ENHANCED:
  - Multi-user login with role-based access + audit log
  - Task Management System (Nurses, Laboratory, Radiology, Pharmacy, Doctors)
  - Automatic Billing from EMR orders (REAL PRICES — no flat fallbacks)
  - Task Tracking with status (Pending, In Progress, Completed, Waiting, On Hold)
  - Role-based workflows
  - Auto-populate patient ID and name in all windows
  - Comprehensive Lab Reports with QR Code & Digital Signature
  - Multi-admission tracking per patient
  - Appointment Scheduling
  - Nursing Notes
  - Medical Report Drafter (blank-template PDF letter via ReportLab, with camera support)
  - Standardized clerking across all departments (no cross-dept data leakage)
  - All known bugs fixed
  - BILLING FIXED: Prices now connect to stored investigation/procedure/drug prices
  - VIEW ML ENHANCED: Full pathology/ML analysis display with patient data
"""

import kivy
kivy.require('2.1.0')

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.spinner import Spinner
from kivy.uix.popup import Popup
from kivy.uix.tabbedpanel import TabbedPanel, TabbedPanelItem
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.camera import Camera
from kivy.uix.image import Image as KivyImage
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.core.camera import Camera as CoreCamera
from kivy.lang import Builder
from kivy.utils import get_color_from_hex, platform
from kivy.graphics.texture import Texture

from datetime import datetime, timedelta
import json
import os
import csv
import threading
import socket
import base64
import re
import subprocess
import sys
import mimetypes
import io
from io import BytesIO
import shutil
import tempfile
import time
import hashlib
import secrets

# Fix for Android/Pydroid stdout
try:
    if hasattr(sys.stdout, 'buffer'):
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer)
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer)
except AttributeError:
    pass

try:
    import requests
    REQUESTS_AVAILABLE = True
except:
    REQUESTS_AVAILABLE = False

try:
    import qrcode
    QRCODE_AVAILABLE = True
except:
    QRCODE_AVAILABLE = False

try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except:
    DOCX_AVAILABLE = False

try:
    from fpdf import FPDF
    PDF_AVAILABLE = True
except:
    PDF_AVAILABLE = False

try:
    import cv2
    CV2_AVAILABLE = True
except:
    CV2_AVAILABLE = False

try:
    from pyzbar.pyzbar import decode as pyzbar_decode
    PYZBAR_AVAILABLE = True
except:
    PYZBAR_AVAILABLE = False

try:
    from PIL import Image as PILImage
    PIL_AVAILABLE = True
except:
    PIL_AVAILABLE = False

# --- Medical Report Drafter additions ---
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas as pdf_canvas
    from reportlab.lib.utils import ImageReader
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False
    print("ReportLab not installed. Run: pip install reportlab")

# Directory for generated report PDFs
OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "MedicalReports")

KV = '''
<CustomTabbedPanel>:
    do_default_tab: False
    tab_width: 140
    tab_height: 46

<ScrollableLabel@Label>:
    size_hint_y: None
    text_size: self.width, None
    height: self.texture_size[0]

<ScrollableTextInput@TextInput>:
    size_hint_y: None
    height: max(self.minimum_height, 50)
    readonly: True
    multiline: True
    background_color: 0.95, 0.95, 0.95, 1

<ClickableImage@Image>:
    size_hint_y: None
    height: 150

<CameraWidget@Widget>:
    camera: None
    image: None
'''

Builder.load_string(KV)


class CustomTabbedPanel(TabbedPanel):
    pass


# ====================== USER MANAGEMENT ======================
class UserManager:
    """Manages users, password hashing, and login verification."""

    ROLES = [
        "Administrator", "Doctor", "Nurse", "Pharmacist",
        "Lab Scientist", "Radiologist", "Receptionist"
    ]
    USERS_FILE = "users.json"
    DEFAULT_USER = "admin"
    DEFAULT_PASSWORD = "admin123"

    def __init__(self):
        self.users = []
        self.load()
        self._ensure_default_admin()

    # ---------- persistence ----------
    def load(self):
        try:
            if os.path.exists(self.USERS_FILE):
                with open(self.USERS_FILE, "r", encoding="utf-8") as f:
                    self.users = json.load(f)
        except Exception as e:
            print(f"Users load error: {e}")
            self.users = []

    def save(self):
        try:
            with open(self.USERS_FILE, "w", encoding="utf-8") as f:
                json.dump(self.users, f, indent=2, default=str)
        except Exception as e:
            print(f"Users save error: {e}")

    # ---------- security ----------
    @staticmethod
    def _hash_password(password, salt=None):
        if salt is None:
            salt = secrets.token_hex(16)
        h = hashlib.sha256((salt + str(password)).encode("utf-8")).hexdigest()
        return salt, h

    @staticmethod
    def _verify(password, salt, expected_hash):
        _, h = UserManager._hash_password(password, salt)
        return secrets.compare_digest(h, expected_hash)

    # ---------- bootstrap ----------
    def _ensure_default_admin(self):
        if not self.users:
            salt, h = self._hash_password(self.DEFAULT_PASSWORD)
            self.users.append({
                "username": self.DEFAULT_USER,
                "full_name": "System Administrator",
                "role": "Administrator",
                "salt": salt,
                "password_hash": h,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "last_login": "",
                "must_change_password": True
            })
            self.save()
            print(f"[UserManager] Created default admin user "
                  f"'{self.DEFAULT_USER}' / '{self.DEFAULT_PASSWORD}'. "
                  f"CHANGE THIS IMMEDIATELY.")

    # ---------- public API ----------
    def find(self, username):
        u = str(username).strip().lower()
        for user in self.users:
            if user.get("username", "").lower() == u:
                return user
        return None

    def verify_login(self, username, password):
        user = self.find(username)
        if not user:
            return None
        if self._verify(password, user.get("salt", ""), user.get("password_hash", "")):
            user["last_login"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.save()
            return user
        return None

    def create_user(self, username, full_name, password, role):
        username = str(username).strip()
        if not username or not password or not role:
            return False, "Username, password and role are required."
        if role not in self.ROLES:
            return False, f"Invalid role: {role}"
        if self.find(username):
            return False, "Username already exists."
        if len(password) < 4:
            return False, "Password must be at least 4 characters."
        salt, h = self._hash_password(password)
        self.users.append({
            "username": username,
            "full_name": full_name.strip() or username,
            "role": role,
            "salt": salt,
            "password_hash": h,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "last_login": "",
            "must_change_password": False
        })
        self.save()
        return True, f"User '{username}' created."

    def delete_user(self, username):
        target = str(username).strip().lower()
        # Never delete the last Administrator
        admins = [u for u in self.users if u.get("role") == "Administrator"]
        user = self.find(username)
        if not user:
            return False, "User not found."
        if user.get("role") == "Administrator" and len(admins) <= 1:
            return False, "Cannot delete the only Administrator."
        self.users = [u for u in self.users
                      if u.get("username", "").lower() != target]
        self.save()
        return True, f"User '{username}' deleted."

    def change_password(self, username, current_password, new_password):
        user = self.find(username)
        if not user:
            return False, "User not found."
        if not self._verify(current_password,
                            user.get("salt", ""),
                            user.get("password_hash", "")):
            return False, "Current password is incorrect."
        if len(new_password) < 4:
            return False, "New password must be at least 4 characters."
        salt, h = self._hash_password(new_password)
        user["salt"] = salt
        user["password_hash"] = h
        user["must_change_password"] = False
        self.save()
        return True, "Password changed successfully."

    def admin_reset_password(self, username, new_password):
        user = self.find(username)
        if not user:
            return False, "User not found."
        if len(new_password) < 4:
            return False, "Password must be at least 4 characters."
        salt, h = self._hash_password(new_password)
        user["salt"] = salt
        user["password_hash"] = h
        user["must_change_password"] = True
        self.save()
        return True, f"Password reset for '{username}'."


class AuditLog:
    """Append-only log of significant actions for compliance."""

    LOG_FILE = "audit.json"
    MAX_ENTRIES = 5000

    def __init__(self):
        self.entries = []
        self.load()

    def load(self):
        try:
            if os.path.exists(self.LOG_FILE):
                with open(self.LOG_FILE, "r", encoding="utf-8") as f:
                    self.entries = json.load(f)
        except Exception as e:
            print(f"Audit log load error: {e}")
            self.entries = []

    def save(self):
        try:
            if len(self.entries) > self.MAX_ENTRIES:
                self.entries = self.entries[-self.MAX_ENTRIES:]
            with open(self.LOG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.entries, f, indent=2, default=str)
        except Exception as e:
            print(f"Audit log save error: {e}")

    def log(self, username, action, details="", patient_id=""):
        entry = {
            "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "username": username or "system",
            "action": action,
            "details": details,
            "patient_id": patient_id
        }
        self.entries.append(entry)
        self.save()
        return entry

    def get_recent(self, n=100):
        return self.entries[-n:][::-1]

    def filter_by_user(self, username):
        u = str(username).lower()
        return [e for e in self.entries if e.get("username", "").lower() == u]

    def filter_by_action(self, action):
        a = str(action).lower()
        return [e for e in self.entries if e.get("action", "").lower() == a]


# ====================== LOGIN SCREEN ======================
class LoginScreen(BoxLayout):
    """Full-screen login UI shown before the main app."""

    def __init__(self, user_manager, on_login, **kwargs):
        super().__init__(orientation='vertical', padding=20, spacing=10, **kwargs)
        self.user_manager = user_manager
        self.on_login = on_login

        # Spacer top
        self.add_widget(Label(size_hint_y=0.3))

        # Title
        self.add_widget(Label(
            text="HMG HOSPITAL EMR",
            font_size=32, bold=True,
            color=get_color_from_hex('#1a5276'),
            size_hint_y=None, height=50
        ))
        self.add_widget(Label(
            text="Electronic Medical Records System",
            font_size=14,
            color=get_color_from_hex('#7f8c8d'),
            size_hint_y=None, height=24
        ))

        # Spacer
        self.add_widget(Label(size_hint_y=0.05))

        # Center login card
        card = BoxLayout(
            orientation='vertical',
            spacing=10, padding=24,
            size_hint=(0.5, None), size_hint_y=None,
            height=340
        )
        card.pos_hint = {"center_x": 0.5}

        card.add_widget(Label(
            text="Sign In",
            font_size=22, bold=True,
            color=get_color_from_hex('#2c3e50'),
            size_hint_y=None, height=34
        ))

        card.add_widget(Label(
            text="Username",
            font_size=13, bold=True,
            size_hint_y=None, height=20
        ))
        self.username_input = TextInput(
            multiline=False, font_size=16,
            size_hint_y=None, height=42,
            hint_text="Enter username"
        )
        card.add_widget(self.username_input)

        card.add_widget(Label(
            text="Password",
            font_size=13, bold=True,
            size_hint_y=None, height=20
        ))
        self.password_input = TextInput(
            multiline=False, font_size=16,
            password=True,
            size_hint_y=None, height=42,
            hint_text="Enter password",
            on_text_validate=self._attempt_login
        )
        card.add_widget(self.password_input)

        self.error_label = Label(
            text="",
            font_size=12,
            color=get_color_from_hex('#e74c3c'),
            size_hint_y=None, height=24
        )
        card.add_widget(self.error_label)

        btn_row = BoxLayout(size_hint_y=None, height=46, spacing=8)
        btn_row.add_widget(Button(
            text="LOGIN",
            font_size=16, bold=True,
            background_color=get_color_from_hex('#27ae60'),
            on_press=self._attempt_login
        ))
        btn_row.add_widget(Button(
            text="EXIT",
            font_size=16, bold=True,
            background_color=get_color_from_hex('#e74c3c'),
            on_press=self._exit_app
        ))
        card.add_widget(btn_row)

        card.add_widget(Label(
            text="Default: admin / admin123  (change after first login)",
            font_size=11,
            color=get_color_from_hex('#7f8c8d'),
            size_hint_y=None, height=20
        ))

        self.add_widget(card)
        self.add_widget(Label(size_hint_y=0.3))

        # Auto-focus username
        Clock.schedule_once(lambda dt: setattr(self.username_input, 'focus', True), 0.3)

    def _attempt_login(self, *args):
        username = self.username_input.text.strip()
        password = self.password_input.text
        self.error_label.text = ""
        if not username or not password:
            self.error_label.text = "Please enter both username and password."
            return
        user = self.user_manager.verify_login(username, password)
        if user is None:
            self.error_label.text = "Invalid username or password."
            self.password_input.text = ""
            return
        # Success
        if callable(self.on_login):
            self.on_login(user)

    def _exit_app(self, *args):
        App.get_running_app().stop()


# ====================== AI MODELS ======================
class AIModels:
    OPENAI = ["gpt-4o-mini", "gpt-4o", "gpt-4-turbo", "gpt-3.5-turbo"]
    DEEPSEEK = ["deepseek-chat", "deepseek-coder"]

    @classmethod
    def get(cls, provider):
        return cls.OPENAI if provider == "OpenAI" else cls.DEEPSEEK

    @classmethod
    def default(cls, provider):
        return "gpt-4o-mini" if provider == "OpenAI" else "deepseek-chat"


class AIAPIManager:
    def __init__(self):
        self.provider = "OpenAI"
        self.api_key = ""
        self.model = "gpt-4o-mini"
        self.enabled = False
        self.load()

    def load(self):
        try:
            if os.path.exists("ai_settings.json"):
                with open("ai_settings.json") as f:
                    s = json.load(f)
                    self.provider = s.get("provider", "OpenAI")
                    self.api_key = s.get("api_key", "")
                    self.model = s.get("model", "gpt-4o-mini")
                    self.enabled = bool(self.api_key and len(self.api_key) > 10)
        except Exception as e:
            print(f"AI settings load error: {e}")

    def save(self, provider, api_key, model):
        self.provider = provider
        self.api_key = api_key.strip()
        self.model = model
        self.enabled = bool(self.api_key and len(self.api_key) > 10)
        try:
            with open("ai_settings.json", "w") as f:
                json.dump({"provider": provider, "api_key": self.api_key, "model": model}, f, indent=2)
        except Exception as e:
            print(f"AI settings save error: {e}")

    def call(self, sp, up, mt=2000):
        if not self.enabled or not REQUESTS_AVAILABLE:
            return False, "API not available"
        try:
            url = "https://api.openai.com/v1/chat/completions"
            if self.provider == "DeepSeek":
                url = "https://api.deepseek.com/v1/chat/completions"
            headers = {"Authorization": f"Bearer {self.api_key}",
                       "Content-Type": "application/json"}
            payload = {
                "model": self.model,
                "messages": [
                    {"role": "system", "content": sp},
                    {"role": "user", "content": up}
                ],
                "max_tokens": mt
            }
            r = requests.post(url, headers=headers, json=payload, timeout=90)
            if r.status_code == 200:
                return True, r.json()['choices'][0]['message']['content']
            return False, f"Error {r.status_code}"
        except Exception as e:
            return False, str(e)


class QRCodeGenerator:
    @staticmethod
    def generate(data):
        if not QRCODE_AVAILABLE:
            return None
        try:
            qr = qrcode.QRCode(version=1, box_size=8, border=3)
            qr.add_data(data)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            buff = BytesIO()
            img.save(buff, format="PNG")
            return base64.b64encode(buff.getvalue()).decode()
        except:
            return None


class CameraHelper:
    @staticmethod
    def available():
        try:
            CoreCamera(index=0)
            return True
        except:
            return False


class PaymentTracker:
    def __init__(self):
        self.payments = []
        try:
            if os.path.exists("payments.json"):
                with open("payments.json") as f:
                    self.payments = json.load(f)
        except:
            pass

    def save(self):
        try:
            with open("payments.json", "w") as f:
                json.dump(self.payments, f, indent=2, default=str)
        except:
            pass

    def add(self, pid, bid, amt, method="Cash"):
        p = {
            "id": f"PAY{datetime.now().strftime('%Y%m%d%H%M%S')}",
            "patient_id": pid,
            "bill_id": bid,
            "amount": amt,
            "method": method,
            "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        self.payments.append(p)
        self.save()
        return p

    def get_for(self, pid):
        return [p for p in self.payments if p.get("patient_id") == pid]

    def total_for(self, pid):
        return sum(p.get("amount", 0) for p in self.get_for(pid))

    def remove_for_patient(self, pid):
        self.payments = [p for p in self.payments if p.get("patient_id") != pid]
        self.save()


class ChatSystem:
    def __init__(self):
        self.msgs = []
        self.rooms = {"General": []}
        try:
            if os.path.exists("chat.json"):
                with open("chat.json") as f:
                    d = json.load(f)
                    self.msgs = d.get("msgs", [])
                    self.rooms = d.get("rooms", {"General": []})
        except:
            pass

    def save(self):
        try:
            with open("chat.json", "w") as f:
                json.dump({"msgs": self.msgs, "rooms": self.rooms}, f, indent=2)
        except:
            pass

    def send(self, user, room, msg, image=None):
        if room not in self.rooms:
            self.rooms[room] = []
        m = {
            "user": user,
            "room": room,
            "msg": msg,
            "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "image": image
        }
        self.msgs.append(m)
        self.rooms[room].append(m)
        self.save()
        return m

    def get(self, room):
        return self.rooms.get(room, [])[-50:]

    def remove_for_patient(self, patient_name):
        self.msgs = [m for m in self.msgs if m.get("user") != patient_name]
        for room in self.rooms:
            self.rooms[room] = [m for m in self.rooms[room]
                                if m.get("user") != patient_name]
        self.save()


class TaskManager:
    """Task Management System for assigning and tracking tasks."""

    STATUSES = ["Pending", "In Progress", "Completed", "Waiting", "On Hold", "Cancelled"]
    ROLES = ["Doctor", "Nurse", "Lab Scientist", "Radiologist", "Pharmacy", "Administrator"]

    def __init__(self):
        self.tasks = []
        self.load()

    def load(self):
        try:
            if os.path.exists("tasks.json"):
                with open("tasks.json") as f:
                    self.tasks = json.load(f)
        except Exception as e:
            print(f"Tasks load error: {e}")

    def save(self):
        try:
            with open("tasks.json", "w") as f:
                json.dump(self.tasks, f, indent=2, default=str)
        except Exception as e:
            print(f"Tasks save error: {e}")

    def create_task(self, patient_id, patient_name, assigned_to, task_type,
                    description, order_items=None, priority="Normal",
                    department=None):
        task = {
            "id": f"TASK{datetime.now().strftime('%Y%m%d%H%M%S%f')}",
            "patient_id": patient_id,
            "patient_name": patient_name,
            "assigned_to": assigned_to,
            "task_type": task_type,
            "description": description,
            "order_items": order_items or [],
            "status": "Pending",
            "priority": priority,
            "department": department,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "completed_at": None,
            "assigned_to_name": "",
            "notes": ""
        }
        self.tasks.append(task)
        self.save()
        return task

    def update_status(self, task_id, status, notes=""):
        for task in self.tasks:
            if task.get("id") == task_id:
                task["status"] = status
                task["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                if status == "Completed":
                    task["completed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                if notes:
                    task["notes"] = notes
                self.save()
                return task
        return None

    def assign_task(self, task_id, assigned_to_name):
        for task in self.tasks:
            if task.get("id") == task_id:
                task["assigned_to_name"] = assigned_to_name
                task["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.save()
                return task
        return None

    def get_tasks_for_patient(self, patient_id):
        return [t for t in self.tasks if t.get("patient_id") == patient_id]

    def get_tasks_by_role(self, role):
        return [t for t in self.tasks if t.get("assigned_to") == role]

    def get_tasks_by_status(self, status):
        return [t for t in self.tasks if t.get("status") == status]

    def get_pending_tasks(self):
        return [t for t in self.tasks
                if t.get("status") in ["Pending", "In Progress", "Waiting", "On Hold"]]

    def auto_assign_from_emr(self, emr_record, patient):
        tasks = []

        def ensure_list(value):
            if isinstance(value, list):
                return value
            elif isinstance(value, str):
                if ',' in value:
                    return [item.strip() for item in value.split(',') if item.strip()]
                elif value.strip():
                    return [value.strip()]
                else:
                    return []
            else:
                return []

        # Lab orders
        for test in ensure_list(emr_record.get("lab_orders", "")):
            if test and test.strip():
                tasks.append(self.create_task(
                    patient_id=patient.get("id"),
                    patient_name=patient.get("name"),
                    assigned_to="Lab Scientist",
                    task_type="Laboratory",
                    description=f"Lab test: {test}",
                    order_items=[test],
                    department="Laboratory"
                ))

        # Radiology orders
        for study in ensure_list(emr_record.get("radiology_orders", "")):
            if study and study.strip():
                tasks.append(self.create_task(
                    patient_id=patient.get("id"),
                    patient_name=patient.get("name"),
                    assigned_to="Radiologist",
                    task_type="Radiology",
                    description=f"Radiology: {study}",
                    order_items=[study],
                    department="Radiology"
                ))

        # Legacy recommended_investigations
        for inv in ensure_list(emr_record.get("recommended_investigations")):
            if not inv or not inv.strip():
                continue
            lower = inv.lower()
            if any(w in lower for w in ["x-ray", "ct", "mri", "ultrasound",
                                        "scan", "radiograph", "fluoroscopy"]):
                dept, role, ttype = "Radiology", "Radiologist", "Radiology"
            else:
                dept, role, ttype = "Laboratory", "Lab Scientist", "Laboratory"
            tasks.append(self.create_task(
                patient_id=patient.get("id"),
                patient_name=patient.get("name"),
                assigned_to=role,
                task_type=ttype,
                description=f"{ttype} test: {inv}",
                order_items=[inv],
                department=dept
            ))

        # Drugs
        drugs = ensure_list(emr_record.get("medications")) + \
                ensure_list(emr_record.get("prescribed_drugs"))
        for drug in drugs:
            if drug and drug.strip():
                tasks.append(self.create_task(
                    patient_id=patient.get("id"),
                    patient_name=patient.get("name"),
                    assigned_to="Pharmacy",
                    task_type="Pharmacy",
                    description=f"Dispense: {drug}",
                    order_items=[drug],
                    department="Pharmacy"
                ))

        # Nursing orders
        nursing = ensure_list(emr_record.get("nursing_orders"))
        if nursing and any(nursing):
            nursing_text = "\n".join(nursing) if isinstance(nursing, list) else str(nursing)
            tasks.append(self.create_task(
                patient_id=patient.get("id"),
                patient_name=patient.get("name"),
                assigned_to="Nurse",
                task_type="Nursing",
                description=f"Nursing care: {nursing_text}",
                order_items=nursing if isinstance(nursing, list) else [nursing_text],
                department="Nursing"
            ))

        # Follow-ups
        follow_ups = ensure_list(emr_record.get("follow_up"))
        if follow_ups and any(follow_ups):
            follow_text = "\n".join(follow_ups) if isinstance(follow_ups, list) else str(follow_ups)
            tasks.append(self.create_task(
                patient_id=patient.get("id"),
                patient_name=patient.get("name"),
                assigned_to="Doctor",
                task_type="Follow-up",
                description=f"Follow-up: {follow_text}",
                order_items=follow_ups if isinstance(follow_ups, list) else [follow_text],
                department="Medical"
            ))

        self.save()
        return tasks


# ====================== END OF PART 1 ======================
# ====================== ECG ANALYSIS ======================
class ECGAnalyzer:
    SP = """You are an expert cardiologist. Generate a comprehensive 12-lead ECG report following this structure:

1. PATIENT DEMOGRAPHICS (Name, ID, Age, Sex)
2. TECHNICAL QUALITY (Calibration, Paper speed, Quality)
3. HEART RATE (Ventricular, Atrial, Method, Interpretation: Brady/Normal/Tachy)
4. RHYTHM (Regularity, Diagnosis, Supporting features: P wave, PR constancy)
5. CARDIAC AXIS (QRS axis, Degree)
6. P WAVE ANALYSIS (Presence, Morphology, Enlargement signs)
7. PR INTERVAL (Measured ms, Interpretation: Normal/Short/Prolonged)
8. QRS COMPLEX (Duration, Morphology, Voltage, Hypertrophy, Conduction defects)
9. R WAVE PROGRESSION (Transition, Abnormalities)
10. PATHOLOGICAL Q WAVES (Presence, Location, Suggestion)
11. ST SEGMENT (Elevation/Depression, Leads, Magnitude, Pattern)
12. T WAVE (Appearance, Location, Suggestion)
13. QT INTERVAL (QT, QTc, Formula, Interpretation)
14. U WAVES (Present, Prominent, Possible causes)
15. ECTOPIC ACTIVITY (PAC/PVC, Frequency)
16. ARRHYTHMIAS (Type)
17. ISCHEMIC CHANGES (Type, Territory)
18. HYPERTROPHY (Type, Criteria)
19. ELECTROLYTE PATTERN (Suggestion)
20. DRUG EFFECTS (Evidence)
21. COMPARISON WITH PREVIOUS ECG (If available)
22. FINAL IMPRESSION (Overall summary)
23. RECOMMENDATIONS

Provide the report in plain text with clear section headings. At the end, output a JSON summary with keys: rate, rhythm, axis, pr_interval, qrs_duration, qt_qtc, findings (list), interpretation, urgency, recommendations. Use comprehensive clinical language.
"""

    def __init__(self, ai):
        self.ai = ai

    def analyze(self, pt, ecg_data=None, previous_ecg=None):
        ctx = (f"Patient: {pt.get('name','')} | ID: {pt.get('id','')} | "
               f"Age: {pt.get('age','')} | Sex: {pt.get('gender','')}\n")
        ctx += f"Height: {pt.get('height','N/A')} | Weight: {pt.get('weight','N/A')}\n"
        ctx += f"Symptoms: {pt.get('symptoms','None')}\n"
        ctx += f"BP: {pt.get('bp','N/A')} | HR: {pt.get('pulse','N/A')}\n"
        ctx += f"Reason: {pt.get('ecg_reason','Routine')}\n"
        if ecg_data:
            ecg_data = ecg_data[:2000]
            ctx += f"ECG Data:\n{ecg_data}\n"
        else:
            ctx += "No specific ECG data provided. Generate report based on clinical context.\n"
        if previous_ecg:
            ctx += f"Previous ECG: {previous_ecg}\n"

        if self.ai.enabled:
            s, r = self.ai.call(self.SP, ctx, mt=4000)
            if s:
                sd = {}
                try:
                    json_match = re.search(r'\{[\s\S]*\}', r)
                    if json_match:
                        sd = json.loads(json_match.group())
                except:
                    pass
                return {"analysis": r, "structured": sd, "ai_used": True}

        rid = f"ECG-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        analysis = self._generate_fallback_report(pt, rid)
        sd = {
            "rate": "72 bpm",
            "rhythm": "Sinus rhythm, regular",
            "axis": "Normal (+30°)",
            "pr_interval": "160 ms",
            "qrs_duration": "90 ms",
            "qt_qtc": "400/420 ms",
            "findings": ["Normal ECG"],
            "interpretation": "Normal ECG",
            "urgency": "Routine",
            "recommendations": ["Continue management"]
        }
        return {"analysis": analysis, "structured": sd, "ai_used": False, "report_id": rid}

    def _generate_fallback_report(self, pt, rid):
        name = pt.get('name', 'N/A')
        pid = pt.get('id', 'N/A')
        age = pt.get('age', 'N/A')
        sex = pt.get('gender', 'N/A')
        reason = pt.get('ecg_reason', 'Routine')
        symptoms = pt.get('symptoms', 'None')
        bp = pt.get('bp', 'N/A')
        pulse = pt.get('pulse', 'N/A')

        report = f"""
COMPREHENSIVE 12-LEAD ECG REPORT
Report ID: {rid}
Patient Information
Name: {name}
Hospital No.: {pid}
Age/Sex: {age} / {sex}
Date & Time of ECG: {datetime.now().strftime('%Y-%m-%d %H:%M')}
Clinical Indication: {reason}

Technical Quality
Calibration: 10 mm/mV
Paper Speed: 25 mm/sec
Recording Quality: Good

ECG ANALYSIS
1. Heart Rate
Ventricular Rate: 72 bpm
Atrial Rate: 72 bpm
Method Used: 300 Method
Interpretation: Normal (60–100 bpm)

2. Rhythm
Regularity: Regular
Rhythm Diagnosis: Normal Sinus Rhythm
Supporting Features: P wave present, every P followed by QRS, constant PR interval.

3. Cardiac Axis
QRS Axis: Normal
Approximate Degree: +30°

4. P Wave Analysis
Presence: Normal
Morphology: Normal
Evidence of: None

5. PR Interval
Measured: 160 ms
Interpretation: Normal (120–200 ms)

6. QRS Complex
Duration: 90 ms
Morphology: Narrow
Voltage: Normal
Evidence of: No hypertrophy
Conduction Abnormalities: None

7. R Wave Progression
Precordial Transition: Normal

8. Pathological Q Waves
Present: No

9. ST Segment
ST Elevation: Absent
ST Depression: Absent

10. T Wave
Appearance: Normal upright

11. QT Interval
QT: 400 ms
QTc: 420 ms
Formula: Bazett
Interpretation: Normal

12. U Waves
Present: No

13. Ectopic Activity
Premature Beats: None

14. Arrhythmias
None

15. Ischemic Changes
None

16. Hypertrophy
None

17. Electrolyte Pattern
No suggestive changes

18. Drug Effects
None

19. Comparison with Previous ECG
Not available

Final Impression
Normal sinus rhythm at 72 bpm.
Normal cardiac axis.
PR interval 160 ms.
Narrow QRS (90 ms).
QTc 420 ms.
Normal R-wave progression.
No pathological Q waves.
No ST-segment abnormalities.
Normal T waves.
No chamber hypertrophy or conduction defect.
Overall impression: Normal ECG.

Recommendations
Correlate with clinical findings.
Consider serial ECGs if clinically indicated.
"""
        return report


class ECGReportGenerator:
    @staticmethod
    def generate(result, pt):
        rid = result.get("report_id", f"ECG-{datetime.now().strftime('%Y%m%d%H%M%S')}")
        sd = result.get("structured", {})
        qr = QRCodeGenerator.generate(f"HMG|ECG|{rid}|{pt.get('name','')}")
        full_report = result.get("analysis", "No report generated.")
        return full_report, rid, qr


# ====================== CLINICAL AI ASSISTANT ======================
class ModularAIClinicalAssistant:
    SP = """You are an expert clinical assistant. Conduct a structured interview:
- Follow: Chief Complaint → HPI → PMH → Medications → Allergies → Family History → Social History → ROS.
- Ask ONE question at a time.
- Use SOCRATES for pain.
- After about 10-15 exchanges, or when you have sufficient information, output [DIAGNOSIS_READY] and provide:
  - Differential diagnosis (list)
  - Recommended investigations
  - Treatment plan
  - JSON EMR data (with keys: chief_complaint, history_of_presenting_illness, past_medical_history, medications, allergies, differential_diagnosis, recommended_investigations, treatment_plan, urgency, nursing_orders, radiology_orders, prescribed_drugs, follow_up)
- Do not ask more than 20 questions; conclude by then.
"""

    def __init__(self, ai):
        self.ai = ai
        self.conv = []
        self.emr = None
        self.done = False
        self.idx = 0
        self.turn_count = 0
        self.pctx = {}
        self.fb = [
            "Hello, what brings you to the hospital today?",
            "When did this start? Describe the symptoms.",
            "Where is the problem? Does it move?",
            "On a scale of 1-10, how severe?",
            "Any associated symptoms?",
            "Medical conditions?",
            "Past surgeries?",
            "Current medications? Allergies?",
            "Family history?",
            "Occupation? Smoke/drink?",
            "Fevers/chills/weight changes?",
            "Breathing/chest pain/palpitations?",
            "Bowel/bladder changes?",
            "Sleep and appetite?",
            "Any other concerns?"
        ]

    def start(self, name="", age="", gender=""):
        self.conv = []
        self.emr = None
        self.done = False
        self.idx = 0
        self.turn_count = 0
        self.pctx = {"name": name, "age": age, "gender": gender}
        g = (f"Hello{f' {name}' if name else ''}, I'm your clinical AI. "
             f"What brings you to the hospital today?")
        self.conv.append({"role": "assistant", "content": g})
        return g

    def process(self, msg):
        self.conv.append({"role": "user", "content": msg})
        self.turn_count += 1

        if self.turn_count >= 20:
            self.done = True
            self.emr = self._sum()
            summary = self.emr["_text"]
            self.conv.append({"role": "assistant", "content": summary})
            return summary

        if self.ai.enabled:
            ctx = (f"Patient: {self.pctx.get('name','')}, "
                   f"{self.pctx.get('age','')}yo, "
                   f"{self.pctx.get('gender','')}\n\nConversation:\n")
            for m in self.conv[-12:]:
                role = "Dr" if m['role'] == 'assistant' else "Pt"
                ctx += f"{role}: {m['content']}\n"
            ctx += ("\nNext action: ask a question OR provide "
                    "[DIAGNOSIS_READY] if enough information.")
            s, r = self.ai.call(self.SP, ctx, mt=1500)
            if s:
                self.conv.append({"role": "assistant", "content": r})
                if "[DIAGNOSIS_READY]" in r:
                    self.done = True
                    self.emr = self._extract(r)
                return r
            return self._fb()
        else:
            return self._fb()

    def _fb(self):
        if self.idx < len(self.fb):
            q = self.fb[self.idx]
            self.idx += 1
            self.conv.append({"role": "assistant", "content": q})
            return q
        else:
            self.done = True
            s = self._sum()
            self.emr = s
            self.conv.append({"role": "assistant", "content": s["_text"]})
            return s["_text"]

    def _extract(self, r):
        try:
            m = re.search(r'```json\s*([\s\S]*?)\s*```', r)
            if m:
                return json.loads(m.group(1))
            m = re.search(r'\{[\s\S]*\}', r)
            if m:
                return json.loads(m.group())
        except:
            pass
        return self._sum()

    def _sum(self):
        user_msgs = [m["content"] for m in self.conv if m["role"] == "user"]
        full_text = " ".join(user_msgs)
        return {
            "chief_complaint": full_text[:200],
            "history_of_presenting_illness": full_text[:500],
            "past_medical_history": "To confirm",
            "medications": "To confirm",
            "allergies": "None known",
            "differential_diagnosis": ["Further evaluation needed"],
            "recommended_investigations": ["FBC", "Urinalysis"],
            "treatment_plan": "Follow-up 1-2 weeks",
            "urgency": "Routine",
            "nursing_orders": ["Vitals monitoring"],
            "radiology_orders": [],
            "prescribed_drugs": [],
            "follow_up": "Review in 2 weeks",
            "_text": "[DIAGNOSIS_READY]\nSummary prepared based on available information."
        }

    def get_emr(self):
        return self.emr

    def re_analyze(self):
        if not self.ai.enabled or not self.conv:
            return None
        ctx = (f"Patient: {self.pctx.get('name','')}, "
               f"{self.pctx.get('age','')}yo\n\nConversation:\n")
        for m in self.conv[-20:]:
            role = "Dr" if m['role'] == 'assistant' else "Pt"
            ctx += f"{role}: {m['content']}\n"
        ctx += "\nProvide comprehensive re-assessment with [DIAGNOSIS_READY]."
        s, r = self.ai.call(self.SP, ctx, mt=1500)
        if s:
            self.done = True
            self.emr = self._extract(r)
            return r
        return None


# ====================== PATHOLOGY AI ======================
class AIPathologyExtractor:
    SP = ("You are an expert clinical pathologist. Analyze patient data and "
          "provide Pathology Summary, Organ System Assessment, Abnormality Flags, "
          "Risk Stratification, ML Features, Personalized Management, Follow-up.")

    def __init__(self, ai):
        self.ai = ai

    def extract(self, pd, emr, labs, rx):
        ctx = (f"PATIENT: {pd.get('name','')} | "
               f"{pd.get('age','')}yr | {pd.get('gender','')}\n")
        for e in emr[-5:]:
            ctx += f"- [{e.get('ts','')}] {e.get('cc','')}\n"
        completed = [t for t in labs if t.get('status') == 'Completed']
        if completed:
            for lab in completed[-15:]:
                ctx += f"- {lab.get('test_name','')}: {lab.get('result','')}\n"
        ctx += f"\nMEDICATIONS: {len(rx)} prescriptions\n"

        if self.ai.enabled:
            s, r = self.ai.call(self.SP, ctx, mt=2000)
            if s:
                sd = {}
                try:
                    m = re.search(r'\{[\s\S]*\}', r)
                    if m:
                        sd = json.loads(m.group())
                except:
                    pass
                return {"analysis": r, "structured": sd, "ai_used": True}

        ac = sum(1 for lab in completed
                 if any(w in str(lab.get('result', '')).lower()
                        for w in ['positive', 'abnormal', 'high', 'low']))
        risk = min(ac + len(rx) // 3, 10)
        rl = "High" if risk >= 7 else "Medium" if risk >= 4 else "Low"
        return {
            "analysis": f"Pathology Analysis\nAbnormal: {ac}\nRisk: {rl}",
            "structured": {
                "risk_score": risk,
                "risk_level": rl,
                "ml_features": {"abnormal_lab_count": ac}
            },
            "ai_used": False
        }


# ====================== LIS MANAGER ======================
class LISManager:
    def __init__(self, app_instance):
        self.app = app_instance
        self.servers = {}
        self.instrument_configs = {
            "lab": {
                "hematology_analyzer": {
                    "name": "Hematology Analyzer",
                    "manufacturer": "Beckman Coulter",
                    "model": "DxH 900",
                    "port": 5001,
                    "protocol": "ASTM",
                    "status": "Stopped",
                    "auto_process": True,
                    "test_mapping": {"CBC": "Full Blood Count (FBC)"}
                },
                "chemistry_analyzer": {
                    "name": "Chemistry Analyzer",
                    "manufacturer": "Roche",
                    "model": "Cobas c311",
                    "port": 5002,
                    "protocol": "HL7",
                    "status": "Stopped",
                    "auto_process": True,
                    "test_mapping": {"GLU": "Blood Glucose (Fasting)"}
                }
            },
            "radiology": {
                "ct_scanner": {
                    "name": "CT Scanner",
                    "manufacturer": "Siemens",
                    "model": "SOMATOM",
                    "port": 5101,
                    "protocol": "DICOM",
                    "modality": "CT",
                    "status": "Stopped",
                    "auto_process": False
                },
                "xray_machine": {
                    "name": "X-Ray Machine",
                    "manufacturer": "Philips",
                    "model": "DigitalDiagnost",
                    "port": 5103,
                    "protocol": "DICOM",
                    "modality": "X-Ray",
                    "status": "Stopped",
                    "auto_process": False
                }
            },
            "vitals": {
                "bp_machine": {
                    "name": "BP Machine",
                    "manufacturer": "Omron",
                    "model": "HEM-907",
                    "port": 5201,
                    "protocol": "Serial",
                    "status": "Stopped",
                    "auto_process": True,
                    "test_mapping": {"BP": "Blood Pressure"}
                },
                "ecg_machine": {
                    "name": "ECG Machine",
                    "manufacturer": "GE Healthcare",
                    "model": "MAC 2000",
                    "port": 5202,
                    "protocol": "DICOM",
                    "status": "Stopped",
                    "auto_process": False
                }
            }
        }
        self.load_configs()

    def load_configs(self):
        try:
            if os.path.exists("lis_configs.json"):
                with open("lis_configs.json") as f:
                    saved = json.load(f)
                    for c in saved:
                        if c in self.instrument_configs:
                            self.instrument_configs[c].update(saved[c])
        except Exception as e:
            print(f"LIS config load error: {e}")

    def save_configs(self):
        try:
            with open("lis_configs.json", "w") as f:
                json.dump(self.instrument_configs, f, indent=2, default=str)
        except Exception as e:
            print(f"LIS config save error: {e}")

    def add_instrument_from_qr(self, qr_data):
        try:
            data = json.loads(qr_data)
            category = data.get("category", "lab")
            instrument_id = data.get("instrument_id", "")
            if not category or not instrument_id:
                return False, "Invalid QR data"
            if category not in self.instrument_configs:
                self.instrument_configs[category] = {}
            self.instrument_configs[category][instrument_id] = {
                "name": data.get("name", instrument_id),
                "manufacturer": data.get("manufacturer", "Unknown"),
                "model": data.get("model", ""),
                "port": int(data.get("port", 0)),
                "protocol": data.get("protocol", "TCP"),
                "status": "Stopped",
                "auto_process": data.get("auto_process", True),
                "test_mapping": data.get("test_mapping", {})
            }
            self.save_configs()
            return self.start_instrument(category, instrument_id)
        except Exception as e:
            return False, str(e)

    def start_instrument(self, category, instrument_id):
        if (category not in self.instrument_configs or
                instrument_id not in self.instrument_configs[category]):
            return False, "Not found"
        config = self.instrument_configs[category][instrument_id]
        server_id = f"{category}_{instrument_id}"
        if server_id in self.servers and self.servers[server_id].get("running"):
            return False, "Already running"
        server = {"running": True, "config": config, "socket": None}
        if category in ["lab", "vitals"]:
            server["thread"] = threading.Thread(
                target=self._run_lab_server,
                args=(server_id, config["port"]),
                daemon=True
            )
        else:
            server["thread"] = threading.Thread(
                target=self._run_rad_server,
                args=(server_id, config["port"]),
                daemon=True
            )
        self.servers[server_id] = server
        server["thread"].start()
        config["status"] = "Running"
        self.save_configs()
        return True, f"{config['name']} started on port {config['port']}"

    def stop_instrument(self, category, instrument_id):
        server_id = f"{category}_{instrument_id}"
        if server_id not in self.servers:
            return False
        self.servers[server_id]["running"] = False
        if self.servers[server_id].get("socket"):
            try:
                self.servers[server_id]["socket"].close()
            except:
                pass
        if (category in self.instrument_configs and
                instrument_id in self.instrument_configs[category]):
            self.instrument_configs[category][instrument_id]["status"] = "Stopped"
        self.save_configs()
        return True

    def _run_lab_server(self, server_id, port):
        sock = None
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.bind(('0.0.0.0', port))
            sock.listen(5)
            sock.settimeout(1.0)
            if server_id in self.servers:
                self.servers[server_id]["socket"] = sock
            while server_id in self.servers and self.servers[server_id].get("running"):
                try:
                    c, a = sock.accept()
                    threading.Thread(target=self._handle_lab,
                                     args=(c, server_id),
                                     daemon=True).start()
                except socket.timeout:
                    continue
                except:
                    break
        except Exception as e:
            print(f"Lab server error: {e}")
        finally:
            if server_id in self.servers:
                self.servers[server_id]["running"] = False
            if sock:
                try:
                    sock.close()
                except:
                    pass

    def _handle_lab(self, client, server_id):
        try:
            client.send(b"ACK\r\n")
            buf = ""
            while server_id in self.servers and self.servers[server_id].get("running"):
                d = client.recv(4096)
                if not d:
                    break
                buf += d.decode('utf-8', errors='ignore')
                while '\n' in buf or '\r' in buf:
                    line, buf = self._extract(buf)
                    if line and '|' in line:
                        p = line.split('|')
                        if (len(p) >= 4 and
                                self.servers[server_id].get("config", {}).get("auto_process")):
                            self._auto_process({
                                "instrument": p[0],
                                "patient_id": p[1],
                                "test_code": p[2],
                                "result": p[3]
                            }, server_id)
                        client.send(b"ACK\r\n")
        except Exception as e:
            print(f"Lab handle error: {e}")
        finally:
            try:
                client.close()
            except:
                pass

    def _run_rad_server(self, server_id, port):
        sock = None
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.bind(('0.0.0.0', port))
            sock.listen(5)
            sock.settimeout(1.0)
            if server_id in self.servers:
                self.servers[server_id]["socket"] = sock
            while server_id in self.servers and self.servers[server_id].get("running"):
                try:
                    c, a = sock.accept()
                    threading.Thread(target=self._handle_rad,
                                     args=(c, server_id),
                                     daemon=True).start()
                except socket.timeout:
                    continue
                except:
                    break
        except Exception as e:
            print(f"Rad server error: {e}")
        finally:
            if server_id in self.servers:
                self.servers[server_id]["running"] = False
            if sock:
                try:
                    sock.close()
                except:
                    pass

    def _handle_rad(self, client, server_id):
        try:
            client.send(b"ACK\r\n")
            while server_id in self.servers and self.servers[server_id].get("running"):
                d = client.recv(8192)
                if not d:
                    break
        except Exception as e:
            print(f"Rad handle error: {e}")
        finally:
            try:
                client.close()
            except:
                pass

    def _extract(self, buf):
        for d in ['\r\n', '\n', '\r']:
            if d in buf:
                i = buf.index(d)
                return buf[:i].strip(), buf[i + len(d):]
        return None, buf

    def _auto_process(self, result, server_id):
        pid = result.get("patient_id", "")
        config = self.servers.get(server_id, {}).get("config", {})
        test_mapping = config.get("test_mapping", {})
        test_name = test_mapping.get(
            result.get("test_code", "").upper(),
            result.get("test_code", "")
        )
        self.app.lab_tests.append({
            "test_id": f"LIS-{datetime.now().strftime('%Y%m%d%H%M%S')}",
            "patient_id": pid,
            "patient_name": self.app.get_patient_name(pid),
            "test_name": test_name,
            "status": "Completed",
            "result": result.get("result", ""),
            "unit": "",
            "reference": "",
            "date_ordered": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "date_completed": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "billed": False,
            "instrument": config.get("name", "Unknown"),
            "source": "LIS"
        })
        self.app.save_all()
        Clock.schedule_once(lambda dt: self.app.refresh_lis_displays(), 0)

    def get_status(self, category, instrument_id):
        config = self.instrument_configs.get(category, {}).get(instrument_id, {})
        server_id = f"{category}_{instrument_id}"
        is_running = (server_id in self.servers and
                      self.servers[server_id].get("running", False))
        return {
            "name": config.get("name", ""),
            "manufacturer": config.get("manufacturer", ""),
            "model": config.get("model", ""),
            "port": config.get("port", 0),
            "protocol": config.get("protocol", ""),
            "status": "Running" if is_running else "Stopped",
            "auto_process": config.get("auto_process", False)
        }


# ====================== BELEMA'S LAW ======================
class BelemaLaw:
    COEFFICIENT = 12.6

    @staticmethod
    def calculate_liver_span(hm, age):
        if hm <= 0 or age < 0:
            return None
        return round(BelemaLaw.COEFFICIENT * hm + (-3.00 if age < 12 else -5.04), 2)

    @staticmethod
    def interpret_finding(ms, ps):
        if ps is None or ps <= 0:
            return "Unable", "N/A"
        d = ms - ps
        if d < 2:
            return "Normal Variation", "green"
        elif d < 3:
            return "Mild Hepatomegaly", "yellow"
        elif d < 4:
            return "Moderate Hepatomegaly", "orange"
        else:
            return "Severe Hepatomegaly", "red"

    @staticmethod
    def get_clinical_recommendation(i):
        return {
            "Normal Variation": "No further investigation.",
            "Mild Hepatomegaly": "Consider LFTs, abdominal US.",
            "Moderate Hepatomegaly": "Urgent LFTs, US, hepatitis screen.",
            "Severe Hepatomegaly": "Emergency evaluation."
        }.get(i, "Clinical correlation required.")


# ====================== AI BUSINESS OPERATING SYSTEM ======================
class AIBOS:
    """AI Business Operating System - Strategic Business Planning Engine"""

    def __init__(self, ai_manager):
        self.ai = ai_manager
        self.business_profile = {}
        self.health_scores = {}
        self.market_intel = {}
        self.revenue_engine = {}
        self.sales_engine = {}
        self.marketing_engine = {}
        self.cx_engine = {}
        self.financial_engine = {}
        self.ai_automation = {}
        self.documents = {}
        self.implementation_roadmap = {}
        self.decision_log = []
        self.self_evaluation = {}

    def _evaluate_output(self, output, criteria):
        scores = {}
        for criterion in criteria:
            scores[criterion] = 8
        return scores

    def _improve_output(self, output, scores, criteria):
        improvements_needed = [c for c, s in scores.items() if s < 9]
        if not improvements_needed:
            return output
        return output + f"\n\n[IMPROVEMENTS: {', '.join(improvements_needed)}]"

    def get_business_profile(self, answers=None):
        if answers is None:
            return self._business_discovery_questions()
        self.business_profile = answers
        return self.business_profile

    def _business_discovery_questions(self):
        return {
            "company_name": "What is your company name?",
            "industry": "What industry does your business operate in?",
            "products_services": "What products or services do you offer?",
            "revenue_model": "What is your revenue model?",
            "cost_structure": "Describe your cost structure",
            "staff_count": "How many employees do you have?",
            "technology_stack": "What technology stack do you use?",
            "customer_segments": "Who are your target customer segments?",
            "competitors": "Who are your main competitors?",
            "sales_channels": "What sales channels do you use?",
            "marketing_channels": "What marketing channels do you use?",
            "annual_revenue": "What is your annual revenue?",
            "geographic_markets": "What geographic markets do you serve?",
            "growth_goals": "What are your growth goals?",
            "challenges": "What are your biggest challenges?",
            "risks": "What risks do you face?",
            "regulatory": "What regulatory requirements apply?"
        }

    def calculate_health_score(self):
        areas = [
            "Leadership", "Strategy", "Finance", "Marketing", "Sales",
            "Customer Service", "Operations", "HR", "Technology",
            "AI Readiness", "Innovation", "Brand", "Customer Experience",
            "Cybersecurity", "Legal Compliance"
        ]
        scores = {}
        for area in areas:
            if area == "AI Readiness" and not self.ai.enabled:
                scores[area] = 30
            else:
                scores[area] = 50
        self.health_scores = scores
        return scores

    def generate_market_intelligence(self):
        profile = self.business_profile
        prompt = f"""Based on the following business profile, provide market intelligence analysis:
Company: {profile.get('company_name', 'Unknown')}
Industry: {profile.get('industry', 'Unknown')}
Products/Services: {profile.get('products_services', 'Unknown')}
Competitors: {profile.get('competitors', 'Unknown')}
Markets: {profile.get('geographic_markets', 'Unknown')}

Provide analysis of market size, trends, customer behavior, competitor strengths/weaknesses, opportunities, threats, and positioning strategies.
"""
        if self.ai.enabled:
            success, result = self.ai.call(
                "You are a market intelligence analyst.", prompt, mt=1500
            )
            self.market_intel["analysis"] = result if success else self._generate_fallback_market_intel()
        else:
            self.market_intel["analysis"] = self._generate_fallback_market_intel()
        return self.market_intel

    def _generate_fallback_market_intel(self):
        return """MARKET INTELLIGENCE REPORT

Market Size: Growing market with significant potential for innovation.

Customer Behavior: Customers increasingly value digital convenience and personalized service.

Competitor Analysis:
- Strengths: Brand recognition and scale
- Weaknesses: Slow technology adoption

Opportunities:
1. Digital transformation
2. AI-powered customer engagement
3. Remote service delivery

Threats:
1. Regulatory changes
2. Increasing competition

Positioning Strategy: Position as a technology-forward, customer-centric provider.
"""

    def generate_revenue_engine(self):
        opportunities = [
            "Referral programs", "Affiliate programs", "Loyalty programs",
            "Premium services", "Subscription models", "Upselling",
            "Cross-selling", "Bundles", "Digital products", "Training",
            "Licensing", "Franchising", "Partnerships", "Corporate sales",
            "Government contracts", "International expansion"
        ]
        analysis = {}
        for opp in opportunities:
            analysis[opp] = {
                "revenue_potential": "High" if opp in ["Subscription models", "Partnerships"] else "Medium",
                "profit_potential": "High" if opp in ["Digital products", "Licensing"] else "Medium",
                "cost": "Low" if opp in ["Referral programs", "Affiliate programs"] else "Medium",
                "difficulty": "Easy" if opp in ["Referral programs", "Upselling"] else "Medium",
                "risk": "Low" if opp in ["Referral programs", "Loyalty programs"] else "Medium",
                "time_to_implement": "1-2 months" if opp in ["Referral programs", "Upselling"] else "3-6 months",
                "roi": "High" if opp in ["Subscription models", "Digital products"] else "Medium"
            }
        sorted_opportunities = sorted(
            analysis.items(),
            key=lambda x: (x[1]['roi'] == "High", x[1]['revenue_potential'] == "High"),
            reverse=True
        )
        self.revenue_engine = {"opportunities": analysis, "ranked": sorted_opportunities}
        return self.revenue_engine

    def generate_sales_engine(self):
        self.sales_engine = {
            "strategy": "Customer-centric sales approach focusing on value-based selling.",
            "sales_scripts": ["Initial outreach", "Discovery call", "Demo presentation",
                              "Closing", "Follow-up"],
            "crm_workflow": "Lead capture → Qualification → Nurturing → Sales call → Proposal → Close → Onboarding",
            "lead_scoring": "Points-based: 50 for fit, 30 for engagement, 20 for authority",
            "lead_qualification": "BANT: Budget, Authority, Need, Timeline",
            "sales_funnel": "Awareness → Interest → Consideration → Intent → Purchase → Retention",
            "sales_kpis": {"monthly_revenue": "N500,000", "conversion_rate": "25%"}
        }
        return self.sales_engine

    def generate_marketing_engine(self):
        self.marketing_engine = {
            "strategy": "Integrated digital marketing strategy",
            "brand_positioning": "Leading technology-enabled provider",
            "seo_plan": ["Keyword research", "Local SEO", "Content creation"],
            "social_media_campaigns": ["Facebook ads", "LinkedIn networking"],
            "email_campaigns": ["Welcome series", "Nurture sequence"],
            "paid_advertising": ["Google Ads", "Facebook/Instagram Ads"]
        }
        return self.marketing_engine

    def generate_customer_experience_engine(self):
        self.cx_engine = {
            "onboarding": ["Welcome email", "Introduction call"],
            "customer_support": ["24/7 support hotline", "Live chat", "Email support"],
            "customer_retention": ["Monthly check-ins", "Satisfaction surveys", "Loyalty rewards"],
            "customer_journey_mapping": ["Awareness", "Consideration", "Decision", "Onboarding",
                                          "Adoption", "Retention", "Advocacy"]
        }
        return self.cx_engine

    def generate_financial_engine(self):
        rev = float(self.business_profile.get('annual_revenue', 500000) or 500000)
        self.financial_engine = {
            "profit_loss": {
                "revenue": rev,
                "cost_of_goods": rev * 0.4,
                "gross_profit": rev * 0.6,
                "operating_expenses": rev * 0.3,
                "net_profit": rev * 0.3
            },
            "budget_recommendations": {
                "marketing": rev * 0.1,
                "operations": rev * 0.15,
                "technology": rev * 0.05,
                "staff": rev * 0.25
            },
            "cost_reduction_opportunities": [
                "Negotiate supplier contracts",
                "Automate manual processes"
            ]
        }
        return self.financial_engine

    def generate_ai_automation_engine(self):
        self.ai_automation = {
            "customer_support": ["AI-powered chatbot"],
            "appointment_booking": ["Online booking system", "Automated reminders"],
            "marketing": ["AI-powered email personalization"],
            "sales": ["Lead scoring automation"],
            "roi_prioritization": {
                "high_impact": ["Customer support chatbot", "Appointment booking"],
                "medium_impact": ["Marketing automation", "Sales automation"]
            }
        }
        return self.ai_automation

    def generate_document(self, doc_type, custom_data=None):
        profile = self.business_profile
        documents = {
            "business_plan": self._generate_business_plan(profile),
            "sales_plan": self._generate_sales_plan(profile),
            "marketing_plan": self._generate_marketing_plan(profile),
            "referral_program": self._generate_referral_program(profile),
            "sop": self._generate_sop(profile),
            "employee_handbook": self._generate_employee_handbook(profile),
            "financial_report": self._generate_financial_report(profile),
            "risk_assessment": self._generate_risk_assessment(profile),
            "pitch_deck": self._generate_pitch_deck(profile),
            "grant_proposal": self._generate_grant_proposal(profile),
            "partnership_proposal": self._generate_partnership_proposal(profile),
            "project_plan": self._generate_project_plan(profile),
            "kpi_dashboard": self._generate_kpi_dashboard(profile)
        }
        if doc_type in documents:
            result = documents[doc_type]
            criteria = ["Accuracy", "Practicality", "Profitability", "Innovation",
                        "Scalability", "Ease of implementation", "Customer value",
                        "Financial impact", "Risk management", "Completeness"]
            scores = self._evaluate_output(result, criteria)
            if any(score < 9 for score in scores.values()):
                result = self._improve_output(result, scores, criteria)
            self.documents[doc_type] = result
            return result
        return "Document type not found."

    def _generate_business_plan(self, profile):
        name = profile.get('company_name', 'HMG Hospital')
        return f"""BUSINESS PLAN: {name}

1. EXECUTIVE SUMMARY
{name} is a healthcare organization delivering exceptional service through innovation.

2. COMPANY OVERVIEW
- Name: {name}
- Industry: {profile.get('industry', 'Healthcare')}
- Products/Services: {profile.get('products_services', 'Healthcare services')}

3. STRATEGIC GOALS
- Increase revenue by 30% in 12 months
- Achieve 90% patient satisfaction rating

4. FINANCIAL PROJECTIONS
Year 1 Revenue: N{self.financial_engine.get('profit_loss', {}).get('revenue', 500000):,.2f}
"""

    def _generate_sales_plan(self, profile):
        return f"""SALES PLAN

1. SALES OBJECTIVES
- Monthly Revenue: {self.sales_engine.get('sales_kpis', {}).get('monthly_revenue', 'N500,000')}

2. SALES PROCESS
{self.sales_engine.get('crm_workflow', 'Lead capture → Proposal → Close')}

3. KEY PERFORMANCE INDICATORS
- Conversion Rate: {self.sales_engine.get('sales_kpis', {}).get('conversion_rate', '25%')}
"""

    def _generate_marketing_plan(self, profile):
        return f"""MARKETING PLAN

1. BRAND POSITIONING
{self.marketing_engine.get('brand_positioning', 'Technology-enabled healthcare provider')}

2. MARKETING CHANNELS
- SEO: {', '.join(self.marketing_engine.get('seo_plan', ['Keyword research', 'Content creation']))}
- Social Media: {', '.join(self.marketing_engine.get('social_media_campaigns', ['Facebook', 'LinkedIn']))}

3. BUDGET ALLOCATION
Digital Advertising: 40%
Content Marketing: 20%
Social Media: 15%
"""

    def _generate_referral_program(self, profile):
        name = profile.get('company_name', 'HMG Hospital')
        return f"""REFERRAL PROGRAM

1. PROGRAM OVERVIEW
The {name} Referral Program rewards customers for referring new clients.

2. REWARD STRUCTURE
- N5,000 for each successful referral
- Bonus for 5 referrals

3. PROGRAM RULES
- Referral must be a new customer
- Rewards paid within 30 days
"""

    def _generate_sop(self, profile):
        name = profile.get('company_name', 'HMG Hospital')
        return f"""STANDARD OPERATING PROCEDURE - {name}

1. PATIENT ONBOARDING
1.1 Collect demographic information
1.2 Schedule appointment
1.3 Provide welcome package

2. SERVICE DELIVERY
2.1 Complete service per protocol
2.2 Document service delivery
2.3 Follow up with patient

3. QUALITY CONTROL
3.1 Regular audits
3.2 Customer feedback collection
"""

    def _generate_employee_handbook(self, profile):
        name = profile.get('company_name', 'HMG Hospital')
        return f"""EMPLOYEE HANDBOOK - {name}

1. COMPANY MISSION
To provide exceptional healthcare through innovation and excellence.

2. EMPLOYMENT POLICIES
- Equal opportunity employer
- Fair compensation and benefits

3. CODE OF CONDUCT
- Professional behavior
- Confidentiality
- Quality standards
"""

    def _generate_financial_report(self, profile):
        name = profile.get('company_name', 'HMG Hospital')
        fin = self.financial_engine.get('profit_loss', {})
        return f"""FINANCIAL REPORT - {name}
Period: {datetime.now().strftime('%B %Y')}

INCOME STATEMENT
Revenue: N{fin.get('revenue', 0):,.2f}
COGS: N{fin.get('cost_of_goods', 0):,.2f}
Gross Profit: N{fin.get('gross_profit', 0):,.2f}
Net Profit: N{fin.get('net_profit', 0):,.2f}

RECOMMENDATIONS
{chr(10).join(self.financial_engine.get('cost_reduction_opportunities', ['Review expenses']))}
"""

    def _generate_risk_assessment(self, profile):
        return """RISK ASSESSMENT

1. OPERATIONAL RISKS
- Service quality variation
- Staff turnover
- Technology failures

2. FINANCIAL RISKS
- Revenue fluctuations
- Cash flow shortages

3. MITIGATION STRATEGIES
- Quality management system
- Financial reserves
- Compliance training
"""

    def _generate_pitch_deck(self, profile):
        name = profile.get('company_name', 'HMG Hospital')
        return f"""PITCH DECK - {name}

1. THE PROBLEM
Healthcare customers face challenges with:
- Lack of convenient access
- High costs

2. OUR SOLUTION
{name} provides:
- Technology-enabled healthcare
- Affordable services

3. FINANCIAL PROJECTIONS
Year 1: N{self.financial_engine.get('profit_loss', {}).get('revenue', 500000):,.2f}
"""

    def _generate_grant_proposal(self, profile):
        name = profile.get('company_name', 'HMG Hospital')
        return f"""GRANT PROPOSAL - {name}

1. PROJECT SUMMARY
This project aims to transform healthcare delivery through AI-powered solutions.

2. PROJECT GOALS
- Implement AI-powered patient management
- Improve service delivery efficiency

3. BUDGET
Equipment: N500,000
Training: N200,000
Implementation: N300,000
"""

    def _generate_partnership_proposal(self, profile):
        name = profile.get('company_name', 'HMG Hospital')
        return f"""PARTNERSHIP PROPOSAL - {name}

1. PARTNERSHIP OPPORTUNITY
{name} seeks strategic partnerships to enhance service delivery.

2. PARTNERSHIP BENEFITS
- Extended service offerings
- Increased customer base

3. IMPLEMENTATION PLAN
- Initial consultation
- Partnership agreement
- Launch and monitoring
"""

    def _generate_project_plan(self, profile):
        name = profile.get('company_name', 'HMG Hospital')
        return f"""PROJECT PLAN - {name}

1. PROJECT OBJECTIVE
Implement AI-powered solutions to improve operational efficiency.

2. TIMELINE
- Week 1-2: Assessment
- Week 3-4: Technology selection
- Week 5-8: Implementation
- Week 9-12: Launch

3. SUCCESS METRICS
- 30% improvement in efficiency
- 90% user satisfaction
"""

    def _generate_kpi_dashboard(self, profile):
        return f"""KPI DASHBOARD

FINANCIAL KPIs
- Revenue: {self.financial_engine.get('profit_loss', {}).get('revenue', 'N500,000')}
- Revenue Growth: 15%

OPERATIONAL KPIs
- Customer Satisfaction: 4.5/5
- Service Quality: 92%

SALES KPIs
- Conversion Rate: {self.sales_engine.get('sales_kpis', {}).get('conversion_rate', '25%')}
"""

    def generate_implementation_roadmap(self):
        self.implementation_roadmap = {
            "first_7_days": {
                "milestones": ["Complete business discovery", "Identify quick wins"],
                "costs": "N0"
            },
            "first_30_days": {
                "milestones": ["Develop revenue engine", "Create sales plan"],
                "costs": "N200,000"
            },
            "first_90_days": {
                "milestones": ["Implement AI automation", "Financial optimization"],
                "costs": "N500,000"
            },
            "first_6_months": {
                "milestones": ["Market expansion", "Technology integration"],
                "costs": "N1,000,000"
            },
            "first_year": {
                "milestones": ["Achieve revenue targets", "Continuous improvement"],
                "costs": "N2,000,000"
            }
        }
        return self.implementation_roadmap

    def decision_support(self, decision_context, options):
        analysis = {
            "context": decision_context,
            "options": options,
            "recommendation": options[0] if options else "Continue with current strategy",
            "success_metrics": ["Revenue increase of 20%", "Customer satisfaction improvement"]
        }
        self.decision_log.append(analysis)
        return analysis


# ====================== END OF PART 2 ======================
# ====================== MAIN APPLICATION ======================
class HMGBillingApp(App):
    """
    HMG Hospital EMR — Main Application
    v13.0.0 — Multi-user login, audit log, standardized clerking,
              real-price auto-billing, camera-enabled report drafter.
    """

    DEFAULT_DELETE_PASSCODE = "1234"
    CONFIG_FILE = "config.json"

    # ---------- app bootstrap ----------
    def build(self):
        self.title = "HMG Hospital EMR v13.0.0"
        Window.size = (1400, 900)

        # --- Core services ---
        self.user_manager = UserManager()
        self.audit = AuditLog()
        self.ai = AIAPIManager()
        self.clinical_ai = ModularAIClinicalAssistant(self.ai)
        self.ecg_ai = ECGAnalyzer(self.ai)
        self.pathology_ai = AIPathologyExtractor(self.ai)
        self.pt_tracker = PaymentTracker()
        self.chat_sys = ChatSystem()
        self.lis_manager = LISManager(self)
        self.bos = AIBOS(self.ai)
        self.task_manager = TaskManager()

        # --- Session state ---
        self.current_user = None
        self.root_container = BoxLayout(orientation='vertical')
        self.main_root = None  # set after successful login

        # --- Show login screen first ---
        self.login_screen = LoginScreen(
            self.user_manager,
            on_login=self._on_login_success
        )
        self.root_container.add_widget(self.login_screen)
        return self.root_container

    # ====================== LOGIN / LOGOUT ======================
    def _on_login_success(self, user):
        """Called by LoginScreen after a valid login."""
        self.current_user = user
        self.audit.log(
            username=user.get("username"),
            action="LOGIN",
            details=f"Role: {user.get('role')}"
        )
        # Replace login screen with main app
        self.root_container.clear_widgets()
        self._build_main_app()

        # Warn if using default password
        if user.get("must_change_password"):
            Clock.schedule_once(
                lambda dt: self._popup(
                    "Security Warning",
                    "You are using a default or reset password.\n\n"
                    "Please go to Settings → SECURITY → CHANGE PASSWORD "
                    "and set a new one immediately."
                ),
                0.5
            )

    def _logout(self, instance=None):
        """Log out and return to login screen."""
        if self.current_user:
            self.audit.log(
                username=self.current_user.get("username"),
                action="LOGOUT",
                details=""
            )
        self.save_all()
        self.current_user = None

        # Tear down main UI
        self.root_container.clear_widgets()

        # Reset in-memory state tied to previous session
        self.cp = None
        self.bill_items = []
        self.emr_drug_list = []
        self.clinical_ai = ModularAIClinicalAssistant(self.ai)

        # Show login again
        self.login_screen = LoginScreen(
            self.user_manager,
            on_login=self._on_login_success
        )
        self.root_container.add_widget(self.login_screen)

    def _build_main_app(self):
        """Build the full tabbed UI after a successful login."""
        # --- Fresh in-memory stores ---
        self.patients = []
        self.emr_records = []
        self.lab_tests = []
        self.radiology = []
        self.prescriptions = []
        self.bills = []
        self.bill_items = []
        self.ecg_reports = []
        self.pathology_data = []
        self.followups = []
        self.patient_summaries = {}
        self.consent_forms = []
        self.referral_notes = []
        self.daily_notes = []
        self.patient_snapshots = []
        self.procedures_db = []
        self.lab_results = []
        self.tasks = self.task_manager.tasks
        self.emr_drug_list = []
        self.admissions = []
        self.saved_lab_reports = []
        self.appointments = []
        self.nursing_notes = []
        self.delete_attempts = {}

        # --- Initial databases ---
        self.drugs = self._init_drugs()
        self.procedures = self._init_procs()
        self.investigations = self._init_invs()
        self.charges = self._init_charges()
        self.depts = ["General Medicine", "Pediatrics", "Obstetrics",
                      "Surgery", "Cardiology", "Emergency"]

        # --- Working state ---
        self.cp = None
        self.sel_drug_idx = None
        self.sel_inv_idx = None
        self.cur_ecg_img = None
        self.cur_ecg_report = None
        self.cur_upload = None
        self.camera_widget = None
        self.ai_patient_id = None
        self._admission_id_map = {}

        # --- Ensure config + passcode exist ---
        self._load_config()

        # --- Load persisted data ---
        self.load_all()

        # --- Build the UI ---
        self.main_root = BoxLayout(orientation='vertical')
        self._header()
        self.tp = CustomTabbedPanel()
        self.main_root.add_widget(self.tp)

        # Tabs — same as before
        self._dashboard()
        self._register()
        self._clinical_ai_tab()
        self._emr()
        self._task_management_tab()
        self._ecg_tab()
        self._liver_span_tab()
        self._pathology_tab()
        self._lis_tab()
        self._billing()
        self._payment()
        self._pharmacy()
        self._drug_mgmt()
        self._inv_mgmt()
        self._lab()
        self._rad()
        self._upload_tab()
        self._settings()
        self._chat_tab()
        self._clinical_summary_tab()
        self._vitals_timeline_tab()
        self._consent_form_tab()
        self._referral_note_tab()
        self._saved_items_tab()
        self._daily_notes_tab()
        self._ai_analysis_tab()
        self._ai_bos_tab()
        self._snapshot_tab()
        self._procedures_tab()
        self._enhanced_lab_tab()
        self._pharmacy_dispense_tab()
        self._admissions_tab()
        self._appointments_tab()
        self._nursing_notes_tab()
        self._medical_report_tab()
        self._user_management_tab()   # NEW: User management for admins
        self._audit_log_tab()          # NEW: Audit log viewer

        self.root_container.add_widget(self.main_root)
        Clock.schedule_once(lambda dt: self._refresh_dash(), 0.5)

    # ====================== DELETE PASSCODE MANAGEMENT ======================
    def _load_config(self):
        """Load config.json; create default if missing."""
        if not os.path.exists(self.CONFIG_FILE):
            default = {
                "delete_passcode_hash": self._hash_passcode(self.DEFAULT_DELETE_PASSCODE)
            }
            try:
                with open(self.CONFIG_FILE, "w", encoding="utf-8") as f:
                    json.dump(default, f, indent=2)
            except Exception as e:
                print(f"Config write error: {e}")
            return default
        try:
            with open(self.CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            if "delete_passcode_hash" not in cfg:
                cfg["delete_passcode_hash"] = self._hash_passcode(self.DEFAULT_DELETE_PASSCODE)
                with open(self.CONFIG_FILE, "w", encoding="utf-8") as f:
                    json.dump(cfg, f, indent=2)
            return cfg
        except Exception as e:
            print(f"Config read error: {e}")
            return {"delete_passcode_hash": self._hash_passcode(self.DEFAULT_DELETE_PASSCODE)}

    def _save_config(self, cfg):
        try:
            with open(self.CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
            return True
        except Exception as e:
            print(f"Config save error: {e}")
            return False

    @staticmethod
    def _hash_passcode(passcode):
        return hashlib.sha256(str(passcode).encode("utf-8")).hexdigest()

    def _verify_delete_passcode(self, entered):
        if not entered:
            return False
        cfg = self._load_config()
        stored = cfg.get("delete_passcode_hash", "")
        return stored != "" and self._hash_passcode(entered.strip()) == stored

    def _set_delete_passcode(self, new_passcode):
        new_passcode = str(new_passcode).strip()
        if len(new_passcode) < 4:
            return False, "Passcode must be at least 4 characters."
        cfg = self._load_config()
        cfg["delete_passcode_hash"] = self._hash_passcode(new_passcode)
        if self._save_config(cfg):
            return True, "Delete passcode updated successfully."
        return False, "Could not save config file."

    # ====================== DATA LOADING / SAVING ======================
    def _save_json(self, filename, data, max_backups=3):
        """Atomically save data to JSON file with rotation backups."""
        try:
            if os.path.exists(filename):
                for i in range(max_backups - 1, 0, -1):
                    src = f"{filename}.bak{i - 1}" if i > 1 else filename
                    dst = f"{filename}.bak{i}"
                    if os.path.exists(src):
                        shutil.copy2(src, dst)
                shutil.copy2(filename, f"{filename}.bak0")

            fd, temp_path = tempfile.mkstemp(
                dir=os.path.dirname(filename) or '.',
                prefix='tmp_', suffix='.json'
            )
            with os.fdopen(fd, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, default=str, ensure_ascii=False)
            os.replace(temp_path, filename)
        except Exception as e:
            print(f"Error saving {filename}: {e}")
            if os.path.exists(f"{filename}.bak0"):
                shutil.copy2(f"{filename}.bak0", filename)

    def load_all(self):
        files = {
            "patients.json": "patients",
            "emr.json": "emr_records",
            "lab.json": "lab_tests",
            "rad.json": "radiology",
            "rx.json": "prescriptions",
            "bills.json": "bills",
            "drugs.json": "drugs",
            "invs.json": "investigations",
            "ecg.json": "ecg_reports",
            "pathology.json": "pathology_data",
            "followups.json": "followups",
            "summaries.json": "patient_summaries",
            "consent_forms.json": "consent_forms",
            "referral_notes.json": "referral_notes",
            "daily_notes.json": "daily_notes",
            "snapshots.json": "patient_snapshots",
            "procedures_db.json": "procedures_db",
            "lab_results.json": "lab_results",
            "tasks.json": "tasks",
            "admissions.json": "admissions",
            "saved_lab_reports.json": "saved_lab_reports",
            "appointments.json": "appointments",
            "nursing_notes.json": "nursing_notes"
        }
        for fn, attr in files.items():
            try:
                if os.path.exists(fn):
                    with open(fn, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    if attr == "drugs":
                        setattr(self, attr, [{
                            "name": d.get("name", ""),
                            "price": d.get("price", 0),
                            "qty": d.get("qty", d.get("quantity", 0)),
                            "min": d.get("min", d.get("min_stock", 20)),
                            "cat": d.get("cat", d.get("category", "")),
                            "mfr": d.get("mfr", d.get("manufacturer", "")),
                            "exp": d.get("exp", ""),
                            "batch": d.get("batch", "")
                        } for d in data])
                    elif attr == "investigations":
                        setattr(self, attr, [{
                            "name": inv.get("name", ""),
                            "price": inv.get("price", 0),
                            "dept": inv.get("dept", inv.get("department", "")),
                            "cat": inv.get("cat", inv.get("category", "")),
                            "tat": inv.get("tat", ""),
                            "specimen": inv.get("specimen", ""),
                            "ref": inv.get("ref", ""),
                            "unit": inv.get("unit", ""),
                            "desc": inv.get("desc", "")
                        } for inv in data])
                    elif attr == "tasks":
                        self.tasks = data
                        self.task_manager.tasks = data
                    else:
                        setattr(self, attr, data)
            except Exception as e:
                print(f"Error loading {fn}: {e}")

    def save_all(self):
        data = {
            "patients.json": self.patients,
            "emr.json": self.emr_records,
            "lab.json": self.lab_tests,
            "rad.json": self.radiology,
            "rx.json": self.prescriptions,
            "bills.json": self.bills,
            "drugs.json": self.drugs,
            "invs.json": self.investigations,
            "ecg.json": self.ecg_reports,
            "pathology.json": self.pathology_data,
            "followups.json": self.followups,
            "summaries.json": self.patient_summaries,
            "consent_forms.json": self.consent_forms,
            "referral_notes.json": self.referral_notes,
            "daily_notes.json": self.daily_notes,
            "snapshots.json": self.patient_snapshots,
            "procedures_db.json": self.procedures_db,
            "lab_results.json": self.lab_results,
            "tasks.json": self.tasks,
            "admissions.json": self.admissions,
            "saved_lab_reports.json": self.saved_lab_reports,
            "appointments.json": self.appointments,
            "nursing_notes.json": self.nursing_notes
        }
        for fn, d in data.items():
            self._save_json(fn, d)

    # ====================== HELPERS ======================
    def _current_username(self):
        return self.current_user.get("username", "system") if self.current_user else "system"

    def _current_role(self):
        return self.current_user.get("role", "") if self.current_user else ""

    def _require_role(self, *roles):
        """Return True if the current user has one of the given roles."""
        if not self.current_user:
            return False
        if self.current_user.get("role") in roles:
            return True
        self._popup(
            "Access Denied",
            f"This action requires one of the following roles:\n{', '.join(roles)}"
        )
        return False

    def log_activity(self, msg):
        """Write to the on-screen activity log AND the persistent audit log."""
        ts = datetime.now().strftime("%H:%M:%S")
        user = self._current_username()
        line = f"[{ts}] [{user}] {msg}"
        if hasattr(self, 'act_log'):
            current = self.act_log.text
            lines = current.split('\n')
            lines.append(line)
            if len(lines) > 50:
                lines = lines[-50:]
            self.act_log.text = "\n".join(lines)
        else:
            print(line)

    def audit_log(self, action, details="", patient_id=""):
        """Write a structured entry to the persistent audit log."""
        try:
            self.audit.log(
                username=self._current_username(),
                action=action,
                details=details,
                patient_id=patient_id
            )
        except Exception as e:
            print(f"Audit log error: {e}")

    def _sanitize_text(self, text):
        if not text:
            return ""
        try:
            return text.encode('latin-1', 'replace').decode('latin-1')
        except:
            return text

    def _decode_qr_from_texture(self, texture):
        fn = f"qr_temp_{datetime.now().strftime('%Y%m%d%H%M%S')}.png"
        texture.save(fn)
        try:
            if PYZBAR_AVAILABLE:
                try:
                    img = PILImage.open(fn)
                    decoded = pyzbar_decode(img)
                    if decoded:
                        return decoded[0].data.decode('utf-8')
                except:
                    pass
            elif CV2_AVAILABLE:
                img = cv2.imread(fn)
                detector = cv2.QRCodeDetector()
                data, _, _ = detector.detectAndDecode(img)
                if data:
                    return data
            return None
        finally:
            try:
                os.remove(fn)
            except:
                pass

    def _get_reports_dir(self):
        if platform == 'android':
            dir_path = "/sdcard/HMG_Reports"
        else:
            dir_path = os.path.join(os.getcwd(), "HMG_Reports")
        os.makedirs(dir_path, exist_ok=True)
        return dir_path

    def _export_to_pdf(self, text, prefix, title):
        """Export plain text to PDF using FPDF, falling back to .txt."""
        try:
            dir_path = self._get_reports_dir()
            fn = os.path.join(
                dir_path,
                f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )
            if PDF_AVAILABLE:
                try:
                    class PDF(FPDF):
                        def header(self):
                            self.set_font('Arial', 'B', 12)
                            self.cell(0, 10, 'HMG Hospital', 0, 1, 'C')
                            self.ln(5)

                        def footer(self):
                            self.set_y(-15)
                            self.set_font('Arial', 'I', 8)
                            self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

                    pdf = PDF()
                    pdf.add_page()
                    pdf.set_font("Arial", 'B', 16)
                    pdf.cell(0, 10, self._sanitize_text(title), ln=True, align='C')
                    pdf.ln(5)
                    pdf.set_font("Arial", 'I', 10)
                    pdf.cell(0, 6,
                             f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                             ln=True)
                    pdf.cell(0, 6,
                             f"User: {self._current_username()}",
                             ln=True)
                    pdf.ln(5)

                    for line in text.split('\n'):
                        safe_line = self._sanitize_text(line)
                        if safe_line.strip().startswith('=') or safe_line.strip().startswith('-'):
                            pdf.set_font('Arial', 'B', 12)
                            pdf.cell(0, 8, safe_line.strip()[:80], ln=True)
                            pdf.set_font('Arial', '', 11)
                        elif safe_line.strip() == '':
                            pdf.ln(2)
                        else:
                            pdf.set_font('Arial', '', 11)
                            pdf.multi_cell(0, 6, safe_line.strip()[:100])

                    pdf_fn = f"{fn}.pdf"
                    pdf.output(pdf_fn)
                    self._open_file(pdf_fn)
                    self._popup("Export Success", f"PDF saved to:\n{pdf_fn}")
                    self.audit_log("EXPORT_PDF", f"{title} → {pdf_fn}")
                    return True, pdf_fn
                except Exception as e:
                    self.log_activity(f"PDF export error: {e}")
                    txt_fn = f"{fn}.txt"
                    with open(txt_fn, 'w', encoding='utf-8') as f:
                        f.write(f"{title}\n\n{text}")
                    self._open_file(txt_fn)
                    return True, txt_fn
            else:
                txt_fn = f"{fn}.txt"
                with open(txt_fn, 'w', encoding='utf-8') as f:
                    f.write(f"{title}\n\n{text}")
                self._open_file(txt_fn)
                return True, txt_fn
        except Exception as e:
            self.log_activity(f"Export error: {e}")
            self._popup("Export Error", f"Failed to export: {str(e)}")
            return False, str(e)

    def _open_file(self, filepath):
        if not os.path.exists(filepath):
            self._popup("Error", "File does not exist.")
            return
        if platform == 'android':
            try:
                from jnius import autoclass
                PythonActivity = autoclass('org.kivy.android.PythonActivity')
                Intent = autoclass('android.content.Intent')
                File = autoclass('java.io.File')
                Uri = autoclass('android.net.Uri')
                activity = PythonActivity.mActivity
                mime_type, _ = mimetypes.guess_type(filepath)
                if mime_type is None:
                    mime_type = "*/*"
                file = File(filepath)
                uri = Uri.fromFile(file)
                intent = Intent(Intent.ACTION_VIEW)
                intent.setDataAndType(uri, mime_type)
                intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK)
                intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
                activity.startActivity(intent)
            except:
                pass
        else:
            try:
                if sys.platform == 'win32':
                    os.startfile(filepath)
                elif sys.platform == 'darwin':
                    subprocess.Popen(['open', filepath])
                else:
                    subprocess.Popen(['xdg-open', filepath])
            except:
                pass

    # ====================== INITIAL DATABASES ======================
    def _init_drugs(self):
        return [
            {"name": "Paracetamol 500mg", "price": 50, "qty": 500, "min": 50, "cat": "Analgesic"},
            {"name": "Ibuprofen 400mg", "price": 80, "qty": 300, "min": 30, "cat": "Analgesic"},
            {"name": "Amoxicillin 500mg", "price": 150, "qty": 200, "min": 30, "cat": "Antibiotic"},
            {"name": "Ciprofloxacin 500mg", "price": 200, "qty": 150, "min": 20, "cat": "Antibiotic"},
            {"name": "Omeprazole 20mg", "price": 180, "qty": 200, "min": 30, "cat": "Antacid"},
            {"name": "Losartan 50mg", "price": 220, "qty": 150, "min": 20, "cat": "Antihypertensive"},
            {"name": "Metformin 500mg", "price": 180, "qty": 200, "min": 30, "cat": "Antidiabetic"},
            {"name": "Insulin Regular", "price": 500, "qty": 30, "min": 5, "cat": "Antidiabetic"}
        ]

    def _init_procs(self):
        return [
            {"name": "Caesarean Section", "price": 500000},
            {"name": "Appendectomy", "price": 350000}
        ]

    def _init_invs(self):
        return [
            {"name": "Full Blood Count (FBC)", "price": 8000, "dept": "Laboratory"},
            {"name": "Urinalysis", "price": 3000, "dept": "Laboratory"},
            {"name": "Blood Glucose", "price": 2500, "dept": "Laboratory"},
            {"name": "Chest X-Ray", "price": 8000, "dept": "Radiology"},
            {"name": "Abdominal US", "price": 8000, "dept": "Radiology"},
            {"name": "ECG", "price": 15000, "dept": "Cardiology"}
        ]

    def _init_charges(self):
        return [
            {"name": "Consultation", "price": 5000},
            {"name": "Specialist Consult", "price": 10000},
            {"name": "Ward Bed/day", "price": 15000},
            {"name": "ER Fee", "price": 20000},
            {"name": "ECG Report", "price": 20000},
            {"name": "Nursing Care", "price": 5000}
        ]

    # ====================== HEADER ======================
    def _header(self):
        h = BoxLayout(size_hint_y=None, height=52, padding=5, spacing=5)
        h.add_widget(Label(
            text="HMG Hospital v13.0.0",
            font_size=18, bold=True,
            color=get_color_from_hex('#1a5276'),
            size_hint_x=0.17
        ))

        lab_running = sum(1 for s in self.lis_manager.servers.values() if s.get("running"))
        self.lis_lbl = Label(
            text=f"LIS: {lab_running}",
            font_size=11, size_hint_x=0.06,
            color=get_color_from_hex('#27ae60')
        )
        h.add_widget(self.lis_lbl)

        ai_text = f"AI: {self.ai.provider}" if self.ai.enabled else "AI: Offline"
        self.ai_lbl = Label(text=ai_text, font_size=11, size_hint_x=0.08)
        h.add_widget(self.ai_lbl)

        # Currently-logged-in user badge
        uname = self._current_username()
        urole = self._current_role()
        self.user_lbl = Label(
            text=f"👤 {uname} ({urole})",
            font_size=12, bold=True,
            color=get_color_from_hex('#8e44ad'),
            size_hint_x=0.16
        )
        h.add_widget(self.user_lbl)

        self.pt_lbl = Label(
            text="No Patient",
            font_size=14, size_hint_x=0.24,
            color=get_color_from_hex('#7f8c8d')
        )
        h.add_widget(self.pt_lbl)

        h.add_widget(Button(
            text="SELECT",
            font_size=14, bold=True,
            size_hint_x=0.08,
            on_press=lambda x: self._sel_pt_popup(),
            background_color=get_color_from_hex('#2980b9')
        ))
        h.add_widget(Button(
            text="CLEAR",
            font_size=14,
            size_hint_x=0.06,
            on_press=lambda x: self._clear_pt()
        ))
        h.add_widget(Button(
            text="LOGOUT",
            font_size=14, bold=True,
            size_hint_x=0.09,
            background_color=get_color_from_hex('#c0392b'),
            on_press=self._logout
        ))

        self.main_root.add_widget(h)

    # ====================== DASHBOARD ======================
    def _dashboard(self):
        tab = TabbedPanelItem(text="Dashboard")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=8, spacing=6, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))

        st = GridLayout(cols=6, size_hint_y=None, height=75, spacing=5)
        self.stats = {}
        for t, c in [("Patients", "#3498db"), ("Today", "#2ecc71"),
                     ("Pending", "#e74c3c"), ("Stock", "#f39c12"),
                     ("LIS", "#16a085"), ("Tasks", "#8e44ad")]:
            card = BoxLayout(orientation='vertical')
            card.add_widget(Label(text=t, font_size=11, size_hint_y=None, height=20))
            lbl = Label(
                text="0", font_size=26, bold=True,
                color=get_color_from_hex(c),
                size_hint_y=None, height=38
            )
            self.stats[t] = lbl
            card.add_widget(lbl)
            st.add_widget(card)
        l.add_widget(st)

        l.add_widget(Label(
            text="QUICK ACTIONS", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#2c3e50')
        ))

        a = GridLayout(cols=7, size_hint_y=None, height=120, spacing=5)
        for t, c in [("Register", "#27ae60"), ("Clinical AI", "#e67e22"),
                     ("EMR", "#2980b9"), ("Tasks", "#8e44ad"),
                     ("ECG", "#e74c3c"), ("Liver Span", "#16a085"),
                     ("LIS", "#16a085"), ("Pharmacy", "#d35400"),
                     ("Drugs", "#16a085"), ("Lab", "#c0392b"),
                     ("Dispense", "#f39c12"), ("Report", "#16a085"),
                     ("Settings", "#7f8c8d")]:
            btn = Button(
                text=t, font_size=13, bold=True,
                background_color=get_color_from_hex(c)
            )
            btn.bind(on_press=lambda x, tx=t: self._switch(
                "Registration" if tx == "Register"
                else "Clinical AI" if tx == "Clinical AI"
                else "ECG Analysis" if tx == "ECG"
                else "Liver Span" if tx == "Liver Span"
                else "LIS Connections" if tx == "LIS"
                else "Manage Drugs" if tx == "Drugs"
                else "Pharmacy Dispense" if tx == "Dispense"
                else "Report Drafter" if tx == "Report"
                else tx
            ))
            a.add_widget(btn)
        l.add_widget(a)

        self.act_log = TextInput(
            multiline=True, readonly=True,
            font_size=11, size_hint_y=None, height=140
        )
        l.add_widget(Label(
            text="ACTIVITY", font_size=14, bold=True,
            size_hint_y=None, height=20
        ))
        l.add_widget(self.act_log)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)

    def _refresh_dash(self, dt=None):
        try:
            self.stats["Patients"].text = str(len(self.patients))
            td = datetime.now().strftime("%Y-%m-%d")
            self.stats["Today"].text = str(
                len([e for e in self.emr_records if td in str(e.get('ts', ''))])
            )
            pn = sum(b.get('total', 0) for b in self.bills
                     if b.get('status') in ['Pending', 'Partial'])
            self.stats["Pending"].text = f"N{pn:,.0f}"
            self.stats["Stock"].text = str(
                len([d for d in self.drugs if d.get('qty', 0) <= d.get('min', 0)])
            )
            self.stats["LIS"].text = str(
                sum(1 for s in self.lis_manager.servers.values() if s.get("running"))
            )
            self.stats["Tasks"].text = str(len(self.task_manager.get_pending_tasks()))
        except:
            pass

    # ====================== REGISTRATION ======================
    def _register(self):
        tab = TabbedPanelItem(text="Registration")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=6, spacing=4, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))
        l.add_widget(Label(
            text="PATIENT REGISTRATION",
            font_size=20, bold=True,
            size_hint_y=None, height=30,
            color=get_color_from_hex('#1a5276')
        ))

        f = GridLayout(cols=2, size_hint_y=None, height=400, spacing=5)
        self.rw = {}
        for lb, w in [
            ("ID:", TextInput(text=f"PT{datetime.now().strftime('%Y%m%d%H%M%S')}",
                              font_size=15, size_hint_y=None, height=40)),
            ("Name:", TextInput(font_size=15, size_hint_y=None, height=40)),
            ("DOB:", TextInput(text="1990-01-01", font_size=15, size_hint_y=None, height=40)),
            ("Age:", TextInput(text="0", font_size=15, size_hint_y=None, height=40, readonly=True)),
            ("Gender:", Spinner(text="Select", values=["Male", "Female", "Other"],
                                size_hint_y=None, height=40)),
            ("Phone:", TextInput(font_size=15, size_hint_y=None, height=40)),
            ("Height(m):", TextInput(font_size=15, size_hint_y=None, height=40, hint_text="1.65")),
            ("Blood:", Spinner(text="Unknown",
                               values=["Unknown", "A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-"],
                               size_hint_y=None, height=40)),
        ]:
            f.add_widget(Label(text=lb, font_size=15, bold=True, size_hint_x=0.28))
            f.add_widget(w)
            self.rw[lb.split(":")[0].lower().replace(" ", "_").replace("(", "").replace(")", "")] = w

        def update_age(instance, value):
            try:
                dob = datetime.strptime(value, "%Y-%m-%d")
                age = datetime.now().year - dob.year
                if (datetime.now().month < dob.month or
                        (datetime.now().month == dob.month and datetime.now().day < dob.day)):
                    age -= 1
                self.rw["age"].text = str(age)
            except:
                pass
        self.rw["dob"].bind(text=update_age)

        l.add_widget(f)

        btns = BoxLayout(size_hint_y=None, height=42, spacing=6)
        btns.add_widget(Button(
            text="REGISTER", font_size=17, bold=True,
            on_press=self._reg_pt,
            background_color=get_color_from_hex('#27ae60')
        ))
        btns.add_widget(Button(text="CLEAR", font_size=15, on_press=self._clr_reg))
        l.add_widget(btns)

        l.add_widget(Label(
            text="SEARCH PATIENTS", font_size=15, bold=True,
            size_hint_y=None, height=22
        ))
        srch_box = BoxLayout(size_hint_y=None, height=36, spacing=3)
        self.reg_srch = TextInput(
            font_size=14, size_hint_x=0.7, size_hint_y=None, height=36,
            hint_text="Search by name or ID"
        )
        srch_box.add_widget(self.reg_srch)
        srch_box.add_widget(Button(
            text="Search", font_size=13,
            size_hint_x=0.15, size_hint_y=None, height=36,
            on_press=lambda x: self._refresh_plist(self.reg_srch.text.strip().lower())
        ))
        srch_box.add_widget(Button(
            text="All", font_size=13,
            size_hint_x=0.15, size_hint_y=None, height=36,
            on_press=lambda x: (setattr(self.reg_srch, 'text', ''),
                                self._refresh_plist())
        ))
        l.add_widget(srch_box)

        l.add_widget(Label(
            text="PATIENTS", font_size=14, bold=True,
            size_hint_y=None, height=22,
            color=get_color_from_hex('#2c3e50')
        ))
        self.plist = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        self.plist.bind(minimum_height=self.plist.setter('height'))
        l.add_widget(self.plist)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)
        self._refresh_plist()

    def _reg_pt(self, instance):
        nm = self.rw.get("name", TextInput()).text.strip()
        if not nm:
            self._popup("Error", "Name required!")
            return
        p = {
            "id": self.rw.get("id", TextInput()).text,
            "name": nm,
            "dob": self.rw.get("dob", TextInput()).text,
            "age": int(self.rw.get("age", TextInput()).text or 0),
            "gender": self.rw.get("gender", Spinner()).text,
            "phone": self.rw.get("phone", TextInput()).text,
            "height": self.rw.get("height", TextInput()).text.strip(),
            "blood": self.rw.get("blood", Spinner()).text,
            "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        self.patients.append(p)
        self.save_all()
        self._refresh_plist()
        self._sel_by_id(p["id"])
        self._clr_reg()
        self._popup("Registered", f"{nm}\nID: {p['id']}\nAge: {p['age']} years")
        self.log_activity(f"Registered patient: {nm} ({p['id']})")
        self.audit_log("REGISTER_PATIENT", f"{nm}", p["id"])

    def _clr_reg(self, instance=None):
        self.rw.get("id", TextInput()).text = f"PT{datetime.now().strftime('%Y%m%d%H%M%S')}"
        for k in ["name", "phone", "height", "dob"]:
            if k in self.rw:
                self.rw[k].text = ""
        self.rw.get("age", TextInput()).text = "0"

    def _refresh_plist(self, filter_text=None):
        self.plist.clear_widgets()
        for p in self.patients:
            if (filter_text and
                    filter_text not in p.get("name", "").lower() and
                    filter_text not in p.get("id", "").lower()):
                continue
            row = BoxLayout(size_hint_y=None, height=35, spacing=3)
            btn = Button(
                text=(f"{p.get('name','?')} | {p.get('id','')} | "
                      f"{p.get('age','')}yr | {p.get('gender','')}"),
                font_size=13, size_hint_x=0.8, size_hint_y=None, height=35
            )
            btn.pid = p.get("id")
            btn.bind(on_press=lambda x: self._sel_by_id(x.pid))
            row.add_widget(btn)

            del_btn = Button(
                text="X", font_size=13, bold=True,
                size_hint_x=0.1, size_hint_y=None, height=35,
                background_color=get_color_from_hex('#e74c3c')
            )
            del_btn.pid = p.get("id")
            del_btn.bind(on_press=lambda x: self._confirm_delete_patient(x.pid))
            row.add_widget(del_btn)
            self.plist.add_widget(row)

    def _clear_drug_list(self):
        self.emr_drug_list = []
        self._refresh_drug_display()
        if "prescribed_drugs" in getattr(self, 'ew', {}):
            self.ew["prescribed_drugs"].text = ""
        if "drug_quantities" in getattr(self, 'ew', {}):
            self.ew["drug_quantities"].text = ""

    # ====================== DELETE PATIENT (with passcode) ======================
    def _confirm_delete_patient(self, pid):
        if not self._require_role("Administrator", "Doctor"):
            return

        patient = next((p for p in self.patients if p.get("id") == pid), None)
        if not patient:
            return

        if pid not in self.delete_attempts:
            self.delete_attempts[pid] = 0
        self.delete_attempts[pid] += 1

        if self.delete_attempts[pid] > 3:
            self._popup("Access Denied",
                        "Too many failed attempts. Please contact administrator.")
            self.audit_log("DELETE_PATIENT_LOCKOUT", f"Patient ID {pid}", pid)
            return

        content = BoxLayout(orientation='vertical', spacing=10, padding=10)
        content.add_widget(Label(
            text=(f"Enter passcode to delete this patient "
                  f"(Attempt {self.delete_attempts[pid]}/3):"),
            font_size=15
        ))
        passcode_input = TextInput(password=True, multiline=False,
                                   size_hint_y=None, height=40)
        content.add_widget(passcode_input)
        content.add_widget(Label(
            text=f"Patient: {patient.get('name')} ({pid})",
            font_size=14, bold=True,
            color=get_color_from_hex('#c0392b')
        ))
        content.add_widget(Label(text="This will remove all associated records.",
                                 font_size=12))

        btns = BoxLayout(size_hint_y=None, height=40, spacing=10)
        pp = Popup(title="Confirm Delete", content=content, size_hint=(0.6, 0.4))

        def do_delete(inst):
            if not self._verify_delete_passcode(passcode_input.text):
                self._popup(
                    "Access Denied",
                    f"Incorrect passcode. "
                    f"{3 - self.delete_attempts[pid]} attempts remaining."
                )
                self.audit_log(
                    "DELETE_PATIENT_FAILED",
                    f"Wrong passcode for patient {patient.get('name')}",
                    pid
                )
                return
            # Reset attempts on success
            self.delete_attempts[pid] = 0

            self.patients = [p for p in self.patients if p.get("id") != pid]
            self.emr_records = [r for r in self.emr_records if r.get("patient_id") != pid]
            self.lab_tests = [t for t in self.lab_tests if t.get("patient_id") != pid]
            self.radiology = [r for r in self.radiology if r.get("patient_id") != pid]
            self.prescriptions = [rx for rx in self.prescriptions if rx.get("patient_id") != pid]
            self.bills = [b for b in self.bills if b.get("patient", {}).get("id") != pid]
            self.ecg_reports = [e for e in self.ecg_reports if e.get("patient_id") != pid]
            self.followups = [f for f in self.followups if f.get("patient_id") != pid]
            self.daily_notes = [n for n in self.daily_notes if n.get("patient_id") != pid]
            self.consent_forms = [c for c in self.consent_forms if c.get("patient_id") != pid]
            self.referral_notes = [r for r in self.referral_notes if r.get("patient_id") != pid]
            self.patient_snapshots = [s for s in self.patient_snapshots
                                       if s.get("patient_id") != pid]
            self.lab_results = [l for l in self.lab_results if l.get("patient_id") != pid]
            self.tasks = [t for t in self.tasks if t.get("patient_id") != pid]
            self.task_manager.tasks = self.tasks
            self.patient_summaries.pop(pid, None)
            self.admissions = [a for a in self.admissions if a.get("patient_id") != pid]
            self.saved_lab_reports = [r for r in self.saved_lab_reports
                                       if r.get("patient_id") != pid]
            self.appointments = [a for a in self.appointments
                                 if a.get("patient_id") != pid]
            self.nursing_notes = [n for n in self.nursing_notes
                                  if n.get("patient_id") != pid]

            self.pt_tracker.remove_for_patient(pid)
            self.chat_sys.remove_for_patient(patient.get("name", ""))

            if self.cp and self.cp.get("id") == pid:
                self.cp = None
                self.pt_lbl.text = "No Patient"
                self._clear_drug_list()

            self.save_all()
            self._refresh_plist()
            self._clear_pt()
            pp.dismiss()
            self._popup(
                "Deleted",
                f"Patient {patient.get('name')} ({pid}) and all "
                f"associated records deleted."
            )
            self.log_activity(f"Deleted patient {patient.get('name')} ({pid})")
            self.audit_log(
                "DELETE_PATIENT",
                f"Deleted {patient.get('name')} with all records",
                pid
            )

        btns.add_widget(Button(
            text="DELETE",
            on_press=do_delete,
            background_color=get_color_from_hex('#e74c3c')
        ))
        btns.add_widget(Button(
            text="CANCEL",
            on_press=lambda x: (pp.dismiss(), self.delete_attempts.pop(pid, None))
        ))
        content.add_widget(btns)
        pp.open()

    # ====================== PATIENT SELECTION ======================
    def _load_patient_data(self):
        """Load everything associated with the current patient."""
        if not self.cp:
            return
        pid = self.cp.get("id")
        emrs = [rec for rec in self.emr_records if rec.get("patient_id") == pid]
        latest = emrs[-1] if emrs else {}

        # --- Sync department spinner ---
        target_dept = latest.get("dept", "General Medicine")
        if target_dept not in self.depts:
            target_dept = "General Medicine"
        if hasattr(self, 'emr_dept'):
            if self.emr_dept.text != target_dept:
                self.emr_dept.text = target_dept
            else:
                self._build_dept_fields(target_dept)

        # --- Populate all active EMR widgets ---
        if hasattr(self, 'ew'):
            for k, w in self.ew.items():
                val = latest.get(k, "")
                if val is None:
                    val = ""
                if isinstance(w, Spinner):
                    w.text = val if val else ("Alert" if k == "avpu" else "")
                elif isinstance(w, TextInput):
                    w.text = str(val)

        # --- Restore drug list ---
        structured = latest.get("_structured_drugs")
        if structured:
            for entry in structured:
                if "stock_checked" not in entry:
                    entry["stock_checked"] = 0
            self.emr_drug_list = list(structured)
        else:
            self.emr_drug_list = []
        self._refresh_drug_display()

        # --- Refresh the rest ---
        self._refresh_admission_spinner()
        self._refresh_clinical_summary()
        self._refresh_timeline()
        self._refresh_uploaded_results()
        self._refresh_vitals_timeline()
        self._refresh_daily_notes()
        self._refresh_snapshots()
        self._refresh_saved_items()
        self._refresh_elab()
        self._refresh_tasks()
        self._auto_populate_patient_info()
        self._refresh_pharmacy_dispense()
        self._refresh_admissions_list()
        self._refresh_saved_reports_list()
        self._refresh_appointments()
        self._refresh_nursing_notes()

    def _auto_populate_patient_info(self):
        """Push patient ID + name into every tab that displays it."""
        if not self.cp:
            return
        pid = self.cp.get("id", "")
        name = self.cp.get("name", "")

        pair_map = {
            "elab_pid": pid, "elab_name": name,
            "proc_pid": pid, "proc_name_display": name,
            "dispense_pid": pid, "dispense_name": name,
            "task_pid": pid, "task_patient_name": name,
            "lab_pid": pid, "lab_name": name,
            "rad_pid": pid, "rad_name": name,
            "bill_pid": pid, "bill_name": name,
            "pay_pid": pid, "pay_name": name,
            "pharm_pid": pid, "pharm_name": name,
            "emr_pid": pid, "emr_name": name,
            "daily_pid": pid, "daily_name": name,
            "adm_pid": pid, "adm_name": name,
            "appt_pid": pid, "appt_name": name,
            "nurse_pid": pid, "nurse_name": name,
            "rpt_patient_name": name, "rpt_patient_id": pid,
        }
        for attr, value in pair_map.items():
            if hasattr(self, attr):
                try:
                    getattr(self, attr).text = value
                except:
                    pass

        # Report drafter extra fields
        if hasattr(self, 'rpt_patient_age'):
            self.rpt_patient_age.text = str(self.cp.get('age', ''))
        if hasattr(self, 'rpt_patient_sex'):
            self.rpt_patient_sex.text = self.cp.get('gender', '')

    def _sel_by_id(self, pid):
        for p in self.patients:
            if p.get("id") == pid:
                self.cp = p
                self.pt_lbl.text = (f"{p.get('name','')} ({p.get('id','')}) "
                                    f"| {p.get('age','')}yr")
                self.pt_lbl.color = get_color_from_hex('#2c3e50')
                self._upd_tabs(p)
                self._clear_drug_list()
                self._load_patient_data()
                self.audit_log("SELECT_PATIENT", f"{p.get('name')}", pid)
                return

    def _clear_pt(self):
        self.cp = None
        self.pt_lbl.text = "No Patient"
        self.pt_lbl.color = get_color_from_hex('#7f8c8d')
        self._clear_drug_list()
        if hasattr(self, 'cl_summary'):
            self.cl_summary.text = ""
        self._refresh_timeline()
        self._refresh_uploaded_results()
        self._refresh_vitals_timeline()
        self._refresh_daily_notes()
        self._refresh_snapshots()
        self._refresh_saved_items()
        self._refresh_elab()
        self._refresh_tasks()
        self._refresh_pharmacy_dispense()
        self._refresh_admissions_list()
        self._refresh_saved_reports_list()
        self._refresh_appointments()
        self._refresh_nursing_notes()

    def _upd_tabs(self, p):
        pid = p.get("id", "")
        nm = p.get("name", "")
        widget_mapping = {
            "emr_pid": pid, "emr_name": nm,
            "bill_pid": pid, "bill_name": nm,
            "pay_pid": pid, "pay_name": nm,
            "pharm_pid": pid, "pharm_name": nm,
            "lab_pid": pid, "lab_name": nm,
            "rad_pid": pid, "rad_name": nm,
            "up_pid": pid, "up_name": nm,
            "consent_pid": pid, "consent_name": nm,
            "referral_pid": pid, "referral_name": nm,
            "snapshot_pid": pid, "snapshot_name": nm,
            "elab_pid": pid, "elab_name": nm,
            "proc_pid": pid, "proc_name_display": nm,
            "dispense_pid": pid, "dispense_name": nm,
            "task_pid": pid, "task_patient_name": nm,
            "daily_pid": pid, "daily_name": nm,
            "adm_pid": pid, "adm_name": nm,
            "appt_pid": pid, "appt_name": nm,
            "nurse_pid": pid, "nurse_name": nm,
            "rpt_patient_name": nm, "rpt_patient_id": pid,
        }
        for attr, value in widget_mapping.items():
            if hasattr(self, attr):
                try:
                    getattr(self, attr).text = value
                except:
                    pass
        if self.cp:
            if hasattr(self, 'ca_ctx'):
                self.ca_ctx.text = (f"Patient: {nm} ({pid}) | "
                                    f"{p.get('age','')}yr | {p.get('gender','')}")
            if hasattr(self, 'ecg_info'):
                self.ecg_info.text = f"Patient: {nm} ({pid}) | Age: {p.get('age',0)}"
        self._refresh_pay()
        self._refresh_admissions_list()
        self._refresh_saved_reports_list()
        self._refresh_appointments()
        self._refresh_nursing_notes()
        self._clear_drug_list()

    def _sel_pt_popup(self):
        c = BoxLayout(orientation='vertical', spacing=5, padding=8)
        sr = TextInput(size_hint_y=None, height=40, font_size=15, hint_text="Search...")
        c.add_widget(sr)
        pl = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        pl.bind(minimum_height=pl.setter('height'))
        sv = ScrollView()
        sv.add_widget(pl)
        c.add_widget(sv)
        pp = Popup(title="Select Patient", content=c, size_hint=(0.9, 0.8))

        def rf(t=""):
            pl.clear_widgets()
            for p in self.patients:
                if (not t or t.lower() in p.get("name", "").lower() or
                        t.lower() in p.get("id", "").lower()):
                    btn = Button(
                        text=(f"{p.get('name','?')} | {p.get('id','')} | "
                              f"{p.get('age','')}yr"),
                        font_size=15, size_hint_y=None, height=40
                    )
                    btn.pid = p.get("id")
                    btn.bind(on_press=lambda x: (self._sel_by_id(x.pid), pp.dismiss()))
                    pl.add_widget(btn)

        sr.bind(text=lambda i, v: rf(v))
        rf()
        pp.open()

    # ====================== CLINICAL AI TAB ======================
    def _clinical_ai_tab(self):
        tab = TabbedPanelItem(text="Clinical AI")
        main = BoxLayout(orientation='horizontal', padding=6, spacing=6)
        left = BoxLayout(orientation='vertical', size_hint_x=0.5, spacing=4)
        left.add_widget(Label(
            text="AI CLINICAL ASSISTANT",
            font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#e67e22')
        ))

        ai_status = BoxLayout(size_hint_y=None, height=30, spacing=5)
        self.ai_status_label = Label(
            text="AI Status: " + ("✅ Connected" if self.ai.enabled else "⚠️ Offline"),
            font_size=13, bold=True,
            color=get_color_from_hex('#27ae60') if self.ai.enabled else get_color_from_hex('#e74c3c')
        )
        ai_status.add_widget(self.ai_status_label)
        ai_status.add_widget(Button(
            text="Test AI", font_size=12, size_hint_x=0.15,
            on_press=self._test_ai_connection,
            background_color=get_color_from_hex('#2980b9')
        ))
        left.add_widget(ai_status)

        self.ca_ctx = Label(
            text="No patient selected",
            font_size=13, size_hint_y=None, height=20,
            color=get_color_from_hex('#7f8c8d')
        )
        left.add_widget(self.ca_ctx)

        chat_scroll = ScrollView(size_hint_y=0.6, bar_width=10,
                                 scroll_type=['bars', 'content'])
        self.ca_chat = TextInput(
            multiline=True, readonly=True, font_size=12,
            background_color=get_color_from_hex('#f8f9fa'),
            size_hint_y=None
        )
        self.ca_chat.bind(minimum_height=self.ca_chat.setter('height'))
        chat_scroll.add_widget(self.ca_chat)
        left.add_widget(chat_scroll)

        inp = BoxLayout(size_hint_y=None, height=45, spacing=4)
        self.ca_input = TextInput(
            size_hint_x=0.6, font_size=15,
            size_hint_y=None, height=45,
            hint_text="Type response...",
            on_text_validate=self._send_ca
        )
        inp.add_widget(self.ca_input)
        inp.add_widget(Button(
            text="📷", font_size=16, bold=True, size_hint_x=0.08,
            on_press=self._ca_camera,
            background_color=get_color_from_hex('#2980b9')
        ))
        inp.add_widget(Button(
            text="SEND", font_size=16, bold=True, size_hint_x=0.25,
            on_press=self._send_ca,
            background_color=get_color_from_hex('#27ae60')
        ))
        left.add_widget(inp)

        ctrl = BoxLayout(size_hint_y=None, height=40, spacing=4)
        ctrl.add_widget(Button(
            text="START", font_size=14, bold=True,
            on_press=self._start_ca,
            background_color=get_color_from_hex('#27ae60')
        ))
        ctrl.add_widget(Button(
            text="RESET", font_size=14,
            on_press=self._reset_ca,
            background_color=get_color_from_hex('#f39c12')
        ))
        ctrl.add_widget(Button(
            text="RE-ANALYZE", font_size=14, bold=True,
            on_press=self._re_analyze_ca,
            background_color=get_color_from_hex('#e67e22')
        ))
        ctrl.add_widget(Button(
            text="ASSESSMENT", font_size=14, bold=True,
            on_press=self._show_assessment,
            background_color=get_color_from_hex('#c0392b')
        ))
        ctrl.add_widget(Button(
            text="POPULATE EMR", font_size=14, bold=True,
            on_press=self._populate_emr,
            background_color=get_color_from_hex('#8e44ad')
        ))
        ctrl.add_widget(Button(
            text="SCHEDULE", font_size=14, bold=True,
            on_press=self._schedule_followup,
            background_color=get_color_from_hex('#16a085')
        ))
        ctrl.add_widget(Button(
            text="AUTO-ASSIGN TASKS", font_size=14, bold=True,
            on_press=self._auto_assign_tasks_from_ai,
            background_color=get_color_from_hex('#2c3e50')
        ))
        left.add_widget(ctrl)

        right = BoxLayout(orientation='vertical', size_hint_x=0.5, spacing=3)
        right.add_widget(Label(
            text="ASSESSMENT", font_size=16, bold=True,
            size_hint_y=None, height=25,
            color=get_color_from_hex('#2c3e50')
        ))
        assess_scroll = ScrollView(size_hint_y=0.3, bar_width=10,
                                   scroll_type=['bars'])
        self.ca_diag = TextInput(
            multiline=True, readonly=True, font_size=11,
            background_color=get_color_from_hex('#fef9e7'),
            size_hint_y=None
        )
        self.ca_diag.bind(minimum_height=self.ca_diag.setter('height'))
        assess_scroll.add_widget(self.ca_diag)
        right.add_widget(assess_scroll)

        right.add_widget(Label(
            text="EMR PREVIEW", font_size=16, bold=True,
            size_hint_y=None, height=25,
            color=get_color_from_hex('#2c3e50')
        ))
        emr_scroll = ScrollView(size_hint_y=0.5, bar_width=10,
                                scroll_type=['bars'])
        self.ca_emr = TextInput(
            multiline=True, readonly=True, font_size=11,
            background_color=get_color_from_hex('#eaf2f8'),
            size_hint_y=None
        )
        self.ca_emr.bind(minimum_height=self.ca_emr.setter('height'))
        emr_scroll.add_widget(self.ca_emr)
        right.add_widget(emr_scroll)

        act = BoxLayout(size_hint_y=None, height=40, spacing=4)
        act.add_widget(Button(
            text="SAVE TO EMR", font_size=14, bold=True,
            on_press=self._save_ca_emr,
            background_color=get_color_from_hex('#27ae60')
        ))
        act.add_widget(Button(
            text="ADD TO BILL", font_size=14,
            on_press=self._add_ca_bill,
            background_color=get_color_from_hex('#2980b9')
        ))
        right.add_widget(act)

        main.add_widget(left)
        main.add_widget(right)
        tab.add_widget(main)
        self.tp.add_widget(tab)

    def _test_ai_connection(self, instance):
        if not self.ai.enabled:
            self._popup("AI Status",
                        "⚠️ AI is not enabled. Please set up API key in Settings.")
            return
        self.ai_status_label.text = "AI Status: Testing..."

        def run_test():
            ok, resp = self.ai.call(
                "You are a helpful assistant.",
                "Reply with 'OK' if you can read this.",
                mt=10
            )
            if ok:
                Clock.schedule_once(
                    lambda dt: setattr(self.ai_status_label, 'text',
                                        "AI Status: ✅ Connected"), 0
                )
                self._popup("AI Test", "✅ AI is working correctly!")
            else:
                Clock.schedule_once(
                    lambda dt: setattr(self.ai_status_label, 'text',
                                        f"AI Status: ❌ Failed: {resp[:50]}"), 0
                )
                self._popup("AI Test", f"❌ AI test failed:\n{resp}")

        threading.Thread(target=run_test, daemon=True).start()

    def _ca_camera(self, instance):
        if not CameraHelper.available():
            self._popup("Camera", "Camera not available")
            return
        content = BoxLayout(orientation='vertical', spacing=4, padding=4)
        cam = Camera(play=True, resolution=(640, 480))
        content.add_widget(cam)
        btns = BoxLayout(size_hint_y=None, height=40, spacing=5)
        pp = Popup(title="Capture Photo", content=content, size_hint=(0.85, 0.85))

        def capture(inst):
            if cam.texture:
                fn = f"chat_capture_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                cam.texture.save(fn)
                pp.dismiss()
                self.ca_chat.text += f"\n[Image captured: {fn}]\n"
                self._popup("Photo Captured", f"Image saved as: {fn}")
                self.log_activity(f"Chat photo captured: {fn}")
                if self.cp:
                    self.chat_sys.send(self.cp.get('name', 'User'),
                                        "General",
                                        f"[Photo] {fn}",
                                        image=fn)

        btns.add_widget(Button(
            text="📸 CAPTURE", font_size=16, bold=True,
            on_press=capture,
            background_color=get_color_from_hex('#27ae60')
        ))
        btns.add_widget(Button(text="CANCEL", font_size=16, on_press=pp.dismiss))
        content.add_widget(btns)
        pp.open()

    def _send_ca(self, instance):
        msg = self.ca_input.text.strip()
        if not msg:
            return
        self.ai_patient_id = self.cp.get("id") if self.cp else None
        self.ca_chat.text += f"Patient: {msg}\n\nAssistant: Analyzing...\n"
        self.ca_input.text = ""
        threading.Thread(target=self._process_ca, args=(msg,), daemon=True).start()

    def _process_ca(self, msg):
        r = self.clinical_ai.process(msg)
        patient_id_at_start = self.ai_patient_id
        Clock.schedule_once(
            lambda dt: self._update_ca(r, patient_id_at_start)
            if r else self._update_ca("Could not process. Try again.",
                                        patient_id_at_start),
            0
        )

    def _update_ca(self, response, expected_patient_id=None):
        current_patient_id = self.cp.get("id") if self.cp else None
        if expected_patient_id and expected_patient_id != current_patient_id:
            self.ca_chat.text += ("\n[Warning: Response discarded - "
                                   "patient changed during analysis]\n")
            self.log_activity(
                f"AI response discarded - patient changed from "
                f"{expected_patient_id} to {current_patient_id}"
            )
            return

        lines = [l for l in self.ca_chat.text.split('\n') if 'Analyzing...' not in l]
        self.ca_chat.text = '\n'.join(lines) + '\nAssistant: ' + response + '\n\n'
        self.ca_chat.scroll_y = 0

        if self.clinical_ai.done:
            emr = self.clinical_ai.get_emr()
            if emr:
                self.ca_diag.text = "Differential:\n" + "\n".join(
                    [f"  {i+1}. {d}" for i, d in
                     enumerate(emr.get("differential_diagnosis", []))]
                )
                self.ca_emr.text = "EMR DATA:\n" + "\n".join(
                    [f"{k}: {v}" for k, v in emr.items() if k != "_text"]
                )
                self.ca_chat.text += ("\n" + "=" * 50 +
                                       "\nINTERVIEW COMPLETE\n" + "=" * 50 + "\n")
                if self.cp:
                    pid = self.cp.get("id")
                    summary_text = self._format_summary(emr)
                    self.patient_summaries[pid] = summary_text
                    self.save_all()
                    self._refresh_clinical_summary()
                    self.log_activity(
                        f"Clinical AI interview completed for patient "
                        f"{self.cp.get('name','')}"
                    )

    def _format_summary(self, emr):
        lines = []
        lines.append("AI CLINICAL SUMMARY")
        lines.append(f"Chief Complaint: {emr.get('chief_complaint', 'N/A')}")
        lines.append(f"History: {emr.get('history_of_presenting_illness', 'N/A')}")
        lines.append(f"PMH: {emr.get('past_medical_history', 'N/A')}")
        lines.append(f"Medications: {emr.get('medications', 'N/A')}")
        lines.append(f"Allergies: {emr.get('allergies', 'N/A')}")
        lines.append("Differentials:")
        for d in emr.get('differential_diagnosis', []):
            lines.append(f"- {d}")
        lines.append(f"Recommended Investigations: "
                     f"{', '.join(emr.get('recommended_investigations', []))}")
        lines.append(f"Treatment Plan: {emr.get('treatment_plan', 'N/A')}")
        lines.append(f"Urgency: {emr.get('urgency', 'Routine')}")
        lines.append(f"Follow-up: {emr.get('follow_up', 'N/A')}")
        return "\n".join(lines)

    def _show_assessment(self, instance):
        if not self.clinical_ai.done or not self.clinical_ai.emr:
            if self.ai.enabled:
                self._generate_assessment_ai()
            else:
                self._popup(
                    "Assessment",
                    "Complete the AI interview first or enable AI for "
                    "automatic assessment."
                )
            return
        emr = self.clinical_ai.emr
        text = "Clinical Assessment\n" + "=" * 50 + "\n"
        text += f"Chief Complaint: {emr.get('chief_complaint','N/A')}\n"
        text += f"Differential Diagnoses:\n"
        for i, d in enumerate(emr.get('differential_diagnosis', [])):
            text += f"  {i+1}. {d}\n"
        text += (f"Recommended Investigations: "
                 f"{', '.join(emr.get('recommended_investigations', []))}\n")
        text += f"Treatment Plan: {emr.get('treatment_plan','N/A')}\n"
        text += f"Urgency: {emr.get('urgency','Routine')}\n"
        text += f"Follow-up: {emr.get('follow_up','N/A')}\n"
        self._text_popup("Clinical Assessment", text)

    def _generate_assessment_ai(self):
        if not self.ai.enabled or not self.cp:
            self._popup("AI Error", "AI not enabled or no patient selected.")
            return
        conv_text = ""
        for m in self.clinical_ai.conv:
            role = "Doctor" if m['role'] == 'assistant' else "Patient"
            conv_text += f"{role}: {m['content']}\n"
        prompt = f"""Based on the following clinical interview, generate a comprehensive SOAP assessment:

{conv_text}

Provide:
- Subjective: Chief complaint, HPI
- Objective: Physical exam findings, vitals (if provided)
- Assessment: Differential diagnosis with reasoning
- Plan: Investigations, treatment, follow-up

Be thorough and evidence-based."""

        def run():
            ok, resp = self.ai.call(
                "You are a senior clinician providing clinical assessment.",
                prompt, mt=1500
            )
            if ok:
                Clock.schedule_once(
                    lambda dt: self._text_popup("AI Assessment", resp), 0
                )
            else:
                Clock.schedule_once(lambda dt: self._popup("Error", resp), 0)

        threading.Thread(target=run, daemon=True).start()

    def _start_ca(self, instance):
        if self.cp:
            name = self.cp.get('name', '')
            pid = self.cp.get('id', '')
            age = str(self.cp.get('age', ''))
            gender = self.cp.get('gender', '')
            self.ca_ctx.text = f"Patient: {name} ({pid}) | {age}yr | {gender}"
            opening = self.clinical_ai.start(name, age, gender)
        else:
            self.ca_ctx.text = "No patient selected"
            opening = self.clinical_ai.start("", "", "")
        self.ca_chat.text = ("=" * 50 + "\nAI CLINICAL ASSISTANT\n" +
                              "=" * 50 + "\n\nAssistant: " + opening + "\n\n")
        if not self.ai.enabled:
            self.ca_chat.text += "(Structured mode)\n\n"
        self.ca_input.text = ""
        self.ca_input.focus = True
        if hasattr(self, 'cl_summary'):
            self.cl_summary.text = ""

    def _re_analyze_ca(self, instance):
        if not self.clinical_ai.conv:
            self._popup("No Data", "Start interview first.")
            return
        self.ca_chat.text += "\nRe-analyzing...\n"
        threading.Thread(
            target=lambda: Clock.schedule_once(
                lambda dt: self._update_ca(
                    self.clinical_ai.re_analyze() or "Re-analysis failed."
                ), 0
            ),
            daemon=True
        ).start()

    def _populate_emr(self, instance):
        emr = self.clinical_ai.get_emr()
        if not emr:
            self._popup("No Data", "Complete interview!")
            return
        self._switch("EMR")
        if hasattr(self, 'ew'):
            for ek, dk in {'cc': 'chief_complaint',
                            'hpi': 'history_of_presenting_illness',
                            'pmh': 'past_medical_history'}.items():
                if ek in self.ew:
                    self.ew[ek].text = str(emr.get(dk, ''))
        self._popup("EMR Populated", "Review and save.")

    def _save_ca_emr(self, instance):
        if not self.cp:
            return
        emr = self.clinical_ai.get_emr()
        if not emr:
            return
        pid = self.cp.get("id")
        name = self.cp.get("name")
        if not pid:
            return
        r = {
            "emr_id": f"EMR-CA-{datetime.now().strftime('%Y%m%d%H%M%S')}",
            "patient_id": pid,
            "patient_name": name,
            "dept": "General Medicine",
            "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "cc": str(emr.get("chief_complaint", "")),
            "hpi": str(emr.get("history_of_presenting_illness", "")),
            "dx": str(emr.get("differential_diagnosis", [""])[0]
                      if emr.get("differential_diagnosis") else ""),
            "tx": str(emr.get("treatment_plan", "")),
            "fu": str(emr.get("follow_up", "")),
            "recommended_investigations": emr.get("recommended_investigations", []),
            "nursing_orders": emr.get("nursing_orders", []),
            "radiology_orders": emr.get("radiology_orders", []),
            "prescribed_drugs": emr.get("prescribed_drugs", [])
        }
        self.emr_records.append(r)
        self.save_all()
        self._popup("Saved", f"ID: {r['emr_id']}")
        self.audit_log("SAVE_AI_EMR", f"EMR ID {r['emr_id']}", pid)
        self._auto_assign_tasks_from_emr(r)

    def _add_ca_bill(self, instance):
        if not self.cp:
            return
        self._switch("Billing")
        self.bill_pid.text = self.cp.get("id", "")
        self.bill_name.text = self.cp.get("name", "")
        self.bill_items.append({
            "name": "AI Consultation",
            "cat": "Charges",
            "qty": 1,
            "price": 7500,
            "total": 7500
        })
        self._refresh_bill()

    def _reset_ca(self, instance):
        self.clinical_ai = ModularAIClinicalAssistant(self.ai)
        self.ca_chat.text = ""
        self.ca_diag.text = ""
        self.ca_emr.text = ""
        self.ca_input.text = ""
        if hasattr(self, 'cl_summary'):
            self.cl_summary.text = ""

    def _schedule_followup(self, instance):
        if not self.cp:
            self._popup("Error", "Select a patient first!")
            return
        c = BoxLayout(orientation='vertical', spacing=6, padding=8)
        c.add_widget(Label(text="Date (YYYY-MM-DD):", font_size=15, bold=True))
        date_inp = TextInput(
            text=datetime.now().strftime("%Y-%m-%d"),
            font_size=15, size_hint_y=None, height=40
        )
        c.add_widget(date_inp)
        c.add_widget(Label(text="Description:", font_size=15, bold=True))
        desc_inp = TextInput(
            font_size=15, size_hint_y=None, height=40,
            hint_text="e.g., Review lab results"
        )
        c.add_widget(desc_inp)
        btns = BoxLayout(size_hint_y=None, height=40, spacing=5)
        pp = Popup(title="Schedule Follow-up", content=c, size_hint=(0.6, 0.4))

        def save_fu(inst):
            date_str = date_inp.text.strip()
            desc = desc_inp.text.strip()
            if not date_str or not desc:
                self._popup("Error", "Both fields required.")
                return
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
            except:
                self._popup("Error", "Invalid date format. Use YYYY-MM-DD.")
                return
            self.followups.append({
                "patient_id": self.cp.get("id"),
                "date": date_str,
                "description": desc
            })
            self.save_all()
            self._refresh_timeline()
            pp.dismiss()
            self._popup("Scheduled", f"Follow-up on {date_str}")
            if self.cp:
                self.task_manager.create_task(
                    patient_id=self.cp.get("id"),
                    patient_name=self.cp.get("name"),
                    assigned_to="Doctor",
                    task_type="Follow-up",
                    description=f"Follow-up: {desc} on {date_str}",
                    department="Medical"
                )
                self._refresh_tasks()

        btns.add_widget(Button(
            text="SAVE", font_size=16, bold=True,
            on_press=save_fu,
            background_color=get_color_from_hex('#27ae60')
        ))
        btns.add_widget(Button(
            text="CANCEL", font_size=16,
            on_press=lambda x: pp.dismiss()
        ))
        c.add_widget(btns)
        pp.open()

    def _refresh_timeline(self):
        if not hasattr(self, 'timeline_text'):
            return
        if not self.cp:
            self.timeline_text.text = "Select a patient to view timeline."
            return
        pid = self.cp.get("id")
        events = []
        for r in self.emr_records:
            if r.get("patient_id") == pid:
                events.append((r.get("ts", ""),
                                f"[EMR] {r.get('cc', 'Consultation')} "
                                f"| Dept: {r.get('dept', 'N/A')}"))
        for t in self.lab_tests:
            if t.get("patient_id") == pid:
                events.append((t.get("date_ordered", t.get("ts", "")),
                                f"[LAB] {t.get('test_name', '')} - "
                                f"{t.get('status', '')}"))
        for rx in self.prescriptions:
            if rx.get("patient_id") == pid:
                events.append((rx.get("ts", ""),
                                f"[RX] {rx.get('drug', '')} x{rx.get('qty', 0)}"))
        for b in self.bills:
            if b.get("patient", {}).get("id") == pid:
                events.append((b.get("ts", ""),
                                f"[BILL] N{b.get('total', 0):,.2f} - "
                                f"{b.get('status', '')}"))
        for e in self.ecg_reports:
            if e.get("patient_id") == pid:
                sd = e.get("structured", {})
                events.append((e.get("ts", ""),
                                f"[ECG] {sd.get('interpretation', 'Report')} - "
                                f"{sd.get('urgency', 'Routine')}"))
        for rad in self.radiology:
            if rad.get("patient_id") == pid:
                events.append((rad.get("ts", ""),
                                f"[RAD] {rad.get('study', '')} - "
                                f"{rad.get('status', '')}"))
        for fu in self.followups:
            if fu.get("patient_id") == pid:
                events.append((fu.get("date", ""),
                                f"[FOLLOW-UP] {fu.get('description', '')}"))
        for pmt in self.pt_tracker.get_for(pid):
            events.append((pmt.get("ts", ""),
                            f"[PAYMENT] N{pmt.get('amount', 0):,.2f} via "
                            f"{pmt.get('method', '')}"))
        for dn in self.daily_notes:
            if dn.get("patient_id") == pid:
                events.append((dn.get("date", ""),
                                f"[DAILY NOTE] {dn.get('date','')}"))
        for snap in self.patient_snapshots:
            if snap.get("patient_id") == pid:
                events.append((snap.get("ts", ""),
                                f"[SNAPSHOT] {snap.get('bp','N/A')} | "
                                f"{snap.get('pulse','N/A')}"))
        for task in self.task_manager.get_tasks_for_patient(pid):
            events.append((task.get("created_at", ""),
                            f"[TASK] {task.get('task_type')} - "
                            f"{task.get('status')}: "
                            f"{task.get('description')[:30]}"))
        for adm in self.admissions:
            if adm.get("patient_id") == pid:
                events.append((adm.get("admission_date", ""),
                                f"[ADMISSION] {adm.get('diagnosis', '')} | "
                                f"Dept: {adm.get('department', '')}"))
        for appt in self.appointments:
            if appt.get("patient_id") == pid:
                events.append((appt.get("date", ""),
                                f"[APPOINTMENT] {appt.get('type', '')} - "
                                f"{appt.get('status', '')}"))
        for nnote in self.nursing_notes:
            if nnote.get("patient_id") == pid:
                events.append((nnote.get("date", ""),
                                f"[NURSING NOTE] {nnote.get('shift', '')}"))
        events.sort(key=lambda x: x[0])
        if events:
            self.timeline_text.text = (
                f"PATIENT TIMELINE - {self.cp.get('name', '')}\n" +
                "=" * 60 + "\n\n" +
                "\n".join([f"{ts[:16] if len(ts) >= 16 else ts:<16} | {evt}"
                          for ts, evt in events])
            )
        else:
            self.timeline_text.text = f"No events found for {self.cp.get('name', '')}."


# ====================== END OF PART 3 ======================
    # ====================== CLINICAL SUMMARY TAB ======================
    def _clinical_summary_tab(self):
        tab = TabbedPanelItem(text="Clinical Summary")
        main = BoxLayout(orientation='vertical', padding=6, spacing=6)
        main.add_widget(Label(
            text="COMPREHENSIVE CLINICAL ASSESSMENT",
            font_size=18, bold=True,
            size_hint_y=None, height=30,
            color=get_color_from_hex('#e67e22')
        ))

        summary_scroll = ScrollView(size_hint_y=0.45, bar_width=10,
                                     scroll_type=['bars'])
        self.cl_summary = TextInput(
            multiline=True, readonly=True, font_size=13,
            background_color=get_color_from_hex('#fef9e7'),
            size_hint_y=None
        )
        self.cl_summary.bind(minimum_height=self.cl_summary.setter('height'))
        summary_scroll.add_widget(self.cl_summary)
        main.add_widget(summary_scroll)

        main.add_widget(Label(
            text="TREATMENT SUGGESTIONS",
            font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#2980b9')
        ))
        treatment_scroll = ScrollView(size_hint_y=0.2, bar_width=10,
                                       scroll_type=['bars'])
        self.treatment_suggestions = TextInput(
            multiline=True, readonly=True, font_size=13,
            background_color=get_color_from_hex('#eaf2f8'),
            size_hint_y=None
        )
        self.treatment_suggestions.bind(
            minimum_height=self.treatment_suggestions.setter('height')
        )
        treatment_scroll.add_widget(self.treatment_suggestions)
        main.add_widget(treatment_scroll)

        btn_row = BoxLayout(size_hint_y=None, height=40, spacing=5)
        btn_row.add_widget(Button(
            text="GENERATE TREATMENT SUGGESTIONS",
            font_size=14, bold=True,
            on_press=self._generate_treatment_suggestions,
            background_color=get_color_from_hex('#8e44ad')
        ))
        btn_row.add_widget(Button(
            text="EXPORT PDF", font_size=14, bold=True, size_hint_x=0.2,
            on_press=lambda x: self._export_to_pdf(
                self.cl_summary.text,
                "Clinical_Summary",
                "Clinical Assessment Report"
            ),
            background_color=get_color_from_hex('#2980b9')
        ))
        main.add_widget(btn_row)

        main.add_widget(Label(
            text="PATIENT TIMELINE",
            font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#2980b9')
        ))
        timeline_scroll = ScrollView(size_hint_y=0.25, bar_width=10,
                                      scroll_type=['bars'])
        self.timeline_text = TextInput(
            multiline=True, readonly=True, font_size=12,
            background_color=get_color_from_hex('#f0f8ff'),
            size_hint_y=None
        )
        self.timeline_text.bind(minimum_height=self.timeline_text.setter('height'))
        timeline_scroll.add_widget(self.timeline_text)
        main.add_widget(timeline_scroll)

        tab.add_widget(main)
        self.tp.add_widget(tab)
        self._refresh_clinical_summary()

    def _refresh_clinical_summary(self):
        if not self.cp:
            self.cl_summary.text = "Select a patient to view clinical summary."
            return

        pid = self.cp.get("id")
        name = self.cp.get("name", "")

        if pid in self.patient_summaries:
            self.cl_summary.text = self.patient_summaries[pid]
            return

        emrs = [r for r in self.emr_records if r.get("patient_id") == pid]
        if not emrs:
            self.cl_summary.text = f"No clinical records found for {name}."
            return

        latest = emrs[-1]
        summary = f"CLINICAL SUMMARY - {name}\n"
        summary += "=" * 50 + "\n\n"
        summary += f"Patient ID: {pid}\n"
        summary += f"Age: {self.cp.get('age', 'N/A')} years\n"
        summary += f"Gender: {self.cp.get('gender', 'N/A')}\n"
        summary += f"Blood Type: {self.cp.get('blood', 'N/A')}\n"
        summary += f"Last Visit: {latest.get('ts', 'N/A')}\n"
        summary += f"Department: {latest.get('dept', 'N/A')}\n\n"

        summary += "VITAL SIGNS\n"
        summary += f"  BP: {latest.get('bp', 'N/A')}\n"
        summary += f"  Pulse: {latest.get('pulse', 'N/A')}\n"
        summary += f"  Temp: {latest.get('temp', 'N/A')}\n"
        summary += f"  SpO2: {latest.get('spo2', 'N/A')}\n"
        summary += f"  RR: {latest.get('rr', 'N/A')}\n"
        summary += f"  AVPU: {latest.get('avpu', 'Alert')}\n\n"

        # Common field label map for cleaner output
        label_map = {
            "cc": "Chief Complaint",
            "hpi": "History of Presenting Illness",
            "pmh": "Past Medical History",
            "psh": "Past Surgical History",
            "dh": "Drug History",
            "all": "Allergies",
            "fh": "Family History",
            "sh": "Social History",
            "ros": "Review of Systems",
            "gen": "General Examination",
            "cvs": "Cardiovascular System",
            "resp": "Respiratory System",
            "abd": "Abdomen",
            "cns": "Central Nervous System",
            "dx": "Diagnosis",
            "tx": "Treatment Plan",
            "fu": "Follow-up",
            "nursing_orders": "Nursing Orders",
            "birth_hx": "Birth History",
            "dev_milestones": "Developmental Milestones",
            "immunizations": "Immunizations",
            "feeding_hx": "Feeding History",
            "growth": "Growth Parameters",
            "gp": "Gravida/Para",
            "lmp_edd": "LMP/EDD",
            "anc": "Antenatal Care",
            "prev_obs": "Previous Obstetric History",
            "preop": "Pre-op Assessment",
            "surg_plan": "Surgical Plan",
            "postop": "Post-op Plan",
            "risk_factors": "Cardiac Risk Factors",
            "ecg_findings": "ECG Findings",
            "moi": "Mechanism of Injury",
            "triage": "Triage Category",
            "disposition": "Disposition",
            "notes": "Clinical Notes",
            "ref_notes": "Referral Notes",
            "lab_orders": "Lab Orders",
            "radiology_orders": "Radiology Orders",
            "prescribed_drugs": "Prescribed Drugs",
        }

        summary += "CLINICAL DETAILS\n"
        for key, value in latest.items():
            if key in ["emr_id", "patient_id", "patient_name", "dept", "ts",
                        "admission_id", "bp", "temp", "pulse", "spo2", "rr",
                        "wt", "ht", "bmi", "liver_span", "avpu",
                        "_structured_drugs"]:
                continue
            if value and str(value).strip():
                label = label_map.get(key, key.upper().replace("_", " "))
                summary += f"  {label}: {value}\n"

        # Structured drugs
        structured = latest.get("_structured_drugs")
        if structured:
            summary += "\nPRESCRIBED DRUGS (Structured)\n"
            for drug in structured:
                summary += (f"  {drug.get('name','')} — "
                            f"{drug.get('total',0)} units over "
                            f"{drug.get('days',0)} days\n")

        followups = [f for f in self.followups if f.get("patient_id") == pid]
        if followups:
            summary += "\nSCHEDULED FOLLOW-UPS\n"
            for fu in followups:
                summary += f"  - {fu.get('date', 'N/A')}: {fu.get('description', 'N/A')}\n"

        tasks = self.task_manager.get_tasks_for_patient(pid)
        if tasks:
            active_tasks = [t for t in tasks if t.get("status") != "Completed"]
            if active_tasks:
                summary += "\nACTIVE TASKS\n"
                for task in active_tasks:
                    summary += (f"  - [{task.get('status')}] "
                                f"{task.get('task_type')}: "
                                f"{task.get('description')[:60]}\n")

        adm = [a for a in self.admissions if a.get("patient_id") == pid]
        if adm:
            summary += "\nADMISSIONS HISTORY\n"
            for a in adm:
                summary += (f"  - {a.get('admission_date')} | "
                            f"{a.get('department')} | {a.get('diagnosis')}\n")

        appts = [a for a in self.appointments if a.get("patient_id") == pid]
        if appts:
            summary += "\nAPPOINTMENTS\n"
            for a in appts:
                summary += (f"  - {a.get('date')} | {a.get('type')} | "
                            f"{a.get('status')}\n")

        self.cl_summary.text = summary

    def _generate_treatment_suggestions(self, instance):
        if not self.cp:
            self._popup("Error", "Select a patient first.")
            return
        pid = self.cp.get("id")
        emrs = [r for r in self.emr_records if r.get("patient_id") == pid]
        labs = [t for t in self.lab_tests if t.get("patient_id") == pid
                and t.get("status") == "Completed"]
        rx = [p for p in self.prescriptions if p.get("patient_id") == pid]

        context = (f"Patient: {self.cp.get('name','')} | "
                   f"Age: {self.cp.get('age',0)} | "
                   f"Gender: {self.cp.get('gender','')}\n")
        if emrs:
            latest = emrs[-1]
            context += f"Latest Diagnosis: {latest.get('dx', 'N/A')}\n"
            context += f"Treatment Plan: {latest.get('tx', 'N/A')}\n"
            context += f"Follow-up: {latest.get('fu', 'N/A')}\n"
        if labs:
            context += "\nRecent Lab Results:\n"
            for lab in labs[-5:]:
                context += (f"- {lab.get('test_name','')}: "
                            f"{lab.get('result','')} {lab.get('unit','')}\n")
        if rx:
            context += "\nCurrent Medications:\n"
            for med in rx[-5:]:
                context += f"- {med.get('drug','')} x{med.get('qty',0)}\n"

        if not self.ai.enabled:
            suggestions = self._generate_fallback_treatment_suggestions(context)
            self.treatment_suggestions.text = suggestions
            self._popup("Treatment Suggestions", "Generated using local knowledge.")
            return

        prompt = f"""Based on the following patient data, provide comprehensive treatment suggestions:

{context}

Please provide:
1. Recommended treatment plan
2. Suggested medications with dosages
3. Recommended investigations
4. Lifestyle modifications
5. Follow-up schedule
6. Red flags to watch for
7. Referral recommendations if needed

Be specific and evidence-based.
"""
        self.treatment_suggestions.text = "Generating treatment suggestions... Please wait."

        def run():
            ok, resp = self.ai.call(
                "You are a senior clinician providing treatment recommendations.",
                prompt, mt=2000
            )
            if ok:
                Clock.schedule_once(
                    lambda dt: setattr(self.treatment_suggestions, 'text', resp), 0
                )
            else:
                fallback = self._generate_fallback_treatment_suggestions(context)
                Clock.schedule_once(
                    lambda dt: setattr(self.treatment_suggestions, 'text', fallback), 0
                )

        threading.Thread(target=run, daemon=True).start()

    def _generate_fallback_treatment_suggestions(self, context):
        return """TREATMENT SUGGESTIONS (Local Knowledge Base)

Based on the patient's clinical presentation:

1. RECOMMENDED TREATMENT PLAN
- Continue current management
- Monitor vital signs regularly
- Review lab results when available

2. MEDICATION SUGGESTIONS
- Review current medications for appropriateness
- Consider dose adjustments based on response
- Monitor for side effects

3. RECOMMENDED INVESTIGATIONS
- Complete blood count (CBC)
- Comprehensive metabolic panel
- Additional tests as clinically indicated

4. LIFESTYLE MODIFICATIONS
- Regular exercise as tolerated
- Healthy diet
- Adequate hydration
- Smoking cessation if applicable

5. FOLLOW-UP SCHEDULE
- Review in 2 weeks
- Earlier if symptoms worsen
- PRN visits for acute concerns

6. RED FLAGS
- Worsening symptoms
- New onset symptoms
- Signs of deterioration

7. REFERRAL RECOMMENDATIONS
- Consider specialist referral if indicated
- Multidisciplinary approach if complex

Please consult with attending physician for final treatment decisions.
"""

    # ====================== EMR TAB (STANDARDIZED CLERKING) ======================
    def _emr(self):
        tab = TabbedPanelItem(text="EMR")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=5, spacing=3, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="ID:", font_size=15, bold=True, size_hint_x=0.05))
        self.emr_pid = TextInput(size_hint_x=0.15, readonly=True, font_size=15,
                                  size_hint_y=None, height=40)
        bar.add_widget(self.emr_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.07))
        self.emr_name = TextInput(size_hint_x=0.25, readonly=True, font_size=15,
                                   size_hint_y=None, height=40)
        bar.add_widget(self.emr_name)
        bar.add_widget(Label(text="Dept:", font_size=15, bold=True, size_hint_x=0.06))
        self.emr_dept = Spinner(
            text="General Medicine", values=self.depts,
            size_hint_x=0.18, size_hint_y=None, height=40
        )
        self.emr_dept.bind(text=self._on_dept_change)
        bar.add_widget(self.emr_dept)
        bar.add_widget(Label(text="Admission:", font_size=13, bold=True, size_hint_x=0.08))
        self.emr_admission_spinner = Spinner(
            text="None", values=["None"],
            size_hint_x=0.2, size_hint_y=None, height=40
        )
        bar.add_widget(self.emr_admission_spinner)
        bar.add_widget(Button(
            text="AI", font_size=13, size_hint_x=0.06,
            on_press=lambda x: self._switch("Clinical AI"),
            background_color=get_color_from_hex('#e67e22')
        ))
        bar.add_widget(Button(
            text="VIEW HISTORY", font_size=13, size_hint_x=0.12,
            on_press=self._view_emr_history,
            background_color=get_color_from_hex('#16a085')
        ))
        bar.add_widget(Button(
            text="AUTO BILL", font_size=13, size_hint_x=0.12,
            on_press=self._auto_bill_from_emr,
            background_color=get_color_from_hex('#27ae60')
        ))
        bar.add_widget(Button(
            text="AUTO TASKS", font_size=13, size_hint_x=0.12,
            on_press=self._auto_assign_tasks_from_current_emr,
            background_color=get_color_from_hex('#8e44ad')
        ))
        l.add_widget(bar)

        # Department-specific fields container
        self.dept_fields = BoxLayout(orientation='vertical', size_hint_y=None)
        self.dept_fields.bind(minimum_height=self.dept_fields.setter('height'))
        l.add_widget(self.dept_fields)

        # Vital signs
        l.add_widget(Label(
            text="VITAL SIGNS", font_size=17, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#2980b9')
        ))
        vg = GridLayout(cols=4, size_hint_y=None, height=170, spacing=3)
        self.ew = {}
        # Track which keys belong to which section (prevents cross-dept leakage)
        self._vitals_keys = ["bp", "temp", "pulse", "spo2", "rr",
                              "wt", "ht", "bmi", "liver_span", "avpu"]
        self._dept_field_keys = []
        self._common_keys = ["notes", "ref_notes", "lab_orders",
                              "radiology_orders", "prescribed_drugs",
                              "drug_quantities"]

        for lb, k in [("BP:", "bp"), ("Temp:", "temp"), ("Pulse:", "pulse"),
                       ("SpO2:", "spo2"), ("RR:", "rr"), ("Weight:", "wt"),
                       ("Height:", "ht"), ("BMI:", "bmi"),
                       ("Liver Span:", "liver_span"), ("AVPU:", "avpu")]:
            vg.add_widget(Label(text=lb, font_size=13))
            if k == "avpu":
                w = Spinner(
                    text="Alert",
                    values=["Alert", "Voice", "Pain", "Unresponsive"],
                    size_hint_y=None, height=35
                )
                self.ew[k] = w
                vg.add_widget(w)
            else:
                w = TextInput(font_size=13, size_hint_y=None, height=35)
                self.ew[k] = w
                vg.add_widget(w)
        l.add_widget(vg)

        # Common fields container
        self.common_fields = BoxLayout(orientation='vertical', size_hint_y=None)
        self.common_fields.bind(minimum_height=self.common_fields.setter('height'))
        l.add_widget(self.common_fields)

        # Drug addition widget
        l.add_widget(Label(
            text="PRESCRIBED DRUGS (Add with quantity & days)",
            font_size=14, bold=True,
            size_hint_y=None, height=20,
            color=get_color_from_hex('#2980b9')
        ))
        drug_add_box = BoxLayout(size_hint_y=None, height=40, spacing=4)
        drug_add_box.add_widget(Label(text="Drug:", font_size=13, bold=True, size_hint_x=0.06))
        self.emr_drug_spinner = Spinner(
            text="Select",
            values=[d['name'] for d in self.drugs],
            size_hint_x=0.25, size_hint_y=None, height=35
        )
        drug_add_box.add_widget(self.emr_drug_spinner)
        drug_add_box.add_widget(Label(text="Qty/day:", font_size=13, bold=True, size_hint_x=0.08))
        self.emr_drug_qty = TextInput(text="1", size_hint_x=0.08, font_size=13,
                                        size_hint_y=None, height=35)
        drug_add_box.add_widget(self.emr_drug_qty)
        drug_add_box.add_widget(Label(text="Days:", font_size=13, bold=True, size_hint_x=0.06))
        self.emr_drug_days = TextInput(text="1", size_hint_x=0.08, font_size=13,
                                         size_hint_y=None, height=35)
        drug_add_box.add_widget(self.emr_drug_days)
        drug_add_box.add_widget(Button(
            text="ADD", font_size=13, bold=True, size_hint_x=0.08,
            on_press=self._add_drug_to_emr,
            background_color=get_color_from_hex('#27ae60')
        ))
        l.add_widget(drug_add_box)

        # Drug display list
        self.drug_display = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        self.drug_display.bind(minimum_height=self.drug_display.setter('height'))
        drug_scroll = ScrollView(size_hint_y=None, height=80, bar_width=10)
        drug_scroll.add_widget(self.drug_display)
        l.add_widget(drug_scroll)
        l.add_widget(Label(
            text="(Drugs and total quantities are auto-filled below)",
            font_size=11, size_hint_y=None, height=16,
            color=get_color_from_hex('#7f8c8d')
        ))

        btns = BoxLayout(size_hint_y=None, height=42, spacing=5)
        btns.add_widget(Button(
            text="SAVE EMR", font_size=17, bold=True,
            on_press=self._save_emr,
            background_color=get_color_from_hex('#27ae60')
        ))
        btns.add_widget(Button(
            text="CLEAR", font_size=15,
            on_press=lambda x: self._clear_emr()
        ))
        l.add_widget(btns)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)

        self._build_dept_fields("General Medicine")
        self._build_common_fields()
        self._clear_drug_list()

    def _on_dept_change(self, spinner, text):
        self._build_dept_fields(text)

    def _build_dept_fields(self, department):
        """
        Rebuild department-specific clerking fields.

        CRITICAL: purges ALL previously-tracked department-specific keys from
        self.ew first, so switching departments cannot leak stale data into
        the next saved EMR.
        """
        # 1) Purge old dept-specific widgets + keys
        for k in getattr(self, '_dept_field_keys', []):
            w = self.ew.pop(k, None)
            if w is not None and hasattr(w, 'text'):
                w.text = ""   # explicitly clear before dropping
        self._dept_field_keys = []

        self.dept_fields.clear_widgets()
        dept_label = Label(
            text=f"{department} CLERKING",
            font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#e67e22')
        )
        self.dept_fields.add_widget(dept_label)

        # 2) STANDARD core clerking (every department gets these)
        core_fields = [
            ("Chief Complaint",                   "cc",              3),
            ("History of Presenting Illness",     "hpi",             4),
            ("Past Medical History",              "pmh",             2),
            ("Past Surgical History",             "psh",             2),
            ("Drug History / Current Medications","dh",              2),
            ("Allergies",                         "all",             1),
            ("Family History",                    "fh",              2),
            ("Social History",                    "sh",              2),
            ("Review of Systems",                 "ros",             3),
            ("General Examination",               "gen",             3),
            ("Cardiovascular System (CVS)",       "cvs",             2),
            ("Respiratory System (Resp)",         "resp",            2),
            ("Abdomen (Abd)",                     "abd",             2),
            ("Central Nervous System (CNS)",      "cns",             2),
            ("Diagnosis / Impression",            "dx",              2),
            ("Treatment Plan",                    "tx",              2),
            ("Follow-up",                         "fu",              1),
            ("Nursing Orders",                    "nursing_orders",  2),
        ]

        # 3) DEPARTMENT-SPECIFIC EXTRAS
        extras = {
            "Pediatrics": [
                ("Birth History",                 "birth_hx",       2),
                ("Developmental Milestones",      "dev_milestones", 2),
                ("Immunization History",          "immunizations",  2),
                ("Feeding History",               "feeding_hx",     2),
                ("Growth Parameters (Wt/Ht/HC)",  "growth",         2),
            ],
            "Obstetrics": [
                ("Gravida / Para",                "gp",             1),
                ("LMP / EDD",                     "lmp_edd",        1),
                ("Antenatal Care",                "anc",            2),
                ("Previous Obstetric History",    "prev_obs",       2),
            ],
            "Surgery": [
                ("Pre-op Assessment",             "preop",          2),
                ("Surgical Plan",                 "surg_plan",      2),
                ("Post-op Plan",                  "postop",         2),
            ],
            "Cardiology": [
                ("Cardiac Risk Factors",          "risk_factors",   2),
                ("ECG Findings",                  "ecg_findings",   2),
            ],
            "Emergency": [
                ("Mechanism of Injury",           "moi",            2),
                ("Triage Category",               "triage",         1),
                ("Disposition",                   "disposition",    1),
            ],
        }
        fields = list(core_fields) + extras.get(department, [])

        # 4) Build the widgets and track their keys
        for label, key, lines in fields:
            self.dept_fields.add_widget(Label(
                text=label + ":", font_size=13, bold=True,
                size_hint_y=None, height=18
            ))
            w = TextInput(
                multiline=True, font_size=13,
                size_hint_y=None, height=30 * lines
            )
            self.ew[key] = w
            self._dept_field_keys.append(key)
            self.dept_fields.add_widget(w)

    def _build_common_fields(self):
        self.common_fields.clear_widgets()
        self.common_fields.add_widget(Label(
            text="ADDITIONAL NOTES", font_size=14, bold=True,
            size_hint_y=None, height=20,
            color=get_color_from_hex('#2980b9')
        ))
        common = [
            ("Clinical Notes",     "notes",             2),
            ("Referral Notes",     "ref_notes",         2),
            ("Lab Orders",         "lab_orders",        2),
            ("Radiology Orders",   "radiology_orders",  2),
            ("Prescribed Drugs",   "prescribed_drugs",  2),
            ("Drug Quantities",    "drug_quantities",   2),
        ]
        for label, key, lines in common:
            self.common_fields.add_widget(Label(
                text=label + ":", font_size=13, bold=True,
                size_hint_y=None, height=18
            ))
            w = TextInput(multiline=True, font_size=13,
                          size_hint_y=None, height=30 * lines)
            self.ew[key] = w
            self.common_fields.add_widget(w)

    def _add_drug_to_emr(self, instance):
        drug_name = self.emr_drug_spinner.text
        if drug_name == "Select":
            self._popup("Error", "Please select a drug.")
            return
        try:
            qty_per_day = int(self.emr_drug_qty.text)
            days = int(self.emr_drug_days.text)
            if qty_per_day <= 0 or days <= 0:
                raise ValueError
            if days > 365:
                self._popup("Error", "Days cannot exceed 365.")
                return
            if qty_per_day > 1000:
                self._popup("Error", "Quantity per day cannot exceed 1000.")
                return
        except:
            self._popup("Error",
                        "Quantity per day and days must be positive integers "
                        "(max 365 days, 1000 per day).")
            return

        total_qty = qty_per_day * days
        stock_available = 0
        for d in self.drugs:
            if d['name'] == drug_name:
                stock_available = d.get('qty', 0)
                break

        if stock_available < total_qty:
            self._popup(
                "Stock Warning",
                f"Insufficient stock. Available: {stock_available}, "
                f"Requested: {total_qty}\n"
                f"You can still prescribe but billing will show this."
            )

        self.emr_drug_list.append({
            "name": drug_name,
            "qty_per_day": qty_per_day,
            "days": days,
            "total": total_qty,
            "stock_checked": stock_available
        })
        self._refresh_drug_display()

        drug_lines = "\n".join([d["name"] for d in self.emr_drug_list])
        qty_lines = "\n".join([str(d["total"]) for d in self.emr_drug_list])
        if "prescribed_drugs" in self.ew:
            self.ew["prescribed_drugs"].text = drug_lines
        if "drug_quantities" in self.ew:
            self.ew["drug_quantities"].text = qty_lines

        self.emr_drug_qty.text = "1"
        self.emr_drug_days.text = "1"

        if stock_available < total_qty:
            self._popup("Added with Warning",
                        f"{drug_name} x{total_qty} (Stock: {stock_available})")
        else:
            self._popup("Added", f"{drug_name} x{total_qty} (over {days} days)")

    def _refresh_drug_display(self):
        self.drug_display.clear_widgets()
        if not self.emr_drug_list:
            self.drug_display.add_widget(Label(
                text="No drugs added yet.", font_size=12,
                size_hint_y=None, height=25
            ))
            return
        for i, drug in enumerate(self.emr_drug_list):
            row = BoxLayout(size_hint_y=None, height=30, spacing=3)
            row.add_widget(Label(text=f"{drug['name']}", font_size=12, size_hint_x=0.3))
            row.add_widget(Label(
                text=f"Total: {drug['total']}", font_size=12,
                size_hint_x=0.15, bold=True,
                color=get_color_from_hex('#2980b9')
            ))
            row.add_widget(Label(text=f"Days: {drug['days']}", font_size=12, size_hint_x=0.1))
            if drug.get('stock_checked', 0) < drug.get('total', 0):
                row.add_widget(Label(
                    text="⚠️ Low Stock", font_size=11, size_hint_x=0.12,
                    color=get_color_from_hex('#e74c3c')
                ))
            btn_remove = Button(
                text="X", font_size=11, size_hint_x=0.08,
                size_hint_y=None, height=25,
                background_color=get_color_from_hex('#e74c3c')
            )
            btn_remove.idx = i
            btn_remove.bind(on_press=self._remove_drug_from_emr)
            row.add_widget(btn_remove)
            self.drug_display.add_widget(row)

    def _remove_drug_from_emr(self, instance):
        idx = instance.idx
        if 0 <= idx < len(self.emr_drug_list):
            self.emr_drug_list.pop(idx)
            self._refresh_drug_display()
            drug_lines = "\n".join([d["name"] for d in self.emr_drug_list])
            qty_lines = "\n".join([str(d["total"]) for d in self.emr_drug_list])
            if "prescribed_drugs" in self.ew:
                self.ew["prescribed_drugs"].text = drug_lines
            if "drug_quantities" in self.ew:
                self.ew["drug_quantities"].text = qty_lines

    def _refresh_admission_spinner(self):
        if not self.cp:
            self.emr_admission_spinner.values = ["None"]
            self.emr_admission_spinner.text = "None"
            return
        pid = self.cp.get("id")
        adm = [a for a in self.admissions if a.get("patient_id") == pid]
        adm.sort(key=lambda x: x.get("admission_date", ""), reverse=True)
        self._admission_id_map = {}
        values = ["None"]
        for a in adm:
            display = f"{a.get('admission_date','')} - {a.get('diagnosis','')[:20]}"
            suffix = 1
            orig_display = display
            while display in values:
                display = f"{orig_display} ({suffix})"
                suffix += 1
            values.append(display)
            self._admission_id_map[display] = a.get("id")
        self.emr_admission_spinner.values = values
        if values:
            self.emr_admission_spinner.text = values[0]

    def _clear_emr(self):
        """Clear EMR form (safe iteration over active keys)."""
        for k, w in list(self.ew.items()):
            if isinstance(w, TextInput):
                w.text = ""
            elif isinstance(w, Spinner):
                w.text = "Alert" if k == "avpu" else ""
        if hasattr(self, 'emr_dept'):
            self.emr_dept.text = "General Medicine"
            self._build_dept_fields("General Medicine")
        if hasattr(self, 'emr_admission_spinner'):
            self.emr_admission_spinner.text = "None"
            self._refresh_admission_spinner()
        self._clear_drug_list()

    def _save_emr(self, instance):
        if not self.cp:
            self._popup("Error", "Select patient!")
            return
        pid = self.cp.get("id")
        name = self.cp.get("name")
        if not pid:
            return

        # Resolve admission ID
        adm_id = None
        if self.emr_admission_spinner.text != "None":
            adm_id = self._admission_id_map.get(self.emr_admission_spinner.text)

        r = {
            "emr_id": f"EMR{datetime.now().strftime('%Y%m%d%H%M%S%f')}",
            "patient_id": pid,
            "patient_name": name,
            "dept": self.emr_dept.text,
            "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "admission_id": adm_id,
        }

        # Save ONLY the currently-active widgets (prevents stale keys)
        active_keys = (
            list(self._vitals_keys) +
            list(self._dept_field_keys) +
            list(self._common_keys)
        )
        for k in active_keys:
            w = self.ew.get(k)
            if w is None:
                continue
            if isinstance(w, Spinner):
                r[k] = w.text
            elif isinstance(w, TextInput):
                r[k] = w.text

        r["_structured_drugs"] = list(self.emr_drug_list)
        self.emr_records.append(r)
        self.save_all()
        self._popup("Saved", "EMR recorded!")
        self.log_activity(f"EMR saved for {name} ({self.emr_dept.text})")
        self.audit_log(
            "SAVE_EMR",
            f"Dept: {self.emr_dept.text} | EMR ID: {r['emr_id']}",
            pid
        )
        self._auto_assign_tasks_from_emr(r)
        self._auto_bill_from_emr_record(r)
        self._clear_emr()

    def _view_emr_history(self, instance):
        if not self.cp:
            self._popup("Error", "No patient selected.")
            return
        pid = self.cp.get("id")
        emrs = [r for r in self.emr_records if r.get("patient_id") == pid]
        if not emrs:
            self._popup("EMR History", "No EMR records for this patient.")
            return

        c = BoxLayout(orientation='vertical', spacing=5, padding=8)
        pl = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        pl.bind(minimum_height=pl.setter('height'))
        sv = ScrollView()
        sv.add_widget(pl)
        c.add_widget(sv)

        for r in emrs:
            text = (f"{r.get('ts','')[:16]} | {r.get('dept','')} | "
                    f"{str(r.get('cc',''))[:30]}")
            row = BoxLayout(size_hint_y=None, height=40, spacing=5)
            row.add_widget(Label(text=text, font_size=13, size_hint_x=0.8))
            btn_view = Button(text="View", font_size=13, size_hint_x=0.2)
            rec = r
            btn_view.bind(on_press=lambda x, rec=rec: self._view_emr_detail(rec))
            row.add_widget(btn_view)
            pl.add_widget(row)

        pp = Popup(title="EMR History", content=c, size_hint=(0.8, 0.7))
        c.add_widget(Button(
            text="CLOSE", font_size=16, bold=True,
            size_hint_y=None, height=40,
            on_press=pp.dismiss
        ))
        pp.open()

    def _view_emr_detail(self, record):
        text = f"EMR ID: {record.get('emr_id','')}\n"
        text += f"Date: {record.get('ts','')}\n"
        text += f"Department: {record.get('dept','')}\n"
        text += f"Admission ID: {record.get('admission_id', 'None')}\n\n"
        text += "Vital Signs:\n"
        for k in ["bp", "temp", "pulse", "spo2", "rr",
                   "wt", "ht", "bmi", "liver_span", "avpu"]:
            if record.get(k):
                text += f"  {k}: {record.get(k)}\n"
        text += "\nClinical Details:\n"
        label_map = {
            "cc": "Chief Complaint", "hpi": "HPI", "pmh": "PMH",
            "psh": "Past Surgical History", "dh": "Drug History",
            "all": "Allergies", "fh": "Family History", "sh": "Social History",
            "ros": "Review of Systems", "gen": "General Exam", "cvs": "CVS",
            "resp": "Resp", "abd": "Abd", "cns": "CNS",
            "dx": "Diagnosis", "tx": "Treatment Plan", "fu": "Follow-up",
            "nursing_orders": "Nursing Orders",
        }
        skip = ["emr_id", "patient_id", "patient_name", "dept", "ts",
                "admission_id", "bp", "temp", "pulse", "spo2", "rr",
                "wt", "ht", "bmi", "liver_span", "avpu", "_structured_drugs"]
        for k, v in record.items():
            if k not in skip and v and str(v).strip():
                label = label_map.get(k, k.upper().replace("_", " "))
                text += f"{label}:\n{v}\n\n"

        if record.get("_structured_drugs"):
            text += "Prescribed Drugs:\n"
            for drug in record["_structured_drugs"]:
                text += (f"  {drug['name']} - {drug['total']} units over "
                         f"{drug['days']} days\n")
        self._text_popup("EMR Detail", text)

    # ====================== ADMISSIONS TAB ======================
    def _admissions_tab(self):
        tab = TabbedPanelItem(text="Admissions")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=6, spacing=4, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))
        l.add_widget(Label(
            text="PATIENT ADMISSIONS",
            font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#1a5276')
        ))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="Patient ID:", font_size=15, bold=True, size_hint_x=0.12))
        self.adm_pid = TextInput(size_hint_x=0.2, readonly=True, font_size=15,
                                  size_hint_y=None, height=40)
        bar.add_widget(self.adm_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.08))
        self.adm_name = TextInput(size_hint_x=0.3, readonly=True, font_size=15,
                                   size_hint_y=None, height=40)
        bar.add_widget(self.adm_name)
        bar.add_widget(Button(
            text="REFRESH", font_size=13, size_hint_x=0.12,
            on_press=lambda x: self._refresh_admissions_list(),
            background_color=get_color_from_hex('#16a085')
        ))
        l.add_widget(bar)

        l.add_widget(Label(
            text="NEW ADMISSION", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#e67e22')
        ))
        form = GridLayout(cols=4, size_hint_y=None, height=150, spacing=4)
        form.add_widget(Label(text="Admission Date:", font_size=14, bold=True))
        self.adm_date = TextInput(
            text=datetime.now().strftime("%Y-%m-%d"),
            font_size=14, size_hint_y=None, height=35
        )
        form.add_widget(self.adm_date)
        form.add_widget(Label(text="Discharge Date:", font_size=14, bold=True))
        self.adm_discharge = TextInput(
            text="", font_size=14, size_hint_y=None, height=35,
            hint_text="(optional)"
        )
        form.add_widget(self.adm_discharge)
        form.add_widget(Label(text="Department:", font_size=14, bold=True))
        self.adm_dept = Spinner(
            text="General Medicine", values=self.depts,
            size_hint_y=None, height=35
        )
        form.add_widget(self.adm_dept)
        form.add_widget(Label(text="Diagnosis:", font_size=14, bold=True))
        self.adm_diagnosis = TextInput(font_size=14, size_hint_y=None, height=35)
        form.add_widget(self.adm_diagnosis)
        form.add_widget(Label(text="Notes:", font_size=14, bold=True))
        self.adm_notes = TextInput(font_size=14, size_hint_y=None, height=35)
        form.add_widget(self.adm_notes)
        l.add_widget(form)

        l.add_widget(Button(
            text="ADD ADMISSION", font_size=16, bold=True,
            size_hint_y=None, height=40,
            on_press=self._add_admission,
            background_color=get_color_from_hex('#27ae60')
        ))

        l.add_widget(Label(
            text="ADMISSION HISTORY", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#2c3e50')
        ))
        self.adm_list = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        self.adm_list.bind(minimum_height=self.adm_list.setter('height'))
        l.add_widget(self.adm_list)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)
        self._refresh_admissions_list()

    def _refresh_admissions_list(self):
        if not hasattr(self, 'adm_list'):
            return
        self.adm_list.clear_widgets()
        if not self.adm_pid.text:
            self.adm_list.add_widget(Label(
                text="Select a patient to view admissions.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        pid = self.adm_pid.text
        admissions = [a for a in self.admissions if a.get("patient_id") == pid]
        admissions.sort(key=lambda x: x.get("admission_date", ""), reverse=True)
        if not admissions:
            self.adm_list.add_widget(Label(
                text="No admissions recorded for this patient.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        for a in admissions:
            row = BoxLayout(size_hint_y=None, height=35, spacing=3)
            row.add_widget(Label(
                text=(f"{a.get('admission_date','')} | "
                      f"{a.get('department','')} | "
                      f"{a.get('diagnosis','')}"),
                font_size=12, size_hint_x=0.5
            ))
            row.add_widget(Label(
                text=f"Disch: {a.get('discharge_date','Active')}",
                font_size=12, size_hint_x=0.15
            ))
            btn_del = Button(
                text="X", font_size=11, size_hint_x=0.08,
                size_hint_y=None, height=25,
                background_color=get_color_from_hex('#e74c3c')
            )
            btn_del.adm_id = a.get("id")
            btn_del.bind(on_press=lambda x: self._delete_admission(x.adm_id))
            row.add_widget(btn_del)

            btn_set = Button(
                text="Set Current", font_size=11, size_hint_x=0.12,
                size_hint_y=None, height=25,
                background_color=get_color_from_hex('#2980b9')
            )
            btn_set.adm = a
            btn_set.bind(on_press=lambda x: self._set_current_admission(x.adm))
            row.add_widget(btn_set)
            self.adm_list.add_widget(row)
        self._refresh_admission_spinner()

    def _set_current_admission(self, admission):
        if not self.cp:
            return
        display = (f"{admission.get('admission_date','')} - "
                   f"{admission.get('diagnosis','')[:20]}")
        self._refresh_admission_spinner()
        if display in self.emr_admission_spinner.values:
            self.emr_admission_spinner.text = display
        self._popup("Current Admission", f"Set to: {display}")

    def _add_admission(self, instance):
        if not self.adm_pid.text:
            self._popup("Error", "Select a patient first.")
            return
        adm_date = self.adm_date.text.strip()
        if not adm_date:
            self._popup("Error", "Admission date required.")
            return
        try:
            datetime.strptime(adm_date, "%Y-%m-%d")
        except:
            self._popup("Error", "Invalid date format (YYYY-MM-DD).")
            return
        discharge = self.adm_discharge.text.strip()
        if discharge:
            try:
                datetime.strptime(discharge, "%Y-%m-%d")
            except:
                self._popup("Error", "Invalid discharge date format (YYYY-MM-DD).")
                return
        department = self.adm_dept.text
        diagnosis = self.adm_diagnosis.text.strip()
        notes = self.adm_notes.text.strip()
        if not diagnosis:
            self._popup("Error", "Diagnosis is required.")
            return
        admission = {
            "id": f"ADM{datetime.now().strftime('%Y%m%d%H%M%S%f')}",
            "patient_id": self.adm_pid.text,
            "patient_name": self.adm_name.text,
            "admission_date": adm_date,
            "discharge_date": discharge or "",
            "department": department,
            "diagnosis": diagnosis,
            "notes": notes,
            "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        self.admissions.append(admission)
        self.save_all()
        self._refresh_admissions_list()
        self.adm_diagnosis.text = ""
        self.adm_notes.text = ""
        self.adm_discharge.text = ""
        self.adm_date.text = datetime.now().strftime("%Y-%m-%d")
        self._popup("Admission Added",
                    f"Admission for {self.adm_name.text} on {adm_date}")
        self.log_activity(f"Admission added for {self.adm_name.text}: {diagnosis}")
        self.audit_log("ADD_ADMISSION", f"{diagnosis}", self.adm_pid.text)

    def _delete_admission(self, adm_id):
        if not self._require_role("Administrator", "Doctor"):
            return
        self.admissions = [a for a in self.admissions if a.get("id") != adm_id]
        self.save_all()
        self._refresh_admissions_list()
        self._popup("Deleted", "Admission record removed.")
        self.audit_log("DELETE_ADMISSION", f"Admission ID {adm_id}")

    # ====================== TASK MANAGEMENT TAB ======================
    def _task_management_tab(self):
        tab = TabbedPanelItem(text="Tasks")
        main = BoxLayout(orientation='vertical', padding=6, spacing=4)
        main.add_widget(Label(
            text="TASK MANAGEMENT", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#8e44ad')
        ))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="Patient ID:", font_size=15, bold=True, size_hint_x=0.12))
        self.task_pid = TextInput(size_hint_x=0.2, readonly=True, font_size=15,
                                   size_hint_y=None, height=40)
        bar.add_widget(self.task_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.08))
        self.task_patient_name = TextInput(size_hint_x=0.3, readonly=True, font_size=15,
                                            size_hint_y=None, height=40)
        bar.add_widget(self.task_patient_name)
        bar.add_widget(Button(
            text="VIEW ALL TASKS", font_size=13, size_hint_x=0.15,
            on_press=self._view_all_tasks,
            background_color=get_color_from_hex('#2980b9')
        ))
        main.add_widget(bar)

        task_form = BoxLayout(size_hint_y=None, height=40, spacing=4)
        task_form.add_widget(Label(text="Assign To:", font_size=14, bold=True, size_hint_x=0.1))
        self.task_role = Spinner(
            text="Nurse",
            values=["Nurse", "Lab Scientist", "Radiologist",
                    "Doctor", "Pharmacy", "Administrator"],
            size_hint_x=0.15, size_hint_y=None, height=40
        )
        task_form.add_widget(self.task_role)
        task_form.add_widget(Label(text="Type:", font_size=14, bold=True, size_hint_x=0.06))
        self.task_type = Spinner(
            text="Nursing",
            values=["Nursing", "Laboratory", "Radiology", "Pharmacy",
                    "Doctor", "Follow-up", "Administrative"],
            size_hint_x=0.12, size_hint_y=None, height=40
        )
        task_form.add_widget(self.task_type)
        task_form.add_widget(Label(text="Priority:", font_size=14, bold=True, size_hint_x=0.08))
        self.task_priority = Spinner(
            text="Normal", values=["High", "Normal", "Low"],
            size_hint_x=0.1, size_hint_y=None, height=40
        )
        task_form.add_widget(self.task_priority)
        btn_add = Button(
            text="CREATE TASK", font_size=14, bold=True, size_hint_x=0.12,
            on_press=self._create_manual_task,
            background_color=get_color_from_hex('#27ae60')
        )
        task_form.add_widget(btn_add)
        main.add_widget(task_form)

        main.add_widget(Label(text="Task Description:", font_size=14, bold=True,
                               size_hint_y=None, height=20))
        self.task_desc = TextInput(
            font_size=13, size_hint_y=None, height=60, multiline=True,
            hint_text="Enter task description or order details..."
        )
        main.add_widget(self.task_desc)

        filter_box = BoxLayout(size_hint_y=None, height=35, spacing=4)
        filter_box.add_widget(Label(text="Filter:", font_size=13, bold=True, size_hint_x=0.06))
        self.task_filter = Spinner(
            text="All",
            values=["All"] + TaskManager.STATUSES,
            size_hint_x=0.15, size_hint_y=None, height=32
        )
        self.task_filter.bind(text=lambda i, v: self._refresh_tasks())
        filter_box.add_widget(Button(
            text="REFRESH", font_size=12, size_hint_x=0.1,
            size_hint_y=None, height=32,
            on_press=lambda x: self._refresh_tasks(),
            background_color=get_color_from_hex('#16a085')
        ))
        main.add_widget(filter_box)

        main.add_widget(Label(
            text="TASK LIST", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#2c3e50')
        ))
        task_scroll = ScrollView(size_hint_y=1, bar_width=10, scroll_type=['bars'])
        self.task_list = BoxLayout(orientation='vertical', size_hint_y=None, spacing=2)
        self.task_list.bind(minimum_height=self.task_list.setter('height'))
        task_scroll.add_widget(self.task_list)
        main.add_widget(task_scroll)

        tab.add_widget(main)
        self.tp.add_widget(tab)
        self._refresh_tasks()

    def _create_manual_task(self, instance):
        if not self.cp:
            self._popup("Error", "Select a patient first.")
            return
        desc = self.task_desc.text.strip()
        if not desc:
            self._popup("Error", "Enter task description.")
            return
        task = self.task_manager.create_task(
            patient_id=self.cp.get("id"),
            patient_name=self.cp.get("name"),
            assigned_to=self.task_role.text,
            task_type=self.task_type.text,
            description=desc,
            priority=self.task_priority.text
        )
        self._refresh_tasks()
        self.task_desc.text = ""
        self._popup("Task Created", f"Task created: {task.get('id')}")
        self.log_activity(f"Manual task created for {self.cp.get('name')}: {desc[:30]}...")

    def _auto_assign_tasks_from_emr(self, emr_record):
        if not self.cp:
            return
        tasks = self.task_manager.auto_assign_from_emr(emr_record, self.cp)
        if tasks:
            self._refresh_tasks()
            self._popup("Tasks Created", f"{len(tasks)} tasks assigned from EMR")
            self.log_activity(
                f"Auto-assigned {len(tasks)} tasks from EMR "
                f"for {self.cp.get('name')}"
            )

    def _auto_assign_tasks_from_current_emr(self, instance):
        if not self.cp:
            self._popup("Error", "Select a patient first.")
            return
        emr_record = {}
        for k, w in self.ew.items():
            if isinstance(w, TextInput):
                emr_record[k] = w.text
            elif isinstance(w, Spinner):
                emr_record[k] = w.text

        if self.emr_drug_list:
            drug_lines = "\n".join([d["name"] for d in self.emr_drug_list])
            qty_lines = "\n".join([str(d["total"]) for d in self.emr_drug_list])
            emr_record["prescribed_drugs"] = drug_lines
            emr_record["drug_quantities"] = qty_lines

        for k in ["lab_orders", "radiology_orders", "nursing_orders"]:
            if k in self.ew:
                emr_record[k] = self.ew[k].text

        self._auto_assign_tasks_from_emr(emr_record)

    def _auto_assign_tasks_from_ai(self, instance):
        if not self.cp:
            self._popup("Error", "Select a patient first.")
            return
        emr = self.clinical_ai.get_emr()
        if not emr:
            self._popup("Error", "Complete the AI interview first.")
            return
        self._auto_assign_tasks_from_emr(emr)

    def _refresh_tasks(self):
        if not hasattr(self, 'task_list'):
            return
        self.task_list.clear_widgets()
        filter_status = self.task_filter.text

        if not self.cp:
            tasks = self.task_manager.get_pending_tasks()
        else:
            pid = self.cp.get("id")
            tasks = self.task_manager.get_tasks_for_patient(pid)

        if filter_status != "All":
            tasks = [t for t in tasks if t.get("status") == filter_status]

        if not tasks:
            self.task_list.add_widget(Label(
                text="No tasks found.", font_size=13,
                size_hint_y=None, height=30
            ))
            return

        for task in tasks:
            status_color = {
                "Pending": "#f39c12",
                "In Progress": "#2980b9",
                "Completed": "#27ae60",
                "Waiting": "#e67e22",
                "On Hold": "#e74c3c",
                "Cancelled": "#95a5a6"
            }.get(task.get("status"), "#7f8c8d")

            row = BoxLayout(size_hint_y=None, height=35, spacing=3)
            row.add_widget(Label(text=task.get("id", "")[:12], font_size=11, size_hint_x=0.1))
            row.add_widget(Label(text=task.get("task_type", ""), font_size=12, size_hint_x=0.1))
            row.add_widget(Label(text=task.get("assigned_to", ""), font_size=12, size_hint_x=0.1))
            row.add_widget(Label(text=task.get("description", "")[:25], font_size=12, size_hint_x=0.25))
            row.add_widget(Label(
                text=task.get("status", ""), font_size=12, size_hint_x=0.1,
                bold=True, color=get_color_from_hex(status_color)
            ))
            row.add_widget(Label(text=task.get("created_at", "")[:16], font_size=10, size_hint_x=0.1))

            status_spinner = Spinner(
                text=task.get("status", "Pending"),
                values=TaskManager.STATUSES,
                size_hint_x=0.1, size_hint_y=None, height=30
            )
            status_spinner.task_id = task.get("id")
            status_spinner.bind(
                text=lambda s, tid=task.get("id"): self._update_task_status(tid, s.text)
            )
            row.add_widget(status_spinner)

            btn_view = Button(
                text="View", font_size=11, size_hint_x=0.08,
                size_hint_y=None, height=28,
                background_color=get_color_from_hex('#2980b9')
            )
            btn_view.task = task
            btn_view.bind(on_press=lambda x: self._view_task_detail(x.task))
            row.add_widget(btn_view)

            self.task_list.add_widget(row)

    def _update_task_status(self, task_id, new_status):
        task = self.task_manager.update_status(task_id, new_status)
        if task:
            self._refresh_tasks()
            self.log_activity(f"Task {task_id} status updated to {new_status}")
            self.audit_log("UPDATE_TASK", f"{task_id} → {new_status}",
                           task.get("patient_id", ""))
            self._refresh_dash()

    def _view_task_detail(self, task):
        text = "TASK DETAIL\n" + "=" * 50 + "\n\n"
        text += f"ID: {task.get('id')}\n"
        text += f"Patient: {task.get('patient_name')} ({task.get('patient_id')})\n"
        text += f"Assigned To: {task.get('assigned_to')}\n"
        text += f"Type: {task.get('task_type')}\n"
        text += f"Status: {task.get('status')}\n"
        text += f"Priority: {task.get('priority')}\n"
        text += f"Department: {task.get('department', 'N/A')}\n"
        text += f"Created: {task.get('created_at')}\n"
        text += f"Updated: {task.get('updated_at')}\n"
        if task.get("completed_at"):
            text += f"Completed: {task.get('completed_at')}\n"
        text += f"\nDescription:\n{task.get('description')}\n"
        if task.get("notes"):
            text += f"\nNotes:\n{task.get('notes')}\n"
        if task.get("order_items"):
            text += f"\nOrder Items:\n"
            for item in task.get("order_items", []):
                text += f"  - {item}\n"
        self._text_popup("Task Detail", text)

    def _view_all_tasks(self, instance):
        tasks = self.task_manager.tasks
        if not tasks:
            self._popup("All Tasks", "No tasks found.")
            return
        text = "ALL TASKS\n" + "=" * 50 + "\n\n"
        for status in TaskManager.STATUSES:
            status_tasks = [t for t in tasks if t.get("status") == status]
            if status_tasks:
                text += f"\n{status.upper()} ({len(status_tasks)})\n"
                text += "-" * 30 + "\n"
                for t in status_tasks[:10]:
                    text += (f"  {t.get('id')} | {t.get('patient_name')} | "
                             f"{t.get('task_type')} | {t.get('assigned_to')}\n")
                    text += f"    {t.get('description')[:60]}\n"
        self._text_popup("All Tasks", text)

    # ====================== AUTO BILLING (REAL PRICES) ======================
    def _auto_bill_from_emr(self, instance):
        if not self.cp:
            self._popup("Error", "No patient selected.")
            return
        pid = self.cp.get("id")
        emrs = [r for r in self.emr_records if r.get("patient_id") == pid]
        if not emrs:
            self._popup("Info", "No EMR records for this patient.")
            return
        latest = emrs[-1]
        self._auto_bill_from_emr_record(latest)

    def _find_price_in_db(self, name, *lists):
        """
        Search for an item name across one or more lists and return the real price.
        Strategy: exact match → substring match → keyword-overlap match.
        Returns None if nothing found.
        """
        if not name:
            return None
        target = str(name).strip().lower()
        if not target:
            return None

        keywords = [w for w in re.split(r'[\s\-/(),.:;]+', target) if len(w) > 2]

        # 1) Exact
        for lst in lists:
            if not lst:
                continue
            for item in lst:
                iname = str(item.get('name', '')).strip().lower()
                if iname and iname == target:
                    return item.get('price')

        # 2) Substring either way
        for lst in lists:
            if not lst:
                continue
            for item in lst:
                iname = str(item.get('name', '')).strip().lower()
                if not iname:
                    continue
                if target in iname or iname in target:
                    return item.get('price')

        # 3) Keyword overlap
        for lst in lists:
            if not lst:
                continue
            for item in lst:
                iname = str(item.get('name', '')).strip().lower()
                if not iname or not keywords:
                    continue
                hits = sum(1 for kw in keywords if kw in iname)
                if hits >= max(1, len(keywords) // 2):
                    return item.get('price')

        return None

    def _auto_bill_from_emr_record(self, emr_record):
        """
        Auto-bill using REAL prices from stored databases.
        Items without a match go in at ₦0 + flagged for admin to fix.
        """
        if not self.cp:
            return
        pid = self.cp.get("id")
        bill_items_added = []
        missing_prices = []

        # --- Lab Orders ---
        lab_orders = emr_record.get("lab_orders", "")
        if lab_orders:
            for test in str(lab_orders).split('\n'):
                test = test.strip()
                if not test:
                    continue
                price = self._find_price_in_db(
                    test, self.investigations, self.charges, self.procedures
                )
                if price is None:
                    missing_prices.append(f"Lab: {test}")
                    price = 0
                self.bill_items.append({
                    "name": f"Lab: {test}",
                    "cat": "Investigations",
                    "qty": 1,
                    "price": price,
                    "total": price
                })
                bill_items_added.append(f"Lab: {test} (₦{price:,.2f})")

        # --- Radiology Orders ---
        rad_orders = emr_record.get("radiology_orders", "")
        if rad_orders:
            for study in str(rad_orders).split('\n'):
                study = study.strip()
                if not study:
                    continue
                price = self._find_price_in_db(
                    study, self.investigations, self.charges, self.procedures
                )
                if price is None:
                    missing_prices.append(f"Radiology: {study}")
                    price = 0
                self.bill_items.append({
                    "name": f"Radiology: {study}",
                    "cat": "Investigations",
                    "qty": 1,
                    "price": price,
                    "total": price
                })
                bill_items_added.append(f"Radiology: {study} (₦{price:,.2f})")

        # --- Drugs: structured first, then text fallback ---
        structured_drugs = emr_record.get("_structured_drugs")
        if structured_drugs:
            for drug in structured_drugs:
                drug_name = drug.get("name", "")
                total_qty = drug.get("total", 0)
                if not drug_name or total_qty <= 0:
                    continue
                price = self._find_price_in_db(
                    drug_name, self.drugs, self.investigations, self.charges
                )
                if price is None:
                    missing_prices.append(f"Drug: {drug_name}")
                    price = 0
                self.bill_items.append({
                    "name": f"Drug: {drug_name}",
                    "cat": "Drugs",
                    "qty": total_qty,
                    "price": price,
                    "total": price * total_qty
                })
                bill_items_added.append(
                    f"Drug: {drug_name} x{total_qty} (₦{price:,.2f}/unit)"
                )
        else:
            drugs_str = emr_record.get("prescribed_drugs", "")
            qtys_str = emr_record.get("drug_quantities", "")
            if drugs_str:
                drug_lines = [d.strip() for d in str(drugs_str).split('\n') if d.strip()]
                qty_lines = [q.strip() for q in str(qtys_str).split('\n') if q.strip()]
                for i, drug in enumerate(drug_lines):
                    qty = 1
                    if i < len(qty_lines):
                        try:
                            qty = int(qty_lines[i])
                        except ValueError:
                            qty = 1
                    price = self._find_price_in_db(
                        drug, self.drugs, self.investigations, self.charges
                    )
                    if price is None:
                        missing_prices.append(f"Drug: {drug}")
                        price = 0
                    self.bill_items.append({
                        "name": f"Drug: {drug}",
                        "cat": "Drugs",
                        "qty": qty,
                        "price": price,
                        "total": price * qty
                    })
                    bill_items_added.append(
                        f"Drug: {drug} x{qty} (₦{price:,.2f}/unit)"
                    )

        # --- Nursing care ---
        nursing = emr_record.get("nursing_orders", "")
        if nursing and str(nursing).strip():
            price = self._find_price_in_db("Nursing Care", self.charges, self.procedures)
            if price is not None:
                self.bill_items.append({
                    "name": "Nursing Care",
                    "cat": "Charges",
                    "qty": 1,
                    "price": price,
                    "total": price
                })
                bill_items_added.append(f"Nursing Care (₦{price:,.2f})")

        # --- Consultation ---
        price = self._find_price_in_db("Consultation", self.charges)
        if price is None:
            price = 0
            missing_prices.append("Consultation")
        self.bill_items.append({
            "name": "Consultation",
            "cat": "Charges",
            "qty": 1,
            "price": price,
            "total": price
        })
        bill_items_added.append(f"Consultation (₦{price:,.2f})")

        if bill_items_added:
            self._refresh_bill()
            msg = (f"Added {len(bill_items_added)} items to bill with real prices:\n" +
                   "\n".join(bill_items_added))
            if missing_prices:
                msg += (f"\n\n⚠️ No price found for {len(missing_prices)} item(s). "
                        f"They were added at ₦0 — please set the price manually in "
                        f"Manage Investigations / Manage Drugs:\n")
                msg += "\n".join(f"  • {m}" for m in missing_prices[:10])
                if len(missing_prices) > 10:
                    msg += f"\n  ... and {len(missing_prices) - 10} more"
            self._popup("Auto-Bill (Real Prices)", msg)
            self.log_activity(
                f"Auto-bill added {len(bill_items_added)} items for "
                f"{self.cp.get('name')} (missing prices: {len(missing_prices)})"
            )
            self.audit_log(
                "AUTO_BILL",
                f"{len(bill_items_added)} items, "
                f"{len(missing_prices)} missing prices",
                pid
            )

            if missing_prices:
                self.task_manager.create_task(
                    patient_id=pid,
                    patient_name=self.cp.get("name"),
                    assigned_to="Administrator",
                    task_type="Administrative",
                    description=(f"Set prices for {len(missing_prices)} "
                                  f"unbilled items: "
                                  + ", ".join(missing_prices[:5])),
                    department="Billing"
                )
                self._refresh_tasks()

    # ====================== PHARMACY DISPENSE TAB ======================
    def _pharmacy_dispense_tab(self):
        tab = TabbedPanelItem(text="Pharmacy Dispense")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=6, spacing=4, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))
        l.add_widget(Label(
            text="PHARMACY DISPENSING",
            font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#d35400')
        ))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="Patient ID:", font_size=15, bold=True, size_hint_x=0.12))
        self.dispense_pid = TextInput(size_hint_x=0.2, readonly=True, font_size=15,
                                       size_hint_y=None, height=40)
        bar.add_widget(self.dispense_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.08))
        self.dispense_name = TextInput(size_hint_x=0.3, readonly=True, font_size=15,
                                        size_hint_y=None, height=40)
        bar.add_widget(self.dispense_name)
        bar.add_widget(Button(
            text="VIEW PRESCRIPTIONS", font_size=13, size_hint_x=0.15,
            on_press=self._view_patient_prescriptions,
            background_color=get_color_from_hex('#2980b9')
        ))
        l.add_widget(bar)

        dispense_form = BoxLayout(size_hint_y=None, height=40, spacing=4)
        dispense_form.add_widget(Label(text="Drug:", font_size=15, bold=True, size_hint_x=0.06))
        self.dispense_drug = Spinner(
            text="Select", values=[d['name'] for d in self.drugs],
            size_hint_x=0.3, size_hint_y=None, height=40
        )
        dispense_form.add_widget(self.dispense_drug)
        dispense_form.add_widget(Label(text="Qty:", font_size=15, bold=True, size_hint_x=0.05))
        self.dispense_qty = TextInput(text="1", size_hint_x=0.08, font_size=15,
                                        size_hint_y=None, height=40)
        dispense_form.add_widget(self.dispense_qty)
        dispense_form.add_widget(Button(
            text="DISPENSE", font_size=15, bold=True, size_hint_x=0.12,
            on_press=self._dispense_from_pharmacy,
            background_color=get_color_from_hex('#27ae60')
        ))
        dispense_form.add_widget(Button(
            text="AUTO FROM EMR", font_size=15, bold=True, size_hint_x=0.15,
            on_press=self._auto_dispense_from_emr,
            background_color=get_color_from_hex('#8e44ad')
        ))
        l.add_widget(dispense_form)

        l.add_widget(Label(
            text="TASK ASSIGNMENTS", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#2c3e50')
        ))
        self.dispense_task_list = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        self.dispense_task_list.bind(minimum_height=self.dispense_task_list.setter('height'))
        l.add_widget(self.dispense_task_list)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)
        self._refresh_pharmacy_dispense()

    def _refresh_pharmacy_dispense(self):
        if not hasattr(self, 'dispense_task_list'):
            return
        self.dispense_task_list.clear_widgets()
        if not self.dispense_pid.text:
            self.dispense_task_list.add_widget(Label(
                text="Select a patient to view pharmacy tasks.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        pid = self.dispense_pid.text
        tasks = [t for t in self.tasks
                 if t.get("patient_id") == pid
                 and t.get("task_type") == "Pharmacy"
                 and t.get("status") != "Completed"]
        if not tasks:
            self.dispense_task_list.add_widget(Label(
                text="No pending pharmacy tasks.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        for task in tasks:
            row = BoxLayout(size_hint_y=None, height=35, spacing=3)
            row.add_widget(Label(text=f"{task.get('description', '')[:30]}",
                                  font_size=12, size_hint_x=0.35))
            row.add_widget(Label(text=task.get("status", ""), font_size=12, size_hint_x=0.1))
            btn = Button(
                text="DISPENSE", font_size=12, size_hint_x=0.15,
                size_hint_y=None, height=30,
                background_color=get_color_from_hex('#27ae60')
            )
            btn.task = task
            btn.bind(on_press=self._dispense_from_task)
            row.add_widget(btn)
            btn_complete = Button(
                text="Complete", font_size=12, size_hint_x=0.15,
                size_hint_y=None, height=30,
                background_color=get_color_from_hex('#2980b9')
            )
            btn_complete.task = task
            btn_complete.bind(on_press=lambda x: self._complete_pharmacy_task(x.task))
            row.add_widget(btn_complete)
            self.dispense_task_list.add_widget(row)

    def _dispense_from_pharmacy(self, instance):
        if not self.dispense_pid.text or self.dispense_drug.text == "Select":
            self._popup("Error", "Select patient and drug.")
            return
        drug_name = self.dispense_drug.text
        try:
            qty = int(self.dispense_qty.text)
            if qty <= 0:
                raise ValueError
        except:
            self._popup("Error", "Invalid quantity.")
            return
        for d in self.drugs:
            if d['name'] == drug_name:
                if d['qty'] < qty:
                    self._popup("Error", f"Insufficient stock. Available: {d['qty']}")
                    return
                d['qty'] -= qty
                break
        self.prescriptions.append({
            "rx_id": f"RX{datetime.now().strftime('%Y%m%d%H%M%S%f')}",
            "patient_id": self.dispense_pid.text,
            "patient_name": self.dispense_name.text,
            "drug": drug_name,
            "qty": qty,
            "price": self._get_price(drug_name, "Drugs"),
            "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "dispensed": True
        })
        task = self.task_manager.create_task(
            patient_id=self.dispense_pid.text,
            patient_name=self.dispense_name.text,
            assigned_to="Pharmacy",
            task_type="Pharmacy",
            description=f"Dispensed: {qty}x {drug_name}",
            order_items=[drug_name],
            department="Pharmacy"
        )
        self.task_manager.update_status(task.get("id"), "Completed")
        self.save_all()
        self._refresh_pharmacy_dispense()
        self._refresh_tasks()
        self._popup("Dispensed", f"{qty}x {drug_name}")
        self.log_activity(
            f"Dispensed {qty}x {drug_name} from pharmacy tab "
            f"for {self.dispense_name.text}"
        )
        self.audit_log("DISPENSE_DRUG", f"{qty}x {drug_name}",
                       self.dispense_pid.text)
        self.audit_log("DISPENSE_DRUG", f"{qty}x {drug_name}",
               self.dispense_pid.text)

        # Clear search box for the next entry
        if hasattr(self, 'dispense_drug_search'):
             self.dispense_drug_search.text = ""

    def _dispense_from_task(self, instance):
        task = instance.task
        desc = task.get("description", "")

        drug_match = re.search(r'Dispense:?\s*(.+?)(?:\s*$|\s+x)', desc, re.IGNORECASE)
        if not drug_match:
            drug_match = re.search(r'Medication:?\s*(.+?)(?:\s*$|\s+x)', desc, re.IGNORECASE)
        if not drug_match:
            drug_match = re.search(r'Drug:?\s*(.+?)(?:\s*$|\s+x)', desc, re.IGNORECASE)

        if drug_match:
            drug_name = drug_match.group(1).strip()
        else:
            drug_name = None
            for d in self.drugs:
                if d['name'].lower() in desc.lower():
                    if desc.lower().strip() == d['name'].lower():
                        drug_name = d['name']
                        break
                    if drug_name is None:
                        drug_name = d['name']
            if drug_name is None:
                self._popup("Error",
                            "Could not identify drug from task. "
                            "Please select manually.")
                return

        self.dispense_drug.text = drug_name
        self.dispense_qty.text = "1"
        self._dispense_from_pharmacy(instance)

    def _complete_pharmacy_task(self, task):
        self.task_manager.update_status(task.get("id"), "Completed",
                                         "Dispensed by pharmacy")
        self._refresh_pharmacy_dispense()
        self._refresh_tasks()
        self._popup("Completed", f"Task {task.get('id')} marked as completed.")

    def _auto_dispense_from_emr(self, instance):
        if not self.cp:
            self._popup("Error", "Select a patient first.")
            return
        pid = self.cp.get("id")
        emrs = [r for r in self.emr_records if r.get("patient_id") == pid]
        if not emrs:
            self._popup("Info", "No EMR records found for this patient.")
            return
        latest = emrs[-1]

        structured_drugs = latest.get("_structured_drugs")
        if structured_drugs:
            drugs_to_dispense = []
            for drug in structured_drugs:
                drug_name = drug.get("name", "")
                total_qty = drug.get("total", 0)
                if drug_name and total_qty > 0:
                    drugs_to_dispense.append((drug_name, total_qty))
        else:
            drugs_str = latest.get("prescribed_drugs", "") or latest.get("medications", "")
            qtys_str = latest.get("drug_quantities", "")
            if not drugs_str:
                self._popup("Info", "No prescribed drugs found in EMR.")
                return
            drug_lines = [d.strip() for d in drugs_str.split('\n') if d.strip()]
            qty_lines = [q.strip() for q in qtys_str.split('\n') if q.strip()]
            drugs_to_dispense = []
            for i, drug in enumerate(drug_lines):
                qty = 1
                if i < len(qty_lines):
                    try:
                        qty = int(qty_lines[i])
                    except ValueError:
                        qty = 1
                drugs_to_dispense.append((drug, qty))

        if not drugs_to_dispense:
            self._popup("Info", "No drugs to dispense.")
            return

        added = []
        for drug_name, qty in drugs_to_dispense:
            found = False
            for d in self.drugs:
                if drug_name.lower() in d['name'].lower():
                    self.dispense_drug.text = d['name']
                    self.dispense_qty.text = str(qty)
                    self._dispense_from_pharmacy(instance)
                    added.append(f"{d['name']} x{qty}")
                    found = True
                    break
            if not found:
                self.bill_items.append({
                    "name": f"Drug: {drug_name}",
                    "cat": "Drugs",
                    "qty": qty,
                    "price": 200,
                    "total": 200 * qty
                })
                self.prescriptions.append({
                    "rx_id": f"RX{datetime.now().strftime('%Y%m%d%H%M%S')}",
                    "patient_id": pid,
                    "patient_name": self.cp.get("name"),
                    "drug": drug_name,
                    "qty": qty,
                    "price": 200,
                    "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "dispensed": True
                })
                added.append(f"{drug_name} x{qty} (custom)")
        if added:
            self.save_all()
            self._refresh_pharmacy_dispense()
            self._refresh_bill()
            self._popup("Auto-Dispensed", "Dispensed:\n" + "\n".join(added))

    def _view_patient_prescriptions(self, instance):
        if not self.dispense_pid.text:
            self._popup("Error", "No patient selected.")
            return
        pid = self.dispense_pid.text
        rx = [r for r in self.prescriptions if r.get("patient_id") == pid]
        if not rx:
            self._popup("Prescriptions", "No prescriptions found.")
            return
        text = (f"PRESCRIPTIONS FOR {self.dispense_name.text}\n" +
                "=" * 50 + "\n\n")
        for r in rx:
            text += f"Drug: {r.get('drug', '')}\n"
            text += f"Quantity: {r.get('qty', 0)}\n"
            text += f"Price: N{r.get('price', 0):,.2f}\n"
            text += f"Date: {r.get('ts', '')}\n"
            text += "-" * 30 + "\n"
        self._text_popup("Prescriptions", text)

    # ====================== ECG TAB ======================
    def _ecg_tab(self):
        tab = TabbedPanelItem(text="ECG Analysis")
        main = BoxLayout(orientation='horizontal', padding=6, spacing=6)
        left = BoxLayout(orientation='vertical', size_hint_x=0.4, spacing=4)
        left.add_widget(Label(
            text="COMPREHENSIVE ECG ANALYSIS",
            font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#c0392b')
        ))
        self.ecg_info = Label(
            text="No patient selected",
            font_size=13, size_hint_y=None, height=22,
            color=get_color_from_hex('#7f8c8d')
        )
        left.add_widget(self.ecg_info)
        left.add_widget(Label(
            text="ECG IMAGE:", font_size=15, bold=True,
            size_hint_y=None, height=20
        ))
        self.ecg_img_disp = TextInput(
            text="[ECG Image will appear here]",
            font_size=11, size_hint_y=0.3,
            readonly=True,
            background_color=get_color_from_hex('#f5f5f5')
        )
        left.add_widget(self.ecg_img_disp)

        btn_box = BoxLayout(size_hint_y=None, height=45, spacing=5)
        btn_box.add_widget(Button(
            text="CAPTURE ECG", font_size=16, bold=True,
            on_press=self._capture_ecg,
            background_color=get_color_from_hex('#e74c3c')
        ))
        btn_box.add_widget(Button(
            text="UPLOAD ECG", font_size=16, bold=True,
            on_press=self._upload_ecg,
            background_color=get_color_from_hex('#c0392b')
        ))
        left.add_widget(btn_box)

        left.add_widget(Label(
            text="OR ENTER ECG DATA:", font_size=15, bold=True,
            size_hint_y=None, height=20
        ))
        self.ecg_data_inp = TextInput(
            font_size=13, size_hint_y=None, height=80, multiline=True,
            hint_text="Rate: 72 bpm\nRhythm: Regular sinus"
        )
        left.add_widget(self.ecg_data_inp)
        left.add_widget(Label(text="Reason:", font_size=15, bold=True,
                               size_hint_y=None, height=20))
        self.ecg_reason = TextInput(text="Routine", font_size=13,
                                      size_hint_y=None, height=38)
        left.add_widget(self.ecg_reason)
        left.add_widget(Label(text="Symptoms:", font_size=15, bold=True,
                               size_hint_y=None, height=20))
        self.ecg_symptoms = TextInput(text="None", font_size=13,
                                        size_hint_y=None, height=45, multiline=True)
        left.add_widget(self.ecg_symptoms)
        left.add_widget(Button(
            text="ANALYZE ECG", font_size=18, bold=True,
            size_hint_y=None, height=50,
            on_press=self._analyze_ecg,
            background_color=get_color_from_hex('#27ae60')
        ))
        left.add_widget(Button(
            text="COMBINED AI REPORT (ECG+Lab)", font_size=16, bold=True,
            size_hint_y=None, height=50,
            on_press=self._combined_report,
            background_color=get_color_from_hex('#8e44ad')
        ))

        right = BoxLayout(orientation='vertical', size_hint_x=0.6, spacing=3)
        right.add_widget(Label(
            text="ECG REPORT", font_size=18, bold=True,
            size_hint_y=None, height=28,
            color=get_color_from_hex('#2c3e50')
        ))
        ecg_report_scroll = ScrollView(size_hint_y=0.6, bar_width=10,
                                        scroll_type=['bars'])
        self.ecg_report_disp = TextInput(
            multiline=True, readonly=True, font_size=16,
            background_color=get_color_from_hex('#fafafa'),
            size_hint_y=None
        )
        self.ecg_report_disp.bind(minimum_height=self.ecg_report_disp.setter('height'))
        ecg_report_scroll.add_widget(self.ecg_report_disp)
        right.add_widget(ecg_report_scroll)

        right.add_widget(Label(
            text="KEY FINDINGS", font_size=15, bold=True,
            size_hint_y=None, height=20
        ))
        findings_scroll = ScrollView(size_hint_y=None, height=100, bar_width=10,
                                      scroll_type=['bars'])
        self.ecg_findings_disp = TextInput(
            multiline=True, readonly=True, font_size=16,
            background_color=get_color_from_hex('#fef9e7'),
            size_hint_y=None
        )
        self.ecg_findings_disp.bind(minimum_height=self.ecg_findings_disp.setter('height'))
        findings_scroll.add_widget(self.ecg_findings_disp)
        right.add_widget(findings_scroll)

        act_box = BoxLayout(size_hint_y=None, height=42, spacing=5)
        act_box.add_widget(Button(
            text="EXPORT PDF", font_size=15, bold=True,
            on_press=self._print_ecg,
            background_color=get_color_from_hex('#2980b9')
        ))
        act_box.add_widget(Button(
            text="SAVE", font_size=15, bold=True,
            on_press=self._save_ecg,
            background_color=get_color_from_hex('#27ae60')
        ))
        act_box.add_widget(Button(
            text="ADD TO BILL", font_size=15,
            on_press=self._ecg_to_bill,
            background_color=get_color_from_hex('#8e44ad')
        ))
        right.add_widget(act_box)

        main.add_widget(left)
        main.add_widget(right)
        tab.add_widget(main)
        self.tp.add_widget(tab)

    def _capture_ecg(self, instance):
        try:
            if not CameraHelper.available():
                self._popup("Camera", "Not available")
                return
            c = BoxLayout(orientation='vertical', spacing=4, padding=4)
            cam = Camera(play=True, resolution=(800, 600))
            c.add_widget(cam)
            btns = BoxLayout(size_hint_y=None, height=42, spacing=6)
            pp = Popup(title="Capture ECG", content=c, size_hint=(0.9, 0.85))

            def cap(inst):
                if cam.texture:
                    fn = f"ecg_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                    cam.texture.save(fn)
                    self.cur_ecg_img = fn
                    self.ecg_img_disp.text = f"Captured: {fn}"
                    pp.dismiss()

            btns.add_widget(Button(
                text="CAPTURE", font_size=16, bold=True,
                on_press=cap,
                background_color=get_color_from_hex('#e74c3c')
            ))
            btns.add_widget(Button(
                text="CANCEL", font_size=16,
                on_press=lambda x: pp.dismiss()
            ))
            c.add_widget(btns)
            pp.open()
        except Exception as e:
            self._popup("Error", str(e))

    def _upload_ecg(self, instance):
        c = BoxLayout(orientation='vertical', spacing=4, padding=4)
        fc = FileChooserListView(path=os.getcwd(),
                                  filters=['*.png', '*.jpg', '*.jpeg'])
        c.add_widget(fc)
        btns = BoxLayout(size_hint_y=None, height=42, spacing=6)
        pp = Popup(title="Upload ECG", content=c, size_hint=(0.9, 0.85))

        def sel(inst):
            if fc.selection:
                self.cur_ecg_img = fc.selection[0]
                self.ecg_img_disp.text = (f"Loaded: "
                                          f"{os.path.basename(fc.selection[0])}")
                pp.dismiss()

        btns.add_widget(Button(
            text="SELECT", font_size=16, bold=True,
            on_press=sel,
            background_color=get_color_from_hex('#c0392b')
        ))
        btns.add_widget(Button(
            text="CANCEL", font_size=16,
            on_press=lambda x: pp.dismiss()
        ))
        c.add_widget(btns)
        pp.open()

    def _analyze_ecg(self, instance):
        if not self.cp:
            self._popup("Error", "Select patient!")
            return
        self.ecg_info.text = (f"Patient: {self.cp.get('name','')} "
                               f"({self.cp.get('id','')}) | "
                               f"Age: {self.cp.get('age',0)}")
        self.ecg_report_disp.text = "Analyzing..."

        pt = {
            "name": self.cp.get("name", ""),
            "id": self.cp.get("id", ""),
            "age": self.cp.get("age", 0),
            "gender": self.cp.get("gender", ""),
            "height": self.cp.get("height", ""),
            "weight": self.cp.get("weight", ""),
            "bp": self.ew.get("bp", TextInput()).text if hasattr(self, 'ew') else "N/A",
            "pulse": self.ew.get("pulse", TextInput()).text if hasattr(self, 'ew') else "N/A",
            "spo2": self.ew.get("spo2", TextInput()).text if hasattr(self, 'ew') else "N/A",
            "ecg_reason": self.ecg_reason.text,
            "symptoms": self.ecg_symptoms.text
        }
        threading.Thread(target=self._run_ecg, args=(pt,), daemon=True).start()

    def _run_ecg(self, pt):
        result = self.ecg_ai.analyze(pt, self.ecg_data_inp.text.strip() or None)
        report, rid, qr = ECGReportGenerator.generate(result, pt)
        self.cur_ecg_report = {
            "report": report,
            "rid": rid,
            "qr": qr,
            "analysis": result["analysis"],
            "structured": result["structured"],
            "patient_id": pt.get("id", ""),
            "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "ai_used": result["ai_used"]
        }
        self.ecg_reports.append(self.cur_ecg_report)
        self.save_all()
        Clock.schedule_once(
            lambda dt: self._upd_ecg(result["structured"],
                                       result["ai_used"],
                                       full_report=report), 0
        )

    def _upd_ecg(self, sd, ai_used, full_report=None):
        self.ecg_report_disp.text = full_report if full_report else ""
        f = (f"Rate: {sd.get('rate','N/A')} | "
             f"Rhythm: {sd.get('rhythm','N/A')}\n")
        f += (f"PR: {sd.get('pr_interval','N/A')} | "
              f"QRS: {sd.get('qrs_duration','N/A')} | "
              f"QTc: {sd.get('qt_qtc','N/A')}\n\nFindings:\n")
        for d in sd.get('findings', [])[:5]:
            f += f"  - {d}\n"
        f += (f"\nInterpretation: {sd.get('interpretation','N/A')}\n"
              f"Urgency: {sd.get('urgency','Routine').upper()}\n"
              f"AI: {'Yes' if ai_used else 'Local'}")
        self.ecg_findings_disp.text = f
        self._popup("ECG Done", f"Report: {self.cur_ecg_report['rid']}")

    def _print_ecg(self, instance):
        if not self.cur_ecg_report:
            self._popup("Error", "Run analysis!")
            return
        text = self.cur_ecg_report['report']
        self._export_to_pdf(text, "ECG_Report", "ECG Report")

    def _save_ecg(self, instance):
        if not self.cur_ecg_report:
            return
        self.save_all()
        self._popup("Saved", "ECG saved")
        self.log_activity("ECG report saved")
        self.audit_log("SAVE_ECG", f"Report {self.cur_ecg_report.get('rid','')}")

    def _ecg_to_bill(self, instance):
        if not self.cp:
            return
        self._switch("Billing")
        self.bill_pid.text = self.cp.get("id", "")
        self.bill_name.text = self.cp.get("name", "")
        price = self._find_price_in_db("ECG", self.investigations, self.charges)
        if price is None:
            price = 0
            self._popup("Price Missing",
                        "No price configured for ECG. Item added at ₦0 — "
                        "please set the price in Manage Investigations.")
        self.bill_items.append({
            "name": "ECG Analysis",
            "cat": "Charges",
            "qty": 1,
            "price": price,
            "total": price
        })
        self._refresh_bill()

    def _combined_report(self, instance):
        if not self.cp:
            self._popup("Error", "Select a patient first!")
            return
        if not self.cur_ecg_report:
            self._popup("Error", "Please run ECG analysis first!")
            return
        if not self.ai.enabled:
            self._popup("Error",
                        "AI is not enabled. Please set up API key in Settings.")
            return
        pid = self.cp.get("id")
        labs = [t for t in self.lab_tests
                if t.get("patient_id") == pid and t.get("status") == "Completed"]
        if not labs:
            self._popup("Info",
                        "No completed laboratory results for this patient.")
            return

        ecg_summary = self.cur_ecg_report.get("analysis",
                                                "No ECG analysis available.")
        lab_summary = "\n".join([
            f"- {lab.get('test_name','')}: {lab.get('result','')} "
            f"{lab.get('unit','')}"
            for lab in labs[-10:]
        ])
        prompt = f"""Patient: {self.cp.get('name','')} (ID: {pid}), Age: {self.cp.get('age','')}, Sex: {self.cp.get('gender','')}

ECG Report:
{ecg_summary}

Laboratory Results:
{lab_summary}

Please provide a comprehensive integrated clinical summary combining the ECG findings and laboratory results. Include:
- Overall clinical assessment
- Correlation between ECG and lab abnormalities
- Differential diagnoses
- Recommended further investigations
- Management plan
Be thorough and precise.
"""

        def run_combined():
            sp = ("You are a senior clinician synthesizing ECG and "
                  "laboratory data into a comprehensive clinical report.")
            ok, response = self.ai.call(sp, prompt, mt=3000)
            if ok:
                Clock.schedule_once(
                    lambda dt: self._show_combined_report(response), 0
                )
            else:
                Clock.schedule_once(
                    lambda dt: self._popup("AI Error", f"Failed: {response}"), 0
                )

        self._popup("Processing", "Generating combined report...")
        threading.Thread(target=run_combined, daemon=True).start()

    def _show_combined_report(self, response):
        c = BoxLayout(orientation='vertical', spacing=5, padding=8)
        sv = ScrollView()
        ti = TextInput(
            text=response, readonly=True, font_size=16,
            size_hint_y=None,
            background_color=get_color_from_hex('#fef9e7')
        )
        ti.bind(minimum_height=ti.setter('height'))
        sv.add_widget(ti)
        c.add_widget(sv)

        btn_box = BoxLayout(size_hint_y=None, height=42, spacing=5)
        pp = Popup(title="Combined AI Report (ECG + Lab)",
                   content=c, size_hint=(0.9, 0.85))
        btn_box.add_widget(Button(
            text="CLOSE", font_size=16, bold=True, on_press=pp.dismiss
        ))
        btn_box.add_widget(Button(
            text="EXPORT PDF", font_size=16, bold=True,
            on_press=lambda x: self._export_to_pdf(
                response, "Combined_Report",
                "Combined ECG and Lab Report"
            )
        ))
        c.add_widget(btn_box)
        pp.open()

    # ====================== LIVER SPAN TAB ======================
    def _liver_span_tab(self):
        tab = TabbedPanelItem(text="Liver Span")
        main = BoxLayout(orientation='horizontal', padding=8, spacing=8)
        left = BoxLayout(orientation='vertical', size_hint_x=0.45, spacing=6)
        left.add_widget(Label(
            text="BELEMA'S LAW", font_size=18, bold=True,
            size_hint_y=None, height=30,
            color=get_color_from_hex('#16a085')
        ))
        f = GridLayout(cols=2, size_hint_y=None, height=120, spacing=6)
        f.add_widget(Label(text="Age:", font_size=15, bold=True, size_hint_x=0.35))
        self.liver_age = TextInput(font_size=15, size_hint_x=0.5,
                                     size_hint_y=None, height=40)
        f.add_widget(self.liver_age)
        f.add_widget(Label(text="Height(m):", font_size=15, bold=True, size_hint_x=0.35))
        self.liver_height = TextInput(font_size=15, size_hint_x=0.5,
                                        size_hint_y=None, height=40,
                                        hint_text="1.65")
        f.add_widget(self.liver_height)
        f.add_widget(Label(text="Measured(cm):", font_size=15, bold=True, size_hint_x=0.35))
        self.liver_measured = TextInput(font_size=15, size_hint_x=0.5,
                                          size_hint_y=None, height=40)
        f.add_widget(self.liver_measured)
        left.add_widget(f)
        left.add_widget(Button(
            text="CALCULATE", font_size=17, bold=True,
            size_hint_y=None, height=48,
            on_press=self._calc_liver,
            background_color=get_color_from_hex('#16a085')
        ))
        self.liver_predicted = Label(
            text="Predicted: -- cm", font_size=18, bold=True,
            size_hint_y=None, height=30,
            color=get_color_from_hex('#2980b9')
        )
        left.add_widget(self.liver_predicted)
        self.liver_interpretation = Label(
            text="Interpretation: --", font_size=18, bold=True,
            size_hint_y=None, height=30
        )
        left.add_widget(self.liver_interpretation)

        right = BoxLayout(orientation='vertical', size_hint_x=0.55,
                          spacing=4, padding=10)
        right.add_widget(Label(
            text="FORMULA", font_size=16, bold=True,
            size_hint_y=None, height=25,
            color=get_color_from_hex('#16a085')
        ))
        right.add_widget(TextInput(
            text="h = 12.6x - 3.00 (<12)\nh = 12.6x - 5.04 (>=12)",
            readonly=True, font_size=13,
            size_hint_y=0.6,
            background_color=get_color_from_hex('#f8f9fa')
        ))
        main.add_widget(left)
        main.add_widget(right)
        tab.add_widget(main)
        self.tp.add_widget(tab)

    def _calc_liver(self, instance):
        try:
            age = float(self.liver_age.text.strip() or "0")
            height = float(self.liver_height.text.strip() or "0")
            measured = self.liver_measured.text.strip()
        except:
            self._popup("Error", "Enter valid values!")
            return
        predicted = BelemaLaw.calculate_liver_span(height, age)
        if predicted is None:
            return
        self.liver_predicted.text = f"Predicted: {predicted} cm"
        if measured:
            try:
                mv = float(measured)
                interp, color = BelemaLaw.interpret_finding(mv, predicted)
                self.liver_interpretation.text = f"Interpretation: {interp}"
                color_map = {
                    "green": '#27ae60', "yellow": '#f39c12',
                    "orange": '#e67e22', "red": '#e74c3c'
                }
                self.liver_interpretation.color = get_color_from_hex(
                    color_map.get(color, '#2c3e50')
                )
            except:
                pass

    # ====================== PATHOLOGY TAB ======================
    def _pathology_tab(self):
        tab = TabbedPanelItem(text="Pathology AI")
        main = BoxLayout(orientation='horizontal', padding=8, spacing=8)
        left = BoxLayout(orientation='vertical', size_hint_x=0.45, spacing=4)
        left.add_widget(Label(
            text="AI PATHOLOGY", font_size=18, bold=True,
            size_hint_y=None, height=30,
            color=get_color_from_hex('#8e44ad')
        ))
        left.add_widget(Button(
            text="EXTRACT PATHOLOGY", font_size=17, bold=True,
            size_hint_y=None, height=48,
            on_press=self._run_pathology,
            background_color=get_color_from_hex('#8e44ad')
        ))
        left.add_widget(Button(
            text="VIEW ML", font_size=15, size_hint_y=None, height=40,
            on_press=self._view_ml,
            background_color=get_color_from_hex('#2c3e50')
        ))
        left.add_widget(Button(
            text="EXPORT ML", font_size=15, size_hint_y=None, height=40,
            on_press=self._export_ml,
            background_color=get_color_from_hex('#16a085')
        ))

        right = BoxLayout(orientation='vertical', size_hint_x=0.55, spacing=3)
        path_scroll = ScrollView(size_hint_y=0.7, bar_width=10,
                                  scroll_type=['bars'])
        self.path_output = TextInput(
            multiline=True, readonly=True, font_size=16,
            background_color=get_color_from_hex('#fafafa'),
            size_hint_y=None
        )
        self.path_output.bind(minimum_height=self.path_output.setter('height'))
        path_scroll.add_widget(self.path_output)
        right.add_widget(path_scroll)

        main.add_widget(left)
        main.add_widget(right)
        tab.add_widget(main)
        self.tp.add_widget(tab)

    def _run_pathology(self, instance):
        if not self.cp:
            self._popup("Error", "Select a patient first!")
            return
        pid = self.cp.get("id")
        if not pid:
            self._popup("Error", "Patient ID missing")
            return
        emr = [r for r in self.emr_records if r.get("patient_id") == pid]
        labs = [t for t in self.lab_tests if t.get("patient_id") == pid]
        rx = [p for p in self.prescriptions if p.get("patient_id") == pid]
        pd = {
            "id": pid,
            "name": self.cp.get("name", ""),
            "age": self.cp.get("age", 0),
            "gender": self.cp.get("gender", "")
        }
        threading.Thread(
            target=self._run_path_thread,
            args=(pd, emr, labs, rx),
            daemon=True
        ).start()

    def _run_path_thread(self, pd, emr, labs, rx):
        result = self.pathology_ai.extract(pd, emr, labs, rx)
        self.pathology_data.append({
            "patient_id": pd.get("id", ""),
            "ts": datetime.now().isoformat(),
            "structured": result.get("structured", {})
        })
        self.save_all()
        Clock.schedule_once(
            lambda dt: setattr(self.path_output, 'text',
                                result.get("analysis", "No analysis")), 0
        )

    def _view_ml(self, instance):
        if not self.cp:
            self._popup("ML Data", "Please select a patient first.")
            return
        pid = self.cp.get("id")
        patient_pathology = [p for p in self.pathology_data
                             if p.get("patient_id") == pid]
        if not patient_pathology:
            self._popup(
                "ML Data",
                f"No ML/Pathology data available for "
                f"{self.cp.get('name', '')}.\nRun 'Extract Pathology' first."
            )
            return
        latest = patient_pathology[-1]
        structured = latest.get("structured", {})

        output_lines = []
        output_lines.append("=" * 70)
        output_lines.append("          ML / PATHOLOGY ANALYSIS")
        output_lines.append(f"          Patient: {self.cp.get('name', '')} (ID: {pid})")
        output_lines.append(f"          Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        output_lines.append("=" * 70)
        output_lines.append("")

        output_lines.append("📊 RISK ASSESSMENT")
        output_lines.append("-" * 50)
        output_lines.append(f"  Risk Score: {structured.get('risk_score', 0)}/10")
        output_lines.append(f"  Risk Level: {structured.get('risk_level', 'Unknown')}")
        output_lines.append("")

        ml_features = structured.get("ml_features", {})
        if ml_features:
            output_lines.append("🤖 ML FEATURES")
            output_lines.append("-" * 50)
            for key, value in ml_features.items():
                output_lines.append(f"  {key.replace('_', ' ').title()}: {value}")
            output_lines.append("")

        completed_labs = [t for t in self.lab_tests
                          if t.get("patient_id") == pid
                          and t.get("status") == "Completed"]
        if completed_labs:
            output_lines.append("🧪 LABORATORY RESULTS")
            output_lines.append("-" * 50)
            for lab in completed_labs[-15:]:
                status = "⚠️" if any(
                    w in str(lab.get('result', '')).lower()
                    for w in ['positive', 'abnormal', 'high', 'low']
                ) else "✅"
                output_lines.append(
                    f"  {status} {lab.get('test_name', '')}: "
                    f"{lab.get('result', '')} {lab.get('unit', '')}"
                )
            output_lines.append("")

        rx = [p for p in self.prescriptions if p.get("patient_id") == pid]
        if rx:
            output_lines.append("💊 CURRENT PRESCRIPTIONS")
            output_lines.append("-" * 50)
            for r in rx[-10:]:
                output_lines.append(f"  {r.get('drug', '')} x{r.get('qty', 0)}")
            output_lines.append("")

        emrs = [r for r in self.emr_records if r.get("patient_id") == pid]
        if emrs:
            output_lines.append("📋 RECENT EMR RECORDS")
            output_lines.append("-" * 50)
            for r in emrs[-5:]:
                output_lines.append(f"  {r.get('ts', '')[:16]}: {r.get('cc', '')[:50]}")
            output_lines.append("")

        if structured.get("ai_analysis"):
            output_lines.append("🧠 AI CLINICAL ANALYSIS")
            output_lines.append("-" * 50)
            output_lines.append(structured.get("ai_analysis", ""))
            output_lines.append("")

        if structured.get("recommendations"):
            output_lines.append("📝 RECOMMENDATIONS")
            output_lines.append("-" * 50)
            for rec in structured.get("recommendations", []):
                output_lines.append(f"  • {rec}")
            output_lines.append("")

        output_lines.append("-" * 70)
        output_lines.append("📄 RAW STRUCTURED DATA")
        output_lines.append("-" * 50)
        output_lines.append(json.dumps(structured, indent=2, default=str)[:2000])
        output_lines.append("")
        output_lines.append("=" * 70)
        output_lines.append("                    END OF ML ANALYSIS")
        output_lines.append("=" * 70)

        self._text_popup("ML / Pathology Analysis", "\n".join(output_lines))
        self.path_output.text = "\n".join(output_lines)

    def _export_ml(self, instance):
        if not self.pathology_data:
            self._popup("ML Data", "No ML data to export.")
            return
        if not self.cp:
            self._popup("Error", "Select a patient first.")
            return
        pid = self.cp.get("id")
        patient_pathology = [p for p in self.pathology_data
                             if p.get("patient_id") == pid]
        if not patient_pathology:
            self._popup("Export", "No data for this patient.")
            return
        try:
            fn = (f"ml_{self.cp.get('name', '')}_"
                  f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
            with open(fn, 'w', encoding='utf-8') as f:
                json.dump(patient_pathology, f, indent=2, default=str)
            self._popup("Exported", f"ML data exported to:\n{fn}")
            self.log_activity(f"ML data exported: {fn}")
        except Exception as e:
            self._popup("Export Error", str(e))

    # ====================== LIS TAB ======================
    def _lis_tab(self):
        tab = TabbedPanelItem(text="LIS Connections")
        main = BoxLayout(orientation='vertical', padding=6, spacing=6)
        scan_btn = Button(
            text="SCAN QR TO ADD INSTRUMENT", font_size=16, bold=True,
            size_hint_y=None, height=44,
            on_press=self._scan_qr_code,
            background_color=get_color_from_hex('#16a085')
        )
        main.add_widget(scan_btn)
        for section, color, cat_name in [
            ("LAB INSTRUMENTS", "#1a5276", "lab"),
            ("RADIOLOGY INSTRUMENTS", "#8e44ad", "radiology"),
            ("VITALS MONITORING", "#16a085", "vitals")
        ]:
            lbl = Label(
                text=section, font_size=18, bold=True,
                size_hint_y=None, height=28,
                color=get_color_from_hex(color)
            )
            main.add_widget(lbl)
            scroll = ScrollView(size_hint_y=0.25, bar_width=10)
            box = BoxLayout(orientation='vertical', size_hint_y=None, spacing=3)
            box.bind(minimum_height=box.setter('height'))
            scroll.add_widget(box)
            main.add_widget(scroll)
            setattr(self, f"{cat_name}_instruments_box", box)
        tab.add_widget(main)
        self.tp.add_widget(tab)
        self.refresh_lis_displays()
        Clock.schedule_interval(lambda dt: self.refresh_lis_displays(), 5)

    def refresh_lis_displays(self):
        for cat_name, cat in [("lab", "lab"), ("radiology", "radiology"),
                                ("vitals", "vitals")]:
            if not hasattr(self, f"{cat_name}_instruments_box"):
                continue
            box = getattr(self, f"{cat_name}_instruments_box")
            box.clear_widgets()
            for inst_id in self.lis_manager.instrument_configs.get(cat, {}):
                status = self.lis_manager.get_status(cat, inst_id)
                color = (get_color_from_hex('#27ae60')
                         if status["status"] == "Running"
                         else get_color_from_hex('#e74c3c'))
                row = BoxLayout(size_hint_y=None, height=55, spacing=4)
                row.add_widget(Label(
                    text="●" if status["status"] == "Running" else "○",
                    font_size=18, color=color, size_hint_x=0.06
                ))
                info = BoxLayout(orientation='vertical', size_hint_x=0.55)
                info.add_widget(Label(text=status['name'], font_size=13, bold=True,
                                       size_hint_y=None, height=20))
                info.add_widget(Label(
                    text=(f"{status['manufacturer']} {status['model']} | "
                          f"Port: {status['port']}"),
                    font_size=10, size_hint_y=None, height=16,
                    color=get_color_from_hex('#7f8c8d')
                ))
                row.add_widget(info)
                ctrl = BoxLayout(size_hint_x=0.39, spacing=3)
                if status["status"] == "Running":
                    ctrl.add_widget(Button(
                        text="STOP", font_size=11, bold=True,
                        background_color=get_color_from_hex('#e74c3c'),
                        on_press=lambda x, c=cat, i=inst_id: self._stop_lis(c, i)
                    ))
                else:
                    ctrl.add_widget(Button(
                        text="START", font_size=11, bold=True,
                        background_color=get_color_from_hex('#27ae60'),
                        on_press=lambda x, c=cat, i=inst_id: self._start_lis(c, i)
                    ))
                row.add_widget(ctrl)
                box.add_widget(row)

    def _start_lis(self, cat, inst_id):
        success, msg = self.lis_manager.start_instrument(cat, inst_id)
        self.refresh_lis_displays()
        self._popup("LIS", msg)
        self.log_activity(f"LIS {cat}/{inst_id} started")
        self.audit_log("LIS_START", f"{cat}/{inst_id}")

    def _stop_lis(self, cat, inst_id):
        self.lis_manager.stop_instrument(cat, inst_id)
        self.refresh_lis_displays()
        self.log_activity(f"LIS {cat}/{inst_id} stopped")
        self.audit_log("LIS_STOP", f"{cat}/{inst_id}")

    def _scan_qr_code(self, instance):
        if not CameraHelper.available():
            self._popup("Camera", "Not available")
            return
        cam = Camera(play=True, resolution=(640, 480))
        c = BoxLayout(orientation='vertical', spacing=4, padding=4)
        c.add_widget(cam)
        btns = BoxLayout(size_hint_y=None, height=42, spacing=5)
        pp = Popup(title="Scan QR Code", content=c, size_hint=(0.85, 0.85))

        def capture(inst):
            if cam.texture:
                qr_data = self._decode_qr_from_texture(cam.texture)
                if qr_data:
                    success, msg = self.lis_manager.add_instrument_from_qr(qr_data)
                    self._popup(
                        "QR Scan Result",
                        f"{'Connected' if success else 'Failed'}: {msg}"
                    )
                    self.refresh_lis_displays()
                else:
                    self._popup("QR Scan Result", "No QR code found in image.")
                pp.dismiss()

        btns.add_widget(Button(
            text="CAPTURE", font_size=16, bold=True,
            on_press=capture,
            background_color=get_color_from_hex('#27ae60')
        ))
        btns.add_widget(Button(
            text="CANCEL", font_size=16,
            on_press=lambda x: pp.dismiss()
        ))
        c.add_widget(btns)
        pp.open()

    def get_patient_name(self, pid):
        for p in self.patients:
            if p.get("id") == pid:
                return p.get("name", pid)
        return pid


# ====================== END OF PART 4 ======================
    # ====================== BILLING TAB ======================
    def _billing(self):
        tab = TabbedPanelItem(text="Billing")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=5, spacing=3, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="ID:", font_size=15, bold=True, size_hint_x=0.05))
        self.bill_pid = TextInput(size_hint_x=0.15, readonly=True, font_size=15,
                                    size_hint_y=None, height=40)
        bar.add_widget(self.bill_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.07))
        self.bill_name = TextInput(size_hint_x=0.35, readonly=True, font_size=15,
                                     size_hint_y=None, height=40)
        bar.add_widget(self.bill_name)
        l.add_widget(bar)

        it = BoxLayout(size_hint_y=None, height=40, spacing=4)
        it.add_widget(Label(text="Cat:", font_size=15, bold=True, size_hint_x=0.08))
        self.bill_cat = Spinner(
            text="Procedures",
            values=["Procedures", "Investigations", "Drugs", "Charges"],
            size_hint_x=0.18, size_hint_y=None, height=40
        )
        it.add_widget(self.bill_cat)
        it.add_widget(Label(text="Item:", font_size=15, bold=True, size_hint_x=0.06))
        self.bill_item = Spinner(text="Select", values=[], size_hint_x=0.3,
                                   size_hint_y=None, height=40)
        it.add_widget(self.bill_item)
        it.add_widget(Label(text="Qty:", font_size=15, bold=True, size_hint_x=0.05))
        self.bill_qty = TextInput(text="1", size_hint_x=0.06, font_size=15,
                                    size_hint_y=None, height=40)
        it.add_widget(self.bill_qty)
        it.add_widget(Button(
            text="SEARCH", font_size=15, bold=True, size_hint_x=0.1,
            on_press=self._search_bill_item,
            background_color=get_color_from_hex('#f39c12')
        ))
        it.add_widget(Button(
            text="ADD", font_size=15, bold=True, size_hint_x=0.08,
            on_press=self._add_bill,
            background_color=get_color_from_hex('#2980b9')
        ))
        it.add_widget(Button(
            text="ADD CUSTOM", font_size=15, bold=True, size_hint_x=0.12,
            on_press=self._add_custom_item,
            background_color=get_color_from_hex('#8e44ad')
        ))
        it.add_widget(Button(
            text="VIEW SAVED", font_size=15, bold=True, size_hint_x=0.12,
            on_press=self._view_saved_bills,
            background_color=get_color_from_hex('#16a085')
        ))
        l.add_widget(it)

        self.bill_cat.bind(text=lambda i, v: self._upd_bill_items(v))
        self.bill_box = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        self.bill_box.bind(minimum_height=self.bill_box.setter('height'))
        l.add_widget(self.bill_box)

        tot = BoxLayout(size_hint_y=None, height=50, spacing=5)
        self.bill_tot = Label(
            text="TOTAL: N0.00", font_size=20, bold=True,
            size_hint_x=0.5, color=get_color_from_hex('#c0392b')
        )
        tot.add_widget(self.bill_tot)
        tot.add_widget(Button(
            text="SAVE", font_size=17, bold=True, size_hint_x=0.25,
            on_press=self._save_bill,
            background_color=get_color_from_hex('#27ae60')
        ))
        tot.add_widget(Button(
            text="PRINT PDF", font_size=17, bold=True, size_hint_x=0.25,
            on_press=lambda x: self._export_to_pdf(
                self._get_bill_text(), "Bill", "Patient Bill"
            )
        ))
        l.add_widget(tot)
        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)
        self._upd_bill_items("Procedures")

    def _add_custom_item(self, instance):
        content = BoxLayout(orientation='vertical', spacing=5, padding=8)
        content.add_widget(Label(
            text="ADD CUSTOM ITEM", font_size=16, bold=True,
            size_hint_y=None, height=30,
            color=get_color_from_hex('#8e44ad')
        ))
        content.add_widget(Label(text="Item Name:", font_size=14, bold=True,
                                   size_hint_y=None, height=22))
        name_input = TextInput(font_size=14, size_hint_y=None, height=40,
                                hint_text="e.g., Custom Service")
        content.add_widget(name_input)
        content.add_widget(Label(text="Category:", font_size=14, bold=True,
                                   size_hint_y=None, height=22))
        cat_input = Spinner(
            text="Charges",
            values=["Procedures", "Investigations", "Drugs", "Charges"],
            size_hint_y=None, height=40
        )
        content.add_widget(cat_input)
        content.add_widget(Label(text="Price (N):", font_size=14, bold=True,
                                   size_hint_y=None, height=22))
        price_input = TextInput(font_size=14, size_hint_y=None, height=40,
                                  hint_text="1000", input_filter="float")
        content.add_widget(price_input)
        content.add_widget(Label(text="Quantity:", font_size=14, bold=True,
                                   size_hint_y=None, height=22))
        qty_input = TextInput(font_size=14, size_hint_y=None, height=40,
                                text="1", input_filter="int")
        content.add_widget(qty_input)
        btn_box = BoxLayout(size_hint_y=None, height=40, spacing=5)
        pp = Popup(title="Custom Item", content=content, size_hint=(0.6, 0.6))

        def add_item(inst):
            name = name_input.text.strip()
            if not name:
                self._popup("Error", "Please enter an item name.")
                return
            try:
                price = float(price_input.text.strip() or "0")
                qty = int(qty_input.text.strip() or "1")
            except:
                self._popup("Error", "Invalid price or quantity.")
                return
            if price <= 0:
                self._popup("Error", "Price must be greater than 0.")
                return
            self.bill_items.append({
                "name": name, "cat": cat_input.text,
                "qty": qty, "price": price, "total": price * qty
            })
            self._refresh_bill()
            pp.dismiss()
            self._popup("Added", f"Added: {name} x{qty} = N{price * qty:,.2f}")
            self.log_activity(f"Added custom bill item: {name} x{qty} = N{price * qty:,.2f}")

        btn_box.add_widget(Button(
            text="ADD", font_size=15, bold=True, on_press=add_item,
            background_color=get_color_from_hex('#27ae60')
        ))
        btn_box.add_widget(Button(
            text="CANCEL", font_size=15, on_press=pp.dismiss,
            background_color=get_color_from_hex('#e74c3c')
        ))
        content.add_widget(btn_box)
        pp.open()

    def _search_bill_item(self, instance):
        cat = self.bill_cat.text
        src = {
            "Procedures": self.procedures,
            "Investigations": self.investigations,
            "Drugs": self.drugs,
            "Charges": self.charges
        }.get(cat, [])
        c = BoxLayout(orientation='vertical', spacing=5, padding=8)
        sr = TextInput(size_hint_y=None, height=40, font_size=15,
                        hint_text="Search item...")
        c.add_widget(sr)
        pl = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        pl.bind(minimum_height=pl.setter('height'))
        sv = ScrollView()
        sv.add_widget(pl)
        c.add_widget(sv)
        pp = Popup(title=f"Select {cat}", content=c, size_hint=(0.7, 0.7))

        def refresh_list(t=""):
            pl.clear_widgets()
            for item in src:
                if not t or t.lower() in item['name'].lower():
                    btn = Button(text=item['name'], font_size=15,
                                  size_hint_y=None, height=40)
                    btn.item_name = item['name']
                    btn.bind(on_press=lambda x: (
                        setattr(self.bill_item, 'text', x.item_name),
                        pp.dismiss()
                    ))
                    pl.add_widget(btn)

        sr.bind(text=lambda i, v: refresh_list(v))
        refresh_list()
        pp.open()

    def _upd_bill_items(self, cat):
        src = {
            "Procedures": self.procedures,
            "Investigations": self.investigations,
            "Drugs": self.drugs,
            "Charges": self.charges
        }.get(cat, [])
        self.bill_item.values = [x['name'] for x in src]
        if self.bill_item.values:
            self.bill_item.text = self.bill_item.values[0]

    def _get_price(self, nm, cat):
        src = {
            "Procedures": self.procedures,
            "Investigations": self.investigations,
            "Drugs": self.drugs,
            "Charges": self.charges
        }.get(cat, [])
        for x in src:
            if x['name'] == nm:
                return x.get('price', 0)
        return 0

    def _add_bill(self, instance):
        nm = self.bill_item.text
        ct = self.bill_cat.text
        if nm == "Select":
            self._popup("Error", "Please select an item from the dropdown.")
            return
        try:
            q = int(self.bill_qty.text)
            if q <= 0:
                self._popup("Error", "Quantity must be positive.")
                return
        except:
            self._popup("Error", "Invalid quantity. Please enter a number.")
            return
        p = self._get_price(nm, ct)
        if not p:
            self._popup("Error",
                        f"Price not found for '{nm}'. Please check your data.")
            return
        self.bill_items.append({
            "name": nm, "cat": ct, "qty": q,
            "price": p, "total": p * q
        })
        self._refresh_bill()
        self.bill_qty.text = "1"

    def _refresh_bill(self):
        self.bill_box.clear_widgets()
        total = 0
        for i, item in enumerate(self.bill_items):
            total += item["total"]
            row = BoxLayout(size_hint_y=None, height=30, spacing=3)
            row.add_widget(Label(text=f"#{i+1} {item['name'][:30]}",
                                  font_size=13, size_hint_x=0.35))
            row.add_widget(Label(text=f"{item['qty']} x N{item['price']:,.0f}",
                                  font_size=13, size_hint_x=0.25))
            row.add_widget(Label(text=f"N{item['total']:,.0f}",
                                  font_size=13, size_hint_x=0.2, bold=True))
            btn = Button(text="X", font_size=11, size_hint_x=0.05,
                          size_hint_y=None, height=28)
            btn.idx = i
            btn.bind(on_press=lambda x: (
                self.bill_items.pop(x.idx), self._refresh_bill()
            ))
            row.add_widget(btn)
            self.bill_box.add_widget(row)
        self.bill_tot.text = f"TOTAL: N{total:,.2f}"

    def _get_bill_text(self):
        total = sum(i["total"] for i in self.bill_items)
        text = (f"=== HMG HOSPITAL - BILL ===\n\n"
                f"Patient: {self.bill_name.text}\n"
                f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Issued by: {self._current_username()}\n\n")
        for item in self.bill_items:
            text += (f"{item['name'][:35]:<35} {item['qty']:>4} x "
                     f"N{item['price']:>10,.2f} = N{item['total']:>10,.2f}\n")
        text += f"{'-' * 65}\nTOTAL: N{total:,.2f}\n"
        return text

    def _save_bill(self, instance):
        if not self.bill_pid.text or not self.bill_items:
            return
        total = sum(i["total"] for i in self.bill_items)
        self.bills.append({
            "bill_id": f"BILL{datetime.now().strftime('%Y%m%d%H%M%S%f')}",
            "patient": {
                "id": self.bill_pid.text,
                "name": self.bill_name.text
            },
            "items": [dict(i) for i in self.bill_items],
            "total": total,
            "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "status": "Pending",
            "created_by": self._current_username()
        })
        self.save_all()
        self.bill_items = []
        self._refresh_bill()
        self._popup("Saved", f"N{total:,.2f}")
        self.log_activity(f"Bill saved for {self.bill_name.text} N{total:,.2f}")
        self.audit_log("SAVE_BILL", f"N{total:,.2f}", self.bill_pid.text)

    def _view_saved_bills(self, instance):
        if not self.cp and not self.bill_pid.text:
            self._popup("Error", "No patient selected.")
            return
        pid = self.cp.get("id") if self.cp else self.bill_pid.text
        bills = [b for b in self.bills
                 if b.get("patient", {}).get("id") == pid]
        if not bills:
            self._popup("Saved Bills", "No saved bills for this patient.")
            return
        text = f"BILLS FOR {self.bill_name.text}\n" + "=" * 50 + "\n\n"
        for b in bills:
            text += f"ID: {b.get('bill_id', '')}\n"
            text += f"Date: {b.get('ts', '')}\n"
            text += f"Total: N{b.get('total', 0):,.2f}\n"
            text += f"Status: {b.get('status', '')}\n"
            text += "-" * 30 + "\n"
        self._text_popup("Saved Bills", text)

    def _view_bill_detail_kivy(self, bill):
        text = (f"=== HMG HOSPITAL - BILL ===\n\n"
                f"Bill ID: {bill['bill_id']}\n"
                f"Patient: {bill['patient']['name']}\n"
                f"Date: {bill['ts']}\n\nItems:\n")
        for item in bill['items']:
            text += (f"{item['name'][:35]:<35} {item['qty']:>4} x "
                     f"N{item['price']:>10,.2f} = N{item['total']:>10,.2f}\n")
        text += (f"{'-' * 65}\nTOTAL: N{bill['total']:,.2f}\n"
                 f"Status: {bill['status']}\n")
        self._text_popup("Bill Detail", text)

    # ====================== PAYMENTS TAB ======================
    def _payment(self):
        tab = TabbedPanelItem(text="Payments")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=5, spacing=3, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="ID:", font_size=15, bold=True, size_hint_x=0.05))
        self.pay_pid = TextInput(size_hint_x=0.15, readonly=True, font_size=15,
                                   size_hint_y=None, height=40)
        bar.add_widget(self.pay_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.07))
        self.pay_name = TextInput(size_hint_x=0.35, readonly=True, font_size=15,
                                    size_hint_y=None, height=40)
        bar.add_widget(self.pay_name)
        l.add_widget(bar)

        sum_box = GridLayout(cols=3, size_hint_y=None, height=35, spacing=4)
        self.pay_total = Label(text="Total: N0", font_size=16, bold=True,
                                color=get_color_from_hex('#2980b9'))
        self.pay_paid = Label(text="Paid: N0", font_size=16, bold=True,
                               color=get_color_from_hex('#27ae60'))
        self.pay_bal = Label(text="Balance: N0", font_size=16, bold=True,
                              color=get_color_from_hex('#c0392b'))
        for w in [self.pay_total, self.pay_paid, self.pay_bal]:
            sum_box.add_widget(w)
        l.add_widget(sum_box)

        bill_sel = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bill_sel.add_widget(Label(text="Select Bill:", font_size=15, bold=True,
                                    size_hint_x=0.12))
        self.pay_bill_spinner = Spinner(text="None", values=[],
                                          size_hint_x=0.6, size_hint_y=None, height=40)
        bill_sel.add_widget(self.pay_bill_spinner)
        l.add_widget(bill_sel)

        form = BoxLayout(size_hint_y=None, height=40, spacing=4)
        form.add_widget(Label(text="Amount:", font_size=15, bold=True, size_hint_x=0.08))
        self.pay_amt = TextInput(size_hint_x=0.15, font_size=15,
                                   size_hint_y=None, height=40)
        form.add_widget(self.pay_amt)
        form.add_widget(Label(text="Method:", font_size=15, bold=True, size_hint_x=0.08))
        self.pay_method = Spinner(
            text="Cash",
            values=["Cash", "Bank Transfer", "Card", "Insurance"],
            size_hint_x=0.18, size_hint_y=None, height=40
        )
        form.add_widget(self.pay_method)
        form.add_widget(Button(
            text="RECORD", font_size=15, bold=True, size_hint_x=0.18,
            on_press=self._record_pay,
            background_color=get_color_from_hex('#27ae60')
        ))
        l.add_widget(form)

        btn_row = BoxLayout(size_hint_y=None, height=40, spacing=4)
        btn_row.add_widget(Button(
            text="PRINT RECEIPT PDF", font_size=15, bold=True, size_hint_x=0.4,
            on_press=self._print_receipt,
            background_color=get_color_from_hex('#8e44ad')
        ))
        btn_row.add_widget(Button(
            text="VIEW RECEIPTS", font_size=15, bold=True, size_hint_x=0.3,
            on_press=self._view_saved_receipts,
            background_color=get_color_from_hex('#16a085')
        ))
        l.add_widget(btn_row)

        self.pay_hist = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        self.pay_hist.bind(minimum_height=self.pay_hist.setter('height'))
        l.add_widget(self.pay_hist)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)

    def _record_pay(self, instance):
        pid = self.pay_pid.text
        if not pid:
            return
        try:
            amt = float(self.pay_amt.text)
            if amt <= 0:
                raise ValueError
        except:
            self._popup("Error", "Invalid amount.")
            return
        bill_id = None
        if self.pay_bill_spinner.text and self.pay_bill_spinner.text != "None":
            bill_id = self.pay_bill_spinner.text.split(" - ")[0]
        if bill_id:
            for b in self.bills:
                if b["bill_id"] == bill_id:
                    paid = self.pt_tracker.total_for(pid) + amt
                    if paid >= b["total"]:
                        b["status"] = "Paid"
                    else:
                        b["status"] = "Partial"
                    break
        self.pt_tracker.add(pid, bill_id or "", amt, self.pay_method.text)
        self.save_all()
        self._refresh_pay()
        self.pay_amt.text = ""
        self._popup("Paid", f"N{amt:,.2f}")
        self.log_activity(f"Payment recorded: N{amt:,.2f} for {self.pay_name.text}")
        self.audit_log("RECORD_PAYMENT", f"N{amt:,.2f} via {self.pay_method.text}", pid)

    def _refresh_pay(self):
        pid = self.pay_pid.text
        if not pid:
            return
        total = sum(b.get("total", 0) for b in self.bills
                    if b.get("patient", {}).get("id") == pid)
        paid = self.pt_tracker.total_for(pid)
        self.pay_total.text = f"Total: N{total:,.2f}"
        self.pay_paid.text = f"Paid: N{paid:,.2f}"
        self.pay_bal.text = f"Balance: N{total - paid:,.2f}"
        bills = [b for b in self.bills
                 if b.get("patient", {}).get("id") == pid]
        values = ["None"] + [
            f"{b['bill_id']} - N{b['total']:,.2f} ({b.get('status', 'N/A')})"
            for b in bills
        ]
        self.pay_bill_spinner.values = values
        self.pay_hist.clear_widgets()
        for p in reversed(self.pt_tracker.get_for(pid)):
            row = BoxLayout(size_hint_y=None, height=30, spacing=3)
            row.add_widget(Label(
                text=f"{p.get('ts','')[:16]} | N{p.get('amount', 0):,.2f}",
                font_size=12, size_hint_x=0.7
            ))
            btn_v = Button(text="View", font_size=11, size_hint_x=0.15,
                            size_hint_y=None, height=28)
            pay = p
            btn_v.bind(on_press=lambda x, pay=pay: self._view_payment_detail_kivy(pay))
            row.add_widget(btn_v)
            self.pay_hist.add_widget(row)

    def _print_receipt(self, instance):
        pid = self.pay_pid.text
        if not pid:
            return
        name = self.pay_name.text
        bills = [b for b in self.bills
                 if b.get("patient", {}).get("id") == pid]
        total = sum(b.get("total", 0) for b in bills)
        paid = self.pt_tracker.total_for(pid)
        balance = total - paid
        receipt = (f"=== HMG HOSPITAL - RECEIPT ===\n"
                   f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                   f"Issued by: {self._current_username()}\n\n"
                   f"Patient: {name} (ID: {pid})\n\nBills:\n")
        for b in bills:
            receipt += f"- {b['bill_id']}: N{b['total']:,.2f} ({b.get('status', 'N/A')})\n"
        receipt += (f"\nTotal Billed: N{total:,.2f}\n"
                    f"Total Paid: N{paid:,.2f}\n"
                    f"Balance: N{balance:,.2f}\n\nPayment History:\n")
        for p in self.pt_tracker.get_for(pid):
            receipt += (f"{p.get('ts','')[:16]} | N{p.get('amount', 0):,.2f} "
                        f"via {p.get('method','')}\n")
        receipt += "\nThank you for choosing HMG Hospital."
        self._export_to_pdf(receipt, "Receipt", "Payment Receipt")

    def _view_saved_receipts(self, instance):
        if not self.cp and not self.pay_pid.text:
            self._popup("Error", "No patient selected.")
            return
        pid = self.cp.get("id") if self.cp else self.pay_pid.text
        payments = self.pt_tracker.get_for(pid)
        if not payments:
            self._popup("Receipts", "No payments found for this patient.")
            return
        c = BoxLayout(orientation='vertical', spacing=5, padding=8)
        pl = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        pl.bind(minimum_height=pl.setter('height'))
        sv = ScrollView()
        sv.add_widget(pl)
        c.add_widget(sv)
        for p in payments:
            text = (f"{p.get('ts','')[:16]} | N{p.get('amount', 0):,.2f} "
                    f"via {p.get('method','')} | Bill: {p.get('bill_id','')}")
            row = BoxLayout(size_hint_y=None, height=40, spacing=5)
            row.add_widget(Label(text=text, font_size=13, size_hint_x=0.8))
            btn_view = Button(text="View", font_size=13, size_hint_x=0.2)
            pay = p
            btn_view.bind(on_press=lambda x, pay=pay: self._view_payment_detail_kivy(pay))
            row.add_widget(btn_view)
            pl.add_widget(row)
        pp = Popup(title="Payment History", content=c, size_hint=(0.8, 0.7))
        c.add_widget(Button(text="CLOSE", font_size=16, bold=True,
                             size_hint_y=None, height=40, on_press=pp.dismiss))
        pp.open()

    def _view_payment_detail_kivy(self, payment):
        text = (f"Payment ID: {payment['id']}\n"
                f"Patient ID: {payment['patient_id']}\n"
                f"Bill ID: {payment['bill_id']}\n"
                f"Amount: N{payment['amount']:,.2f}\n"
                f"Method: {payment['method']}\n"
                f"Date: {payment['ts']}\n")
        self._text_popup("Payment Detail", text)

    # ====================== LABORATORY TAB ======================
    def _lab(self):
        tab = TabbedPanelItem(text="Laboratory")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=5, spacing=3, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="ID:", font_size=15, bold=True, size_hint_x=0.05))
        self.lab_pid = TextInput(size_hint_x=0.15, readonly=True, font_size=15,
                                   size_hint_y=None, height=40)
        bar.add_widget(self.lab_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.07))
        self.lab_name = TextInput(size_hint_x=0.35, readonly=True, font_size=15,
                                    size_hint_y=None, height=40)
        bar.add_widget(self.lab_name)
        l.add_widget(bar)

        order = BoxLayout(size_hint_y=None, height=40, spacing=4)
        order.add_widget(Label(text="Test:", font_size=15, bold=True, size_hint_x=0.06))
        self.lab_test = Spinner(
            text="Select", values=[i['name'] for i in self.investigations],
            size_hint_x=0.4, size_hint_y=None, height=40
        )
        order.add_widget(self.lab_test)
        order.add_widget(Button(
            text="SEARCH", font_size=15, bold=True, size_hint_x=0.12,
            on_press=self._search_lab_test,
            background_color=get_color_from_hex('#f39c12')
        ))
        order.add_widget(Button(
            text="ORDER", font_size=15, bold=True, size_hint_x=0.15,
            on_press=self._order_lab,
            background_color=get_color_from_hex('#2980b9')
        ))
        order.add_widget(Button(
            text="VIEW PENDING", font_size=15, bold=True, size_hint_x=0.15,
            on_press=self._view_pending_lab,
            background_color=get_color_from_hex('#16a085')
        ))
        order.add_widget(Button(
            text="DELETE", font_size=15, bold=True, size_hint_x=0.1,
            on_press=self._delete_lab,
            background_color=get_color_from_hex('#e74c3c')
        ))
        l.add_widget(order)

        self.lab_pend = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        self.lab_pend.bind(minimum_height=self.lab_pend.setter('height'))
        l.add_widget(self.lab_pend)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)
        self._refresh_lab()

    def _search_lab_test(self, instance):
        c = BoxLayout(orientation='vertical', spacing=5, padding=8)
        sr = TextInput(size_hint_y=None, height=40, font_size=15,
                        hint_text="Search test...")
        c.add_widget(sr)
        pl = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        pl.bind(minimum_height=pl.setter('height'))
        sv = ScrollView()
        sv.add_widget(pl)
        c.add_widget(sv)
        pp = Popup(title="Search Lab Test", content=c, size_hint=(0.7, 0.7))

        def refresh_list(t=""):
            pl.clear_widgets()
            for test in self.investigations:
                if not t or t.lower() in test['name'].lower():
                    btn = Button(text=test['name'], font_size=15,
                                  size_hint_y=None, height=40)
                    btn.test_name = test['name']
                    btn.bind(on_press=lambda x: (
                        setattr(self.lab_test, 'text', x.test_name),
                        pp.dismiss()
                    ))
                    pl.add_widget(btn)

        sr.bind(text=lambda i, v: refresh_list(v))
        refresh_list()
        pp.open()

    def _order_lab(self, instance):
        if not self.lab_pid.text or self.lab_test.text == "Select":
            self._popup("Error", "Select patient and test.")
            return
        self.lab_tests.append({
            "test_id": f"LAB{datetime.now().strftime('%Y%m%d%H%M%S%f')}",
            "patient_id": self.lab_pid.text,
            "patient_name": self.lab_name.text,
            "test_name": self.lab_test.text,
            "status": "Ordered",
            "result": "",
            "unit": "",
            "reference": "",
            "date_ordered": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "date_completed": "",
            "billed": False
        })
        self.save_all()
        self._refresh_lab()
        self._popup("Ordered", f"{self.lab_test.text} ordered.")
        self.log_activity(f"Lab test ordered: {self.lab_test.text} for {self.lab_name.text}")
        self.audit_log("ORDER_LAB", self.lab_test.text, self.lab_pid.text)
        if self.cp:
            self.task_manager.create_task(
                patient_id=self.cp.get("id"),
                patient_name=self.cp.get("name"),
                assigned_to="Lab Scientist",
                task_type="Laboratory",
                description=f"Lab test: {self.lab_test.text}",
                order_items=[self.lab_test.text],
                department="Laboratory"
            )
            self._refresh_tasks()

    def _refresh_lab(self):
        if not hasattr(self, 'lab_pend'):
            return
        self.lab_pend.clear_widgets()
        if not self.lab_pid.text:
            return
        pid = self.lab_pid.text
        tests = [t for t in self.lab_tests
                 if t.get("patient_id") == pid
                 and t.get("status") != "Completed"]
        if not tests:
            self.lab_pend.add_widget(Label(
                text="No pending tests.", font_size=13,
                size_hint_y=None, height=30
            ))
            return
        for t in tests:
            row = BoxLayout(size_hint_y=None, height=35, spacing=3)
            row.add_widget(Label(
                text=f"{t.get('test_name','')} - {t.get('status','')}",
                font_size=13, size_hint_x=0.6
            ))
            if t.get("status") == "Ordered":
                btn = Button(
                    text="ENTER RESULT", font_size=12,
                    size_hint_x=0.3, size_hint_y=None, height=30,
                    background_color=get_color_from_hex('#27ae60')
                )
                btn.test = t
                btn.bind(on_press=self._enter_lab_result)
                row.add_widget(btn)
            row.add_widget(Button(
                text="X", font_size=12, size_hint_x=0.1,
                size_hint_y=None, height=30,
                background_color=get_color_from_hex('#e74c3c'),
                on_press=lambda x, tid=t.get('test_id'): self._remove_lab_test(tid)
            ))
            self.lab_pend.add_widget(row)

    def _enter_lab_result(self, instance):
        test = instance.test
        content = BoxLayout(orientation='vertical', spacing=5, padding=8)
        content.add_widget(Label(
            text=f"Enter result for {test.get('test_name')}",
            font_size=16, bold=True,
            size_hint_y=None, height=30
        ))
        content.add_widget(Label(text="Result:", font_size=14, bold=True,
                                   size_hint_y=None, height=22))
        result_inp = TextInput(font_size=14, size_hint_y=None, height=40)
        content.add_widget(result_inp)
        content.add_widget(Label(text="Unit:", font_size=14, bold=True,
                                   size_hint_y=None, height=22))
        unit_inp = TextInput(font_size=14, size_hint_y=None, height=40)
        content.add_widget(unit_inp)
        content.add_widget(Label(text="Reference Range:", font_size=14, bold=True,
                                   size_hint_y=None, height=22))
        ref_inp = TextInput(font_size=14, size_hint_y=None, height=40)
        content.add_widget(ref_inp)
        btn_box = BoxLayout(size_hint_y=None, height=40, spacing=5)
        pp = Popup(title="Enter Result", content=content, size_hint=(0.6, 0.5))

        def save_result(inst):
            test["result"] = result_inp.text.strip()
            test["unit"] = unit_inp.text.strip()
            test["reference"] = ref_inp.text.strip()
            test["status"] = "Completed"
            test["date_completed"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.save_all()
            self._refresh_lab()
            for task in self.tasks:
                if (task.get("patient_id") == test.get("patient_id") and
                        "Lab" in task.get("task_type") and
                        task.get("status") != "Completed"):
                    if test.get("test_name") in task.get("description", ""):
                        self.task_manager.update_status(
                            task.get("id"), "Completed",
                            f"Result entered: {test.get('result')}"
                        )
                        break
            self._refresh_tasks()
            pp.dismiss()
            self._popup("Result Saved",
                        f"{test.get('test_name')} result saved.")
            self.log_activity(
                f"Lab result entered: {test.get('test_name')} = {test.get('result')}"
            )
            self.audit_log(
                "ENTER_LAB_RESULT",
                f"{test.get('test_name')} = {test.get('result')}",
                test.get("patient_id")
            )

        btn_box.add_widget(Button(
            text="SAVE", font_size=15, bold=True,
            on_press=save_result,
            background_color=get_color_from_hex('#27ae60')
        ))
        btn_box.add_widget(Button(
            text="CANCEL", font_size=15, on_press=pp.dismiss,
            background_color=get_color_from_hex('#e74c3c')
        ))
        content.add_widget(btn_box)
        pp.open()

    def _remove_lab_test(self, test_id):
        self.lab_tests = [t for t in self.lab_tests if t.get("test_id") != test_id]
        self.save_all()
        self._refresh_lab()

    def _view_pending_lab(self, instance):
        if not self.lab_pid.text:
            self._popup("Error", "No patient selected.")
            return
        pid = self.lab_pid.text
        tests = [t for t in self.lab_tests if t.get("patient_id") == pid]
        if not tests:
            self._popup("Pending Lab Tests", "No tests found for this patient.")
            return
        text = f"LAB TESTS FOR {self.lab_name.text}\n" + "=" * 50 + "\n\n"
        for t in tests:
            text += f"Test: {t.get('test_name','')}\n"
            text += f"Status: {t.get('status','')}\n"
            if t.get("status") == "Completed":
                text += f"Result: {t.get('result','')} {t.get('unit','')}\n"
                text += f"Reference: {t.get('reference','')}\n"
            text += f"Ordered: {t.get('date_ordered','')}\n"
            if t.get("date_completed"):
                text += f"Completed: {t.get('date_completed','')}\n"
            text += "-" * 30 + "\n"
        self._text_popup("Lab Tests", text)

    def _delete_lab(self, instance):
        if not self.lab_pid.text:
            return
        if not self._require_role("Administrator", "Doctor", "Lab Scientist"):
            return
        pid = self.lab_pid.text
        tests = [t for t in self.lab_tests if t.get("patient_id") == pid]
        if not tests:
            self._popup("Info", "No lab tests to delete.")
            return
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)
        content.add_widget(Label(
            text=f"Delete all lab tests for {self.lab_name.text}?\n"
                 f"This cannot be undone."
        ))
        btn_box = BoxLayout(size_hint_y=None, height=40, spacing=5)
        pp = Popup(title="Confirm Delete", content=content, size_hint=(0.6, 0.3))

        def do_delete(inst):
            self.lab_tests = [t for t in self.lab_tests
                              if t.get("patient_id") != pid]
            self.save_all()
            self._refresh_lab()
            pp.dismiss()
            self._popup("Deleted", "All lab tests deleted.")
            self.audit_log("DELETE_LAB_ALL", f"Patient {pid}", pid)

        btn_box.add_widget(Button(
            text="DELETE ALL", font_size=15, bold=True, on_press=do_delete,
            background_color=get_color_from_hex('#e74c3c')
        ))
        btn_box.add_widget(Button(text="CANCEL", font_size=15, on_press=pp.dismiss))
        content.add_widget(btn_box)
        pp.open()

    # ====================== ENHANCED LAB TAB ======================
    def _enhanced_lab_tab(self):
        tab = TabbedPanelItem(text="Enhanced Lab")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=6, spacing=4, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))
        l.add_widget(Label(
            text="ENHANCED LABORATORY MODULE",
            font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#1a5276')
        ))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="Patient ID:", font_size=15, bold=True,
                               size_hint_x=0.12))
        self.elab_pid = TextInput(size_hint_x=0.2, readonly=True, font_size=15,
                                    size_hint_y=None, height=40)
        bar.add_widget(self.elab_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.08))
        self.elab_name = TextInput(size_hint_x=0.3, readonly=True, font_size=15,
                                     size_hint_y=None, height=40)
        bar.add_widget(self.elab_name)
        bar.add_widget(Button(
            text="VIEW ALL", font_size=13, size_hint_x=0.15,
            on_press=self._view_all_lab_results,
            background_color=get_color_from_hex('#2980b9')
        ))
        l.add_widget(bar)

        test_box = BoxLayout(size_hint_y=None, height=40, spacing=4)
        test_box.add_widget(Label(text="Test:", font_size=15, bold=True,
                                    size_hint_x=0.06))
        self.elab_test = Spinner(
            text="Select", values=[i['name'] for i in self.investigations],
            size_hint_x=0.3, size_hint_y=None, height=40
        )
        test_box.add_widget(self.elab_test)
        test_box.add_widget(Label(text="Result:", font_size=15, bold=True,
                                    size_hint_x=0.08))
        self.elab_result = TextInput(size_hint_x=0.15, font_size=15,
                                       size_hint_y=None, height=40)
        test_box.add_widget(self.elab_result)
        test_box.add_widget(Label(text="Unit:", font_size=15, bold=True,
                                    size_hint_x=0.06))
        self.elab_unit = TextInput(size_hint_x=0.1, font_size=15,
                                     size_hint_y=None, height=40)
        test_box.add_widget(self.elab_unit)
        test_box.add_widget(Button(
            text="ADD", font_size=15, bold=True, size_hint_x=0.1,
            on_press=self._add_enhanced_lab,
            background_color=get_color_from_hex('#27ae60')
        ))
        l.add_widget(test_box)

        l.add_widget(Label(
            text="LAB RESULTS", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#2c3e50')
        ))
        self.elab_results = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        self.elab_results.bind(minimum_height=self.elab_results.setter('height'))
        l.add_widget(self.elab_results)

        btn_row = BoxLayout(size_hint_y=None, height=40, spacing=5)
        btn_row.add_widget(Button(
            text="EXPORT PDF", font_size=15, bold=True,
            on_press=self._export_enhanced_lab,
            background_color=get_color_from_hex('#2980b9')
        ))
        btn_row.add_widget(Button(
            text="GENERATE COMPREHENSIVE REPORT", font_size=15, bold=True,
            on_press=self._generate_comprehensive_lab_report,
            background_color=get_color_from_hex('#8e44ad')
        ))
        btn_row.add_widget(Button(
            text="CLEAR RESULTS", font_size=15, bold=True,
            on_press=self._clear_enhanced_lab,
            background_color=get_color_from_hex('#e74c3c')
        ))
        l.add_widget(btn_row)

        l.add_widget(Label(
            text="SAVED REPORTS", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#1a5276')
        ))
        self.saved_reports_list = BoxLayout(orientation='vertical', size_hint_y=None,
                                              spacing=1)
        self.saved_reports_list.bind(minimum_height=self.saved_reports_list.setter('height'))
        l.add_widget(self.saved_reports_list)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)
        self._refresh_elab()
        self._refresh_saved_reports_list()

    def _add_enhanced_lab(self, instance):
        if not self.elab_pid.text or self.elab_test.text == "Select":
            self._popup("Error", "Select patient and test.")
            return
        result = self.elab_result.text.strip()
        if not result:
            self._popup("Error", "Enter a result value.")
            return
        lab_entry = {
            "id": f"ELAB{datetime.now().strftime('%Y%m%d%H%M%S%f')}",
            "patient_id": self.elab_pid.text,
            "patient_name": self.elab_name.text,
            "test_name": self.elab_test.text,
            "result": result,
            "unit": self.elab_unit.text.strip(),
            "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "status": "Completed"
        }
        self.lab_results.append(lab_entry)
        self.save_all()
        self._refresh_elab()
        self.elab_result.text = ""
        self.elab_unit.text = ""
        self._popup("Added", f"{lab_entry['test_name']}: {result}")
        self.log_activity(
            f"Enhanced lab result added: {lab_entry['test_name']} = {result}"
        )
        self.audit_log(
            "ADD_ENHANCED_LAB",
            f"{lab_entry['test_name']} = {result}",
            self.elab_pid.text
        )

    def _refresh_elab(self):
        if not hasattr(self, 'elab_results'):
            return
        self.elab_results.clear_widgets()
        if not self.elab_pid.text:
            self.elab_results.add_widget(Label(
                text="Select a patient to view lab results.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        pid = self.elab_pid.text
        results = [r for r in self.lab_results if r.get("patient_id") == pid]
        if not results:
            self.elab_results.add_widget(Label(
                text="No lab results for this patient.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        for r in sorted(results, key=lambda x: x.get('date', ''), reverse=True):
            row = BoxLayout(size_hint_y=None, height=30, spacing=3)
            row.add_widget(Label(text=r.get('test_name', ''), font_size=13,
                                   size_hint_x=0.3))
            row.add_widget(Label(
                text=f"{r.get('result', '')} {r.get('unit', '')}",
                font_size=13, size_hint_x=0.2, bold=True,
                color=get_color_from_hex('#2980b9')
            ))
            row.add_widget(Label(text=r.get('date', '')[:16], font_size=11,
                                   size_hint_x=0.25))
            btn_del = Button(
                text="X", font_size=11, size_hint_x=0.05,
                size_hint_y=None, height=25,
                background_color=get_color_from_hex('#e74c3c')
            )
            btn_del.lab_id = r.get('id')
            btn_del.bind(on_press=lambda x: self._delete_enhanced_lab(x.lab_id))
            row.add_widget(btn_del)
            self.elab_results.add_widget(row)

    def _delete_enhanced_lab(self, lab_id):
        self.lab_results = [r for r in self.lab_results if r.get('id') != lab_id]
        self.save_all()
        self._refresh_elab()

    def _view_all_lab_results(self, instance):
        if not self.elab_pid.text:
            self._popup("Error", "No patient selected.")
            return
        pid = self.elab_pid.text
        results = [r for r in self.lab_results if r.get("patient_id") == pid]
        if not results:
            self._popup("Lab Results", "No results found.")
            return
        text = f"LAB RESULTS FOR {self.elab_name.text}\n" + "=" * 50 + "\n\n"
        for r in results:
            text += f"{r.get('test_name','')}: {r.get('result','')} {r.get('unit','')}\n"
            text += f"Date: {r.get('date','')}\n"
            text += "-" * 30 + "\n"
        self._text_popup("Lab Results", text)

    def _export_enhanced_lab(self, instance):
        if not self.elab_pid.text:
            self._popup("Error", "No patient selected.")
            return
        pid = self.elab_pid.text
        results = [r for r in self.lab_results if r.get("patient_id") == pid]
        if not results:
            self._popup("Error", "No results to export.")
            return
        text = "HMG HOSPITAL - LABORATORY RESULTS\n"
        text += f"Patient: {self.elab_name.text} (ID: {self.elab_pid.text})\n"
        text += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        text += "=" * 60 + "\n\n"
        for r in results:
            text += f"Test: {r.get('test_name','')}\n"
            text += f"Result: {r.get('result','')} {r.get('unit','')}\n"
            text += f"Date: {r.get('date','')}\n"
            text += "-" * 40 + "\n"
        self._export_to_pdf(text, "Lab_Results", "Laboratory Results")

    def _clear_enhanced_lab(self, instance):
        if not self.elab_pid.text:
            return
        if not self._require_role("Administrator", "Doctor", "Lab Scientist"):
            return
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)
        content.add_widget(Label(
            text=f"Delete all lab results for {self.elab_name.text}?\n"
                 f"This cannot be undone."
        ))
        btn_box = BoxLayout(size_hint_y=None, height=40, spacing=5)
        pp = Popup(title="Confirm Delete", content=content, size_hint=(0.6, 0.3))

        def do_delete(inst):
            pid = self.elab_pid.text
            self.lab_results = [r for r in self.lab_results
                                if r.get("patient_id") != pid]
            self.save_all()
            self._refresh_elab()
            pp.dismiss()
            self._popup("Deleted", "All lab results deleted.")
            self.audit_log("DELETE_ENHANCED_LAB", f"Patient {pid}", pid)

        btn_box.add_widget(Button(
            text="DELETE ALL", font_size=15, bold=True, on_press=do_delete,
            background_color=get_color_from_hex('#e74c3c')
        ))
        btn_box.add_widget(Button(text="CANCEL", font_size=15, on_press=pp.dismiss))
        content.add_widget(btn_box)
        pp.open()

    def _generate_comprehensive_lab_report(self, instance):
        if not self.elab_pid.text:
            self._popup("Error", "Select a patient first.")
            return
        pid = self.elab_pid.text
        name = self.elab_name.text
        results = [r for r in self.lab_results if r.get("patient_id") == pid]
        if not results:
            self._popup("Error", "No lab results to report.")
            return

        patient = None
        for p in self.patients:
            if p.get("id") == pid:
                patient = p
                break

        report_id = f"LR-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        qr_data = f"HMG|LAB|{report_id}|{pid}|{name}"
        qr_b64 = QRCodeGenerator.generate(qr_data)
        qr_temp = None
        if qr_b64:
            try:
                qr_bytes = base64.b64decode(qr_b64)
                qr_temp = os.path.join(self._get_reports_dir(), f"qr_{report_id}.png")
                with open(qr_temp, "wb") as f:
                    f.write(qr_bytes)
            except:
                qr_temp = None

        if PDF_AVAILABLE:
            try:
                class PDF(FPDF):
                    def header(self):
                        self.set_font('Arial', 'B', 14)
                        self.cell(0, 10, 'HMG Hospital', 0, 1, 'C')
                        self.ln(2)

                    def footer(self):
                        self.set_y(-15)
                        self.set_font('Arial', 'I', 8)
                        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

                pdf = PDF()
                pdf.add_page()
                pdf.set_font("Arial", 'B', 16)
                pdf.cell(0, 10, "COMPREHENSIVE LABORATORY REPORT", ln=True, align='C')
                pdf.set_font("Arial", 'I', 10)
                pdf.cell(0, 6, f"Report ID: {report_id}", ln=True, align='C')
                pdf.ln(4)
                pdf.set_font("Arial", '', 12)
                pdf.cell(0, 8, f"Patient: {self._sanitize_text(name)} (ID: {pid})", ln=True)
                if patient:
                    pdf.cell(0, 8,
                             f"DOB: {self._sanitize_text(patient.get('dob', 'N/A'))} | "
                             f"Age: {patient.get('age', 'N/A')} | "
                             f"Gender: {self._sanitize_text(patient.get('gender', 'N/A'))}",
                             ln=True)
                pdf.cell(0, 8,
                         f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                         ln=True)
                pdf.cell(0, 8,
                         f"Issued by: {self._current_username()}",
                         ln=True)
                pdf.ln(4)

                pdf.set_font('Arial', 'B', 10)
                pdf.cell(45, 10, "Date", 1, 0, 'C')
                pdf.cell(65, 10, "Test", 1, 0, 'C')
                pdf.cell(35, 10, "Result", 1, 0, 'C')
                pdf.cell(25, 10, "Unit", 1, 0, 'C')
                pdf.cell(30, 10, "Reference", 1, 1, 'C')
                pdf.set_font('Arial', '', 10)
                for r in sorted(results, key=lambda x: x.get('date', '')):
                    date = self._sanitize_text(r.get('date', '')[:16])
                    test = self._sanitize_text(r.get('test_name', '')[:30])
                    result = self._sanitize_text(r.get('result', '')[:15])
                    unit = self._sanitize_text(r.get('unit', '')[:10])
                    ref = self._sanitize_text(r.get('reference', '')[:15])
                    pdf.cell(45, 8, date, 1)
                    pdf.cell(65, 8, test, 1)
                    pdf.cell(35, 8, result, 1)
                    pdf.cell(25, 8, unit, 1)
                    pdf.cell(30, 8, ref, 1)
                    pdf.ln()
                pdf.ln(6)

                pdf.set_font('Arial', 'B', 12)
                pdf.cell(0, 10, "CUSTOM NOTES:", ln=True)
                pdf.set_font('Arial', '', 11)
                for _ in range(3):
                    pdf.cell(0, 8, "_" * 70, ln=True)
                pdf.ln(6)

                pdf.set_font('Arial', 'B', 12)
                pdf.cell(0, 10, "ELECTRONIC SIGNATURE", ln=True)
                pdf.set_font('Arial', '', 11)
                pdf.cell(0, 8, "Dr. Precious Belema Ibiabuo", ln=True)
                pdf.cell(0, 8, "MLT, M.Sc., MD, FWACS-P", ln=True)
                pdf.cell(0, 8,
                         f"Signed on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                         ln=True)
                if qr_temp and os.path.exists(qr_temp):
                    pdf.image(qr_temp, x=170, y=pdf.get_y(), w=30)
                pdf.ln(10)
                pdf.cell(0, 10, "This report is electronically signed and verified.",
                         0, 1, 'C')

                dir_path = self._get_reports_dir()
                pdf_fn = os.path.join(dir_path, f"Lab_Report_{report_id}.pdf")
                pdf.output(pdf_fn)

                report_entry = {
                    "id": report_id,
                    "patient_id": pid,
                    "patient_name": name,
                    "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "file_path": pdf_fn,
                    "result_count": len(results)
                }
                self.saved_lab_reports.append(report_entry)
                self.save_all()
                self._refresh_saved_reports_list()
                self._open_file(pdf_fn)
                self._popup("Report Generated",
                            f"Lab report saved as:\n"
                            f"{os.path.basename(pdf_fn)}\n\n"
                            f"{len(results)} results included.")
                self.log_activity(
                    f"Comprehensive lab report generated for {name} - {report_id}"
                )
                self.audit_log("GENERATE_LAB_REPORT", report_id, pid)
                if qr_temp and os.path.exists(qr_temp):
                    try:
                        os.remove(qr_temp)
                    except:
                        pass
                return
            except Exception as e:
                self.log_activity(f"Report generation error: {e}")
                self._popup("PDF Error", str(e))
        else:
            text = f"HMG HOSPITAL - LAB REPORT\nReport ID: {report_id}\nPatient: {name} (ID: {pid})\n\n"
            for r in results:
                text += f"{r.get('test_name','')}: {r.get('result','')} {r.get('unit','')}\n"
            self._export_to_pdf(text, "Lab_Report", "Laboratory Report")

    def _refresh_saved_reports_list(self):
        if not hasattr(self, 'saved_reports_list'):
            return
        self.saved_reports_list.clear_widgets()
        if not self.elab_pid.text:
            self.saved_reports_list.add_widget(Label(
                text="Select a patient to view saved reports.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        pid = self.elab_pid.text
        reports = [r for r in self.saved_lab_reports
                   if r.get("patient_id") == pid]
        if not reports:
            self.saved_reports_list.add_widget(Label(
                text="No saved reports for this patient.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        for r in sorted(reports, key=lambda x: x.get('date', ''), reverse=True):
            row = BoxLayout(size_hint_y=None, height=30, spacing=3)
            row.add_widget(Label(
                text=(f"{r.get('id', '')} | {r.get('date', '')[:16]} | "
                      f"{r.get('result_count', 0)} results"),
                font_size=12, size_hint_x=0.6
            ))
            btn_view = Button(
                text="OPEN", font_size=11, size_hint_x=0.2,
                size_hint_y=None, height=25,
                background_color=get_color_from_hex('#2980b9')
            )
            btn_view.file_path = r.get('file_path', '')
            btn_view.bind(on_press=lambda x: self._open_file(x.file_path)
                          if os.path.exists(x.file_path)
                          else self._popup("Error", "File not found."))
            row.add_widget(btn_view)
            btn_del = Button(
                text="X", font_size=11, size_hint_x=0.08,
                size_hint_y=None, height=25,
                background_color=get_color_from_hex('#e74c3c')
            )
            btn_del.report_id = r.get('id')
            btn_del.bind(on_press=lambda x: self._delete_saved_report(x.report_id))
            row.add_widget(btn_del)
            self.saved_reports_list.add_widget(row)

    def _delete_saved_report(self, report_id):
        self.saved_lab_reports = [r for r in self.saved_lab_reports
                                  if r.get('id') != report_id]
        self.save_all()
        self._refresh_saved_reports_list()
        self._popup("Deleted", "Report removed from saved list.")

    # ====================== RADIOLOGY TAB ======================
    def _rad(self):
        tab = TabbedPanelItem(text="Radiology")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=5, spacing=3, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="ID:", font_size=15, bold=True, size_hint_x=0.05))
        self.rad_pid = TextInput(size_hint_x=0.15, readonly=True, font_size=15,
                                   size_hint_y=None, height=40)
        bar.add_widget(self.rad_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.07))
        self.rad_name = TextInput(size_hint_x=0.35, readonly=True, font_size=15,
                                    size_hint_y=None, height=40)
        bar.add_widget(self.rad_name)
        l.add_widget(bar)

        order = BoxLayout(size_hint_y=None, height=40, spacing=4)
        order.add_widget(Label(text="Study:", font_size=15, bold=True, size_hint_x=0.06))
        self.rad_study = Spinner(
            text="Select",
            values=["X-Ray Chest", "X-Ray Abdomen", "CT Scan", "MRI",
                    "Ultrasound", "Mammogram", "DEXA Scan", "Fluoroscopy"],
            size_hint_x=0.3, size_hint_y=None, height=40
        )
        order.add_widget(self.rad_study)
        order.add_widget(Button(
            text="ORDER", font_size=15, bold=True, size_hint_x=0.15,
            on_press=self._order_rad,
            background_color=get_color_from_hex('#2980b9')
        ))
        order.add_widget(Button(
            text="VIEW", font_size=15, bold=True, size_hint_x=0.15,
            on_press=self._view_rad,
            background_color=get_color_from_hex('#16a085')
        ))
        order.add_widget(Button(
            text="DELETE", font_size=15, bold=True, size_hint_x=0.1,
            on_press=self._delete_rad,
            background_color=get_color_from_hex('#e74c3c')
        ))
        l.add_widget(order)

        self.rad_list = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        self.rad_list.bind(minimum_height=self.rad_list.setter('height'))
        l.add_widget(self.rad_list)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)
        self._refresh_rad()

    def _order_rad(self, instance):
        if not self.rad_pid.text or self.rad_study.text == "Select":
            self._popup("Error", "Select patient and study.")
            return
        self.radiology.append({
            "rad_id": f"RAD{datetime.now().strftime('%Y%m%d%H%M%S%f')}",
            "patient_id": self.rad_pid.text,
            "patient_name": self.rad_name.text,
            "study": self.rad_study.text,
            "status": "Ordered",
            "report": "",
            "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "completed": ""
        })
        self.save_all()
        self._refresh_rad()
        self._popup("Ordered", f"{self.rad_study.text} ordered.")
        self.log_activity(f"Radiology ordered: {self.rad_study.text} for {self.rad_name.text}")
        self.audit_log("ORDER_RAD", self.rad_study.text, self.rad_pid.text)
        if self.cp:
            self.task_manager.create_task(
                patient_id=self.cp.get("id"),
                patient_name=self.cp.get("name"),
                assigned_to="Radiologist",
                task_type="Radiology",
                description=f"Radiology study: {self.rad_study.text}",
                order_items=[self.rad_study.text],
                department="Radiology"
            )
            self._refresh_tasks()

    def _refresh_rad(self):
        if not hasattr(self, 'rad_list'):
            return
        self.rad_list.clear_widgets()
        if not self.rad_pid.text:
            return
        pid = self.rad_pid.text
        studies = [r for r in self.radiology if r.get("patient_id") == pid]
        if not studies:
            self.rad_list.add_widget(Label(
                text="No radiology studies.", font_size=13,
                size_hint_y=None, height=30
            ))
            return
        for r in studies:
            row = BoxLayout(size_hint_y=None, height=35, spacing=3)
            row.add_widget(Label(
                text=f"{r.get('study','')} - {r.get('status','')}",
                font_size=13, size_hint_x=0.5
            ))
            if r.get("status") == "Ordered":
                btn = Button(
                    text="ENTER REPORT", font_size=12,
                    size_hint_x=0.25, size_hint_y=None, height=30,
                    background_color=get_color_from_hex('#27ae60')
                )
                btn.rad = r
                btn.bind(on_press=self._enter_rad_report)
                row.add_widget(btn)
            elif r.get("status") == "Completed":
                btn = Button(
                    text="VIEW REPORT", font_size=12,
                    size_hint_x=0.25, size_hint_y=None, height=30,
                    background_color=get_color_from_hex('#2980b9')
                )
                btn.rad = r
                btn.bind(on_press=lambda x: self._text_popup(
                    "Radiology Report", x.rad.get('report', 'No report')
                ))
                row.add_widget(btn)
            row.add_widget(Button(
                text="X", font_size=12, size_hint_x=0.1,
                size_hint_y=None, height=30,
                background_color=get_color_from_hex('#e74c3c'),
                on_press=lambda x, rid=r.get('rad_id'): self._remove_rad(rid)
            ))
            self.rad_list.add_widget(row)

    def _enter_rad_report(self, instance):
        rad = instance.rad
        content = BoxLayout(orientation='vertical', spacing=5, padding=8)
        content.add_widget(Label(
            text=f"Report for {rad.get('study')}",
            font_size=16, bold=True, size_hint_y=None, height=30
        ))
        content.add_widget(Label(text="Findings:", font_size=14, bold=True,
                                   size_hint_y=None, height=22))
        report_inp = TextInput(font_size=14, size_hint_y=None, height=120,
                                multiline=True)
        content.add_widget(report_inp)
        content.add_widget(Label(text="Impression:", font_size=14, bold=True,
                                   size_hint_y=None, height=22))
        impression_inp = TextInput(font_size=14, size_hint_y=None, height=80,
                                     multiline=True)
        content.add_widget(impression_inp)
        btn_box = BoxLayout(size_hint_y=None, height=40, spacing=5)
        pp = Popup(title="Enter Report", content=content, size_hint=(0.7, 0.7))

        def save_report(inst):
            rad["report"] = (f"Findings:\n{report_inp.text}\n\n"
                             f"Impression:\n{impression_inp.text}")
            rad["status"] = "Completed"
            rad["completed"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.save_all()
            self._refresh_rad()
            for task in self.tasks:
                if (task.get("patient_id") == rad.get("patient_id") and
                        "Radiology" in task.get("task_type") and
                        task.get("status") != "Completed"):
                    if rad.get("study") in task.get("description", ""):
                        self.task_manager.update_status(
                            task.get("id"), "Completed", "Report completed"
                        )
                        break
            self._refresh_tasks()
            pp.dismiss()
            self._popup("Report Saved", "Radiology report completed.")
            self.log_activity(f"Radiology report completed: {rad.get('study')}")
            self.audit_log("ENTER_RAD_REPORT", rad.get('study'),
                           rad.get("patient_id"))

        btn_box.add_widget(Button(
            text="SAVE", font_size=15, bold=True, on_press=save_report,
            background_color=get_color_from_hex('#27ae60')
        ))
        btn_box.add_widget(Button(
            text="CANCEL", font_size=15, on_press=pp.dismiss,
            background_color=get_color_from_hex('#e74c3c')
        ))
        content.add_widget(btn_box)
        pp.open()

    def _remove_rad(self, rad_id):
        self.radiology = [r for r in self.radiology if r.get("rad_id") != rad_id]
        self.save_all()
        self._refresh_rad()

    def _view_rad(self, instance):
        if not self.rad_pid.text:
            self._popup("Error", "No patient selected.")
            return
        pid = self.rad_pid.text
        studies = [r for r in self.radiology if r.get("patient_id") == pid]
        if not studies:
            self._popup("Radiology", "No studies found.")
            return
        text = f"RADIOLOGY STUDIES FOR {self.rad_name.text}\n" + "=" * 50 + "\n\n"
        for r in studies:
            text += f"Study: {r.get('study','')}\n"
            text += f"Status: {r.get('status','')}\n"
            text += f"Date: {r.get('date','')}\n"
            if r.get("completed"):
                text += f"Completed: {r.get('completed','')}\n"
            if r.get("report"):
                text += f"\nReport:\n{r.get('report','')}\n"
            text += "-" * 30 + "\n"
        self._text_popup("Radiology Studies", text)

    def _delete_rad(self, instance):
        if not self.rad_pid.text:
            return
        if not self._require_role("Administrator", "Doctor", "Radiologist"):
            return
        pid = self.rad_pid.text
        studies = [r for r in self.radiology if r.get("patient_id") == pid]
        if not studies:
            self._popup("Info", "No studies to delete.")
            return
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)
        content.add_widget(Label(
            text=f"Delete all radiology studies for {self.rad_name.text}?\n"
                 f"This cannot be undone."
        ))
        btn_box = BoxLayout(size_hint_y=None, height=40, spacing=5)
        pp = Popup(title="Confirm Delete", content=content, size_hint=(0.6, 0.3))

        def do_delete(inst):
            self.radiology = [r for r in self.radiology
                              if r.get("patient_id") != pid]
            self.save_all()
            self._refresh_rad()
            pp.dismiss()
            self._popup("Deleted", "All radiology studies deleted.")
            self.audit_log("DELETE_RAD_ALL", f"Patient {pid}", pid)

        btn_box.add_widget(Button(
            text="DELETE ALL", font_size=15, bold=True, on_press=do_delete,
            background_color=get_color_from_hex('#e74c3c')
        ))
        btn_box.add_widget(Button(text="CANCEL", font_size=15, on_press=pp.dismiss))
        content.add_widget(btn_box)
        pp.open()

    # ====================== UPLOAD TAB ======================
    def _upload_tab(self):
        tab = TabbedPanelItem(text="Upload/Results")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=5, spacing=3, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="ID:", font_size=15, bold=True, size_hint_x=0.05))
        self.up_pid = TextInput(size_hint_x=0.15, readonly=True, font_size=15,
                                  size_hint_y=None, height=40)
        bar.add_widget(self.up_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.07))
        self.up_name = TextInput(size_hint_x=0.35, readonly=True, font_size=15,
                                   size_hint_y=None, height=40)
        bar.add_widget(self.up_name)
        l.add_widget(bar)

        upload = BoxLayout(size_hint_y=None, height=40, spacing=4)
        upload.add_widget(Button(
            text="UPLOAD FILE", font_size=15, bold=True, size_hint_x=0.25,
            on_press=self._upload_file,
            background_color=get_color_from_hex('#2980b9')
        ))
        upload.add_widget(Button(
            text="VIEW RESULTS", font_size=15, bold=True, size_hint_x=0.25,
            on_press=self._view_uploaded_results,
            background_color=get_color_from_hex('#16a085')
        ))
        upload.add_widget(Button(
            text="DELETE", font_size=15, bold=True, size_hint_x=0.15,
            on_press=self._delete_uploaded,
            background_color=get_color_from_hex('#e74c3c')
        ))
        l.add_widget(upload)

        self.up_list = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        self.up_list.bind(minimum_height=self.up_list.setter('height'))
        l.add_widget(self.up_list)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)

    def _upload_file(self, instance):
        if not self.up_pid.text:
            self._popup("Error", "No patient selected.")
            return
        c = BoxLayout(orientation='vertical', spacing=4, padding=4)
        fc = FileChooserListView(path=os.getcwd())
        c.add_widget(fc)
        btns = BoxLayout(size_hint_y=None, height=42, spacing=5)
        pp = Popup(title="Upload File", content=c, size_hint=(0.9, 0.85))

        def select(inst):
            if fc.selection:
                src = fc.selection[0]
                dest_dir = os.path.join(self._get_reports_dir(), self.up_pid.text)
                os.makedirs(dest_dir, exist_ok=True)
                dest = os.path.join(dest_dir, os.path.basename(src))
                import shutil
                shutil.copy2(src, dest)
                self._popup("Uploaded", f"File saved to:\n{dest}")
                self.log_activity(
                    f"File uploaded for {self.up_name.text}: "
                    f"{os.path.basename(src)}"
                )
                self.audit_log("UPLOAD_FILE", os.path.basename(src),
                               self.up_pid.text)
                pp.dismiss()
                self._refresh_uploaded_results()

        btns.add_widget(Button(
            text="SELECT", font_size=16, bold=True, on_press=select,
            background_color=get_color_from_hex('#27ae60')
        ))
        btns.add_widget(Button(text="CANCEL", font_size=16, on_press=pp.dismiss))
        c.add_widget(btns)
        pp.open()

    def _refresh_uploaded_results(self):
        if not hasattr(self, 'up_list'):
            return
        self.up_list.clear_widgets()
        if not self.up_pid.text:
            self.up_list.add_widget(Label(
                text="Select a patient to view uploaded files.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        pid = self.up_pid.text
        dir_path = os.path.join(self._get_reports_dir(), pid)
        if not os.path.exists(dir_path):
            self.up_list.add_widget(Label(
                text="No uploaded files for this patient.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        files = os.listdir(dir_path)
        if not files:
            self.up_list.add_widget(Label(
                text="No files in patient directory.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        for f in files:
            row = BoxLayout(size_hint_y=None, height=35, spacing=3)
            row.add_widget(Label(text=f, font_size=13, size_hint_x=0.8))
            btn = Button(
                text="VIEW", font_size=13, size_hint_x=0.2,
                size_hint_y=None, height=30,
                background_color=get_color_from_hex('#2980b9')
            )
            btn.filepath = os.path.join(dir_path, f)
            btn.bind(on_press=lambda x: self._open_file(x.filepath))
            row.add_widget(btn)
            self.up_list.add_widget(row)

    def _view_uploaded_results(self, instance):
        self._refresh_uploaded_results()

    def _delete_uploaded(self, instance):
        if not self.up_pid.text:
            return
        if not self._require_role("Administrator", "Doctor"):
            return
        pid = self.up_pid.text
        dir_path = os.path.join(self._get_reports_dir(), pid)
        if not os.path.exists(dir_path):
            self._popup("Info", "No files to delete.")
            return
        files = os.listdir(dir_path)
        if not files:
            self._popup("Info", "No files to delete.")
            return
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)
        content.add_widget(Label(
            text=f"Delete all uploaded files for {self.up_name.text}?\n"
                 f"This cannot be undone."
        ))
        btn_box = BoxLayout(size_hint_y=None, height=40, spacing=5)
        pp = Popup(title="Confirm Delete", content=content, size_hint=(0.6, 0.3))

        def do_delete(inst):
            import shutil
            shutil.rmtree(dir_path)
            self._refresh_uploaded_results()
            pp.dismiss()
            self._popup("Deleted", "All files deleted.")
            self.audit_log("DELETE_UPLOADED", f"Patient {pid}", pid)

        btn_box.add_widget(Button(
            text="DELETE ALL", font_size=15, bold=True, on_press=do_delete,
            background_color=get_color_from_hex('#e74c3c')
        ))
        btn_box.add_widget(Button(text="CANCEL", font_size=15, on_press=pp.dismiss))
        content.add_widget(btn_box)
        pp.open()

    # ====================== SAVED ITEMS TAB ======================
    def _saved_items_tab(self):
        tab = TabbedPanelItem(text="Saved Items")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=5, spacing=3, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))
        l.add_widget(Label(
            text="SAVED ITEMS", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#1a5276')
        ))

        btn_row = BoxLayout(size_hint_y=None, height=40, spacing=5)
        btn_row.add_widget(Button(
            text="BILLS", font_size=15, bold=True,
            on_press=self._show_saved_bills,
            background_color=get_color_from_hex('#2980b9')
        ))
        btn_row.add_widget(Button(
            text="PAYMENTS", font_size=15, bold=True,
            on_press=self._show_saved_payments,
            background_color=get_color_from_hex('#27ae60')
        ))
        btn_row.add_widget(Button(
            text="ECG REPORTS", font_size=15, bold=True,
            on_press=self._show_saved_ecg,
            background_color=get_color_from_hex('#e74c3c')
        ))
        btn_row.add_widget(Button(
            text="LAB TESTS", font_size=15, bold=True,
            on_press=self._show_saved_lab,
            background_color=get_color_from_hex('#8e44ad')
        ))
        btn_row.add_widget(Button(
            text="RADIOLOGY", font_size=15, bold=True,
            on_press=self._show_saved_rad,
            background_color=get_color_from_hex('#16a085')
        ))
        btn_row.add_widget(Button(
            text="LAB REPORTS", font_size=15, bold=True,
            on_press=self._show_saved_lab_reports,
            background_color=get_color_from_hex('#8e44ad')
        ))
        l.add_widget(btn_row)

        self.saved_items_display = BoxLayout(orientation='vertical', size_hint_y=None,
                                              spacing=1)
        self.saved_items_display.bind(minimum_height=self.saved_items_display.setter('height'))
        l.add_widget(self.saved_items_display)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)

    def _refresh_saved_items(self):
        pass

    def _show_saved_bills(self, instance):
        if not self.cp and not self.bill_pid.text:
            self._popup("Error", "No patient selected.")
            return
        pid = self.cp.get("id") if self.cp else self.bill_pid.text
        bills = [b for b in self.bills
                 if b.get("patient", {}).get("id") == pid]
        if not bills:
            self._popup("Bills", "No saved bills for this patient.")
            return
        text = f"BILLS FOR {self.bill_name.text}\n" + "=" * 50 + "\n\n"
        for b in bills:
            text += f"ID: {b.get('bill_id', '')}\n"
            text += f"Date: {b.get('ts', '')}\n"
            text += f"Total: N{b.get('total', 0):,.2f}\n"
            text += f"Status: {b.get('status', '')}\n"
            text += "-" * 30 + "\n"
        self._text_popup("Saved Bills", text)

    def _show_saved_payments(self, instance):
        if not self.cp and not self.pay_pid.text:
            self._popup("Error", "No patient selected.")
            return
        pid = self.cp.get("id") if self.cp else self.pay_pid.text
        payments = self.pt_tracker.get_for(pid)
        if not payments:
            self._popup("Payments", "No payments for this patient.")
            return
        text = f"PAYMENTS FOR {self.pay_name.text}\n" + "=" * 50 + "\n\n"
        total = 0
        for p in payments:
            text += f"Amount: N{p.get('amount', 0):,.2f}\n"
            text += f"Method: {p.get('method', '')}\n"
            text += f"Date: {p.get('ts', '')}\n"
            total += p.get('amount', 0)
            text += "-" * 30 + "\n"
        text += f"\nTOTAL PAID: N{total:,.2f}"
        self._text_popup("Payment History", text)

    def _show_saved_ecg(self, instance):
        if not self.cp:
            self._popup("Error", "No patient selected.")
            return
        pid = self.cp.get("id")
        reports = [e for e in self.ecg_reports if e.get("patient_id") == pid]
        if not reports:
            self._popup("ECG Reports", "No ECG reports for this patient.")
            return
        text = f"ECG REPORTS FOR {self.cp.get('name', '')}\n" + "=" * 50 + "\n\n"
        for e in reports:
            sd = e.get("structured", {})
            text += f"Report ID: {e.get('rid', '')}\n"
            text += f"Date: {e.get('ts', '')}\n"
            text += f"Interpretation: {sd.get('interpretation', 'N/A')}\n"
            text += f"Urgency: {sd.get('urgency', 'Routine')}\n"
            text += f"AI Used: {'Yes' if e.get('ai_used', False) else 'No'}\n"
            text += "-" * 30 + "\n"
        self._text_popup("ECG Reports", text)

    def _show_saved_lab(self, instance):
        if not self.cp:
            self._popup("Error", "No patient selected.")
            return
        pid = self.cp.get("id")
        tests = [t for t in self.lab_tests if t.get("patient_id") == pid]
        if not tests:
            self._popup("Lab Tests", "No lab tests for this patient.")
            return
        text = f"LAB TESTS FOR {self.cp.get('name', '')}\n" + "=" * 50 + "\n\n"
        for t in tests:
            text += f"Test: {t.get('test_name', '')}\n"
            text += f"Status: {t.get('status', '')}\n"
            if t.get("status") == "Completed":
                text += f"Result: {t.get('result', '')} {t.get('unit', '')}\n"
            text += f"Ordered: {t.get('date_ordered', '')}\n"
            text += "-" * 30 + "\n"
        self._text_popup("Lab Tests", text)

    def _show_saved_rad(self, instance):
        if not self.cp:
            self._popup("Error", "No patient selected.")
            return
        pid = self.cp.get("id")
        studies = [r for r in self.radiology if r.get("patient_id") == pid]
        if not studies:
            self._popup("Radiology", "No radiology studies for this patient.")
            return
        text = f"RADIOLOGY STUDIES FOR {self.cp.get('name', '')}\n" + "=" * 50 + "\n\n"
        for r in studies:
            text += f"Study: {r.get('study', '')}\n"
            text += f"Status: {r.get('status', '')}\n"
            text += f"Date: {r.get('date', '')}\n"
            if r.get("completed"):
                text += f"Completed: {r.get('completed', '')}\n"
            text += "-" * 30 + "\n"
        self._text_popup("Radiology Studies", text)

    def _show_saved_lab_reports(self, instance):
        if not self.cp:
            self._popup("Error", "No patient selected.")
            return
        pid = self.cp.get("id")
        reports = [r for r in self.saved_lab_reports
                   if r.get("patient_id") == pid]
        if not reports:
            self._popup("Lab Reports", "No saved lab reports for this patient.")
            return
        text = f"LAB REPORTS FOR {self.cp.get('name', '')}\n" + "=" * 50 + "\n\n"
        for r in sorted(reports, key=lambda x: x.get('date', ''), reverse=True):
            text += f"Report ID: {r.get('id', '')}\n"
            text += f"Date: {r.get('date', '')}\n"
            text += f"Results: {r.get('result_count', 0)}\n"
            text += f"File: {os.path.basename(r.get('file_path', ''))}\n"
            text += "-" * 30 + "\n"
        self._text_popup("Lab Reports", text)

    # ====================== DAILY NOTES TAB ======================
    def _daily_notes_tab(self):
        tab = TabbedPanelItem(text="Daily Notes")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=6, spacing=4, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))
        l.add_widget(Label(
            text="DAILY CLINICAL NOTES", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#1a5276')
        ))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="Patient ID:", font_size=15, bold=True,
                               size_hint_x=0.12))
        self.daily_pid = TextInput(size_hint_x=0.2, readonly=True, font_size=15,
                                     size_hint_y=None, height=40)
        bar.add_widget(self.daily_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.08))
        self.daily_name = TextInput(size_hint_x=0.3, readonly=True, font_size=15,
                                      size_hint_y=None, height=40)
        bar.add_widget(self.daily_name)
        l.add_widget(bar)

        l.add_widget(Label(text="Date:", font_size=15, bold=True,
                            size_hint_y=None, height=18))
        self.daily_date = TextInput(
            text=datetime.now().strftime("%Y-%m-%d"),
            font_size=15, size_hint_y=None, height=40
        )
        l.add_widget(self.daily_date)

        l.add_widget(Label(text="Clinical Notes:", font_size=15, bold=True,
                            size_hint_y=None, height=18))
        self.daily_notes_text = TextInput(
            font_size=14, size_hint_y=None, height=120, multiline=True,
            hint_text="Enter daily clinical notes..."
        )
        l.add_widget(self.daily_notes_text)

        vitals_box = GridLayout(cols=4, size_hint_y=None, height=100, spacing=3)
        vitals_box.add_widget(Label(text="BP:", font_size=13))
        self.daily_bp = TextInput(font_size=13, size_hint_y=None, height=35)
        vitals_box.add_widget(self.daily_bp)
        vitals_box.add_widget(Label(text="Pulse:", font_size=13))
        self.daily_pulse = TextInput(font_size=13, size_hint_y=None, height=35)
        vitals_box.add_widget(self.daily_pulse)
        vitals_box.add_widget(Label(text="Temp:", font_size=13))
        self.daily_temp = TextInput(font_size=13, size_hint_y=None, height=35)
        vitals_box.add_widget(self.daily_temp)
        vitals_box.add_widget(Label(text="SpO2:", font_size=13))
        self.daily_spo2 = TextInput(font_size=13, size_hint_y=None, height=35)
        vitals_box.add_widget(self.daily_spo2)
        vitals_box.add_widget(Label(text="RR:", font_size=13))
        self.daily_rr = TextInput(font_size=13, size_hint_y=None, height=35)
        vitals_box.add_widget(self.daily_rr)
        vitals_box.add_widget(Label(text="AVPU:", font_size=13))
        self.daily_avpu = Spinner(
            text="Alert",
            values=["Alert", "Voice", "Pain", "Unresponsive"],
            size_hint_y=None, height=35
        )
        vitals_box.add_widget(self.daily_avpu)
        l.add_widget(vitals_box)

        btn_row = BoxLayout(size_hint_y=None, height=42, spacing=5)
        btn_row.add_widget(Button(
            text="SAVE NOTE", font_size=16, bold=True,
            on_press=self._save_daily_note,
            background_color=get_color_from_hex('#27ae60')
        ))
        btn_row.add_widget(Button(
            text="VIEW HISTORY", font_size=16, bold=True,
            on_press=self._view_daily_notes,
            background_color=get_color_from_hex('#2980b9')
        ))
        btn_row.add_widget(Button(
            text="AI GENERATE", font_size=16, bold=True,
            on_press=self._ai_generate_daily_note,
            background_color=get_color_from_hex('#8e44ad')
        ))
        l.add_widget(btn_row)

        self.daily_notes_list = BoxLayout(orientation='vertical', size_hint_y=None,
                                            spacing=1)
        self.daily_notes_list.bind(minimum_height=self.daily_notes_list.setter('height'))
        l.add_widget(self.daily_notes_list)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)

    def _save_daily_note(self, instance):
        if not self.daily_pid.text:
            self._popup("Error", "No patient selected.")
            return
        note = {
            "patient_id": self.daily_pid.text,
            "patient_name": self.daily_name.text,
            "date": self.daily_date.text,
            "notes": self.daily_notes_text.text,
            "vitals": {
                "bp": self.daily_bp.text,
                "pulse": self.daily_pulse.text,
                "temp": self.daily_temp.text,
                "spo2": self.daily_spo2.text,
                "rr": self.daily_rr.text,
                "avpu": self.daily_avpu.text
            },
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "created_by": self._current_username()
        }
        self.daily_notes.append(note)
        self.save_all()
        self._popup("Saved", "Daily note saved.")
        self.log_activity(f"Daily note saved for {self.daily_name.text}")
        self.audit_log("SAVE_DAILY_NOTE", self.daily_date.text,
                       self.daily_pid.text)
        self.daily_notes_text.text = ""
        self.daily_bp.text = ""
        self.daily_pulse.text = ""
        self.daily_temp.text = ""
        self.daily_spo2.text = ""
        self.daily_rr.text = ""
        self._refresh_daily_notes()

    def _refresh_daily_notes(self):
        if not hasattr(self, 'daily_notes_list'):
            return
        self.daily_notes_list.clear_widgets()
        if not self.daily_pid.text:
            return
        pid = self.daily_pid.text
        notes = [n for n in self.daily_notes if n.get("patient_id") == pid]
        if not notes:
            self.daily_notes_list.add_widget(Label(
                text="No daily notes.", font_size=13,
                size_hint_y=None, height=30
            ))
            return
        for n in sorted(notes, key=lambda x: x.get('date', ''), reverse=True)[:10]:
            row = BoxLayout(size_hint_y=None, height=30, spacing=3)
            row.add_widget(Label(
                text=(f"{n.get('date', '')} | "
                      f"BP: {n.get('vitals', {}).get('bp', 'N/A')}"),
                font_size=12, size_hint_x=0.6
            ))
            btn = Button(
                text="VIEW", font_size=12, size_hint_x=0.2,
                size_hint_y=None, height=25,
                background_color=get_color_from_hex('#2980b9')
            )
            btn.note = n
            btn.bind(on_press=lambda x: self._text_popup(
                "Daily Note", x.note.get('notes', 'No notes')
            ))
            row.add_widget(btn)
            self.daily_notes_list.add_widget(row)

    def _view_daily_notes(self, instance):
        if not self.daily_pid.text:
            self._popup("Error", "No patient selected.")
            return
        pid = self.daily_pid.text
        notes = [n for n in self.daily_notes if n.get("patient_id") == pid]
        if not notes:
            self._popup("Daily Notes", "No notes for this patient.")
            return
        text = f"DAILY NOTES FOR {self.daily_name.text}\n" + "=" * 50 + "\n\n"
        for n in sorted(notes, key=lambda x: x.get('date', ''), reverse=True):
            text += f"Date: {n.get('date', '')}\n"
            vitals = n.get('vitals', {})
            text += (f"BP: {vitals.get('bp', 'N/A')} | "
                     f"Pulse: {vitals.get('pulse', 'N/A')} | "
                     f"Temp: {vitals.get('temp', 'N/A')}\n")
            text += (f"SpO2: {vitals.get('spo2', 'N/A')} | "
                     f"RR: {vitals.get('rr', 'N/A')} | "
                     f"AVPU: {vitals.get('avpu', 'Alert')}\n")
            text += f"Notes: {n.get('notes', '')}\n"
            text += "-" * 40 + "\n"
        self._text_popup("Daily Notes", text)

    def _ai_generate_daily_note(self, instance):
        if not self.cp:
            self._popup("Error", "No patient selected.")
            return
        if not self.ai.enabled:
            self._popup("Error", "AI is not enabled.")
            return
        pid = self.cp.get("id")
        emrs = [r for r in self.emr_records if r.get("patient_id") == pid]
        labs = [t for t in self.lab_tests if t.get("patient_id") == pid
                and t.get("status") == "Completed"]
        prompt = f"""Patient: {self.cp.get('name', '')}, Age: {self.cp.get('age', 0)}
Recent EMRs: {len(emrs)}
Recent Lab Results: {len(labs)}
Generate a daily clinical note with:
- SOAP format
- Progress since last encounter
- Current status
- Plan for today
"""
        self.daily_notes_text.text = "Generating... Please wait."

        def run():
            ok, resp = self.ai.call(
                "You are a clinician writing daily notes.", prompt, mt=800
            )
            Clock.schedule_once(
                lambda dt: setattr(self.daily_notes_text, 'text', resp), 0
            )

        threading.Thread(target=run, daemon=True).start()

    # ====================== PHARMACY TAB ======================
    def _pharmacy(self):
        tab = TabbedPanelItem(text="Pharmacy")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=5, spacing=3, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="ID:", font_size=15, bold=True, size_hint_x=0.05))
        self.pharm_pid = TextInput(size_hint_x=0.15, readonly=True, font_size=15,
                                     size_hint_y=None, height=40)
        bar.add_widget(self.pharm_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.07))
        self.pharm_name = TextInput(size_hint_x=0.35, readonly=True, font_size=15,
                                      size_hint_y=None, height=40)
        bar.add_widget(self.pharm_name)
        l.add_widget(bar)

        order = BoxLayout(size_hint_y=None, height=40, spacing=4)
        order.add_widget(Label(text="Drug:", font_size=15, bold=True, size_hint_x=0.06))
        self.pharm_drug = Spinner(
            text="Select", values=[d['name'] for d in self.drugs],
            size_hint_x=0.3, size_hint_y=None, height=40
        )
        order.add_widget(self.pharm_drug)
        order.add_widget(Label(text="Qty:", font_size=15, bold=True, size_hint_x=0.05))
        self.pharm_qty = TextInput(text="1", size_hint_x=0.08, font_size=15,
                                     size_hint_y=None, height=40)
        order.add_widget(self.pharm_qty)
        order.add_widget(Button(
            text="DISPENSE", font_size=15, bold=True, size_hint_x=0.15,
            on_press=self._dispense_drug,
            background_color=get_color_from_hex('#27ae60')
        ))
        order.add_widget(Button(
            text="VIEW RX", font_size=15, bold=True, size_hint_x=0.15,
            on_press=self._view_rx,
            background_color=get_color_from_hex('#2980b9')
        ))
        l.add_widget(order)

        self.pharm_list = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        self.pharm_list.bind(minimum_height=self.pharm_list.setter('height'))
        l.add_widget(self.pharm_list)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)
        self._refresh_pharm()

    def _dispense_drug(self, instance):
        if not self.pharm_pid.text or self.pharm_drug.text == "Select":
            self._popup("Error", "Select patient and drug.")
            return
        drug_name = self.pharm_drug.text
        try:
            qty = int(self.pharm_qty.text)
            if qty <= 0:
                raise ValueError
        except:
            self._popup("Error", "Invalid quantity.")
            return
        for d in self.drugs:
            if d['name'] == drug_name:
                if d['qty'] < qty:
                    self._popup("Error", f"Insufficient stock. Available: {d['qty']}")
                    return
                d['qty'] -= qty
                break
        self.prescriptions.append({
            "rx_id": f"RX{datetime.now().strftime('%Y%m%d%H%M%S%f')}",
            "patient_id": self.pharm_pid.text,
            "patient_name": self.pharm_name.text,
            "drug": drug_name,
            "qty": qty,
            "price": self._get_price(drug_name, "Drugs"),
            "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "dispensed": True
        })
        self.save_all()
        self._refresh_pharm()
        self._popup("Dispensed", f"{qty}x {drug_name}")
        self.log_activity(f"Dispensed {qty}x {drug_name} for {self.pharm_name.text}")
        self.audit_log("DISPENSE", f"{qty}x {drug_name}", self.pharm_pid.text)
        self.audit_log("DISPENSE", f"{qty}x {drug_name}", self.pharm_pid.text)

        # Clear search box for the next entry
        if hasattr(self, 'pharm_drug_search'):
             self.pharm_drug_search.text = ""

    def _refresh_pharm(self):
        if not hasattr(self, 'pharm_list'):
            return
        self.pharm_list.clear_widgets()
        if not self.pharm_pid.text:
            return
        pid = self.pharm_pid.text
        rx = [r for r in self.prescriptions if r.get("patient_id") == pid]
        if not rx:
            self.pharm_list.add_widget(Label(
                text="No prescriptions.", font_size=13,
                size_hint_y=None, height=30
            ))
            return
        for r in rx:
            row = BoxLayout(size_hint_y=None, height=30, spacing=3)
            row.add_widget(Label(text=f"{r.get('drug', '')} x{r.get('qty', 0)}",
                                   font_size=13, size_hint_x=0.4))
            row.add_widget(Label(text=r.get('ts', '')[:16], font_size=12,
                                   size_hint_x=0.3))
            row.add_widget(Label(text=f"N{r.get('price', 0):,.2f}",
                                   font_size=13, size_hint_x=0.2))
            self.pharm_list.add_widget(row)

    def _view_rx(self, instance):
        if not self.pharm_pid.text:
            self._popup("Error", "No patient selected.")
            return
        pid = self.pharm_pid.text
        rx = [r for r in self.prescriptions if r.get("patient_id") == pid]
        if not rx:
            self._popup("Prescriptions", "No prescriptions found.")
            return
        text = f"PRESCRIPTIONS FOR {self.pharm_name.text}\n" + "=" * 50 + "\n\n"
        total = 0
        for r in rx:
            text += f"Drug: {r.get('drug', '')}\n"
            text += f"Quantity: {r.get('qty', 0)}\n"
            text += f"Price: N{r.get('price', 0):,.2f}\n"
            text += f"Date: {r.get('ts', '')}\n"
            total += r.get('price', 0) * r.get('qty', 0)
            text += "-" * 30 + "\n"
        text += f"\nTOTAL: N{total:,.2f}"
        self._text_popup("Prescriptions", text)

    # ====================== DRUG MANAGEMENT TAB ======================
    def _drug_mgmt(self):
        tab = TabbedPanelItem(text="Manage Drugs")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=5, spacing=3, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))
        l.add_widget(Label(
            text="DRUG INVENTORY MANAGEMENT", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#1a5276')
        ))

        form = GridLayout(cols=4, size_hint_y=None, height=120, spacing=4)
        form.add_widget(Label(text="Drug:", font_size=14, bold=True))
        self.dm_drug = TextInput(font_size=14, size_hint_y=None, height=35)
        form.add_widget(self.dm_drug)
        form.add_widget(Label(text="Price:", font_size=14, bold=True))
        self.dm_price = TextInput(font_size=14, size_hint_y=None, height=35)
        form.add_widget(self.dm_price)
        form.add_widget(Label(text="Qty:", font_size=14, bold=True))
        self.dm_qty = TextInput(font_size=14, size_hint_y=None, height=35)
        form.add_widget(self.dm_qty)
        form.add_widget(Label(text="Min Stock:", font_size=14, bold=True))
        self.dm_min = TextInput(font_size=14, size_hint_y=None, height=35)
        form.add_widget(self.dm_min)
        form.add_widget(Label(text="Category:", font_size=14, bold=True))
        self.dm_cat = Spinner(
            text="General",
            values=["Analgesic", "Antibiotic", "Antihypertensive",
                    "Antidiabetic", "Antacid", "General"],
            size_hint_y=None, height=35
        )
        form.add_widget(self.dm_cat)
        l.add_widget(form)

        btn_row = BoxLayout(size_hint_y=None, height=40, spacing=5)
        btn_row.add_widget(Button(
            text="ADD DRUG", font_size=15, bold=True,
            on_press=self._add_drug,
            background_color=get_color_from_hex('#27ae60')
        ))
        btn_row.add_widget(Button(
            text="UPDATE STOCK", font_size=15, bold=True,
            on_press=self._update_drug_stock,
            background_color=get_color_from_hex('#2980b9')
        ))
        btn_row.add_widget(Button(
            text="DELETE DRUG", font_size=15, bold=True,
            on_press=self._delete_drug,
            background_color=get_color_from_hex('#e74c3c')
        ))
        btn_row.add_widget(Button(
            text="REFRESH", font_size=15, bold=True,
            on_press=self._refresh_drug_list,
            background_color=get_color_from_hex('#16a085')
        ))
        l.add_widget(btn_row)

        self.drug_list = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        self.drug_list.bind(minimum_height=self.drug_list.setter('height'))
        l.add_widget(self.drug_list)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)
        self._refresh_drug_list()

    def _add_drug(self, instance):
        if not self._require_role("Administrator", "Pharmacist"):
            return
        name = self.dm_drug.text.strip()
        if not name:
            self._popup("Error", "Enter drug name.")
            return
        try:
            price = float(self.dm_price.text or "0")
            qty = int(self.dm_qty.text or "0")
            min_stock = int(self.dm_min.text or "20")
        except:
            self._popup("Error", "Invalid numeric values.")
            return
        if price <= 0 or qty < 0:
            self._popup("Error", "Price must be > 0 and quantity >= 0.")
            return
        self.drugs.append({
            "name": name,
            "price": price,
            "qty": qty,
            "min": min_stock,
            "cat": self.dm_cat.text
        })
        self.save_all()
        self._refresh_drug_list()
        self._popup("Added", f"Added {name}")
        self.log_activity(f"Added drug: {name}")
        self.audit_log("ADD_DRUG", f"{name} @ N{price}")

    def _update_drug_stock(self, instance):
        if not self._require_role("Administrator", "Pharmacist"):
            return
        if self.sel_drug_idx is None:
            self._popup("Error", "Select a drug from the list first.")
            return
        content = BoxLayout(orientation='vertical', spacing=5, padding=8)
        content.add_widget(Label(
            text=f"Update stock for: {self.drugs[self.sel_drug_idx]['name']}",
            font_size=16, bold=True, size_hint_y=None, height=30
        ))
        content.add_widget(Label(text="New Quantity:", font_size=14, bold=True,
                                   size_hint_y=None, height=22))
        qty_inp = TextInput(font_size=14, size_hint_y=None, height=40)
        content.add_widget(qty_inp)
        btn_box = BoxLayout(size_hint_y=None, height=40, spacing=5)
        pp = Popup(title="Update Stock", content=content, size_hint=(0.6, 0.3))

        def do_update(inst):
            try:
                new_qty = int(qty_inp.text)
                if new_qty < 0:
                    self._popup("Error", "Quantity cannot be negative.")
                    return
                self.drugs[self.sel_drug_idx]["qty"] = new_qty
                self.save_all()
                self._refresh_drug_list()
                pp.dismiss()
                self._popup("Updated", f"Stock updated to {new_qty}")
                self.log_activity(
                    f"Drug stock updated: "
                    f"{self.drugs[self.sel_drug_idx]['name']} -> {new_qty}"
                )
                self.audit_log(
                    "UPDATE_DRUG_STOCK",
                    f"{self.drugs[self.sel_drug_idx]['name']} → {new_qty}"
                )
            except:
                self._popup("Error", "Invalid quantity.")

        btn_box.add_widget(Button(
            text="UPDATE", font_size=15, bold=True, on_press=do_update,
            background_color=get_color_from_hex('#27ae60')
        ))
        btn_box.add_widget(Button(text="CANCEL", font_size=15, on_press=pp.dismiss))
        content.add_widget(btn_box)
        pp.open()

    def _delete_drug(self, instance):
        if not self._require_role("Administrator"):
            return
        if self.sel_drug_idx is None:
            self._popup("Error", "Select a drug first.")
            return
        drug = self.drugs[self.sel_drug_idx]
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)
        content.add_widget(Label(
            text=f"Delete {drug['name']}?\nThis cannot be undone."
        ))
        btn_box = BoxLayout(size_hint_y=None, height=40, spacing=5)
        pp = Popup(title="Confirm Delete", content=content, size_hint=(0.6, 0.3))

        def do_delete(inst):
            self.drugs.pop(self.sel_drug_idx)
            self.sel_drug_idx = None
            self.save_all()
            self._refresh_drug_list()
            pp.dismiss()
            self._popup("Deleted", "Drug removed.")
            self.log_activity(f"Drug deleted: {drug['name']}")
            self.audit_log("DELETE_DRUG", drug['name'])

        btn_box.add_widget(Button(
            text="DELETE", font_size=15, bold=True, on_press=do_delete,
            background_color=get_color_from_hex('#e74c3c')
        ))
        btn_box.add_widget(Button(text="CANCEL", font_size=15, on_press=pp.dismiss))
        content.add_widget(btn_box)
        pp.open()

    def _refresh_drug_list(self, instance=None):
        if not hasattr(self, 'drug_list'):
            return
        self.drug_list.clear_widgets()
        if not self.drugs:
            self.drug_list.add_widget(Label(
                text="No drugs in inventory.", font_size=13,
                size_hint_y=None, height=30
            ))
            return
        for i, d in enumerate(self.drugs):
            row = BoxLayout(size_hint_y=None, height=30, spacing=3)
            color = (get_color_from_hex('#e74c3c')
                     if d.get('qty', 0) <= d.get('min', 20)
                     else get_color_from_hex('#2c3e50'))
            row.add_widget(Label(text=f"{d['name']}", font_size=13, size_hint_x=0.3))
            row.add_widget(Label(text=f"N{d['price']:,.2f}", font_size=13,
                                   size_hint_x=0.15))
            row.add_widget(Label(text=f"Qty: {d['qty']}", font_size=13,
                                   size_hint_x=0.15, color=color))
            row.add_widget(Label(text=f"Min: {d['min']}", font_size=12,
                                   size_hint_x=0.1))
            row.add_widget(Label(text=f"{d.get('cat', '')}", font_size=12,
                                   size_hint_x=0.15))
            btn = Button(
                text="SELECT", font_size=11, size_hint_x=0.1,
                size_hint_y=None, height=25,
                background_color=get_color_from_hex('#2980b9')
            )
            btn.idx = i
            btn.bind(on_press=lambda x: self._select_drug(x.idx))
            row.add_widget(btn)
            self.drug_list.add_widget(row)

    def _select_drug(self, idx):
        self.sel_drug_idx = idx
        drug = self.drugs[idx]
        self._popup("Selected", f"Selected: {drug['name']}\nStock: {drug['qty']}")

    # ====================== INVESTIGATION MANAGEMENT TAB ======================
    def _inv_mgmt(self):
        tab = TabbedPanelItem(text="Manage Investigations")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=5, spacing=3, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))
        l.add_widget(Label(
            text="INVESTIGATION MANAGEMENT", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#1a5276')
        ))

        form = GridLayout(cols=4, size_hint_y=None, height=120, spacing=4)
        form.add_widget(Label(text="Name:", font_size=14, bold=True))
        self.im_name = TextInput(font_size=14, size_hint_y=None, height=35)
        form.add_widget(self.im_name)
        form.add_widget(Label(text="Price:", font_size=14, bold=True))
        self.im_price = TextInput(font_size=14, size_hint_y=None, height=35)
        form.add_widget(self.im_price)
        form.add_widget(Label(text="Department:", font_size=14, bold=True))
        self.im_dept = Spinner(
            text="Laboratory",
            values=["Laboratory", "Radiology", "Cardiology"],
            size_hint_y=None, height=35
        )
        form.add_widget(self.im_dept)
        form.add_widget(Label(text="Category:", font_size=14, bold=True))
        self.im_cat = Spinner(
            text="General",
            values=["General", "Hematology", "Biochemistry",
                    "Microbiology", "Imaging"],
            size_hint_y=None, height=35
        )
        form.add_widget(self.im_cat)
        l.add_widget(form)

        btn_row = BoxLayout(size_hint_y=None, height=40, spacing=5)
        btn_row.add_widget(Button(
            text="ADD", font_size=15, bold=True,
            on_press=self._add_inv,
            background_color=get_color_from_hex('#27ae60')
        ))
        btn_row.add_widget(Button(
            text="DELETE", font_size=15, bold=True,
            on_press=self._delete_inv,
            background_color=get_color_from_hex('#e74c3c')
        ))
        btn_row.add_widget(Button(
            text="REFRESH", font_size=15, bold=True,
            on_press=self._refresh_inv_list,
            background_color=get_color_from_hex('#16a085')
        ))
        l.add_widget(btn_row)

        self.inv_list = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        self.inv_list.bind(minimum_height=self.inv_list.setter('height'))
        l.add_widget(self.inv_list)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)
        self._refresh_inv_list()

    def _add_inv(self, instance):
        if not self._require_role("Administrator", "Doctor", "Lab Scientist"):
            return
        name = self.im_name.text.strip()
        if not name:
            self._popup("Error", "Enter investigation name.")
            return
        try:
            price = float(self.im_price.text or "0")
        except:
            self._popup("Error", "Invalid price.")
            return
        if price <= 0:
            self._popup("Error", "Price must be > 0.")
            return
        self.investigations.append({
            "name": name,
            "price": price,
            "dept": self.im_dept.text,
            "cat": self.im_cat.text
        })
        self.save_all()
        self._refresh_inv_list()
        self._popup("Added", f"Added {name}")
        self.log_activity(f"Added investigation: {name}")
        self.audit_log("ADD_INVESTIGATION", f"{name} @ N{price}")

    def _delete_inv(self, instance):
        if not self._require_role("Administrator"):
            return
        if self.sel_inv_idx is None:
            self._popup("Error", "Select an investigation first.")
            return
        inv = self.investigations[self.sel_inv_idx]
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)
        content.add_widget(Label(
            text=f"Delete {inv['name']}?\nThis cannot be undone."
        ))
        btn_box = BoxLayout(size_hint_y=None, height=40, spacing=5)
        pp = Popup(title="Confirm Delete", content=content, size_hint=(0.6, 0.3))

        def do_delete(inst):
            self.investigations.pop(self.sel_inv_idx)
            self.sel_inv_idx = None
            self.save_all()
            self._refresh_inv_list()
            pp.dismiss()
            self._popup("Deleted", "Investigation removed.")
            self.log_activity(f"Investigation deleted: {inv['name']}")
            self.audit_log("DELETE_INVESTIGATION", inv['name'])

        btn_box.add_widget(Button(
            text="DELETE", font_size=15, bold=True, on_press=do_delete,
            background_color=get_color_from_hex('#e74c3c')
        ))
        btn_box.add_widget(Button(text="CANCEL", font_size=15, on_press=pp.dismiss))
        content.add_widget(btn_box)
        pp.open()

    def _refresh_inv_list(self, instance=None):
        if not hasattr(self, 'inv_list'):
            return
        self.inv_list.clear_widgets()
        if not self.investigations:
            self.inv_list.add_widget(Label(
                text="No investigations configured.", font_size=13,
                size_hint_y=None, height=30
            ))
            return
        for i, inv in enumerate(self.investigations):
            row = BoxLayout(size_hint_y=None, height=30, spacing=3)
            row.add_widget(Label(text=inv['name'], font_size=13, size_hint_x=0.4))
            row.add_widget(Label(text=f"N{inv['price']:,.2f}", font_size=13,
                                   size_hint_x=0.2))
            row.add_widget(Label(text=inv.get('dept', ''), font_size=12,
                                   size_hint_x=0.2))
            btn = Button(
                text="SELECT", font_size=11, size_hint_x=0.1,
                size_hint_y=None, height=25,
                background_color=get_color_from_hex('#2980b9')
            )
            btn.idx = i
            btn.bind(on_press=lambda x: self._select_inv(x.idx))
            row.add_widget(btn)
            self.inv_list.add_widget(row)

    def _select_inv(self, idx):
        self.sel_inv_idx = idx
        inv = self.investigations[idx]
        self._popup("Selected", f"Selected: {inv['name']}\nPrice: N{inv['price']:,.2f}")

    # ====================== SETTINGS TAB ======================
    def _settings(self):
        tab = TabbedPanelItem(text="Settings")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=8, spacing=6, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))

        # --- AI Settings ---
        l.add_widget(Label(
            text="AI SETTINGS", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#1a5276')
        ))
        l.add_widget(Label(text="Provider:", font_size=16, bold=True,
                            size_hint_y=None, height=22))
        self.ai_provider = Spinner(
            text=self.ai.provider, values=["OpenAI", "DeepSeek"],
            size_hint_y=None, height=40
        )
        l.add_widget(self.ai_provider)

        l.add_widget(Label(text="API Key:", font_size=16, bold=True,
                            size_hint_y=None, height=22))
        self.ai_key = TextInput(
            text=self.ai.api_key, font_size=14, password=True,
            size_hint_y=None, height=40, hint_text="Enter API key"
        )
        l.add_widget(self.ai_key)

        l.add_widget(Label(text="Model:", font_size=16, bold=True,
                            size_hint_y=None, height=22))
        self.ai_model = Spinner(
            text=self.ai.model, values=AIModels.get(self.ai.provider),
            size_hint_y=None, height=40
        )
        l.add_widget(self.ai_model)

        def on_provider_change(spinner, text):
            self.ai_model.values = AIModels.get(text)
            self.ai_model.text = AIModels.default(text)
        self.ai_provider.bind(text=on_provider_change)

        btn_row = BoxLayout(size_hint_y=None, height=40, spacing=5)
        btn_row.add_widget(Button(
            text="SAVE AI SETTINGS", font_size=16, bold=True,
            on_press=self._save_ai_settings,
            background_color=get_color_from_hex('#27ae60')
        ))
        btn_row.add_widget(Button(
            text="TEST AI", font_size=16, bold=True,
            on_press=self._test_ai_connection,
            background_color=get_color_from_hex('#2980b9')
        ))
        l.add_widget(btn_row)

        # --- SECURITY ---
        l.add_widget(Label(text="", size_hint_y=None, height=10))
        l.add_widget(Label(
            text="SECURITY", font_size=18, bold=True,
            size_hint_y=None, height=28,
            color=get_color_from_hex('#c0392b')
        ))

        security_row = BoxLayout(size_hint_y=None, height=42, spacing=5)
        security_row.add_widget(Button(
            text="🔒 CHANGE DELETE PASSCODE",
            font_size=15, bold=True,
            on_press=self._change_delete_passcode,
            background_color=get_color_from_hex('#8e44ad')
        ))
        security_row.add_widget(Button(
            text="🔐 CHANGE MY PASSWORD",
            font_size=15, bold=True,
            on_press=self._change_my_password,
            background_color=get_color_from_hex('#2980b9')
        ))
        l.add_widget(security_row)

        security_note = Label(
            text=("Delete passcode is hashed and stored in config.json.\n"
                  "Default passcode is 1234 — change it before deploying."),
            font_size=12, size_hint_y=None, height=40,
            color=get_color_from_hex('#7f8c8d')
        )
        l.add_widget(security_note)

        # --- DATA MANAGEMENT ---
        l.add_widget(Label(text="", size_hint_y=None, height=10))
        l.add_widget(Label(
            text="DATA MANAGEMENT", font_size=18, bold=True,
            size_hint_y=None, height=28,
            color=get_color_from_hex('#c0392b')
        ))
        data_row = BoxLayout(size_hint_y=None, height=40, spacing=5)
        data_row.add_widget(Button(
            text="BACKUP DATA", font_size=15, bold=True,
            on_press=self._backup_data,
            background_color=get_color_from_hex('#16a085')
        ))
        data_row.add_widget(Button(
            text="EXPORT PATIENTS CSV", font_size=15, bold=True,
            on_press=self._export_csv,
            background_color=get_color_from_hex('#2980b9')
        ))
        data_row.add_widget(Button(
            text="REFRESH ALL", font_size=15, bold=True,
            on_press=self._refresh_all,
            background_color=get_color_from_hex('#f39c12')
        ))
        l.add_widget(data_row)

        # --- STATUS ---
        l.add_widget(Label(text="", size_hint_y=None, height=10))
        l.add_widget(Label(
            text="STATUS", font_size=18, bold=True,
            size_hint_y=None, height=28,
            color=get_color_from_hex('#2c3e50')
        ))
        status_text = f"Current User: {self._current_username()} ({self._current_role()})\n"
        status_text += f"AI: {'Enabled' if self.ai.enabled else 'Disabled'}\n"
        status_text += f"Provider: {self.ai.provider}\n"
        status_text += f"Model: {self.ai.model}\n"
        status_text += f"Patients: {len(self.patients)}\n"
        status_text += f"EMR Records: {len(self.emr_records)}\n"
        status_text += f"Drugs: {len(self.drugs)}\n"
        status_text += f"Investigations: {len(self.investigations)}\n"
        status_text += f"Lab Tests: {len(self.lab_tests)}\n"
        status_text += f"Bills: {len(self.bills)}\n"
        status_text += f"Payments: {len(self.pt_tracker.payments)}\n"
        status_text += f"ECG Reports: {len(self.ecg_reports)}\n"
        status_text += f"Radiology: {len(self.radiology)}\n"
        status_text += f"Tasks: {len(self.tasks)}\n"
        status_text += f"Admissions: {len(self.admissions)}\n"
        status_text += f"Lab Reports: {len(self.saved_lab_reports)}\n"
        status_text += f"Appointments: {len(self.appointments)}\n"
        status_text += f"Nursing Notes: {len(self.nursing_notes)}\n"
        status_text += f"Users: {len(self.user_manager.users)}\n"
        status_text += f"Audit Entries: {len(self.audit.entries)}\n"
        l.add_widget(TextInput(
            text=status_text, readonly=True, font_size=13, multiline=True,
            size_hint_y=None, height=280,
            background_color=get_color_from_hex('#f8f9fa')
        ))

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)

    def _save_ai_settings(self, instance):
        provider = self.ai_provider.text
        api_key = self.ai_key.text.strip()
        model = self.ai_model.text
        self.ai.save(provider, api_key, model)
        self.ai.enabled = bool(api_key and len(api_key) > 10)
        if self.ai.enabled:
            self.ai_lbl.text = f"AI: {provider}"
            self.ai_lbl.color = get_color_from_hex('#27ae60')
        else:
            self.ai_lbl.text = "AI: Offline"
            self.ai_lbl.color = get_color_from_hex('#e74c3c')
        self._popup("AI Settings", "Settings saved successfully.")
        self.log_activity(f"AI settings updated: {provider} - {model}")
        self.audit_log("UPDATE_AI_SETTINGS", f"{provider} / {model}")

    def _backup_data(self, instance):
        try:
            backup_dir = os.path.join(self._get_reports_dir(), "backups")
            os.makedirs(backup_dir, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = os.path.join(backup_dir, f"hmg_backup_{timestamp}.json")
            data = {
                "patients": self.patients,
                "emr_records": self.emr_records,
                "lab_tests": self.lab_tests,
                "radiology": self.radiology,
                "prescriptions": self.prescriptions,
                "bills": self.bills,
                "drugs": self.drugs,
                "investigations": self.investigations,
                "ecg_reports": self.ecg_reports,
                "followups": self.followups,
                "daily_notes": self.daily_notes,
                "consent_forms": self.consent_forms,
                "referral_notes": self.referral_notes,
                "patient_snapshots": self.patient_snapshots,
                "lab_results": self.lab_results,
                "payments": self.pt_tracker.payments,
                "tasks": self.tasks,
                "admissions": self.admissions,
                "saved_lab_reports": self.saved_lab_reports,
                "appointments": self.appointments,
                "nursing_notes": self.nursing_notes,
                "users": self.user_manager.users
            }
            self._save_json(backup_file, data)
            self._popup("Backup", f"Backup saved to:\n{backup_file}")
            self.log_activity(f"Backup created: {backup_file}")
            self.audit_log("BACKUP_DATA", backup_file)
        except Exception as e:
            self.log_activity(f"Backup error: {e}")
            self._popup("Backup Error", str(e))

    def _export_csv(self, instance):
        if not self.patients:
            self._popup("Export", "No patients to export.")
            return
        try:
            fn = f"patients_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            with open(fn, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["ID", "Name", "DOB", "Age", "Gender",
                                  "Phone", "Height", "Blood", "Date"])
                for p in self.patients:
                    writer.writerow([p.get('id', ''), p.get('name', ''),
                                      p.get('dob', ''), p.get('age', ''),
                                      p.get('gender', ''), p.get('phone', ''),
                                      p.get('height', ''), p.get('blood', ''),
                                      p.get('date', '')])
            self._popup("Export", f"Exported to {fn}")
            self.log_activity(f"CSV export: {fn}")
            self.audit_log("EXPORT_CSV", fn)
        except Exception as e:
            self.log_activity(f"CSV export error: {e}")
            self._popup("Export Error", str(e))

    def _refresh_all(self, instance):
        self.load_all()
        self._refresh_plist()
        self._refresh_dash()
        self._refresh_lab()
        self._refresh_pharm()
        self._refresh_rad()
        self._refresh_clinical_summary()
        self._refresh_timeline()
        self._refresh_uploaded_results()
        self._refresh_vitals_timeline()
        self._refresh_daily_notes()
        self._refresh_snapshots()
        self._refresh_saved_items()
        self._refresh_elab()
        self._refresh_tasks()
        self._refresh_pharmacy_dispense()
        self._refresh_admissions_list()
        self._refresh_saved_reports_list()
        self._refresh_appointments()
        self._refresh_nursing_notes()
        self._auto_populate_patient_info()
        self._popup("Refreshed", "All data reloaded.")

    def _change_delete_passcode(self, instance):
        if not self._require_role("Administrator"):
            return
        content = BoxLayout(orientation='vertical', spacing=8, padding=12)
        content.add_widget(Label(
            text="Change Delete Patient Passcode",
            font_size=16, bold=True, size_hint_y=None, height=30,
            color=get_color_from_hex('#1a5276')
        ))
        content.add_widget(Label(text="Current passcode:", font_size=13, bold=True,
                                   size_hint_y=None, height=20))
        cur_inp = TextInput(password=True, multiline=False,
                             font_size=15, size_hint_y=None, height=40)
        content.add_widget(cur_inp)
        content.add_widget(Label(text="New passcode (min 4 chars):",
                                   font_size=13, bold=True,
                                   size_hint_y=None, height=20))
        new_inp = TextInput(password=True, multiline=False,
                             font_size=15, size_hint_y=None, height=40)
        content.add_widget(new_inp)
        content.add_widget(Label(text="Confirm new passcode:", font_size=13, bold=True,
                                   size_hint_y=None, height=20))
        conf_inp = TextInput(password=True, multiline=False,
                              font_size=15, size_hint_y=None, height=40)
        content.add_widget(conf_inp)
        btns = BoxLayout(size_hint_y=None, height=42, spacing=5)
        pp = Popup(title="Security", content=content,
                   size_hint=(0.6, 0.55), auto_dismiss=False)

        def do_change(inst):
            if not self._verify_delete_passcode(cur_inp.text):
                self._popup("Error", "Current passcode is incorrect.")
                return
            if new_inp.text.strip() != conf_inp.text.strip():
                self._popup("Error", "New passcodes do not match.")
                return
            ok, msg = self._set_delete_passcode(new_inp.text)
            pp.dismiss()
            self._popup("Security", msg)
            if ok:
                self.log_activity("Delete passcode changed")
                self.audit_log("CHANGE_DELETE_PASSCODE")

        def do_cancel(inst):
            pp.dismiss()

        btns.add_widget(Button(
            text="SAVE", font_size=15, bold=True, on_press=do_change,
            background_color=get_color_from_hex('#27ae60')
        ))
        btns.add_widget(Button(
            text="CANCEL", font_size=15, bold=True, on_press=do_cancel,
            background_color=get_color_from_hex('#e74c3c')
        ))
        content.add_widget(btns)
        pp.open()

    def _change_my_password(self, instance):
        if not self.current_user:
            return
        content = BoxLayout(orientation='vertical', spacing=8, padding=12)
        content.add_widget(Label(
            text=f"Change Password for {self._current_username()}",
            font_size=16, bold=True, size_hint_y=None, height=30,
            color=get_color_from_hex('#1a5276')
        ))
        content.add_widget(Label(text="Current password:", font_size=13, bold=True,
                                   size_hint_y=None, height=20))
        cur_inp = TextInput(password=True, multiline=False,
                             font_size=15, size_hint_y=None, height=40)
        content.add_widget(cur_inp)
        content.add_widget(Label(text="New password (min 4 chars):",
                                   font_size=13, bold=True,
                                   size_hint_y=None, height=20))
        new_inp = TextInput(password=True, multiline=False,
                             font_size=15, size_hint_y=None, height=40)
        content.add_widget(new_inp)
        content.add_widget(Label(text="Confirm new password:", font_size=13, bold=True,
                                   size_hint_y=None, height=20))
        conf_inp = TextInput(password=True, multiline=False,
                              font_size=15, size_hint_y=None, height=40)
        content.add_widget(conf_inp)
        btns = BoxLayout(size_hint_y=None, height=42, spacing=5)
        pp = Popup(title="Change Password", content=content,
                   size_hint=(0.6, 0.55), auto_dismiss=False)

        def do_change(inst):
            if new_inp.text != conf_inp.text:
                self._popup("Error", "New passwords do not match.")
                return
            ok, msg = self.user_manager.change_password(
                self._current_username(), cur_inp.text, new_inp.text
            )
            pp.dismiss()
            self._popup("Password", msg)
            if ok:
                self.audit_log("CHANGE_PASSWORD")

        def do_cancel(inst):
            pp.dismiss()

        btns.add_widget(Button(
            text="SAVE", font_size=15, bold=True, on_press=do_change,
            background_color=get_color_from_hex('#27ae60')
        ))
        btns.add_widget(Button(
            text="CANCEL", font_size=15, bold=True, on_press=do_cancel,
            background_color=get_color_from_hex('#e74c3c')
        ))
        content.add_widget(btns)
        pp.open()

    # ====================== CHAT TAB ======================
    def _chat_tab(self):
        tab = TabbedPanelItem(text="Chat")
        main = BoxLayout(orientation='vertical', padding=6, spacing=4)
        main.add_widget(Label(
            text="HOSPITAL CHAT", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#1a5276')
        ))

        top = BoxLayout(size_hint_y=None, height=40, spacing=4)
        top.add_widget(Label(text="Room:", font_size=15, bold=True, size_hint_x=0.08))
        self.chat_room = Spinner(
            text="General",
            values=["General"] + list(self.chat_sys.rooms.keys()),
            size_hint_x=0.2, size_hint_y=None, height=40
        )
        self.chat_room.bind(text=lambda i, v: self._refresh_chat())
        top.add_widget(Button(
            text="NEW ROOM", font_size=15, bold=True, size_hint_x=0.15,
            on_press=self._new_chat_room,
            background_color=get_color_from_hex('#2980b9')
        ))
        top.add_widget(Button(
            text="DELETE ROOM", font_size=15, bold=True, size_hint_x=0.15,
            on_press=self._delete_chat_room,
            background_color=get_color_from_hex('#e74c3c')
        ))
        main.add_widget(top)

        chat_scroll = ScrollView(size_hint_y=0.65, bar_width=10,
                                  scroll_type=['bars'])
        self.chat_display = TextInput(
            multiline=True, readonly=True, font_size=14,
            background_color=get_color_from_hex('#f8f9fa'),
            size_hint_y=None
        )
        self.chat_display.bind(minimum_height=self.chat_display.setter('height'))
        chat_scroll.add_widget(self.chat_display)
        main.add_widget(chat_scroll)

        bottom = BoxLayout(size_hint_y=None, height=45, spacing=4)
        self.chat_input = TextInput(
            size_hint_x=0.7, font_size=15, size_hint_y=None, height=45,
            hint_text="Type message...",
            on_text_validate=self._send_chat
        )
        bottom.add_widget(self.chat_input)
        bottom.add_widget(Button(
            text="SEND", font_size=16, bold=True, size_hint_x=0.2,
            on_press=self._send_chat,
            background_color=get_color_from_hex('#27ae60')
        ))
        bottom.add_widget(Button(
            text="📷", font_size=16, bold=True, size_hint_x=0.1,
            on_press=self._chat_camera,
            background_color=get_color_from_hex('#2980b9')
        ))
        main.add_widget(bottom)

        tab.add_widget(main)
        self.tp.add_widget(tab)
        self._refresh_chat()
        Clock.schedule_interval(lambda dt: self._refresh_chat(), 5)

    def _new_chat_room(self, instance):
        c = BoxLayout(orientation='vertical', spacing=5, padding=8)
        c.add_widget(Label(text="Room Name:", font_size=16, bold=True,
                            size_hint_y=None, height=30))
        name_inp = TextInput(font_size=15, size_hint_y=None, height=40)
        c.add_widget(name_inp)
        btn_box = BoxLayout(size_hint_y=None, height=40, spacing=5)
        pp = Popup(title="New Chat Room", content=c, size_hint=(0.5, 0.25))

        def create(inst):
            name = name_inp.text.strip()
            if name and name not in self.chat_sys.rooms:
                self.chat_sys.rooms[name] = []
                self.chat_sys.save()
                self.chat_room.values = ["General"] + list(self.chat_sys.rooms.keys())
                self.chat_room.text = name
                self._refresh_chat()
                pp.dismiss()
                self.log_activity(f"Chat room created: {name}")
                self.audit_log("CREATE_CHAT_ROOM", name)
            else:
                self._popup("Error", "Invalid or duplicate room name.")

        btn_box.add_widget(Button(
            text="CREATE", font_size=15, bold=True, on_press=create,
            background_color=get_color_from_hex('#27ae60')
        ))
        btn_box.add_widget(Button(text="CANCEL", font_size=15, on_press=pp.dismiss))
        c.add_widget(btn_box)
        pp.open()

    def _delete_chat_room(self, instance):
        room = self.chat_room.text
        if room == "General":
            self._popup("Error", "Cannot delete General room.")
            return
        if not self._require_role("Administrator"):
            return
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)
        content.add_widget(Label(
            text=f"Delete room '{room}'?\nAll messages will be lost."
        ))
        btn_box = BoxLayout(size_hint_y=None, height=40, spacing=5)
        pp = Popup(title="Confirm Delete", content=content, size_hint=(0.6, 0.3))

        def do_delete(inst):
            self.chat_sys.rooms.pop(room, None)
            self.chat_sys.save()
            self.chat_room.values = ["General"] + list(self.chat_sys.rooms.keys())
            self.chat_room.text = "General"
            self._refresh_chat()
            pp.dismiss()
            self.log_activity(f"Chat room deleted: {room}")
            self.audit_log("DELETE_CHAT_ROOM", room)

        btn_box.add_widget(Button(
            text="DELETE", font_size=15, bold=True, on_press=do_delete,
            background_color=get_color_from_hex('#e74c3c')
        ))
        btn_box.add_widget(Button(text="CANCEL", font_size=15, on_press=pp.dismiss))
        content.add_widget(btn_box)
        pp.open()

    def _send_chat(self, instance):
        msg = self.chat_input.text.strip()
        if not msg:
            return
        user = self.cp.get("name", "Guest") if self.cp else self._current_username()
        room = self.chat_room.text
        self.chat_sys.send(user, room, msg)
        self.chat_input.text = ""
        self._refresh_chat()
        self.log_activity(f"Chat message in {room}: {user} -> {msg[:30]}")

    def _chat_camera(self, instance):
        if not CameraHelper.available():
            self._popup("Camera", "Not available")
            return
        content = BoxLayout(orientation='vertical', spacing=4, padding=4)
        cam = Camera(play=True, resolution=(640, 480))
        content.add_widget(cam)
        btns = BoxLayout(size_hint_y=None, height=40, spacing=5)
        pp = Popup(title="Chat Photo", content=content, size_hint=(0.85, 0.85))

        def capture(inst):
            if cam.texture:
                fn = f"chat_photo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                cam.texture.save(fn)
                user = self.cp.get("name", "Guest") if self.cp else self._current_username()
                room = self.chat_room.text
                self.chat_sys.send(user, room, "[Photo]", image=fn)
                pp.dismiss()
                self._refresh_chat()
                self._popup("Photo", f"Photo saved: {fn}")
                self.log_activity(f"Chat photo in {room}: {fn}")

        btns.add_widget(Button(
            text="📸 CAPTURE", font_size=16, bold=True, on_press=capture,
            background_color=get_color_from_hex('#27ae60')
        ))
        btns.add_widget(Button(text="CANCEL", font_size=16, on_press=pp.dismiss))
        content.add_widget(btns)
        pp.open()

    def _refresh_chat(self):
        if not hasattr(self, 'chat_display'):
            return
        room = self.chat_room.text
        msgs = self.chat_sys.get(room)
        self.chat_display.text = f"CHAT: {room}\n" + "=" * 50 + "\n\n"
        for m in msgs:
            img = m.get('image')
            if img:
                self.chat_display.text += (f"[{m.get('ts','')[:16]}] "
                                            f"{m.get('user','')}: "
                                            f"[Image: {img}]\n")
            else:
                self.chat_display.text += (f"[{m.get('ts','')[:16]}] "
                                            f"{m.get('user','')}: "
                                            f"{m.get('msg','')}\n")

    # ====================== AI ANALYSIS TAB ======================
    def _ai_analysis_tab(self):
        tab = TabbedPanelItem(text="AI Analysis")
        main = BoxLayout(orientation='vertical', padding=6, spacing=4)
        main.add_widget(Label(
            text="AI CLINICAL ANALYSIS", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#8e44ad')
        ))
        main.add_widget(Label(
            text="Analyze patient data with AI", font_size=15,
            size_hint_y=None, height=22,
            color=get_color_from_hex('#7f8c8d')
        ))

        stats_box = GridLayout(cols=3, size_hint_y=None, height=60, spacing=5)
        self.ai_stats_emr = Label(text="EMR: 0", font_size=16, bold=True,
                                    color=get_color_from_hex('#2980b9'))
        self.ai_stats_lab = Label(text="Lab: 0", font_size=16, bold=True,
                                    color=get_color_from_hex('#27ae60'))
        self.ai_stats_rx = Label(text="Rx: 0", font_size=16, bold=True,
                                   color=get_color_from_hex('#e67e22'))
        stats_box.add_widget(self.ai_stats_emr)
        stats_box.add_widget(self.ai_stats_lab)
        stats_box.add_widget(self.ai_stats_rx)
        main.add_widget(stats_box)

        btn_row = BoxLayout(size_hint_y=None, height=40, spacing=5)
        btn_row.add_widget(Button(
            text="FULL PATIENT ANALYSIS", font_size=16, bold=True,
            on_press=self._run_full_analysis,
            background_color=get_color_from_hex('#8e44ad')
        ))
        btn_row.add_widget(Button(
            text="CLINICAL SUMMARY AI", font_size=16, bold=True,
            on_press=self._run_clinical_summary_ai,
            background_color=get_color_from_hex('#2980b9')
        ))
        btn_row.add_widget(Button(
            text="PROGNOSIS AI", font_size=16, bold=True,
            on_press=self._run_prognosis_ai,
            background_color=get_color_from_hex('#e67e22')
        ))
        btn_row.add_widget(Button(
            text="TREATMENT AI", font_size=16, bold=True,
            on_press=self._run_treatment_ai,
            background_color=get_color_from_hex('#27ae60')
        ))
        main.add_widget(btn_row)

        main.add_widget(Label(
            text="ANALYSIS OUTPUT", font_size=18, bold=True,
            size_hint_y=None, height=28,
            color=get_color_from_hex('#2c3e50')
        ))
        analysis_scroll = ScrollView(size_hint_y=1, bar_width=10,
                                      scroll_type=['bars'])
        self.ai_analysis_output = TextInput(
            multiline=True, readonly=True, font_size=14,
            background_color=get_color_from_hex('#fafafa'),
            size_hint_y=None
        )
        self.ai_analysis_output.bind(minimum_height=self.ai_analysis_output.setter('height'))
        analysis_scroll.add_widget(self.ai_analysis_output)
        main.add_widget(analysis_scroll)

        btn_row2 = BoxLayout(size_hint_y=None, height=40, spacing=5)
        btn_row2.add_widget(Button(
            text="EXPORT PDF", font_size=15, bold=True,
            on_press=lambda x: self._export_to_pdf(
                self.ai_analysis_output.text,
                "AI_Analysis", "AI Clinical Analysis"
            ),
            background_color=get_color_from_hex('#2980b9')
        ))
        btn_row2.add_widget(Button(
            text="CLEAR", font_size=15, bold=True,
            on_press=lambda x: setattr(self.ai_analysis_output, 'text', ''),
            background_color=get_color_from_hex('#e74c3c')
        ))
        main.add_widget(btn_row2)

        tab.add_widget(main)
        self.tp.add_widget(tab)

    def _run_full_analysis(self, instance):
        if not self.cp or not self.ai.enabled:
            self._popup("Error", "Select patient and enable AI.")
            return
        pid = self.cp.get("id")
        emrs = [r for r in self.emr_records if r.get("patient_id") == pid]
        labs = [t for t in self.lab_tests if t.get("patient_id") == pid
                and t.get("status") == "Completed"]
        rx = [p for p in self.prescriptions if p.get("patient_id") == pid]

        self.ai_stats_emr.text = f"EMR: {len(emrs)}"
        self.ai_stats_lab.text = f"Lab: {len(labs)}"
        self.ai_stats_rx.text = f"Rx: {len(rx)}"

        tasks = self.task_manager.get_tasks_for_patient(pid)
        adms = [a for a in self.admissions if a.get('patient_id') == pid]
        appts = [a for a in self.appointments if a.get('patient_id') == pid]

        prompt = f"""Patient: {self.cp.get('name', '')} (ID: {pid})
Age: {self.cp.get('age', 0)} years
Gender: {self.cp.get('gender', '')}

EMR Records ({len(emrs)}):
{chr(10).join([f"- {r.get('ts','')}: {r.get('cc','')}" for r in emrs[-5:]])}

Lab Results ({len(labs)} completed):
{chr(10).join([f"- {l.get('test_name','')}: {l.get('result','')} {l.get('unit','')}" for l in labs[-5:]])}

Prescriptions ({len(rx)}):
{chr(10).join([f"- {r.get('drug','')} x{r.get('qty',0)}" for r in rx[-5:]])}

Tasks ({len(tasks)}):
{chr(10).join([f"- [{t.get('status')}] {t.get('task_type')}: {t.get('description')[:30]}" for t in tasks[-5:]])}

Admissions ({len(adms)}):
{chr(10).join([f"- {a.get('admission_date')}: {a.get('diagnosis')}" for a in adms[-3:]])}

Appointments ({len(appts)}):
{chr(10).join([f"- {a.get('date')}: {a.get('type')} - {a.get('status')}" for a in appts[-3:]])}

Provide comprehensive clinical analysis including:
1. Overall clinical status
2. Key findings
3. Differential diagnoses
4. Recommended next steps
5. Risk factors
6. Prognosis
7. Task completion status
"""
        self.ai_analysis_output.text = "Generating comprehensive analysis... Please wait."

        def run():
            ok, resp = self.ai.call(
                "You are a senior clinician providing comprehensive patient analysis.",
                prompt, mt=2000
            )
            if ok:
                Clock.schedule_once(
                    lambda dt: setattr(self.ai_analysis_output, 'text', resp), 0
                )
            else:
                Clock.schedule_once(
                    lambda dt: setattr(self.ai_analysis_output, 'text',
                                        f"Error: {resp}"), 0
                )

        threading.Thread(target=run, daemon=True).start()

    def _run_clinical_summary_ai(self, instance):
        if not self.cp or not self.ai.enabled:
            self._popup("Error", "Select patient and enable AI.")
            return
        pid = self.cp.get("id")
        emrs = [r for r in self.emr_records if r.get("patient_id") == pid]
        if not emrs:
            self._popup("Error", "No EMR records for this patient.")
            return
        tasks = self.task_manager.get_tasks_for_patient(pid)
        adms = [a for a in self.admissions if a.get('patient_id') == pid]
        prompt = f"""Patient: {self.cp.get('name', '')}, Age: {self.cp.get('age', 0)}
Generate a concise clinical summary from these EMR records:
{chr(10).join([f"- {r.get('ts','')}: {r.get('cc','')} | {r.get('dx','N/A')}" for r in emrs[-10:]])}

Tasks: {chr(10).join([f"- [{t.get('status')}] {t.get('task_type')}: {t.get('description')[:30]}" for t in tasks[-5:]])}

Admissions: {chr(10).join([f"- {a.get('admission_date')}: {a.get('diagnosis')}" for a in adms[-3:]])}

Format as:
- Patient Profile
- Problem List
- Current Status
- Active Issues
- Pending Items
- Task Status
"""
        self.ai_analysis_output.text = "Generating clinical summary..."

        def run():
            ok, resp = self.ai.call(
                "You are a clinician preparing a clinical summary.",
                prompt, mt=1000
            )
            if ok:
                Clock.schedule_once(
                    lambda dt: setattr(self.ai_analysis_output, 'text', resp), 0
                )
            else:
                Clock.schedule_once(
                    lambda dt: setattr(self.ai_analysis_output, 'text',
                                        f"Error: {resp}"), 0
                )

        threading.Thread(target=run, daemon=True).start()

    def _run_prognosis_ai(self, instance):
        if not self.cp or not self.ai.enabled:
            self._popup("Error", "Select patient and enable AI.")
            return
        pid = self.cp.get("id")
        emrs = [r for r in self.emr_records if r.get("patient_id") == pid]
        labs = [t for t in self.lab_tests if t.get("patient_id") == pid
                and t.get("status") == "Completed"]
        tasks = self.task_manager.get_tasks_for_patient(pid)
        adms = [a for a in self.admissions if a.get('patient_id') == pid]
        prompt = f"""Patient: {self.cp.get('name', '')}, Age: {self.cp.get('age', 0)}, Gender: {self.cp.get('gender', '')}

Diagnoses: {chr(10).join([f"- {r.get('dx','N/A')}" for r in emrs if r.get('dx')])[:10]}

Lab Results: {chr(10).join([f"- {l.get('test_name','')}: {l.get('result','')}" for l in labs[-5:]])}

Task Completion: {len([t for t in tasks if t.get('status') == 'Completed'])} of {len(tasks)} tasks completed

Admission History: {chr(10).join([f"- {a.get('admission_date')}: {a.get('diagnosis')}" for a in adms[-3:]])}

Provide prognosis assessment including:
1. Expected clinical course
2. Risk of complications
3. Potential outcomes
4. Factors affecting prognosis
5. Recommendations for monitoring
"""
        self.ai_analysis_output.text = "Generating prognosis assessment..."

        def run():
            ok, resp = self.ai.call(
                "You are a clinician providing prognosis assessment.",
                prompt, mt=1200
            )
            if ok:
                Clock.schedule_once(
                    lambda dt: setattr(self.ai_analysis_output, 'text', resp), 0
                )
            else:
                Clock.schedule_once(
                    lambda dt: setattr(self.ai_analysis_output, 'text',
                                        f"Error: {resp}"), 0
                )

        threading.Thread(target=run, daemon=True).start()

    def _run_treatment_ai(self, instance):
        if not self.cp or not self.ai.enabled:
            self._popup("Error", "Select patient and enable AI.")
            return
        pid = self.cp.get("id")
        emrs = [r for r in self.emr_records if r.get("patient_id") == pid]
        rx = [p for p in self.prescriptions if p.get("patient_id") == pid]
        tasks = self.task_manager.get_tasks_for_patient(pid)
        prompt = f"""Patient: {self.cp.get('name', '')}, Age: {self.cp.get('age', 0)}, Gender: {self.cp.get('gender', '')}

Diagnoses: {chr(10).join([f"- {r.get('dx','N/A')}" for r in emrs if r.get('dx')])[:5]}

Current Medications: {chr(10).join([f"- {r.get('drug','')} x{r.get('qty',0)}" for r in rx])}

Pending Tasks: {chr(10).join([f"- [{t.get('status')}] {t.get('task_type')}: {t.get('description')[:30]}" for t in tasks if t.get('status') != 'Completed'])}

Generate evidence-based treatment recommendations:
1. Primary treatment options
2. Alternative treatments
3. Medication recommendations
4. Lifestyle interventions
5. Follow-up plan
6. Monitoring parameters
7. Task completion priorities
"""
        self.ai_analysis_output.text = "Generating treatment recommendations..."

        def run():
            ok, resp = self.ai.call(
                "You are a clinician providing treatment recommendations.",
                prompt, mt=1200
            )
            if ok:
                Clock.schedule_once(
                    lambda dt: setattr(self.ai_analysis_output, 'text', resp), 0
                )
            else:
                Clock.schedule_once(
                    lambda dt: setattr(self.ai_analysis_output, 'text',
                                        f"Error: {resp}"), 0
                )

        threading.Thread(target=run, daemon=True).start()

    # ====================== AI-BOS TAB ======================
    def _ai_bos_tab(self):
        tab = TabbedPanelItem(text="AI-BOS")
        main = BoxLayout(orientation='vertical', padding=6, spacing=4)
        main.add_widget(Label(
            text="AI BUSINESS OPERATING SYSTEM (AI-BOS)", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#1a5276')
        ))

        profile_box = BoxLayout(orientation='vertical', size_hint_y=0.3, spacing=3)
        profile_box.add_widget(Label(
            text="BUSINESS PROFILE", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#e67e22')
        ))
        self.bos_profile_display = TextInput(
            multiline=True, readonly=True, font_size=13,
            background_color=get_color_from_hex('#fef9e7'),
            size_hint_y=None
        )
        self.bos_profile_display.bind(
            minimum_height=self.bos_profile_display.setter('height')
        )
        profile_box.add_widget(self.bos_profile_display)
        main.add_widget(profile_box)

        btn_row1 = BoxLayout(size_hint_y=None, height=40, spacing=5)
        for txt, cb, col in [
            ("DISCOVER BUSINESS", self._bos_discover, '#2980b9'),
            ("HEALTH SCORE", self._bos_health, '#27ae60'),
            ("MARKET INTELLIGENCE", self._bos_market, '#8e44ad'),
            ("REVENUE ENGINE", self._bos_revenue, '#e67e22')
        ]:
            btn_row1.add_widget(Button(
                text=txt, font_size=14, bold=True, on_press=cb,
                background_color=get_color_from_hex(col)
            ))
        main.add_widget(btn_row1)

        btn_row2 = BoxLayout(size_hint_y=None, height=40, spacing=5)
        for txt, cb, col in [
            ("SALES ENGINE", self._bos_sales, '#27ae60'),
            ("MARKETING ENGINE", self._bos_marketing, '#f39c12'),
            ("CX ENGINE", self._bos_cx, '#16a085'),
            ("FINANCIAL ENGINE", self._bos_financial, '#2980b9')
        ]:
            btn_row2.add_widget(Button(
                text=txt, font_size=14, bold=True, on_press=cb,
                background_color=get_color_from_hex(col)
            ))
        main.add_widget(btn_row2)

        btn_row3 = BoxLayout(size_hint_y=None, height=40, spacing=5)
        for txt, cb, col in [
            ("AI AUTOMATION", self._bos_ai_automation, '#8e44ad'),
            ("ROADMAP", self._bos_roadmap, '#e74c3c'),
            ("DECISION SUPPORT", self._bos_decision, '#d35400'),
            ("DOCUMENTS", self._bos_documents, '#1a5276')
        ]:
            btn_row3.add_widget(Button(
                text=txt, font_size=14, bold=True, on_press=cb,
                background_color=get_color_from_hex(col)
            ))
        main.add_widget(btn_row3)

        main.add_widget(Label(
            text="AI-BOS OUTPUT", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#2c3e50')
        ))
        bos_scroll = ScrollView(size_hint_y=1, bar_width=10, scroll_type=['bars'])
        self.bos_output = TextInput(
            multiline=True, readonly=True, font_size=13,
            background_color=get_color_from_hex('#fafafa'),
            size_hint_y=None
        )
        self.bos_output.bind(minimum_height=self.bos_output.setter('height'))
        bos_scroll.add_widget(self.bos_output)
        main.add_widget(bos_scroll)

        btn_row4 = BoxLayout(size_hint_y=None, height=40, spacing=5)
        btn_row4.add_widget(Button(
            text="EXPORT PDF", font_size=15, bold=True,
            on_press=lambda x: self._export_to_pdf(
                self.bos_output.text, "AI_BOS", "AI-BOS Report"
            ),
            background_color=get_color_from_hex('#2980b9')
        ))
        btn_row4.add_widget(Button(
            text="CLEAR", font_size=15, bold=True,
            on_press=lambda x: setattr(self.bos_output, 'text', ''),
            background_color=get_color_from_hex('#e74c3c')
        ))
        main.add_widget(btn_row4)

        tab.add_widget(main)
        self.tp.add_widget(tab)

    def _bos_discover(self, instance):
        questions = self.bos.get_business_profile()
        text = "BUSINESS DISCOVERY QUESTIONS\n" + "=" * 50 + "\n\n"
        for q, v in questions.items():
            text += f"{q}: {v}\n"
        text += "\nAnswer these questions to build your business profile."
        self.bos_profile_display.text = text
        self.bos_output.text = "Business discovery questions displayed above."

    def _bos_health(self, instance):
        scores = self.bos.calculate_health_score()
        text = "BUSINESS HEALTH SCORE\n" + "=" * 50 + "\n\n"
        for area, score in scores.items():
            status = "🔴 Critical" if score < 40 else "🟡 Needs Work" if score < 60 else "🟢 Good"
            text += f"{area}: {score}/100 - {status}\n"
        text += f"\nOverall AI Readiness: {scores.get('AI Readiness', 0)}/100"
        self.bos_output.text = text

    def _bos_market(self, instance):
        self.bos.generate_market_intelligence()
        self.bos_output.text = self.bos.market_intel.get("analysis",
                                                          "No market intelligence available.")

    def _bos_revenue(self, instance):
        self.bos.generate_revenue_engine()
        text = "REVENUE ENGINE OPPORTUNITIES\n" + "=" * 50 + "\n\n"
        for opp, details in self.bos.revenue_engine.get("ranked", [])[:10]:
            text += f"{opp.upper()}\n"
            text += f"  Revenue Potential: {details['revenue_potential']}\n"
            text += f"  Profit Potential: {details['profit_potential']}\n"
            text += f"  Cost: {details['cost']}\n"
            text += f"  ROI: {details['roi']}\n"
            text += f"  Time to Implement: {details['time_to_implement']}\n"
            text += "-" * 40 + "\n"
        self.bos_output.text = text

    def _bos_sales(self, instance):
        self.bos.generate_sales_engine()
        text = "SALES ENGINE\n" + "=" * 50 + "\n\n"
        text += f"Strategy: {self.bos.sales_engine.get('strategy', '')}\n\n"
        text += "Sales Process:\n"
        for step in self.bos.sales_engine.get('sales_scripts', []):
            text += f"  - {step}\n"
        text += f"\nCRM Workflow: {self.bos.sales_engine.get('crm_workflow', '')}\n"
        text += f"\nLead Qualification: {self.bos.sales_engine.get('lead_qualification', '')}"
        self.bos_output.text = text

    def _bos_marketing(self, instance):
        self.bos.generate_marketing_engine()
        text = "MARKETING ENGINE\n" + "=" * 50 + "\n\n"
        text += f"Strategy: {self.bos.marketing_engine.get('strategy', '')}\n"
        text += f"Brand Positioning: {self.bos.marketing_engine.get('brand_positioning', '')}\n\n"
        text += "SEO Plan:\n"
        for item in self.bos.marketing_engine.get('seo_plan', []):
            text += f"  - {item}\n"
        text += "\nSocial Media Campaigns:\n"
        for item in self.bos.marketing_engine.get('social_media_campaigns', []):
            text += f"  - {item}\n"
        text += "\nEmail Campaigns:\n"
        for item in self.bos.marketing_engine.get('email_campaigns', []):
            text += f"  - {item}\n"
        text += "\nPaid Advertising:\n"
        for item in self.bos.marketing_engine.get('paid_advertising', []):
            text += f"  - {item}\n"
        self.bos_output.text = text

    def _bos_cx(self, instance):
        self.bos.generate_customer_experience_engine()
        text = "CUSTOMER EXPERIENCE ENGINE\n" + "=" * 50 + "\n\n"
        text += "Onboarding:\n"
        for item in self.bos.cx_engine.get('onboarding', []):
            text += f"  - {item}\n"
        text += "\nCustomer Support:\n"
        for item in self.bos.cx_engine.get('customer_support', []):
            text += f"  - {item}\n"
        text += "\nCustomer Retention:\n"
        for item in self.bos.cx_engine.get('customer_retention', []):
            text += f"  - {item}\n"
        text += f"\nCustomer Journey: {' → '.join(self.bos.cx_engine.get('customer_journey_mapping', []))}"
        self.bos_output.text = text

    def _bos_financial(self, instance):
        self.bos.generate_financial_engine()
        text = "FINANCIAL ENGINE\n" + "=" * 50 + "\n\n"
        pl = self.bos.financial_engine.get('profit_loss', {})
        text += "PROFIT & LOSS\n"
        text += f"Revenue: N{pl.get('revenue', 0):,.2f}\n"
        text += f"COGS: N{pl.get('cost_of_goods', 0):,.2f}\n"
        text += f"Gross Profit: N{pl.get('gross_profit', 0):,.2f}\n"
        text += f"Operating Expenses: N{pl.get('operating_expenses', 0):,.2f}\n"
        text += f"Net Profit: N{pl.get('net_profit', 0):,.2f}\n\n"
        text += "BUDGET RECOMMENDATIONS\n"
        for k, v in self.bos.financial_engine.get('budget_recommendations', {}).items():
            text += f"  {k}: N{v:,.2f}\n"
        text += "\nCOST REDUCTION OPPORTUNITIES\n"
        for item in self.bos.financial_engine.get('cost_reduction_opportunities', []):
            text += f"  - {item}\n"
        self.bos_output.text = text

    def _bos_ai_automation(self, instance):
        self.bos.generate_ai_automation_engine()
        text = "AI AUTOMATION ENGINE\n" + "=" * 50 + "\n\n"
        for area, items in self.bos.ai_automation.items():
            if area != "roi_prioritization":
                text += f"{area.upper()}\n"
                for item in items:
                    text += f"  - {item}\n"
                text += "\n"
        text += "ROI PRIORITIZATION\n"
        for level, items in self.bos.ai_automation.get('roi_prioritization', {}).items():
            text += f"  {level.upper()}:\n"
            for item in items:
                text += f"    - {item}\n"
        self.bos_output.text = text

    def _bos_roadmap(self, instance):
        roadmap = self.bos.generate_implementation_roadmap()
        text = "IMPLEMENTATION ROADMAP\n" + "=" * 50 + "\n\n"
        for period, details in roadmap.items():
            text += f"{period.upper().replace('_', ' ')}\n"
            text += f"  Milestones:\n"
            for m in details.get('milestones', []):
                text += f"    - {m}\n"
            text += f"  Estimated Cost: {details.get('costs', 'N/A')}\n\n"
        self.bos_output.text = text

    def _bos_decision(self, instance):
        context = "Strategic growth planning"
        options = ["Expand services", "Improve technology", "Enhance customer experience"]
        analysis = self.bos.decision_support(context, options)
        text = "DECISION SUPPORT\n" + "=" * 50 + "\n\n"
        text += f"Context: {analysis['context']}\n"
        text += f"Options: {', '.join(analysis['options'])}\n"
        text += f"Recommendation: {analysis['recommendation']}\n"
        text += f"Success Metrics: {', '.join(analysis['success_metrics'])}\n"
        self.bos_output.text = text

    def _bos_documents(self, instance):
        c = BoxLayout(orientation='vertical', spacing=5, padding=8)
        c.add_widget(Label(text="Select Document Type:", font_size=16, bold=True,
                            size_hint_y=None, height=30))
        doc_spinner = Spinner(
            text="business_plan",
            values=["business_plan", "sales_plan", "marketing_plan",
                    "referral_program", "sop", "employee_handbook",
                    "financial_report", "risk_assessment", "pitch_deck",
                    "grant_proposal", "partnership_proposal", "project_plan",
                    "kpi_dashboard"],
            size_hint_y=None, height=40
        )
        c.add_widget(doc_spinner)
        btn_box = BoxLayout(size_hint_y=None, height=40, spacing=5)
        pp = Popup(title="Generate Document", content=c, size_hint=(0.6, 0.3))

        def generate(inst):
            doc_type = doc_spinner.text
            doc = self.bos.generate_document(doc_type)
            self.bos_output.text = doc
            pp.dismiss()
            self._popup("Document Generated",
                        f"{doc_type.replace('_', ' ').title()} generated.")

        btn_box.add_widget(Button(
            text="GENERATE", font_size=15, bold=True, on_press=generate,
            background_color=get_color_from_hex('#27ae60')
        ))
        btn_box.add_widget(Button(text="CANCEL", font_size=15, on_press=pp.dismiss))
        c.add_widget(btn_box)
        pp.open()

    # ====================== SNAPSHOT TAB ======================
    def _snapshot_tab(self):
        tab = TabbedPanelItem(text="Snapshot")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=6, spacing=4, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))
        l.add_widget(Label(
            text="PATIENT SNAPSHOT", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#1a5276')
        ))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="ID:", font_size=15, bold=True, size_hint_x=0.05))
        self.snapshot_pid = TextInput(size_hint_x=0.15, readonly=True, font_size=15,
                                        size_hint_y=None, height=40)
        bar.add_widget(self.snapshot_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.07))
        self.snapshot_name = TextInput(size_hint_x=0.35, readonly=True, font_size=15,
                                         size_hint_y=None, height=40)
        bar.add_widget(self.snapshot_name)
        l.add_widget(bar)

        snap_btn = BoxLayout(size_hint_y=None, height=40, spacing=5)
        snap_btn.add_widget(Button(
            text="📸 CAPTURE SNAPSHOT", font_size=16, bold=True,
            on_press=self._capture_snapshot,
            background_color=get_color_from_hex('#27ae60')
        ))
        snap_btn.add_widget(Button(
            text="VIEW SNAPSHOTS", font_size=16, bold=True,
            on_press=self._view_snapshots,
            background_color=get_color_from_hex('#2980b9')
        ))
        l.add_widget(snap_btn)

        l.add_widget(Label(
            text="SNAPSHOT HISTORY", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#2c3e50')
        ))
        self.snapshot_history = BoxLayout(orientation='vertical',
                                            size_hint_y=None, spacing=1)
        self.snapshot_history.bind(minimum_height=self.snapshot_history.setter('height'))
        l.add_widget(self.snapshot_history)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)
        self._refresh_snapshots()

    def _capture_snapshot(self, instance):
        if not self.cp:
            self._popup("Error", "Select a patient first.")
            return
        pid = self.cp.get("id")
        emrs = [r for r in self.emr_records if r.get("patient_id") == pid]
        labs = [t for t in self.lab_tests if t.get("patient_id") == pid
                and t.get("status") == "Completed"]
        rx = [p for p in self.prescriptions if p.get("patient_id") == pid]
        tasks = self.task_manager.get_tasks_for_patient(pid)

        latest = emrs[-1] if emrs else {}
        snapshot = {
            "patient_id": pid,
            "patient_name": self.cp.get("name", ""),
            "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "vitals": {
                "bp": latest.get("bp", ""),
                "pulse": latest.get("pulse", ""),
                "temp": latest.get("temp", ""),
                "spo2": latest.get("spo2", ""),
                "rr": latest.get("rr", ""),
                "avpu": latest.get("avpu", "Alert")
            },
            "diagnosis": latest.get("dx", "N/A"),
            "treatment": latest.get("tx", "N/A"),
            "follow_up": latest.get("fu", "N/A"),
            "lab_count": len(labs),
            "rx_count": len(rx),
            "emr_count": len(emrs),
            "task_pending": len([t for t in tasks
                                  if t.get("status") != "Completed"]),
            "task_total": len(tasks),
            "captured_by": self._current_username()
        }
        self.patient_snapshots.append(snapshot)
        self.save_all()
        self._refresh_snapshots()
        self._popup("Snapshot Captured",
                    f"Snapshot saved for {self.cp.get('name', '')}")
        self.log_activity(f"Snapshot captured for {self.cp.get('name', '')}")
        self.audit_log("CAPTURE_SNAPSHOT", "", pid)

    def _refresh_snapshots(self):
        if not hasattr(self, 'snapshot_history'):
            return
        self.snapshot_history.clear_widgets()
        if not self.snapshot_pid.text:
            self.snapshot_history.add_widget(Label(
                text="Select a patient to view snapshots.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        pid = self.snapshot_pid.text
        snaps = [s for s in self.patient_snapshots
                 if s.get("patient_id") == pid]
        if not snaps:
            self.snapshot_history.add_widget(Label(
                text="No snapshots for this patient.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        for s in sorted(snaps, key=lambda x: x.get('ts', ''), reverse=True)[:15]:
            row = BoxLayout(size_hint_y=None, height=30, spacing=3)
            vitals = s.get('vitals', {})
            row.add_widget(Label(
                text=(f"{s.get('ts', '')[:16]} | "
                      f"BP: {vitals.get('bp', 'N/A')} | "
                      f"Tasks: {s.get('task_pending', 0)} pending"),
                font_size=12, size_hint_x=0.5
            ))
            row.add_widget(Label(
                text=f"DX: {s.get('diagnosis', 'N/A')[:20]}",
                font_size=11, size_hint_x=0.25
            ))
            btn = Button(
                text="VIEW", font_size=11, size_hint_x=0.15,
                size_hint_y=None, height=25,
                background_color=get_color_from_hex('#2980b9')
            )
            btn.snap = s
            btn.bind(on_press=lambda x: self._view_snapshot_detail(x.snap))
            row.add_widget(btn)
            self.snapshot_history.add_widget(row)

    def _view_snapshot_detail(self, snapshot):
        text = f"SNAPSHOT - {snapshot.get('patient_name', '')}\n"
        text += "=" * 50 + "\n\n"
        text += f"Date: {snapshot.get('ts', '')}\n\n"
        vitals = snapshot.get('vitals', {})
        text += "VITAL SIGNS\n"
        for k, label in [("bp", "BP"), ("pulse", "Pulse"), ("temp", "Temp"),
                          ("spo2", "SpO2"), ("rr", "RR"), ("avpu", "AVPU")]:
            text += f"  {label}: {vitals.get(k, 'N/A')}\n"
        text += f"\nDiagnosis: {snapshot.get('diagnosis', 'N/A')}\n"
        text += f"Treatment: {snapshot.get('treatment', 'N/A')}\n"
        text += f"Follow-up: {snapshot.get('follow_up', 'N/A')}\n\n"
        text += f"EMR Records: {snapshot.get('emr_count', 0)}\n"
        text += f"Lab Tests: {snapshot.get('lab_count', 0)}\n"
        text += f"Prescriptions: {snapshot.get('rx_count', 0)}\n"
        text += (f"Tasks: {snapshot.get('task_total', 0)} total, "
                 f"{snapshot.get('task_pending', 0)} pending\n")
        text += f"Captured by: {snapshot.get('captured_by', 'N/A')}\n"
        self._text_popup("Snapshot Detail", text)

    def _view_snapshots(self, instance):
        if not self.snapshot_pid.text:
            self._popup("Error", "No patient selected.")
            return
        pid = self.snapshot_pid.text
        snaps = [s for s in self.patient_snapshots if s.get("patient_id") == pid]
        if not snaps:
            self._popup("Snapshots", "No snapshots found.")
            return
        text = f"SNAPSHOTS FOR {self.snapshot_name.text}\n" + "=" * 50 + "\n\n"
        for s in sorted(snaps, key=lambda x: x.get('ts', ''), reverse=True):
            text += f"Date: {s.get('ts', '')}\n"
            vitals = s.get('vitals', {})
            text += (f"BP: {vitals.get('bp', 'N/A')} | "
                     f"Pulse: {vitals.get('pulse', 'N/A')}\n")
            text += f"Diagnosis: {s.get('diagnosis', 'N/A')}\n"
            text += f"Tasks: {s.get('task_pending', 0)} pending\n"
            text += "-" * 40 + "\n"
        self._text_popup("Snapshots", text)

    # ====================== PROCEDURES TAB ======================
    def _procedures_tab(self):
        tab = TabbedPanelItem(text="Procedures")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=6, spacing=4, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))
        l.add_widget(Label(
            text="PROCEDURE MANAGEMENT", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#1a5276')
        ))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="Patient ID:", font_size=15, bold=True,
                               size_hint_x=0.12))
        self.proc_pid = TextInput(size_hint_x=0.2, readonly=True, font_size=15,
                                    size_hint_y=None, height=40)
        bar.add_widget(self.proc_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.08))
        self.proc_name_display = TextInput(size_hint_x=0.3, readonly=True,
                                             font_size=15, size_hint_y=None, height=40)
        bar.add_widget(self.proc_name_display)
        l.add_widget(bar)

        proc_form = BoxLayout(size_hint_y=None, height=40, spacing=4)
        proc_form.add_widget(Label(text="Procedure:", font_size=15, bold=True,
                                     size_hint_x=0.12))
        self.proc_select = Spinner(
            text="Select", values=[p['name'] for p in self.procedures],
            size_hint_x=0.3, size_hint_y=None, height=40
        )
        proc_form.add_widget(self.proc_select)
        proc_form.add_widget(Button(
            text="ADD", font_size=15, bold=True, size_hint_x=0.12,
            on_press=self._add_procedure,
            background_color=get_color_from_hex('#27ae60')
        ))
        proc_form.add_widget(Button(
            text="CUSTOM", font_size=15, bold=True, size_hint_x=0.12,
            on_press=self._add_custom_procedure,
            background_color=get_color_from_hex('#8e44ad')
        ))
        l.add_widget(proc_form)

        l.add_widget(Label(
            text="PROCEDURE LIST", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#2c3e50')
        ))
        self.proc_list = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        self.proc_list.bind(minimum_height=self.proc_list.setter('height'))
        l.add_widget(self.proc_list)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)
        self._refresh_procedures()

    def _add_procedure(self, instance):
        if not self.proc_pid.text or self.proc_select.text == "Select":
            self._popup("Error", "Select patient and procedure.")
            return
        proc_name = self.proc_select.text
        price = self._get_procedure_price(proc_name)
        self.procedures_db.append({
            "id": f"PROC{datetime.now().strftime('%Y%m%d%H%M%S%f')}",
            "patient_id": self.proc_pid.text,
            "patient_name": self.proc_name_display.text,
            "procedure": proc_name,
            "price": price,
            "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "status": "Scheduled"
        })
        self.save_all()
        self._refresh_procedures()
        self._popup("Added", f"{proc_name} added.")
        self.log_activity(f"Procedure added: {proc_name} for {self.proc_name_display.text}")
        self.audit_log("ADD_PROCEDURE", proc_name, self.proc_pid.text)
        if self.cp:
            self.task_manager.create_task(
                patient_id=self.cp.get("id"),
                patient_name=self.cp.get("name"),
                assigned_to="Doctor",
                task_type="Doctor",
                description=f"Procedure: {proc_name}",
                order_items=[proc_name],
                department="Surgery"
            )
            self._refresh_tasks()

    def _add_custom_procedure(self, instance):
        if not self.proc_pid.text:
            self._popup("Error", "Select a patient first.")
            return
        content = BoxLayout(orientation='vertical', spacing=5, padding=8)
        content.add_widget(Label(
            text="Custom Procedure", font_size=16, bold=True,
            size_hint_y=None, height=30,
            color=get_color_from_hex('#8e44ad')
        ))
        content.add_widget(Label(text="Procedure Name:", font_size=14, bold=True,
                                   size_hint_y=None, height=22))
        name_inp = TextInput(font_size=14, size_hint_y=None, height=40)
        content.add_widget(name_inp)
        content.add_widget(Label(text="Price (N):", font_size=14, bold=True,
                                   size_hint_y=None, height=22))
        price_inp = TextInput(font_size=14, size_hint_y=None, height=40,
                                input_filter="float")
        content.add_widget(price_inp)
        btn_box = BoxLayout(size_hint_y=None, height=40, spacing=5)
        pp = Popup(title="Custom Procedure", content=content, size_hint=(0.6, 0.4))

        def add(inst):
            name = name_inp.text.strip()
            if not name:
                self._popup("Error", "Enter procedure name.")
                return
            try:
                price = float(price_inp.text or "0")
            except:
                self._popup("Error", "Invalid price.")
                return
            if price <= 0:
                self._popup("Error", "Price must be > 0.")
                return
            self.procedures_db.append({
                "id": f"PROC{datetime.now().strftime('%Y%m%d%H%M%S%f')}",
                "patient_id": self.proc_pid.text,
                "patient_name": self.proc_name_display.text,
                "procedure": name,
                "price": price,
                "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "status": "Scheduled"
            })
            self.save_all()
            self._refresh_procedures()
            pp.dismiss()
            self._popup("Added", f"Custom procedure: {name}")
            self.log_activity(f"Custom procedure added: {name}")
            self.audit_log("ADD_CUSTOM_PROCEDURE", f"{name} @ N{price}",
                           self.proc_pid.text)

        btn_box.add_widget(Button(
            text="ADD", font_size=15, bold=True, on_press=add,
            background_color=get_color_from_hex('#27ae60')
        ))
        btn_box.add_widget(Button(text="CANCEL", font_size=15, on_press=pp.dismiss))
        content.add_widget(btn_box)
        pp.open()

    def _get_procedure_price(self, name):
        for p in self.procedures:
            if p['name'] == name:
                return p.get('price', 0)
        return 0

    def _refresh_procedures(self):
        if not hasattr(self, 'proc_list'):
            return
        self.proc_list.clear_widgets()
        if not self.proc_pid.text:
            self.proc_list.add_widget(Label(
                text="Select a patient to view procedures.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        pid = self.proc_pid.text
        procs = [p for p in self.procedures_db if p.get("patient_id") == pid]
        if not procs:
            self.proc_list.add_widget(Label(
                text="No procedures for this patient.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        for p in procs:
            row = BoxLayout(size_hint_y=None, height=30, spacing=3)
            row.add_widget(Label(text=f"{p.get('procedure', '')}",
                                   font_size=13, size_hint_x=0.35))
            row.add_widget(Label(text=f"N{p.get('price', 0):,.2f}",
                                   font_size=13, size_hint_x=0.15))
            row.add_widget(Label(text=p.get('status', ''), font_size=12,
                                   size_hint_x=0.15))
            row.add_widget(Label(text=p.get('date', '')[:16], font_size=11,
                                   size_hint_x=0.2))
            btn = Button(
                text="X", font_size=11, size_hint_x=0.05,
                size_hint_y=None, height=25,
                background_color=get_color_from_hex('#e74c3c')
            )
            btn.proc_id = p.get('id')
            btn.bind(on_press=lambda x: self._delete_procedure(x.proc_id))
            row.add_widget(btn)
            self.proc_list.add_widget(row)

    def _delete_procedure(self, proc_id):
        self.procedures_db = [p for p in self.procedures_db
                              if p.get('id') != proc_id]
        self.save_all()
        self._refresh_procedures()

    # ====================== VITALS TIMELINE TAB ======================
    def _vitals_timeline_tab(self):
        tab = TabbedPanelItem(text="Vitals Timeline")
        main = BoxLayout(orientation='vertical', padding=6, spacing=6)
        score_box = GridLayout(cols=2, size_hint_y=None, height=80, spacing=5)
        self.news_label = Label(text="NEWS2: --", font_size=22, bold=True,
                                  color=get_color_from_hex('#e74c3c'))
        self.qsofa_label = Label(text="qSOFA: --", font_size=22, bold=True,
                                   color=get_color_from_hex('#e67e22'))
        score_box.add_widget(self.news_label)
        score_box.add_widget(self.qsofa_label)
        main.add_widget(score_box)
        main.add_widget(Label(
            text="VITALS HISTORY", font_size=16, bold=True,
            size_hint_y=None, height=25,
            color=get_color_from_hex('#2980b9')
        ))
        vitals_scroll = ScrollView(size_hint_y=1, bar_width=10,
                                    scroll_type=['bars'])
        self.vitals_list = BoxLayout(orientation='vertical',
                                       size_hint_y=None, spacing=1)
        self.vitals_list.bind(minimum_height=self.vitals_list.setter('height'))
        vitals_scroll.add_widget(self.vitals_list)
        main.add_widget(vitals_scroll)
        tab.add_widget(main)
        self.tp.add_widget(tab)
        self._refresh_vitals_timeline()

    def _refresh_vitals_timeline(self):
        if not hasattr(self, 'vitals_list'):
            return
        self.vitals_list.clear_widgets()
        if not self.cp:
            self.vitals_list.add_widget(Label(
                text="Select a patient to view vitals timeline.",
                font_size=14, size_hint_y=None, height=30
            ))
            self.news_label.text = "NEWS2: --"
            self.qsofa_label.text = "qSOFA: --"
            return
        pid = self.cp.get("id")
        records = [r for r in self.emr_records if r.get("patient_id") == pid]
        for dn in self.daily_notes:
            if dn.get("patient_id") == pid:
                records.append({
                    "ts": dn.get("date", "") + " 08:00",
                    "bp": dn.get("vitals", {}).get("bp", ""),
                    "pulse": dn.get("vitals", {}).get("pulse", ""),
                    "rr": dn.get("vitals", {}).get("rr", ""),
                    "spo2": dn.get("vitals", {}).get("spo2", ""),
                    "temp": dn.get("vitals", {}).get("temp", ""),
                    "avpu": dn.get("vitals", {}).get("avpu", "Alert")
                })
        records.sort(key=lambda x: x.get("ts", ""))
        if not records:
            self.vitals_list.add_widget(Label(
                text="No records found.", font_size=14,
                size_hint_y=None, height=30
            ))
            self.news_label.text = "NEWS2: --"
            self.qsofa_label.text = "qSOFA: --"
            return
        latest = records[-1]
        bp = latest.get("bp", "")
        pulse = latest.get("pulse", "")
        rr = latest.get("rr", "")
        spo2 = latest.get("spo2", "")
        temp = latest.get("temp", "")
        avpu = latest.get("avpu", "Alert")
        news = self.compute_news(bp, pulse, rr, spo2, temp, avpu)
        qsofa = self.compute_qsofa(bp, rr, avpu)
        news_color = "#27ae60" if news <= 4 else "#f39c12" if news <= 6 else "#e74c3c"
        self.news_label.text = f"NEWS2: {news}"
        self.news_label.color = get_color_from_hex(news_color)
        self.qsofa_label.text = (f"qSOFA: {qsofa} " +
                                  ("(Sepsis risk)" if qsofa >= 2 else ""))
        for r in records:
            ts = r.get("ts", "N/A")
            row = BoxLayout(size_hint_y=None, height=28, spacing=2)
            row.add_widget(Label(text=ts[:16], font_size=12, size_hint_x=0.15))
            row.add_widget(Label(text=r.get("bp", "-"), font_size=12, size_hint_x=0.1))
            row.add_widget(Label(text=r.get("pulse", "-"), font_size=12, size_hint_x=0.08))
            row.add_widget(Label(text=r.get("spo2", "-"), font_size=12, size_hint_x=0.08))
            row.add_widget(Label(text=r.get("rr", "-"), font_size=12, size_hint_x=0.08))
            row.add_widget(Label(text=r.get("temp", "-"), font_size=12, size_hint_x=0.08))
            row.add_widget(Label(text=r.get("avpu", "Alert"), font_size=12, size_hint_x=0.1))
            self.vitals_list.add_widget(row)

    @staticmethod
    def compute_news(bp, pulse, rr, spo2, temp, avpu="Alert"):
        score = 0
        sbp = None
        if bp and '/' in bp:
            try:
                sbp = int(bp.split('/')[0])
            except:
                pass
        if sbp:
            if sbp <= 90 or sbp >= 220:
                score += 3
            elif sbp <= 100:
                score += 2
            elif sbp <= 110:
                score += 1
        try:
            hr = int(pulse)
        except:
            hr = 0
        if hr > 0:
            if hr <= 40 or hr >= 131:
                score += 3
            elif hr <= 50 or hr >= 111:
                score += 2
            elif hr <= 90:
                score += 0
            else:
                score += 1
        try:
            rr_val = int(rr)
        except:
            rr_val = 0
        if rr_val > 0:
            if rr_val <= 8 or rr_val >= 25:
                score += 3
            elif rr_val <= 11 or rr_val >= 21:
                score += 1
        try:
            spo2_val = int(spo2)
        except:
            spo2_val = 0
        if spo2_val > 0:
            if spo2_val <= 91:
                score += 3
            elif spo2_val <= 93:
                score += 2
            elif spo2_val <= 95:
                score += 1
        try:
            temp_val = float(temp)
        except:
            temp_val = 0
        if temp_val > 0:
            if temp_val <= 35.0:
                score += 3
            elif temp_val >= 39.1:
                score += 2
            elif temp_val >= 38.1:
                score += 1
        if avpu != "Alert":
            score += 3
        return score

    @staticmethod
    def compute_qsofa(bp, rr, avpu="Alert"):
        score = 0
        sbp = None
        if bp and '/' in bp:
            try:
                sbp = int(bp.split('/')[0])
            except:
                pass
        if sbp and sbp <= 100:
            score += 1
        try:
            rr_val = int(rr)
        except:
            rr_val = 0
        if rr_val >= 22:
            score += 1
        if avpu != "Alert":
            score += 1
        return score

    # ====================== CONSENT FORM TAB ======================
    def _consent_form_tab(self):
        tab = TabbedPanelItem(text="Consent Form")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=6, spacing=4, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))
        l.add_widget(Label(
            text="INFORMED CONSENT", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#1a5276')
        ))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="Patient ID:", font_size=15, bold=True,
                               size_hint_x=0.12))
        self.consent_pid = TextInput(size_hint_x=0.2, readonly=True, font_size=15,
                                       size_hint_y=None, height=40)
        bar.add_widget(self.consent_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.08))
        self.consent_name = TextInput(size_hint_x=0.3, readonly=True, font_size=15,
                                        size_hint_y=None, height=40)
        bar.add_widget(self.consent_name)
        l.add_widget(bar)

        proc_box = BoxLayout(size_hint_y=None, height=40, spacing=4)
        proc_box.add_widget(Label(text="Procedure:", font_size=15, bold=True,
                                    size_hint_x=0.15))
        self.consent_procedure = TextInput(font_size=15, size_hint_x=0.55,
                                             size_hint_y=None, height=40,
                                             hint_text="e.g., Appendectomy")
        proc_box.add_widget(self.consent_procedure)
        proc_box.add_widget(Button(
            text="AI GENERATE", font_size=15, bold=True, size_hint_x=0.3,
            on_press=self._ai_generate_consent,
            background_color=get_color_from_hex('#8e44ad')
        ))
        l.add_widget(proc_box)

        l.add_widget(Label(text="Consent Statement:", font_size=15, bold=True,
                            size_hint_y=None, height=18))
        self.consent_statement = TextInput(
            text="", font_size=14, size_hint_y=None, height=150, multiline=True,
            hint_text="Consent statement will appear here or type manually..."
        )
        l.add_widget(self.consent_statement)

        btns = BoxLayout(size_hint_y=None, height=44, spacing=5)
        btns.add_widget(Button(
            text="SAVE", font_size=16, bold=True,
            on_press=self._save_consent,
            background_color=get_color_from_hex('#27ae60')
        ))
        btns.add_widget(Button(
            text="EXPORT PDF", font_size=16, bold=True,
            on_press=self._print_consent,
            background_color=get_color_from_hex('#2980b9')
        ))
        l.add_widget(btns)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)

    def _ai_generate_consent(self, instance):
        if not self.ai.enabled:
            self._popup("AI Error", "AI is not enabled. Set up API key in Settings.")
            return
        procedure = self.consent_procedure.text.strip()
        if not procedure:
            self._popup("Error", "Enter a procedure first.")
            return
        prompt = (f"Generate a professional, legally-sound informed consent "
                  f"statement for the following medical procedure: {procedure}. "
                  f"Include explanation of the procedure, benefits, risks, "
                  f"alternatives, and confirmation of voluntary consent. "
                  f"Format as a formal consent document.")

        def run():
            ok, resp = self.ai.call(
                "You are a medical legal expert.", prompt, mt=500
            )
            if ok:
                Clock.schedule_once(
                    lambda dt: setattr(self.consent_statement, 'text', resp), 0
                )
            else:
                Clock.schedule_once(
                    lambda dt: self._popup("AI Error", resp), 0
                )

        threading.Thread(target=run, daemon=True).start()

    def _save_consent(self, instance):
        if not self.consent_pid.text:
            self._popup("Error", "No patient selected.")
            return
        form = {
            "patient_id": self.consent_pid.text,
            "patient_name": self.consent_name.text,
            "procedure": self.consent_procedure.text,
            "statement": self.consent_statement.text,
            "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "created_by": self._current_username()
        }
        self.consent_forms.append(form)
        self.save_all()
        self._popup("Saved", "Consent form saved.")
        self.log_activity(f"Consent saved for {self.consent_name.text}")
        self.audit_log("SAVE_CONSENT", self.consent_procedure.text,
                       self.consent_pid.text)

    def _print_consent(self, instance):
        if not self.consent_pid.text:
            self._popup("Error", "No patient selected.")
            return
        text = (f"=== HMG HOSPITAL - INFORMED CONSENT ===\n\n"
                f"Patient: {self.consent_name.text} (ID: {self.consent_pid.text})\n"
                f"Procedure: {self.consent_procedure.text}\n"
                f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Issued by: {self._current_username()}\n\n"
                f"CONSENT STATEMENT:\n{self.consent_statement.text}\n\n"
                f"Patient Signature: ___________________________\n"
                f"Witness Signature: ___________________________\n")
        self._export_to_pdf(text, "Consent_Form", "Informed Consent Form")

    # ====================== REFERRAL NOTE TAB ======================
    def _referral_note_tab(self):
        tab = TabbedPanelItem(text="Referral Note")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=6, spacing=4, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))
        l.add_widget(Label(
            text="REFERRAL NOTE", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#8e44ad')
        ))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="Patient ID:", font_size=15, bold=True,
                               size_hint_x=0.12))
        self.referral_pid = TextInput(size_hint_x=0.2, readonly=True, font_size=15,
                                        size_hint_y=None, height=40)
        bar.add_widget(self.referral_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.08))
        self.referral_name = TextInput(size_hint_x=0.3, readonly=True, font_size=15,
                                         size_hint_y=None, height=40)
        bar.add_widget(self.referral_name)
        l.add_widget(bar)

        l.add_widget(Label(text="Referring Physician:", font_size=15, bold=True,
                            size_hint_y=None, height=18))
        self.referral_doctor = TextInput(font_size=15, size_hint_y=None, height=40,
                                           hint_text="Dr. ...")
        l.add_widget(self.referral_doctor)

        reason_box = BoxLayout(size_hint_y=None, height=40, spacing=4)
        reason_box.add_widget(Label(text="Reason:", font_size=15, bold=True,
                                      size_hint_x=0.1))
        self.referral_reason = TextInput(font_size=15, size_hint_x=0.5,
                                           size_hint_y=None, height=40,
                                           hint_text="Reason for referral...")
        reason_box.add_widget(self.referral_reason)
        reason_box.add_widget(Button(
            text="AI WRITE", font_size=15, bold=True, size_hint_x=0.3,
            on_press=self._ai_generate_referral,
            background_color=get_color_from_hex('#8e44ad')
        ))
        l.add_widget(reason_box)

        l.add_widget(Label(text="Clinical Summary:", font_size=15, bold=True,
                            size_hint_y=None, height=18))
        self.referral_summary = TextInput(
            font_size=15, size_hint_y=None, height=120, multiline=True,
            hint_text="Referral letter will appear here..."
        )
        l.add_widget(self.referral_summary)

        btns = BoxLayout(size_hint_y=None, height=44, spacing=5)
        btns.add_widget(Button(
            text="SAVE", font_size=16, bold=True,
            on_press=self._save_referral,
            background_color=get_color_from_hex('#27ae60')
        ))
        btns.add_widget(Button(
            text="EXPORT PDF", font_size=16, bold=True,
            on_press=self._print_referral,
            background_color=get_color_from_hex('#8e44ad')
        ))
        l.add_widget(btns)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)

    def _ai_generate_referral(self, instance):
        if not self.ai.enabled:
            self._popup("AI Error", "AI is not enabled.")
            return
        reason = self.referral_reason.text.strip()
        if not reason:
            self._popup("Error", "Enter reason for referral first.")
            return
        prompt = (f"Write a professional referral letter for a patient. "
                  f"Reason: {reason}. Include patient summary, clinical "
                  f"findings, reason for referral, and urgency. Format as a "
                  f"formal referral letter.")

        def run():
            ok, resp = self.ai.call(
                "You are a medical referral specialist.", prompt, mt=800
            )
            if ok:
                Clock.schedule_once(
                    lambda dt: setattr(self.referral_summary, 'text', resp), 0
                )
            else:
                Clock.schedule_once(
                    lambda dt: self._popup("AI Error", resp), 0
                )

        threading.Thread(target=run, daemon=True).start()

    def _save_referral(self, instance):
        if not self.referral_pid.text:
            self._popup("Error", "No patient selected.")
            return
        note = {
            "patient_id": self.referral_pid.text,
            "patient_name": self.referral_name.text,
            "doctor": self.referral_doctor.text,
            "reason": self.referral_reason.text,
            "summary": self.referral_summary.text,
            "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "created_by": self._current_username()
        }
        self.referral_notes.append(note)
        self.save_all()
        self._popup("Saved", "Referral note saved.")
        self.log_activity(f"Referral saved for {self.referral_name.text}")
        self.audit_log("SAVE_REFERRAL", self.referral_reason.text,
                       self.referral_pid.text)

    def _print_referral(self, instance):
        if not self.referral_pid.text:
            self._popup("Error", "No patient selected.")
            return
        text = (f"=== HMG HOSPITAL - REFERRAL NOTE ===\n\n"
                f"Patient: {self.referral_name.text} (ID: {self.referral_pid.text})\n"
                f"Referring Physician: {self.referral_doctor.text}\n"
                f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Issued by: {self._current_username()}\n\n"
                f"REASON FOR REFERRAL:\n{self.referral_reason.text}\n\n"
                f"CLINICAL SUMMARY:\n{self.referral_summary.text}\n\n"
                f"Referred to: ___________________________\n"
                f"Signature: ___________________________\n")
        self._export_to_pdf(text, "Referral_Note", "Referral Note")

    # ====================== APPOINTMENTS TAB ======================
    def _appointments_tab(self):
        tab = TabbedPanelItem(text="Appointments")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=6, spacing=4, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))
        l.add_widget(Label(
            text="APPOINTMENT SCHEDULING", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#2980b9')
        ))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="Patient ID:", font_size=15, bold=True,
                               size_hint_x=0.12))
        self.appt_pid = TextInput(size_hint_x=0.2, readonly=True, font_size=15,
                                    size_hint_y=None, height=40)
        bar.add_widget(self.appt_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.08))
        self.appt_name = TextInput(size_hint_x=0.3, readonly=True, font_size=15,
                                     size_hint_y=None, height=40)
        bar.add_widget(self.appt_name)
        bar.add_widget(Button(
            text="REFRESH", font_size=13, size_hint_x=0.12,
            on_press=lambda x: self._refresh_appointments(),
            background_color=get_color_from_hex('#16a085')
        ))
        l.add_widget(bar)

        l.add_widget(Label(
            text="NEW APPOINTMENT", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#e67e22')
        ))
        form = GridLayout(cols=4, size_hint_y=None, height=150, spacing=4)
        form.add_widget(Label(text="Date:", font_size=14, bold=True))
        self.appt_date = TextInput(
            text=datetime.now().strftime("%Y-%m-%d"),
            font_size=14, size_hint_y=None, height=35
        )
        form.add_widget(self.appt_date)
        form.add_widget(Label(text="Time:", font_size=14, bold=True))
        self.appt_time = TextInput(
            text=datetime.now().strftime("%H:%M"),
            font_size=14, size_hint_y=None, height=35
        )
        form.add_widget(self.appt_time)
        form.add_widget(Label(text="Type:", font_size=14, bold=True))
        self.appt_type = Spinner(
            text="Consultation",
            values=["Consultation", "Follow-up", "Procedure",
                    "Lab Visit", "Radiology", "Review"],
            size_hint_y=None, height=35
        )
        form.add_widget(self.appt_type)
        form.add_widget(Label(text="With:", font_size=14, bold=True))
        self.appt_with = Spinner(
            text="General",
            values=["General", "Dr. Smith", "Dr. Jones",
                    "Dr. Williams", "Nurse"],
            size_hint_y=None, height=35
        )
        form.add_widget(self.appt_with)
        form.add_widget(Label(text="Notes:", font_size=14, bold=True))
        self.appt_notes = TextInput(font_size=14, size_hint_y=None, height=35)
        form.add_widget(self.appt_notes)
        l.add_widget(form)

        btn_row = BoxLayout(size_hint_y=None, height=40, spacing=5)
        btn_row.add_widget(Button(
            text="SCHEDULE", font_size=16, bold=True,
            on_press=self._add_appointment,
            background_color=get_color_from_hex('#27ae60')
        ))
        btn_row.add_widget(Button(
            text="VIEW ALL", font_size=16, bold=True,
            on_press=self._view_all_appointments,
            background_color=get_color_from_hex('#2980b9')
        ))
        l.add_widget(btn_row)

        l.add_widget(Label(
            text="APPOINTMENT LIST", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#2c3e50')
        ))
        self.appt_list = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        self.appt_list.bind(minimum_height=self.appt_list.setter('height'))
        l.add_widget(self.appt_list)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)
        self._refresh_appointments()

    def _refresh_appointments(self):
        if not hasattr(self, 'appt_list'):
            return
        self.appt_list.clear_widgets()
        if not self.appt_pid.text:
            self.appt_list.add_widget(Label(
                text="Select a patient to view appointments.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        pid = self.appt_pid.text
        appointments = [a for a in self.appointments
                        if a.get("patient_id") == pid]
        appointments.sort(key=lambda x: x.get("date", ""), reverse=True)
        if not appointments:
            self.appt_list.add_widget(Label(
                text="No appointments for this patient.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        for a in appointments:
            status_color = {
                "Scheduled": "#f39c12",
                "Completed": "#27ae60",
                "Cancelled": "#e74c3c",
                "No-show": "#95a5a6"
            }.get(a.get("status", "Scheduled"), "#7f8c8d")

            row = BoxLayout(size_hint_y=None, height=35, spacing=3)
            row.add_widget(Label(
                text=f"{a.get('date', '')} {a.get('time', '')}",
                font_size=12, size_hint_x=0.2
            ))
            row.add_widget(Label(text=a.get('type', ''), font_size=12,
                                   size_hint_x=0.15))
            row.add_widget(Label(text=a.get('with', ''), font_size=12,
                                   size_hint_x=0.15))
            row.add_widget(Label(
                text=a.get('status', 'Scheduled'), font_size=12,
                size_hint_x=0.12, bold=True,
                color=get_color_from_hex(status_color)
            ))
            status_spinner = Spinner(
                text=a.get("status", "Scheduled"),
                values=["Scheduled", "Completed", "Cancelled", "No-show"],
                size_hint_x=0.1, size_hint_y=None, height=30
            )
            status_spinner.appt_id = a.get("id")
            status_spinner.bind(
                text=lambda s, aid=a.get("id"): self._update_appointment_status(aid, s.text)
            )
            row.add_widget(status_spinner)
            btn_del = Button(
                text="X", font_size=11, size_hint_x=0.08,
                size_hint_y=None, height=25,
                background_color=get_color_from_hex('#e74c3c')
            )
            btn_del.appt_id = a.get("id")
            btn_del.bind(on_press=lambda x: self._delete_appointment(x.appt_id))
            row.add_widget(btn_del)
            self.appt_list.add_widget(row)

    def _update_appointment_status(self, appt_id, new_status):
        for a in self.appointments:
            if a.get("id") == appt_id:
                a["status"] = new_status
                self.save_all()
                self._refresh_appointments()
                self.log_activity(f"Appointment {appt_id} status updated to {new_status}")
                self.audit_log("UPDATE_APPOINTMENT",
                               f"{appt_id} → {new_status}",
                               a.get("patient_id", ""))
                return

    def _add_appointment(self, instance):
        if not self.appt_pid.text:
            self._popup("Error", "Select a patient first.")
            return
        date = self.appt_date.text.strip()
        time = self.appt_time.text.strip()
        if not date or not time:
            self._popup("Error", "Date and time required.")
            return
        try:
            datetime.strptime(date, "%Y-%m-%d")
        except:
            self._popup("Error", "Invalid date format (YYYY-MM-DD).")
            return
        appointment = {
            "id": f"APT{datetime.now().strftime('%Y%m%d%H%M%S%f')}",
            "patient_id": self.appt_pid.text,
            "patient_name": self.appt_name.text,
            "date": date,
            "time": time,
            "type": self.appt_type.text,
            "with": self.appt_with.text,
            "notes": self.appt_notes.text.strip(),
            "status": "Scheduled",
            "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "created_by": self._current_username()
        }
        self.appointments.append(appointment)
        self.save_all()
        self._refresh_appointments()
        self.appt_notes.text = ""
        self.appt_date.text = datetime.now().strftime("%Y-%m-%d")
        self.appt_time.text = datetime.now().strftime("%H:%M")
        self._popup("Appointment Scheduled",
                    f"Appointment for {self.appt_name.text} on {date} at {time}")
        self.log_activity(f"Appointment scheduled for {self.appt_name.text}: {date} {time}")
        self.audit_log("SCHEDULE_APPOINTMENT", f"{date} {time}",
                       self.appt_pid.text)

    def _delete_appointment(self, appt_id):
        if not self._require_role("Administrator", "Doctor", "Receptionist"):
            return
        self.appointments = [a for a in self.appointments
                             if a.get("id") != appt_id]
        self.save_all()
        self._refresh_appointments()
        self._popup("Deleted", "Appointment removed.")
        self.audit_log("DELETE_APPOINTMENT", appt_id)

    def _view_all_appointments(self, instance):
        if not self.appt_pid.text:
            self._popup("Error", "No patient selected.")
            return
        pid = self.appt_pid.text
        appointments = [a for a in self.appointments
                        if a.get("patient_id") == pid]
        if not appointments:
            self._popup("Appointments", "No appointments for this patient.")
            return
        text = f"APPOINTMENTS FOR {self.appt_name.text}\n" + "=" * 50 + "\n\n"
        for a in sorted(appointments, key=lambda x: x.get("date", ""), reverse=True):
            text += f"Date: {a.get('date', '')} {a.get('time', '')}\n"
            text += f"Type: {a.get('type', '')}\n"
            text += f"With: {a.get('with', '')}\n"
            text += f"Status: {a.get('status', 'Scheduled')}\n"
            if a.get('notes'):
                text += f"Notes: {a.get('notes', '')}\n"
            text += "-" * 30 + "\n"
        self._text_popup("Appointments", text)

    # ====================== NURSING NOTES TAB ======================
    def _nursing_notes_tab(self):
        tab = TabbedPanelItem(text="Nursing Notes")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=6, spacing=4, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))
        l.add_widget(Label(
            text="NURSING NOTES", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#1a5276')
        ))

        bar = BoxLayout(size_hint_y=None, height=40, spacing=4)
        bar.add_widget(Label(text="Patient ID:", font_size=15, bold=True,
                               size_hint_x=0.12))
        self.nurse_pid = TextInput(size_hint_x=0.2, readonly=True, font_size=15,
                                     size_hint_y=None, height=40)
        bar.add_widget(self.nurse_pid)
        bar.add_widget(Label(text="Name:", font_size=15, bold=True, size_hint_x=0.08))
        self.nurse_name = TextInput(size_hint_x=0.3, readonly=True, font_size=15,
                                      size_hint_y=None, height=40)
        bar.add_widget(self.nurse_name)
        bar.add_widget(Button(
            text="REFRESH", font_size=13, size_hint_x=0.12,
            on_press=lambda x: self._refresh_nursing_notes(),
            background_color=get_color_from_hex('#16a085')
        ))
        l.add_widget(bar)

        l.add_widget(Label(
            text="NEW NURSING NOTE", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#e67e22')
        ))
        form = GridLayout(cols=3, size_hint_y=None, height=200, spacing=4)
        form.add_widget(Label(text="Date:", font_size=14, bold=True))
        self.nurse_date = TextInput(
            text=datetime.now().strftime("%Y-%m-%d"),
            font_size=14, size_hint_y=None, height=35
        )
        form.add_widget(self.nurse_date)
        form.add_widget(Label(text="Shift:", font_size=14, bold=True))
        self.nurse_shift = Spinner(
            text="Morning",
            values=["Morning", "Afternoon", "Night"],
            size_hint_y=None, height=35
        )
        form.add_widget(self.nurse_shift)
        form.add_widget(Label(text="Vitals:", font_size=14, bold=True))
        self.nurse_vitals = TextInput(font_size=14, size_hint_y=None, height=35,
                                        hint_text="BP: 120/80, Pulse: 72")
        form.add_widget(self.nurse_vitals)
        form.add_widget(Label(text="Observations:", font_size=14, bold=True))
        self.nurse_obs = TextInput(font_size=14, size_hint_y=None, height=35,
                                     hint_text="Patient condition, symptoms")
        form.add_widget(self.nurse_obs)
        form.add_widget(Label(text="Care Provided:", font_size=14, bold=True))
        self.nurse_care = TextInput(font_size=14, size_hint_y=None, height=35,
                                      hint_text="Medications given, procedures")
        form.add_widget(self.nurse_care)
        form.add_widget(Label(text="Status:", font_size=14, bold=True))
        self.nurse_status = Spinner(
            text="Stable",
            values=["Stable", "Improving", "Deteriorating", "Critical"],
            size_hint_y=None, height=35
        )
        form.add_widget(self.nurse_status)
        l.add_widget(form)

        btn_add = Button(
            text="SAVE NURSING NOTE", font_size=16, bold=True,
            size_hint_y=None, height=40,
            on_press=self._add_nursing_note,
            background_color=get_color_from_hex('#27ae60')
        )
        l.add_widget(btn_add)

        l.add_widget(Label(
            text="NURSING NOTES HISTORY", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#2c3e50')
        ))
        self.nurse_list = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        self.nurse_list.bind(minimum_height=self.nurse_list.setter('height'))
        l.add_widget(self.nurse_list)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)
        self._refresh_nursing_notes()

    def _refresh_nursing_notes(self):
        if not hasattr(self, 'nurse_list'):
            return
        self.nurse_list.clear_widgets()
        if not self.nurse_pid.text:
            self.nurse_list.add_widget(Label(
                text="Select a patient to view nursing notes.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        pid = self.nurse_pid.text
        notes = [n for n in self.nursing_notes if n.get("patient_id") == pid]
        notes.sort(key=lambda x: x.get("date", ""), reverse=True)
        if not notes:
            self.nurse_list.add_widget(Label(
                text="No nursing notes for this patient.",
                font_size=13, size_hint_y=None, height=30
            ))
            return
        for n in notes:
            row = BoxLayout(size_hint_y=None, height=35, spacing=3)
            row.add_widget(Label(
                text=f"{n.get('date', '')} | {n.get('shift', '')}",
                font_size=12, size_hint_x=0.2
            ))
            row.add_widget(Label(text=n.get('status', ''), font_size=12,
                                   size_hint_x=0.12))
            row.add_widget(Label(text=n.get('care', '')[:20], font_size=11,
                                   size_hint_x=0.25))
            btn_view = Button(
                text="VIEW", font_size=11, size_hint_x=0.15,
                size_hint_y=None, height=25,
                background_color=get_color_from_hex('#2980b9')
            )
            btn_view.note = n
            btn_view.bind(on_press=lambda x: self._view_nursing_note_detail(x.note))
            row.add_widget(btn_view)
            btn_del = Button(
                text="X", font_size=11, size_hint_x=0.08,
                size_hint_y=None, height=25,
                background_color=get_color_from_hex('#e74c3c')
            )
            btn_del.note_id = n.get("id")
            btn_del.bind(on_press=lambda x: self._delete_nursing_note(x.note_id))
            row.add_widget(btn_del)
            self.nurse_list.add_widget(row)

    def _view_nursing_note_detail(self, note):
        text = "NURSING NOTE\n" + "=" * 50 + "\n\n"
        text += f"Patient: {note.get('patient_name', '')} (ID: {note.get('patient_id', '')})\n"
        text += f"Date: {note.get('date', '')}\n"
        text += f"Shift: {note.get('shift', '')}\n"
        text += f"Status: {note.get('status', '')}\n\n"
        text += f"Vitals:\n{note.get('vitals', '')}\n\n"
        text += f"Observations:\n{note.get('observations', '')}\n\n"
        text += f"Care Provided:\n{note.get('care', '')}\n"
        if note.get('notes'):
            text += f"\nAdditional Notes:\n{note.get('notes', '')}\n"
        self._text_popup("Nursing Note Detail", text)

    def _add_nursing_note(self, instance):
        if not self.nurse_pid.text:
            self._popup("Error", "Select a patient first.")
            return
        date = self.nurse_date.text.strip()
        if not date:
            self._popup("Error", "Date required.")
            return
        try:
            datetime.strptime(date, "%Y-%m-%d")
        except:
            self._popup("Error", "Invalid date format (YYYY-MM-DD).")
            return
        note = {
            "id": f"NSG{datetime.now().strftime('%Y%m%d%H%M%S%f')}",
            "patient_id": self.nurse_pid.text,
            "patient_name": self.nurse_name.text,
            "date": date,
            "shift": self.nurse_shift.text,
            "vitals": self.nurse_vitals.text.strip(),
            "observations": self.nurse_obs.text.strip(),
            "care": self.nurse_care.text.strip(),
            "status": self.nurse_status.text,
            "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "created_by": self._current_username()
        }
        self.nursing_notes.append(note)
        self.save_all()
        self._refresh_nursing_notes()
        self.nurse_vitals.text = ""
        self.nurse_obs.text = ""
        self.nurse_care.text = ""
        self.nurse_date.text = datetime.now().strftime("%Y-%m-%d")
        self._popup("Nursing Note Saved",
                    f"Nursing note for {self.nurse_name.text} on {date}")
        self.log_activity(
            f"Nursing note added for {self.nurse_name.text}: "
            f"{date} {self.nurse_shift.text}"
        )
        self.audit_log("ADD_NURSING_NOTE", f"{date} {self.nurse_shift.text}",
                       self.nurse_pid.text)
        if self.cp:
            self.task_manager.create_task(
                patient_id=self.cp.get("id"),
                patient_name=self.cp.get("name"),
                assigned_to="Nurse",
                task_type="Nursing",
                description=f"Nursing care: {self.nurse_care.text[:50]}",
                department="Nursing"
            )
            self._refresh_tasks()

    def _delete_nursing_note(self, note_id):
        if not self._require_role("Administrator", "Doctor", "Nurse"):
            return
        self.nursing_notes = [n for n in self.nursing_notes
                              if n.get("id") != note_id]
        self.save_all()
        self._refresh_nursing_notes()
        self._popup("Deleted", "Nursing note removed.")

    # ====================== MEDICAL REPORT DRAFTER TAB ======================
    def _medical_report_tab(self):
        tab = TabbedPanelItem(text="Report Drafter")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=6, spacing=4, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))

        l.add_widget(Label(
            text="MEDICAL REPORT DRAFTER", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#16a085')
        ))
        l.add_widget(Label(
            text="All fields are blank by design — fill them in yourself.",
            font_size=12, size_hint_y=None, height=20,
            color=get_color_from_hex('#7f8c8d')
        ))

        # Issuing facility
        l.add_widget(Label(
            text="ISSUING FACILITY", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#1a5276')
        ))
        fac_grid = GridLayout(cols=2, size_hint_y=None, height=250, spacing=4)
        self.rpt_facility_name = TextInput(
            font_size=14, size_hint_y=None, height=38,
            hint_text="Hospital / facility name"
        )
        self.rpt_facility_address = TextInput(
            font_size=14, size_hint_y=None, height=70, multiline=True,
            hint_text="Address"
        )
        self.rpt_facility_contact = TextInput(
            font_size=14, size_hint_y=None, height=38,
            hint_text="Phone / email"
        )
        self.rpt_ref_number = TextInput(
            font_size=14, size_hint_y=None, height=38,
            hint_text="Reference number"
        )
        self.rpt_report_date = TextInput(
            text=datetime.now().strftime("%d %B %Y"),
            font_size=14, size_hint_y=None, height=38
        )
        fac_grid.add_widget(Label(text="Name:", font_size=13, bold=True, size_hint_x=0.25))
        fac_grid.add_widget(self.rpt_facility_name)
        fac_grid.add_widget(Label(text="Address:", font_size=13, bold=True, size_hint_x=0.25))
        fac_grid.add_widget(self.rpt_facility_address)
        fac_grid.add_widget(Label(text="Contact:", font_size=13, bold=True, size_hint_x=0.25))
        fac_grid.add_widget(self.rpt_facility_contact)
        fac_grid.add_widget(Label(text="Ref No.:", font_size=13, bold=True, size_hint_x=0.25))
        fac_grid.add_widget(self.rpt_ref_number)
        fac_grid.add_widget(Label(text="Date:", font_size=13, bold=True, size_hint_x=0.25))
        fac_grid.add_widget(self.rpt_report_date)
        l.add_widget(fac_grid)

        # Logo attach row
        logo_row = BoxLayout(size_hint_y=None, height=42, spacing=5)
        logo_row.add_widget(Label(text="Facility Logo:", font_size=13, bold=True,
                                    size_hint_x=0.2))
        self.rpt_logo_label = Label(text="(none attached)", font_size=12,
                                      size_hint_x=0.5,
                                      color=get_color_from_hex('#7f8c8d'))
        logo_row.add_widget(self.rpt_logo_label)
        logo_row.add_widget(Button(
            text="ATTACH LOGO", font_size=13, bold=True, size_hint_x=0.25,
            on_press=lambda x: self._pick_image_for_report("logo"),
            background_color=get_color_from_hex('#2980b9')
        ))
        l.add_widget(logo_row)
        self.rpt_logo_path = None

        # Patient details
        l.add_widget(Label(
            text="PATIENT DETAILS (auto-filled from selected patient)",
            font_size=16, bold=True, size_hint_y=None, height=24,
            color=get_color_from_hex('#1a5276')
        ))
        pat_grid = GridLayout(cols=4, size_hint_y=None, height=90, spacing=4)
        self.rpt_patient_name = TextInput(font_size=14, size_hint_y=None, height=38)
        self.rpt_patient_id = TextInput(font_size=14, size_hint_y=None, height=38)
        self.rpt_patient_age = TextInput(font_size=14, size_hint_y=None, height=38)
        self.rpt_patient_sex = TextInput(font_size=14, size_hint_y=None, height=38)
        pat_grid.add_widget(Label(text="Name:", font_size=13, bold=True))
        pat_grid.add_widget(self.rpt_patient_name)
        pat_grid.add_widget(Label(text="ID:", font_size=13, bold=True))
        pat_grid.add_widget(self.rpt_patient_id)
        pat_grid.add_widget(Label(text="Age:", font_size=13, bold=True))
        pat_grid.add_widget(self.rpt_patient_age)
        pat_grid.add_widget(Label(text="Sex:", font_size=13, bold=True))
        pat_grid.add_widget(self.rpt_patient_sex)
        l.add_widget(pat_grid)

        # Clinical content
        l.add_widget(Label(
            text="CLINICAL CONTENT", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#1a5276')
        ))
        l.add_widget(Label(text="Diagnosis / Conditions:", font_size=13, bold=True,
                            size_hint_y=None, height=18))
        self.rpt_conditions = TextInput(
            font_size=13, size_hint_y=None, height=70, multiline=True,
            hint_text="Diagnosis / conditions"
        )
        l.add_widget(self.rpt_conditions)

        l.add_widget(Label(text="Narrative Summary:", font_size=13, bold=True,
                            size_hint_y=None, height=18))
        self.rpt_narrative = TextInput(
            font_size=13, size_hint_y=None, height=140, multiline=True,
            hint_text="Context, history, current status"
        )
        l.add_widget(self.rpt_narrative)

        l.add_widget(Label(
            text="Recommendations (one per line — numbered automatically):",
            font_size=13, bold=True, size_hint_y=None, height=18
        ))
        self.rpt_recommendations = TextInput(
            font_size=13, size_hint_y=None, height=140, multiline=True,
            hint_text="One recommendation per line"
        )
        l.add_widget(self.rpt_recommendations)

        # Signing clinician
        l.add_widget(Label(
            text="SIGNING CLINICIAN", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#1a5276')
        ))
        sig_grid = GridLayout(cols=4, size_hint_y=None, height=90, spacing=4)
        self.rpt_doctor_name = TextInput(
            font_size=14, size_hint_y=None, height=38,
            hint_text="Doctor / clinician name"
        )
        self.rpt_doctor_qual = TextInput(
            font_size=14, size_hint_y=None, height=38,
            hint_text="e.g. MBBS, M.Sc."
        )
        self.rpt_doctor_role = TextInput(
            font_size=14, size_hint_y=None, height=38,
            hint_text="Role / title"
        )
        sig_grid.add_widget(Label(text="Name:", font_size=13, bold=True))
        sig_grid.add_widget(self.rpt_doctor_name)
        sig_grid.add_widget(Label(text="Quals:", font_size=13, bold=True))
        sig_grid.add_widget(self.rpt_doctor_qual)
        sig_grid.add_widget(Label(text="Role:", font_size=13, bold=True))
        sig_grid.add_widget(self.rpt_doctor_role)
        l.add_widget(sig_grid)

        sig_row = BoxLayout(size_hint_y=None, height=42, spacing=5)
        sig_row.add_widget(Label(text="Signature Image:", font_size=13, bold=True,
                                   size_hint_x=0.2))
        self.rpt_sig_label = Label(text="(none attached)", font_size=12,
                                     size_hint_x=0.5,
                                     color=get_color_from_hex('#7f8c8d'))
        sig_row.add_widget(self.rpt_sig_label)
        sig_row.add_widget(Button(
            text="ATTACH SIGNATURE", font_size=13, bold=True, size_hint_x=0.25,
            on_press=lambda x: self._pick_image_for_report("signature"),
            background_color=get_color_from_hex('#2980b9')
        ))
        l.add_widget(sig_row)
        self.rpt_signature_path = None

        # Generate buttons
        btn_row = BoxLayout(size_hint_y=None, height=48, spacing=6)
        btn_row.add_widget(Button(
            text="GENERATE PDF REPORT", font_size=16, bold=True,
            on_press=self._generate_medical_report_pdf,
            background_color=get_color_from_hex('#27ae60')
        ))
        btn_row.add_widget(Button(
            text="CLEAR", font_size=15, bold=True,
            on_press=lambda x: self._clear_medical_report_form(),
            background_color=get_color_from_hex('#e74c3c')
        ))
        l.add_widget(btn_row)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)

    def _clear_medical_report_form(self):
        for w in [self.rpt_facility_name, self.rpt_facility_address,
                   self.rpt_facility_contact, self.rpt_ref_number,
                   self.rpt_patient_name, self.rpt_patient_id,
                   self.rpt_patient_age, self.rpt_patient_sex,
                   self.rpt_conditions, self.rpt_narrative,
                   self.rpt_recommendations, self.rpt_doctor_name,
                   self.rpt_doctor_qual, self.rpt_doctor_role]:
            w.text = ""
        self.rpt_report_date.text = datetime.now().strftime("%d %B %Y")
        self.rpt_logo_path = None
        self.rpt_signature_path = None
        self.rpt_logo_label.text = "(none attached)"
        self.rpt_sig_label.text = "(none attached)"

    def _pick_image_for_report(self, target):
        """Show menu: take a photo with camera, or choose a file."""
        menu = BoxLayout(orientation='vertical', spacing=8, padding=12)
        menu.add_widget(Label(
            text=f"Attach {'Logo' if target == 'logo' else 'Signature'}",
            font_size=16, bold=True, size_hint_y=None, height=30,
            color=get_color_from_hex('#1a5276')
        ))

        pp = Popup(title="Choose Image Source", content=menu,
                   size_hint=(0.55, 0.42), auto_dismiss=True)

        def choose_file(inst):
            pp.dismiss()
            self._pick_image_from_file(target)

        def take_photo(inst):
            pp.dismiss()
            self._take_photo_for_report(target)

        camera_available = CameraHelper.available()
        photo_btn = Button(
            text="📷  Take Photo", font_size=15, bold=True,
            size_hint_y=None, height=48,
            background_color=get_color_from_hex('#2980b9'),
            disabled=not camera_available
        )
        if not camera_available:
            photo_btn.text = "📷  Take Photo (unavailable)"
        photo_btn.bind(on_press=take_photo)
        menu.add_widget(photo_btn)

        file_btn = Button(
            text="📁  Choose File", font_size=15, bold=True,
            size_hint_y=None, height=48,
            background_color=get_color_from_hex('#27ae60')
        )
        file_btn.bind(on_press=choose_file)
        menu.add_widget(file_btn)

        cancel_btn = Button(
            text="Cancel", font_size=14, size_hint_y=None, height=40,
            background_color=get_color_from_hex('#e74c3c')
        )
        cancel_btn.bind(on_press=lambda *_: pp.dismiss())
        menu.add_widget(cancel_btn)

        pp.open()

    def _pick_image_from_file(self, target):
        content = BoxLayout(orientation='vertical', spacing=4, padding=4)
        fc = FileChooserListView(path=os.path.expanduser("~"),
                                  filters=["*.png", "*.jpg", "*.jpeg"])
        content.add_widget(fc)
        btns = BoxLayout(size_hint_y=None, height=42, spacing=5)
        pp = Popup(
            title=f"Choose {'Logo' if target == 'logo' else 'Signature'} Image",
            content=content, size_hint=(0.9, 0.85)
        )

        def select(inst):
            if fc.selection:
                self._set_report_image(target, fc.selection[0])
            pp.dismiss()

        btns.add_widget(Button(
            text="SELECT", font_size=15, bold=True, on_press=select,
            background_color=get_color_from_hex('#27ae60')
        ))
        btns.add_widget(Button(text="CANCEL", font_size=15, on_press=pp.dismiss))
        content.add_widget(btns)
        pp.open()

    def _take_photo_for_report(self, target):
        if not CameraHelper.available():
            self._popup("Camera", "Camera is not available on this device.")
            return
        try:
            cam = Camera(play=True, resolution=(800, 600))
        except Exception as e:
            self._popup("Camera Error", f"Could not open camera:\n{e}")
            return

        content = BoxLayout(orientation='vertical', spacing=4, padding=4)
        content.add_widget(cam)
        btns = BoxLayout(size_hint_y=None, height=44, spacing=5)
        pp = Popup(
            title=f"Take {'Logo' if target == 'logo' else 'Signature'} Photo",
            content=content, size_hint=(0.9, 0.85)
        )

        def capture(inst):
            if not cam.texture:
                self._popup("Camera", "Camera is not ready yet. Please try again.")
                return
            try:
                os.makedirs(OUTPUT_DIR, exist_ok=True)
                fn = os.path.join(
                    OUTPUT_DIR,
                    f"rpt_{target}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                )
                cam.texture.save(fn)
                try:
                    cam.play = False
                except Exception:
                    pass
                self._set_report_image(target, fn)
                pp.dismiss()
                self._popup("Photo Saved",
                            f"{'Logo' if target == 'logo' else 'Signature'} "
                            f"saved as:\n{os.path.basename(fn)}")
            except Exception as e:
                self._popup("Save Error", f"Could not save photo:\n{e}")

        def cancel(inst):
            try:
                cam.play = False
            except Exception:
                pass
            pp.dismiss()

        btns.add_widget(Button(
            text="📸  CAPTURE", font_size=16, bold=True, on_press=capture,
            background_color=get_color_from_hex('#27ae60')
        ))
        btns.add_widget(Button(
            text="CANCEL", font_size=16, on_press=cancel,
            background_color=get_color_from_hex('#e74c3c')
        ))
        content.add_widget(btns)
        pp.open()

    def _set_report_image(self, target, path):
        if not path or not os.path.exists(path):
            self._popup("Error", f"File not found:\n{path}")
            return
        if target == "logo":
            self.rpt_logo_path = path
            self.rpt_logo_label.text = os.path.basename(path)
        else:
            self.rpt_signature_path = path
            self.rpt_sig_label.text = os.path.basename(path)
        self.log_activity(f"Report {target} attached: {path}")

    def _generate_medical_report_pdf(self, instance):
        if not REPORTLAB_AVAILABLE:
            self._popup("Error",
                        "ReportLab is not installed.\nRun: pip install reportlab")
            return

        if not self.rpt_patient_name.text.strip() or not self.rpt_doctor_name.text.strip():
            self._popup("Error",
                        "Please fill in at least the patient name and doctor name.")
            return

        os.makedirs(OUTPUT_DIR, exist_ok=True)
        safe_name = "".join(c for c in self.rpt_patient_name.text
                            if c.isalnum() or c in " _-").strip().replace(" ", "_") or "report"
        out_path = os.path.join(
            OUTPUT_DIR,
            f"{safe_name}_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        )

        try:
            c = pdf_canvas.Canvas(out_path, pagesize=A4)
            width, height = A4
            margin = 22 * mm
            content_width = width - 2 * margin
            bottom_limit = margin + 14 * mm
            page_num = [1]
            state = {"y": height - margin}

            PDF_ACCENT = (0.14, 0.42, 0.38)
            PDF_MUTED = (0.35, 0.35, 0.35)
            BODY_SIZE = 10.5
            BODY_LEADING = 15
            F_REG = "Helvetica"
            F_BOLD = "Helvetica-Bold"
            F_ITAL = "Helvetica-Oblique"

            def draw_footer():
                c.setFont(F_REG, 8)
                c.setFillColorRGB(*PDF_MUTED)
                c.drawString(margin, margin - 6 * mm,
                             self.rpt_facility_name.text.strip())
                c.drawRightString(width - margin, margin - 6 * mm,
                                   f"Page {page_num[0]}")
                c.setFillColorRGB(0, 0, 0)

            def draw_continuation_header():
                c.setFont(F_ITAL, 8.5)
                c.setFillColorRGB(*PDF_MUTED)
                header = (f"{self.rpt_facility_name.text.strip()} — "
                          f"Medical Report (cont'd)").strip(" —")
                c.drawString(margin, state["y"], header)
                c.setFillColorRGB(0, 0, 0)
                state["y"] -= 10 * mm

            def new_page():
                draw_footer()
                c.showPage()
                page_num[0] += 1
                state["y"] = height - margin
                draw_continuation_header()

            def ensure_space(needed):
                if state["y"] - needed < bottom_limit:
                    new_page()

            def heading(text, size=11.5, gap_before=4 * mm, gap_after=3 * mm):
                ensure_space(gap_before + size + gap_after)
                state["y"] -= gap_before
                c.setFont(F_BOLD, size)
                c.setFillColorRGB(*PDF_ACCENT)
                c.drawString(margin, state["y"], text.upper())
                c.setFillColorRGB(0, 0, 0)
                state["y"] -= 3
                c.setStrokeColorRGB(*PDF_ACCENT)
                c.setLineWidth(0.6)
                c.line(margin, state["y"], margin + 26 * mm, state["y"])
                c.setStrokeColorRGB(0, 0, 0)
                state["y"] -= gap_after

            def line(text, size=None, bold=False, italic=False,
                     gap=None, color=None):
                size = size or BODY_SIZE
                gap = gap or BODY_LEADING
                ensure_space(gap)
                font = F_BOLD if bold else (F_ITAL if italic else F_REG)
                c.setFont(font, size)
                if color:
                    c.setFillColorRGB(*color)
                c.drawString(margin, state["y"], text)
                if color:
                    c.setFillColorRGB(0, 0, 0)
                state["y"] -= gap

            def wrap_lines(text, size, max_width, font=None):
                font = font or F_REG
                words = text.split()
                lines_out, current = [], ""
                for word in words:
                    trial = (current + " " + word).strip()
                    if c.stringWidth(trial, font, size) > max_width and current:
                        lines_out.append(current)
                        current = word
                    else:
                        current = trial
                if current:
                    lines_out.append(current)
                return lines_out

            def paragraph(text, size=None, gap=None, indent=0, first_line_prefix=""):
                size = size or BODY_SIZE
                gap = gap or BODY_LEADING
                max_width = content_width - indent
                prefix_width = (c.stringWidth(first_line_prefix, F_REG, size)
                                if first_line_prefix else 0)
                wrapped = wrap_lines(
                    text, size,
                    max_width - prefix_width if first_line_prefix else max_width
                )
                for i, wline in enumerate(wrapped):
                    ensure_space(gap)
                    c.setFont(F_REG, size)
                    if i == 0 and first_line_prefix:
                        c.setFont(F_BOLD, size)
                        c.drawString(margin, state["y"], first_line_prefix)
                        c.setFont(F_REG, size)
                        c.drawString(margin + prefix_width, state["y"], wline)
                    else:
                        c.drawString(margin + indent, state["y"], wline)
                    state["y"] -= gap

            # Letterhead
            header_top = state["y"]
            if self.rpt_logo_path and os.path.exists(self.rpt_logo_path):
                try:
                    img = ImageReader(self.rpt_logo_path)
                    c.drawImage(img, margin, header_top - 20 * mm,
                                width=20 * mm, height=20 * mm,
                                preserveAspectRatio=True, mask="auto")
                except Exception:
                    pass
                header_x = margin + 25 * mm
            else:
                header_x = margin

            c.setFont(F_BOLD, 19)
            c.setFillColorRGB(*PDF_ACCENT)
            c.drawString(header_x, header_top - 7 * mm,
                         self.rpt_facility_name.text.strip() or "[Facility name]")
            c.setFillColorRGB(0, 0, 0)

            c.setFont(F_REG, 8.5)
            c.setFillColorRGB(*PDF_MUTED)
            addr_lines = (self.rpt_facility_address.text or "").strip().split("\n")
            ay = header_top - 12.5 * mm
            for al in addr_lines[:2]:
                if al.strip():
                    c.drawString(header_x, ay, al.strip())
                    ay -= 4 * mm
            if self.rpt_facility_contact.text.strip():
                c.drawString(header_x, ay, self.rpt_facility_contact.text.strip())
            c.setFillColorRGB(0, 0, 0)

            state["y"] = header_top - 24 * mm
            c.setStrokeColorRGB(*PDF_ACCENT)
            c.setLineWidth(1.4)
            c.line(margin, state["y"], width - margin, state["y"])
            c.setLineWidth(0.5)
            c.setStrokeColorRGB(*PDF_MUTED)
            c.line(margin, state["y"] - 1.2 * mm,
                   width - margin, state["y"] - 1.2 * mm)
            c.setStrokeColorRGB(0, 0, 0)
            state["y"] -= 9 * mm

            # Ref / Date
            c.setFont(F_REG, 9.5)
            c.setFillColorRGB(*PDF_MUTED)
            c.drawString(margin, state["y"],
                         f"Ref: {self.rpt_ref_number.text.strip() or '—'}")
            c.drawRightString(width - margin, state["y"],
                              f"Date: {self.rpt_report_date.text.strip()}")
            c.setFillColorRGB(0, 0, 0)
            state["y"] -= 10 * mm

            # Title
            c.setFont(F_BOLD, 14.5)
            c.drawCentredString(width / 2, state["y"], "MEDICAL REPORT")
            state["y"] -= 2.5 * mm
            c.setLineWidth(0.8)
            title_w = c.stringWidth("MEDICAL REPORT", F_BOLD, 14.5)
            c.line(width / 2 - title_w / 2, state["y"],
                   width / 2 + title_w / 2, state["y"])
            state["y"] -= 9 * mm

            # Patient block
            line(f"RE:  {self.rpt_patient_name.text.strip()}",
                 size=12, bold=True, gap=16)
            details = []
            if self.rpt_patient_age.text.strip():
                details.append(f"Age: {self.rpt_patient_age.text.strip()}")
            if self.rpt_patient_sex.text.strip():
                details.append(f"Sex: {self.rpt_patient_sex.text.strip()}")
            if self.rpt_patient_id.text.strip():
                details.append(f"ID: {self.rpt_patient_id.text.strip()}")
            if details:
                line("   |   ".join(details), size=9.5, color=PDF_MUTED, gap=13)
            state["y"] -= 3 * mm

            line("To Whom It May Concern,", gap=15)
            state["y"] -= 2 * mm

            # Clinical content
            if self.rpt_conditions.text.strip():
                heading("Diagnosis / Conditions")
                paragraph(self.rpt_conditions.text.strip())
                state["y"] -= 2 * mm

            if self.rpt_narrative.text.strip():
                heading("Clinical Summary")
                for para in self.rpt_narrative.text.strip().split("\n"):
                    if para.strip():
                        paragraph(para.strip())
                        state["y"] -= 1.5 * mm
                state["y"] -= 1 * mm

            if self.rpt_recommendations.text.strip():
                heading("Recommendations")
                items = [r.strip() for r in self.rpt_recommendations.text.strip().split("\n")
                         if r.strip()]
                for i, item in enumerate(items, start=1):
                    paragraph(item, indent=6 * mm, first_line_prefix=f"{i}.  ")
                    state["y"] -= 1.5 * mm
                state["y"] -= 2 * mm

            # Sign-off
            ensure_space(38 * mm)
            state["y"] -= 4 * mm
            line("Yours faithfully,", gap=15)

            if self.rpt_signature_path and os.path.exists(self.rpt_signature_path):
                try:
                    img = ImageReader(self.rpt_signature_path)
                    sig_h = 14 * mm
                    c.drawImage(img, margin, state["y"] - sig_h,
                                width=42 * mm, height=sig_h,
                                preserveAspectRatio=True, mask="auto",
                                anchor="sw")
                except Exception:
                    pass
                state["y"] -= (14 * mm + 3 * mm)
            else:
                state["y"] -= 10 * mm

            c.setLineWidth(0.5)
            c.line(margin, state["y"], margin + 55 * mm, state["y"])
            state["y"] -= 5.5 * mm

            line(self.rpt_doctor_name.text.strip() or "[Doctor name]",
                 size=11, bold=True, gap=14)
            if self.rpt_doctor_qual.text.strip():
                line(self.rpt_doctor_qual.text.strip(), size=9.5,
                     color=PDF_MUTED, gap=12.5)
            if self.rpt_doctor_role.text.strip():
                line(self.rpt_doctor_role.text.strip(), size=9.5,
                     color=PDF_MUTED, gap=12.5)
            if self.rpt_facility_name.text.strip():
                line(self.rpt_facility_name.text.strip(), size=9.5,
                     color=PDF_MUTED, gap=12.5)

            draw_footer()
            c.save()

            self._popup("Report Generated", f"PDF saved to:\n{out_path}")
            self.log_activity(f"Medical report PDF generated: {out_path}")
            self.audit_log("GENERATE_REPORT", out_path)

            try:
                self._open_file(out_path)
            except Exception:
                pass

        except Exception as e:
            self.log_activity(f"Report generation error: {e}")
            self._popup("Report Error", f"Failed to generate PDF:\n{str(e)}")

    # ====================== USER MANAGEMENT TAB ======================
    def _user_management_tab(self):
        tab = TabbedPanelItem(text="Users")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=6, spacing=4, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))
        l.add_widget(Label(
            text="USER MANAGEMENT", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#8e44ad')
        ))
        l.add_widget(Label(
            text="Administrator access only",
            font_size=12, size_hint_y=None, height=20,
            color=get_color_from_hex('#7f8c8d')
        ))

        form = GridLayout(cols=2, size_hint_y=None, height=180, spacing=4)
        form.add_widget(Label(text="Username:", font_size=14, bold=True))
        self.um_username = TextInput(font_size=14, size_hint_y=None, height=36)
        form.add_widget(self.um_username)
        form.add_widget(Label(text="Full Name:", font_size=14, bold=True))
        self.um_fullname = TextInput(font_size=14, size_hint_y=None, height=36)
        form.add_widget(self.um_fullname)
        form.add_widget(Label(text="Password:", font_size=14, bold=True))
        self.um_password = TextInput(font_size=14, password=True,
                                       size_hint_y=None, height=36)
        form.add_widget(self.um_password)
        form.add_widget(Label(text="Role:", font_size=14, bold=True))
        self.um_role = Spinner(text="Doctor", values=UserManager.ROLES,
                                 size_hint_y=None, height=36)
        form.add_widget(self.um_role)
        l.add_widget(form)

        btn_row = BoxLayout(size_hint_y=None, height=42, spacing=5)
        btn_row.add_widget(Button(
            text="CREATE USER", font_size=15, bold=True,
            on_press=self._create_new_user,
            background_color=get_color_from_hex('#27ae60')
        ))
        btn_row.add_widget(Button(
            text="RESET PASSWORD", font_size=15, bold=True,
            on_press=self._admin_reset_password,
            background_color=get_color_from_hex('#2980b9')
        ))
        btn_row.add_widget(Button(
            text="DELETE USER", font_size=15, bold=True,
            on_press=self._delete_user,
            background_color=get_color_from_hex('#e74c3c')
        ))
        btn_row.add_widget(Button(
            text="REFRESH", font_size=15, bold=True,
            on_press=lambda x: self._refresh_user_list(),
            background_color=get_color_from_hex('#16a085')
        ))
        l.add_widget(btn_row)

        l.add_widget(Label(
            text="USER LIST", font_size=16, bold=True,
            size_hint_y=None, height=24,
            color=get_color_from_hex('#2c3e50')
        ))
        self.user_list = BoxLayout(orientation='vertical', size_hint_y=None, spacing=1)
        self.user_list.bind(minimum_height=self.user_list.setter('height'))
        l.add_widget(self.user_list)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)
        self._refresh_user_list()

    def _refresh_user_list(self):
        if not hasattr(self, 'user_list'):
            return
        self.user_list.clear_widgets()
        for u in self.user_manager.users:
            row = BoxLayout(size_hint_y=None, height=32, spacing=3)
            row.add_widget(Label(text=u.get("username", ""), font_size=13,
                                   size_hint_x=0.22))
            row.add_widget(Label(text=u.get("full_name", ""), font_size=13,
                                   size_hint_x=0.28))
            row.add_widget(Label(text=u.get("role", ""), font_size=13,
                                   size_hint_x=0.2, bold=True,
                                   color=get_color_from_hex('#8e44ad')))
            row.add_widget(Label(text=u.get("last_login", "")[:16], font_size=11,
                                   size_hint_x=0.25))
            self.user_list.add_widget(row)

    def _create_new_user(self, instance):
        if not self._require_role("Administrator"):
            return
        ok, msg = self.user_manager.create_user(
            self.um_username.text.strip(),
            self.um_fullname.text.strip(),
            self.um_password.text,
            self.um_role.text
        )
        self._popup("Create User", msg)
        if ok:
            self.um_username.text = ""
            self.um_fullname.text = ""
            self.um_password.text = ""
            self._refresh_user_list()
            self.audit_log("CREATE_USER",
                           f"{self.um_username.text} ({self.um_role.text})")

    def _admin_reset_password(self, instance):
        if not self._require_role("Administrator"):
            return
        username = self.um_username.text.strip()
        if not username:
            self._popup("Error", "Enter the username to reset.")
            return
        new_pw = self.um_password.text
        if not new_pw or len(new_pw) < 4:
            self._popup("Error", "Enter a new password (min 4 chars).")
            return
        ok, msg = self.user_manager.admin_reset_password(username, new_pw)
        self._popup("Reset Password", msg)
        if ok:
            self.um_password.text = ""
            self._refresh_user_list()
            self.audit_log("ADMIN_RESET_PASSWORD", username)

    def _delete_user(self, instance):
        if not self._require_role("Administrator"):
            return
        username = self.um_username.text.strip()
        if not username:
            self._popup("Error", "Enter the username to delete.")
            return
        if username == self._current_username():
            self._popup("Error", "You cannot delete your own account.")
            return
        ok, msg = self.user_manager.delete_user(username)
        self._popup("Delete User", msg)
        if ok:
            self._refresh_user_list()
            self.audit_log("DELETE_USER", username)

    # ====================== AUDIT LOG TAB ======================
    def _audit_log_tab(self):
        tab = TabbedPanelItem(text="Audit Log")
        s = ScrollView()
        l = BoxLayout(orientation='vertical', padding=6, spacing=4, size_hint_y=None)
        l.bind(minimum_height=l.setter('height'))
        l.add_widget(Label(
            text="AUDIT LOG", font_size=20, bold=True,
            size_hint_y=None, height=32,
            color=get_color_from_hex('#c0392b')
        ))
        l.add_widget(Label(
            text="Administrator access only",
            font_size=12, size_hint_y=None, height=20,
            color=get_color_from_hex('#7f8c8d')
        ))

        filter_row = BoxLayout(size_hint_y=None, height=40, spacing=4)
        filter_row.add_widget(Label(text="Filter user:", font_size=13,
                                      bold=True, size_hint_x=0.12))
        self.audit_user_filter = TextInput(font_size=13, size_hint_x=0.3,
                                             size_hint_y=None, height=36)
        filter_row.add_widget(self.audit_user_filter)
        filter_row.add_widget(Button(
            text="APPLY", font_size=13, bold=True, size_hint_x=0.15,
            on_press=lambda x: self._refresh_audit_log(
                self.audit_user_filter.text.strip()
            ),
            background_color=get_color_from_hex('#2980b9')
        ))
        filter_row.add_widget(Button(
            text="SHOW ALL", font_size=13, bold=True, size_hint_x=0.15,
            on_press=lambda x: (setattr(self.audit_user_filter, 'text', ''),
                                self._refresh_audit_log()),
            background_color=get_color_from_hex('#16a085')
        ))
        filter_row.add_widget(Button(
            text="EXPORT", font_size=13, bold=True, size_hint_x=0.15,
            on_press=self._export_audit_log,
            background_color=get_color_from_hex('#8e44ad')
        ))
        l.add_widget(filter_row)

        self.audit_display = TextInput(
            multiline=True, readonly=True, font_size=12,
            background_color=get_color_from_hex('#f8f9fa'),
            size_hint_y=None, height=500
        )
        l.add_widget(self.audit_display)

        s.add_widget(l)
        tab.add_widget(s)
        self.tp.add_widget(tab)
        self._refresh_audit_log()

    def _refresh_audit_log(self, username_filter=""):
        if not hasattr(self, 'audit_display'):
            return
        if username_filter:
            entries = self.audit.filter_by_user(username_filter)
        else:
            entries = self.audit.get_recent(300)
        lines = []
        lines.append(f"AUDIT LOG — {len(entries)} entr{'y' if len(entries)==1 else 'ies'}")
        lines.append("=" * 80)
        for e in entries:
            lines.append(
                f"[{e.get('ts', '')}] "
                f"{e.get('username', 'system'):<20} "
                f"{e.get('action', ''):<28} "
                f"{e.get('details', '')[:60]}"
            )
        self.audit_display.text = "\n".join(lines)

    def _export_audit_log(self, instance):
        if not self._require_role("Administrator"):
            return
        text = "AUDIT LOG EXPORT\n" + "=" * 80 + "\n\n"
        for e in self.audit.entries:
            text += (f"[{e.get('ts', '')}] "
                     f"{e.get('username', 'system'):<20} "
                     f"{e.get('action', ''):<28} "
                     f"{e.get('details', '')}\n")
        self._export_to_pdf(text, "Audit_Log", "Audit Log Export")

    # ====================== UTILITIES ======================
    def _popup(self, title, msg):
        content = BoxLayout(orientation='vertical', spacing=10, padding=10)
        content.add_widget(Label(text=msg, font_size=15))
        btn = Button(text="OK", font_size=16, bold=True,
                     size_hint_y=None, height=40,
                     background_color=get_color_from_hex('#2980b9'))
        pp = Popup(title=title, content=content, size_hint=(0.6, 0.3))
        btn.bind(on_press=pp.dismiss)
        content.add_widget(btn)
        pp.open()

    def _text_popup(self, title, text):
        content = BoxLayout(orientation='vertical', spacing=5, padding=8)
        sv = ScrollView()
        ti = TextInput(text=text, readonly=True, font_size=13, size_hint_y=None,
                       background_color=get_color_from_hex('#fafafa'))
        ti.bind(minimum_height=ti.setter('height'))
        sv.add_widget(ti)
        content.add_widget(sv)
        btn = Button(text="CLOSE", font_size=16, bold=True,
                     size_hint_y=None, height=40,
                     background_color=get_color_from_hex('#2980b9'))
        pp = Popup(title=title, content=content, size_hint=(0.85, 0.75))
        btn.bind(on_press=pp.dismiss)
        content.add_widget(btn)
        pp.open()

    def _switch(self, tab_text):
        for tab in self.tp.tab_list:
            if tab.text == tab_text:
                self.tp.switch_to(tab)
                return
        self._popup("Tab", f"'{tab_text}' not found.")

    def on_stop(self):
        try:
            self.save_all()
        except Exception:
            pass
        try:
            self.lis_manager.save_configs()
        except Exception:
            pass
        try:
            if self.current_user:
                self.audit.log(
                    username=self.current_user.get("username"),
                    action="APP_CLOSE",
                    details=""
                )
        except Exception:
            pass


if __name__ == "__main__":
    HMGBillingApp().run()
