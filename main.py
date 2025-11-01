import tkinter as tk
from tkinter import messagebox, ttk
import pyttsx3
import json
import os
import cv2
import mediapipe as mp
import threading
import time
from datetime import datetime
import webbrowser
import re

# Optional libraries
try:
    import dateparser
    DATEPARSER_AVAILABLE = True
except Exception:
    DATEPARSER_AVAILABLE = False

try:
    import pywhatkit as kit
    PYWHATKIT_AVAILABLE = True
except Exception:
    PYWHATKIT_AVAILABLE = False

# pycaw (audio control) - attempt safe init, otherwise volume will be None
volume = None
try:
    from ctypes import cast, POINTER
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    try:
        devices = AudioUtilities.GetSpeakers()
        # Many pycaw setups use Activate; wrap in try/except
        try:
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
        except Exception:
            # fallback: try to get default endpoint (may not be available in all builds)
            try:
                interface = AudioUtilities.GetAudioEndpointVolume()
                volume = cast(interface, POINTER(IAudioEndpointVolume))
            except Exception as inner_e:
                print("Audio control fallback failed:", inner_e)
                volume = None
    except Exception as e:
        print("AudioUtilities init failed:", e)
        volume = None
except Exception as e:
    print("pycaw not available or failed:", e)
    volume = None

# ---------------- IDEAL TIMES FOR EFFICIENCY ----------------
IDEAL_FACE_AUTH = 2.0
IDEAL_SCHEDULE = 0.5
IDEAL_DELETE = 0.2
IDEAL_RESCHEDULE = 0.6
IDEAL_HAND_GESTURE = 0.05

face_eff = schedule_eff = delete_eff = reschedule_eff = hand_eff = 0.0

# ---------------- DATA STORAGE ----------------
filename = "calendar.json"
working_hours = ["9:00 AM","10:00 AM","11:00 AM","12:00 PM",
                 "1:00 PM","2:00 PM","3:00 PM","4:00 PM","5:00 PM"]

if os.path.exists(filename):
    try:
        with open(filename, "r") as f:
            calendar = json.load(f)
    except Exception:
        calendar = {}
else:
    calendar = {}

# ---------------- VOICE (pyttsx3) ----------------
engine = pyttsx3.init()
def speak(text):
    """Speak text asynchronously so GUI doesn't freeze."""
    def _s():
        try:
            engine.say(text)
            engine.runAndWait()
        except Exception as e:
            print("TTS error:", e)
    threading.Thread(target=_s, daemon=True).start()

# ---------------- FACE AUTHENTICATION ----------------
def face_authentication():
    global face_eff
    mp_face_detection = mp.solutions.face_detection
    mp_drawing = mp.solutions.drawing_utils

    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        messagebox.showerror("Camera Error", "Cannot access camera.")
        return False

    detected = False
    start_time = time.time()
    with mp_face_detection.FaceDetection(model_selection=0, min_detection_confidence=0.5) as face_detection:
        while True:
            ret, frame = cap.read()
            if not ret:
                break
            frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            results = face_detection.process(frame_rgb)
            if results.detections:
                detected = True
                for detection in results.detections:
                    mp_drawing.draw_detection(frame, detection)
                cv2.putText(frame, "Face Detected - Access Granted", (30, 40),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0,255,0), 2)
                cv2.imshow("Face Authentication", frame)
                cv2.waitKey(1200)
                messagebox.showinfo("Face Authentication", "Face Detected! Access Granted.")
                break
            else:
                cv2.putText(frame, "No Face Detected - Please Look at Camera", (20, 40),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0,0,255), 2)
            cv2.imshow("Face Authentication", frame)
            if cv2.waitKey(5) & 0xFF == 27:
                break

    cap.release()
    cv2.destroyAllWindows()
    elapsed = time.time() - start_time
    if elapsed <= 0:
        elapsed = 0.001
    face_eff = min((IDEAL_FACE_AUTH / elapsed) * 100, 100)
    update_efficiency_panel()
    return detected

# ---------------- HAND GESTURE VOLUME CONTROL ----------------
def hand_volume_control():
    global hand_eff, volume
    if volume is None:
        messagebox.showerror("Audio Error", "System audio control not available on this machine.")
        return
    mp_hands = mp.solutions.hands
    hands = mp_hands.Hands(max_num_hands=1, min_detection_confidence=0.7)
    cap = cv2.VideoCapture(0)
    prev_y = None
    while True:
        ret, frame = cap.read()
        if not ret:
            break
        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        results = hands.process(frame_rgb)
        if results.multi_hand_landmarks:
            hand_landmarks = results.multi_hand_landmarks[0]
            wrist_y = hand_landmarks.landmark[mp_hands.HandLandmark.WRIST].y
            if prev_y is not None:
                start_t = time.time()
                if wrist_y < prev_y - 0.02:
                    vol = min(volume.GetMasterVolumeLevelScalar() + 0.05, 1.0)
                    volume.SetMasterVolumeLevelScalar(vol, None)
                elif wrist_y > prev_y + 0.02:
                    vol = max(volume.GetMasterVolumeLevelScalar() - 0.05, 0.0)
                    volume.SetMasterVolumeLevelScalar(vol, None)
                gesture_time = time.time() - start_t
                if gesture_time <= 0:
                    gesture_time = 0.001
                hand_eff = min((IDEAL_HAND_GESTURE / gesture_time) * 100, 100)
                update_efficiency_panel()
            prev_y = wrist_y
            mp.solutions.drawing_utils.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)
        try:
            cv2.putText(frame, f"Volume: {int(volume.GetMasterVolumeLevelScalar()*100)}%", (10,30),
                        cv2.FONT_HERSHEY_SIMPLEX, 1, (0,255,0), 2)
        except Exception:
            pass
        cv2.imshow("Hand Volume Control", frame)
        if cv2.waitKey(5) & 0xFF == ord('q'):
            break
    cap.release()
    cv2.destroyAllWindows()

# ---------------- HELPERS ----------------
def save_calendar():
    with open(filename, "w") as f:
        json.dump(calendar, f, indent=4)

def update_calendar_display():
    # clear tree
    for i in calendar_tree.get_children():
        calendar_tree.delete(i)
    current_time = datetime.now().strftime("%I:00 %p").lstrip("0")
    for slot in working_hours:
        if slot in calendar:
            calendar_tree.insert("", "end", values=(slot, calendar[slot], "Busy"), tags=('busy',))
        else:
            calendar_tree.insert("", "end", values=(slot, "-", "Free"), tags=('free',))
        if slot == current_time:
            try:
                calendar_tree.item(calendar_tree.get_children()[-1], tags=('current',))
            except Exception:
                pass
    calendar_tree.tag_configure('busy', background='lightcoral')
    calendar_tree.tag_configure('free', background='lightgreen')
    calendar_tree.tag_configure('current', background='yellow')
    update_efficiency_panel()

def update_efficiency_panel():
    def color_label(lbl, val):
        if val >= 80:
            lbl.config(bg="lightgreen")
        elif val < 50:
            lbl.config(bg="lightcoral")
        else:
            lbl.config(bg="lightyellow")
    try:
        face_label.config(text=f"Face Auth: {face_eff:.1f}%")
        schedule_label.config(text=f"Schedule: {schedule_eff:.1f}%")
        delete_label.config(text=f"Delete: {delete_eff:.1f}%")
        reschedule_label.config(text=f"Reschedule: {reschedule_eff:.1f}%")
        hand_label.config(text=f"Hand Gesture: {hand_eff:.1f}%")
        color_label(face_label, face_eff)
        color_label(schedule_label, schedule_eff)
        color_label(delete_label, delete_eff)
        color_label(reschedule_label, reschedule_eff)
        color_label(hand_label, hand_eff)
    except Exception:
        pass

# ---------------- SCHEDULING & AI BRIEF ----------------
def schedule_manual():
    """Schedule using manual fields (meeting_entry + time_entry)."""
    global schedule_eff
    start_t = time.time()
    name = meeting_entry.get().strip()
    time_val = time_entry.get().strip()
    if not name or not time_val:
        messagebox.showwarning("Input Error", "Enter name and time.")
        return
    if time_val in calendar:
        messagebox.showerror("Conflict", f"{time_val} busy.")
        return
    calendar[time_val] = name
    save_calendar()
    update_calendar_display()
    speak(f"Meeting {name} scheduled at {time_val}.")
    # generate link & open
    link = f"https://meet.jit.si/{name.replace(' ','')}{int(time.time())}"
    webbrowser.open(link)
    root.clipboard_clear(); root.clipboard_append(link)
    elapsed = time.time() - start_t
    if elapsed <= 0: elapsed = 0.001
    schedule_eff = min((IDEAL_SCHEDULE / elapsed) * 100, 100)
    update_efficiency_panel()

def delete_manual():
    global delete_eff
    start_t = time.time()
    time_val = time_entry.get().strip()
    if time_val in calendar:
        name = calendar.pop(time_val)
        save_calendar()
        update_calendar_display()
        speak(f"Deleted meeting {name} at {time_val}")
    else:
        messagebox.showerror("Error", f"No meeting at {time_val}")
    elapsed = time.time() - start_t
    if elapsed <= 0: elapsed = 0.001
    delete_eff = min((IDEAL_DELETE / elapsed) * 100, 100)
    update_efficiency_panel()

def reschedule_manual():
    global reschedule_eff
    start_t = time.time()
    old = time_entry.get().strip()
    new = meeting_entry.get().strip()
    if old in calendar:
        name = calendar.pop(old)
        if new in calendar:
            messagebox.showerror("Conflict", f"{new} busy.")
            calendar[old] = name
        else:
            calendar[new] = name
            save_calendar()
            update_calendar_display()
            speak(f"Rescheduled {name} from {old} to {new}")
    else:
        messagebox.showerror("Error", f"No meeting at {old}")
    elapsed = time.time() - start_t
    if elapsed <= 0: elapsed = 0.001
    reschedule_eff = min((IDEAL_RESCHEDULE / elapsed) * 100, 100)
    update_efficiency_panel()

def ai_brief():
    """Main AI brief flow: parse brief, pick time, schedule, open link, notify participants & speak summary."""
    global schedule_eff
    start_t = time.time()
    brief = brief_textbox.get("1.0", "end").strip()
    if not brief:
        messagebox.showwarning("Input Error", "Please enter a brief.")
        return

    # detect time with dateparser if available
    time_val = None
    if DATEPARSER_AVAILABLE:
        try:
            dt = dateparser.parse(brief, settings={'PREFER_DATES_FROM': 'future'})
            if dt:
                # map to nearest hour slot in working hours if same day; if not, format time
                # we'll use only hour:00 AM/PM format
                time_val = dt.strftime("%I:00 %p").lstrip("0")
        except Exception:
            time_val = None

    # if time not found, pick next available slot
    if not time_val:
        for s in working_hours:
            if s not in calendar:
                time_val = s
                break

    if not time_val:
        messagebox.showerror("No Slot", "No free slots today.")
        return

    # extract topic/name (look for "about ..." or first sentence)
    m = re.search(r"about\s+(.+?)(?:\.|$)", brief, re.IGNORECASE)
    if m:
        meeting_name = m.group(1).strip()
    else:
        # fallback: first 4-6 words as name
        meeting_name = " ".join(brief.split()[:6])
        if not meeting_name:
            meeting_name = "General Meeting"

    # schedule
    calendar[time_val] = meeting_name
    save_calendar()
    update_calendar_display()

    # make jitsi link
    jitsi_link = f"https://meet.jit.si/{meeting_name.replace(' ','')}{int(time.time())}"
    try:
        webbrowser.open(jitsi_link)
    except Exception:
        pass
    root.clipboard_clear(); root.clipboard_append(jitsi_link)

    # parse participants
    participants_raw = participants_entry.get().strip()
    participants = [p.strip() for p in participants_raw.split(",") if p.strip()]
    emails = [p for p in participants if "@" in p]
    phones = [re.sub(r"\D", "", p) for p in participants if re.sub(r"\D", "", p).isdigit()]

    # speak summary: names and time
    names_for_speech = []
    for p in participants:
        if "@" in p:
            names_for_speech.append(p.split("@")[0])
        else:
            names_for_speech.append(p)
    names_text = ", ".join(names_for_speech) if names_for_speech else "no participants"

    summary_text = f"Scheduled {meeting_name} at {time_val}. Notifying: {names_text}."
    speak(summary_text)

    # send email(s)
    if emails:
        try:
            send_email(emails, f"Meeting: {meeting_name}", f"Meeting '{meeting_name}' at {time_val}\nJoin: {jitsi_link}")
        except Exception as e:
            messagebox.showwarning("Email", f"Email send error: {e}")

    # send whatsapp messages (if available)
    if phones and PYWHATKIT_AVAILABLE:
        for ph in phones:
            # pywhatkit send instant sometimes needs +<countrycode>; user can include country code in entry
            try:
                kit.sendwhatmsg_instantly(f"+{ph}", f"Meeting '{meeting_name}' at {time_val}. Link: {jitsi_link}")
            except Exception as e:
                print("WhatsApp send error:", e)

    try:
        messagebox.showinfo("Meeting Created", f"Meeting '{meeting_name}' scheduled at {time_val}\nLink copied to clipboard.")
    except Exception:
        print("Meeting Created:", meeting_name, time_val, jitsi_link)

    elapsed = time.time() - start_t
    if elapsed <= 0: elapsed = 0.001
    schedule_eff = min((IDEAL_SCHEDULE / elapsed) * 100, 100)
    update_efficiency_panel()

# ---------------- EMAIL / WHATSAPP helper ----------------
def send_email(to_list, subject, body):
    # NOTE: user must set sender email and app password before use
    sender_email = "youremail@gmail.com"
    sender_pass = "your_app_password"
    try:
        import smtplib
        from email.message import EmailMessage
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = sender_email
        msg["To"] = ", ".join(to_list)
        msg.set_content(body)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, sender_pass)
            server.send_message(msg)
    except Exception as e:
        raise e

# ---------------- GUI BUILD ----------------
root = tk.Tk()
root.title("MeetMate AI")
root.geometry("950x750")
root.configure(bg="#1e1e1e")

title_font = ("Helvetica", 20, "bold")
label_font = ("Helvetica", 12)
button_font = ("Helvetica", 11, "bold")

tk.Label(root, text="MeetMate AI", font=title_font, bg="#1e1e1e", fg="white").pack(pady=10)

# top frames
top_frame = tk.Frame(root, bg="#1e1e1e")
top_frame.pack(pady=8)
tk.Button(top_frame, text="Face Authentication", font=button_font, bg="#4caf50", fg="white",
          command=lambda: threading.Thread(target=face_authentication, daemon=True).start(), width=18).grid(row=0, column=0, padx=8)
tk.Button(top_frame, text="Hand Volume Control", font=button_font, bg="#2196f3", fg="white",
          command=lambda: threading.Thread(target=hand_volume_control, daemon=True).start(), width=20).grid(row=0, column=1, padx=8)

# participants
tk.Label(root, text="Participants (emails or phone numbers separated by commas)", font=label_font, bg="#1e1e1e", fg="white").pack(pady=6)
participants_entry = tk.Entry(root, width=95, font=label_font, bg="#333", fg="white", insertbackground="white")
participants_entry.pack(pady=4)

# AI brief
tk.Label(root, text="AI Brief (describe the meeting, e.g., 'Tomorrow 3 PM about project demo')", font=label_font, bg="#1e1e1e", fg="white").pack(pady=6)
brief_textbox = tk.Text(root, width=95, height=5, font=label_font, bg="#333", fg="white", insertbackground="white")
brief_textbox.pack(pady=4)
tk.Button(root, text="Submit Brief to MeetMate AI", font=button_font, bg="#4caf50", fg="white", command=lambda: threading.Thread(target=ai_brief, daemon=True).start(), width=28).pack(pady=8)

# manual schedule area
middle_frame = tk.Frame(root, bg="#1e1e1e")
middle_frame.pack(pady=6)
tk.Label(middle_frame, text="Manual: Meeting Name / Time (e.g., 9:00 AM)", font=label_font, bg="#1e1e1e", fg="white").grid(row=0, column=0, columnspan=2)
meeting_entry = tk.Entry(middle_frame, width=40, font=label_font)
meeting_entry.grid(row=1, column=0, padx=6, pady=6)
time_entry = tk.Entry(middle_frame, width=20, font=label_font)
time_entry.grid(row=1, column=1, padx=6, pady=6)

btn_frame = tk.Frame(middle_frame, bg="#1e1e1e")
btn_frame.grid(row=2, column=0, columnspan=2, pady=6)
tk.Button(btn_frame, text="Schedule", bg="#4caf50", fg="white", width=12, command=schedule_manual).grid(row=0, column=0, padx=6)
tk.Button(btn_frame, text="Delete", bg="#f44336", fg="white", width=12, command=delete_manual).grid(row=0, column=1, padx=6)
tk.Button(btn_frame, text="Reschedule", bg="#ff9800", fg="white", width=12, command=reschedule_manual).grid(row=0, column=2, padx=6)

# calendar view
tree_frame = tk.Frame(root)
tree_frame.pack(pady=10)
calendar_tree = ttk.Treeview(tree_frame, columns=("Time","Meeting","Status"), show="headings", height=10)
calendar_tree.heading("Time", text="Time")
calendar_tree.heading("Meeting", text="Meeting")
calendar_tree.heading("Status", text="Status")
calendar_tree.pack(side="left")
tree_scroll = tk.Scrollbar(tree_frame)
tree_scroll.pack(side="right", fill="y")
calendar_tree.config(yscrollcommand=tree_scroll.set)
tree_scroll.config(command=calendar_tree.yview)
calendar_tree.bind("<Double-1>", lambda e: "break")

# efficiency panel
eff_frame = tk.Frame(root, bg="#1e1e1e")
eff_frame.pack(pady=12)
face_label = tk.Label(eff_frame, text="Face Auth: 0%", font=label_font, width=18)
face_label.grid(row=0, column=0, padx=4)
schedule_label = tk.Label(eff_frame, text="Schedule: 0%", font=label_font, width=18)
schedule_label.grid(row=0, column=1, padx=4)
delete_label = tk.Label(eff_frame, text="Delete: 0%", font=label_font, width=18)
delete_label.grid(row=0, column=2, padx=4)
reschedule_label = tk.Label(eff_frame, text="Reschedule: 0%", font=label_font, width=18)
reschedule_label.grid(row=0, column=3, padx=4)
hand_label = tk.Label(eff_frame, text="Hand Gesture: 0%", font=label_font, width=18)
hand_label.grid(row=0, column=4, padx=4)

# ---------------- STARTUP LOGIN & VOICE ----------------
def show_login_and_start():
    """Show login popup; on success, welcome voice and show main window."""
    login = tk.Toplevel(root)
    login.title("Login - MeetMate AI")
    login.geometry("360x160")
    login.configure(bg="#1e1e1e")
    login.resizable(False, False)
    tk.Label(login, text="Enter password to open MeetMate AI", bg="#1e1e1e", fg="white").pack(pady=(14,8))
    pwd_var = tk.StringVar()
    pwd_entry = tk.Entry(login, textvariable=pwd_var, show="*", width=28, bg="#333", fg="white", insertbackground="white")
    pwd_entry.pack(pady=(0,10))
    pwd_entry.focus_set()
    attempts = {"count": 0}
    def try_login():
        if pwd_var.get() == "shruti0707":
            login.destroy()
            speak("Hello Shruti. MeetMate AI is ready. Please brief me about the meeting you want to schedule.")
            try:
                messagebox.showinfo("Welcome", "Welcome to MeetMate AI — please enter your meeting brief.")
            except Exception:
                pass
        else:
            attempts["count"] += 1
            left = 3 - attempts["count"]
            if left > 0:
                messagebox.showerror("Login Failed", f"Incorrect password. {left} attempts left.")
                pwd_entry.delete(0, tk.END)
                pwd_entry.focus_set()
            else:
                messagebox.showerror("Locked", "Maximum attempts reached. Exiting.")
                login.destroy()
                root.destroy()
    btns = tk.Frame(login, bg="#1e1e1e")
    btns.pack()
    tk.Button(btns, text="Login", bg="#3f51b5", fg="white", width=12, command=try_login).grid(row=0, column=0, padx=6)
    tk.Button(btns, text="Exit", bg="#b33a3a", fg="white", width=12, command=lambda: (login.destroy(), root.destroy())).grid(row=0, column=1, padx=6)
    login.transient(root)
    login.grab_set()
    root.wait_window(login)

# ---------------- REMINDER LOOP ----------------
def meeting_reminder_loop():
    while True:
        now = datetime.now().strftime("%I:00 %p").lstrip("0")
        if now in calendar:
            meeting = calendar[now]
            speak(f"Reminder: Meeting {meeting} at {now}")
            try:
                messagebox.showinfo("Meeting Reminder", f"Meeting '{meeting}' at {now}")
            except Exception:
                print("Reminder:", meeting, now)
            time.sleep(60)
        time.sleep(10)

# run login then start reminder
root.withdraw()  # hide main until login
root.after(100, show_login_and_start)  # show login shortly after start
root.deiconify()  # show main (login will block until done)
update_calendar_display()
threading.Thread(target=meeting_reminder_loop, daemon=True).start()

root.mainloop()
