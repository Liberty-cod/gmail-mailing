import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog
import smtplib, csv, time, json, os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ---------------- ЛОГ ----------------
def log_message(msg):
    log_box.insert(tk.END, msg + "\n")
    log_box.see(tk.END)

# ---------------- ШАБЛОНЫ ----------------
TEMPLATES_FILE = "templates.json"

default_templates = {
    "Простой": "<h2>Здравствуйте, {{name}}!</h2><p>Будем рады видеть вас на мероприятии.</p>",
    "С картинкой": "<h2>Привет, {{name}}!</h2><p>Скоро состоится событие, приглашаем!</p><img src='https://i.ibb.co/Z6JBxRHF/aa.png' width='300'>",
    "Официальный": "<p>Уважаемый(ая) {{name}},</p><p>Приглашаем Вас принять участие в нашем мероприятии.</p>"
}

def load_templates():
    if not os.path.exists(TEMPLATES_FILE):
        save_templates(default_templates)
        return default_templates
    try:
        with open(TEMPLATES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        save_templates(default_templates)
        return default_templates

def save_templates(data):
    with open(TEMPLATES_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# ---------------- ПОЧТА ----------------
def send_email(to_email, subject, body, gmail_user, gmail_pass):
    try:
        msg = MIMEMultipart()
        msg["From"] = gmail_user
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "html", "utf-8"))

        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(gmail_user, gmail_pass)
        server.sendmail(gmail_user, to_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        log_message(f"❌ Ошибка {to_email}: {e}")
        return False

# ---------------- ОТПРАВКА ----------------
def send_test_email():
    gmail_user = gmail_entry.get().strip()
    gmail_pass = pass_entry.get().strip()
    if not gmail_user or not gmail_pass:
        messagebox.showerror("Ошибка", "Введите Gmail и пароль приложения")
        return
    body = text_body.get("1.0", tk.END)
    log_message("⏳ Отправка теста...")
    if send_email(gmail_user, "Тестовое письмо", body, gmail_user, gmail_pass):
        log_message("✅ Тестовое письмо отправлено")

def send_bulk():
    gmail_user = gmail_entry.get().strip()
    gmail_pass = pass_entry.get().strip()
    delay = int(delay_entry.get())
    if not gmail_user or not gmail_pass:
        messagebox.showerror("Ошибка", "Введите Gmail и пароль приложения")
        return
    if not guests:
        messagebox.showerror("Ошибка", "Сначала загрузите CSV")
        return

    success, fail = 0, 0
    for guest in guests:
        name, email = guest
        body = text_body.get("1.0", tk.END).replace("{{name}}", name)
        log_message(f"⏳ Отправка: {name} <{email}>")
        if send_email(email, "Приглашение", body, gmail_user, gmail_pass):
            log_message(f"✅ Успешно: {email}")
            success += 1
        else:
            fail += 1
        time.sleep(delay)

    log_message(f"📨 Готово! Успешно: {success}, Ошибки: {fail}")

# ---------------- CSV ----------------
def load_csv():
    global guests
    filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if filename:
        with open(filename, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            guests = [(row[0], row[1]) for row in reader if len(row) >= 2]
        log_message(f"✅ CSV загружен: {len(guests)} гостей")

# ---------------- ШАБЛОНЫ В GUI ----------------
def apply_template(*args):
    text_body.delete("1.0", tk.END)
    text_body.insert(tk.END, templates[template_var.get()])

def save_new_template():
    name = simpledialog.askstring("Новый шаблон", "Введите название шаблона:")
    if not name:
        return
    templates[name] = text_body.get("1.0", tk.END).strip()
    save_templates(templates)
    menu = template_menu["menu"]
    menu.add_command(label=name, command=tk._setit(template_var, name))
    log_message(f"💾 Шаблон '{name}' сохранён")

# ---------------- GUI ----------------
root = tk.Tk()
root.title("📨 Рассылка приглашений")
root.geometry("900x800")

guests = []
templates = load_templates()

tk.Label(root, text="Ваш Gmail:").grid(row=0, column=0, sticky="w")
gmail_entry = tk.Entry(root, width=40)
gmail_entry.grid(row=0, column=1)

tk.Label(root, text="Пароль приложения:").grid(row=1, column=0, sticky="w")
pass_entry = tk.Entry(root, width=40, show="*")
pass_entry.grid(row=1, column=1)

tk.Button(root, text="Загрузить CSV", command=load_csv).grid(row=2, column=0, pady=5)

tk.Label(root, text="Задержка (сек) между письмами:").grid(row=3, column=0, sticky="w")
delay_entry = tk.Entry(root, width=5)
delay_entry.insert(0, "3")
delay_entry.grid(row=3, column=1, sticky="w")

tk.Label(root, text="Шаблон:").grid(row=4, column=0, sticky="w")
template_var = tk.StringVar(value=list(templates.keys())[0])
template_menu = tk.OptionMenu(root, template_var, *templates.keys())
template_menu.grid(row=4, column=1, sticky="w")

tk.Button(root, text="Сохранить новый шаблон", command=save_new_template).grid(row=4, column=2, padx=5)

template_var.trace("w", apply_template)

tk.Label(root, text="Текст письма (HTML):").grid(row=5, column=0, sticky="w")
text_body = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=80, height=15)
text_body.grid(row=6, column=0, columnspan=3, pady=5)

tk.Label(root, text="Логи:").grid(row=7, column=0, sticky="w")
log_box = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=80, height=10, state="normal")
log_box.grid(row=8, column=0, columnspan=3, pady=5)

# Кнопки внизу
button_frame = tk.Frame(root)
button_frame.grid(row=9, column=0, columnspan=3, pady=15)

tk.Button(button_frame, text="Предварительный просмотр",
          command=lambda: log_message("📋 " + text_body.get("1.0", tk.END)[:200]),
          bg="blue", fg="white", width=25).pack(side="left", padx=5)

tk.Button(button_frame, text="Отправить тестовое письмо",
          command=send_test_email, bg="orange", fg="white", width=25).pack(side="left", padx=5)

tk.Button(button_frame, text="Отправить всем",
          command=send_bulk, bg="green", fg="white", width=25).pack(side="left", padx=5)

root.mainloop()
