import logging
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ctypes
from ctypes import wintypes
import threading
import csv
import time
import concurrent.futures
import json
import os
from datetime import datetime

# PID do próprio processo — usado para auto-exclusão nas buscas
import os as _os
OWN_PID = _os.getpid()

# 1. Configura o log (só isso, sem _setup_log)
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("hotkey.log", encoding="utf-8")
    ]
)

# 2. Importa as funções do módulo de teste
from test_process_result import test_process_suspension, TestResult, _check_hotkey_available

# 3. Importa o módulo de sondagem por hwnd
from hwnd_probe import probe_hotkey_owner, probe_by_kill_reopen, ProbeResult, KILL_BLACKLIST

# ─── DEPENDÊNCIA OPCIONAL ───────────────────────────────────────────────────────
try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False
    import subprocess, sys
    # Tenta instalar automaticamente se não encontrar
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "psutil"],
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        import psutil
        HAS_PSUTIL = True
    except Exception:
        pass

# ─── WIN API ───────────────────────────────────────────────────────────────────
user32   = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

PROCESS_QUERY_INFORMATION = 0x0400
PROCESS_VM_READ           = 0x0010

MOD_NOREPEAT = 0x4000
MODIFIERS = {
    "None":           0x0000,
    "Alt":            0x0001,
    "Ctrl":           0x0002,
    "Shift":          0x0004,
    "Win":            0x0008,
    "Ctrl+Alt":       0x0003,
    "Ctrl+Shift":     0x0006,
    "Alt+Shift":      0x0005,
    "Ctrl+Alt+Shift": 0x0007,
}

VK_CODES = {
    "A":0x41,"B":0x42,"C":0x43,"D":0x44,"E":0x45,"F":0x46,"G":0x47,"H":0x48,
    "I":0x49,"J":0x4A,"K":0x4B,"L":0x4C,"M":0x4D,"N":0x4E,"O":0x4F,"P":0x50,
    "Q":0x51,"R":0x52,"S":0x53,"T":0x54,"U":0x55,"V":0x56,"W":0x57,"X":0x58,
    "Y":0x59,"Z":0x5A,
    "0":0x30,"1":0x31,"2":0x32,"3":0x33,"4":0x34,
    "5":0x35,"6":0x36,"7":0x37,"8":0x38,"9":0x39,
    "F1":0x70,"F2":0x71,"F3":0x72,"F4":0x73,"F5":0x74,"F6":0x75,
    "F7":0x76,"F8":0x77,"F9":0x78,"F10":0x79,"F11":0x7A,"F12":0x7B,
    "PrtScn":0x2C,"Ins":0x2D,"Del":0x2E,"Home":0x24,"End":0x23,
    "PgUp":0x21,"PgDn":0x22,"Esc":0x1B,"Tab":0x09,"Space":0x20,
    "Enter":0x0D,"Backspace":0x08,"CapsLock":0x14,
    "Left":0x25,"Up":0x26,"Right":0x27,"Down":0x28,
    "Num0":0x60,"Num1":0x61,"Num2":0x62,"Num3":0x63,"Num4":0x64,
    "Num5":0x65,"Num6":0x66,"Num7":0x67,"Num8":0x68,"Num9":0x69,
    "Num*":0x6A,"Num+":0x6B,"Num-":0x6D,"Num.":0x6E,"Num/":0x6F,
    ";":0xBA,"=":0xBB,",":0xBC,"-":0xBD,".":0xBE,"/":0xBF,"`":0xC0,
}

# ─── BANCO DE ATALHOS CONHECIDOS ───────────────────────────────────────────────
# Formato: (mod_value, vk_value): [(app_name, exe_name, descricao)]
KNOWN_HOTKEYS = {
    # Windows nativos
    (0x0008, 0x44): [("Windows","explorer.exe","Win+D - Mostrar Desktop")],
    (0x0008, 0x45): [("Windows","explorer.exe","Win+E - Abre Explorador")],
    (0x0008, 0x52): [("Windows","explorer.exe","Win+R - Executar")],
    (0x0008, 0x4C): [("Windows","explorer.exe","Win+L - Bloquear PC")],
    (0x0008, 0x53): [("Windows","SearchHost.exe","Win+S - Pesquisar")],
    (0x0008, 0x49): [("Windows","explorer.exe","Win+I - Configurações")],
    (0x0008, 0x58): [("Windows","explorer.exe","Win+X - Menu rápido")],
    (0x0008, 0x26): [("Windows","explorer.exe","Win+Up - Maximizar")],
    (0x0008, 0x28): [("Windows","explorer.exe","Win+Down - Restaurar")],
    (0x0008, 0x70): [("Windows","explorer.exe","Win+F1 - Ajuda")],
    (0x0008, 0x2C): [("Windows / Snipping Tool","SnippingTool.exe","Win+PrtScn - Captura de tela")],
    (0x0008, 0x4E): [("Windows","explorer.exe","Win+N - Central de Notificações")],
    (0x0008, 0x41): [("Windows","explorer.exe","Win+A - Ações rápidas")],
    (0x0008, 0x57): [("Windows","explorer.exe","Win+W - Widgets")],
    (0x0008, 0x5A): [("Windows","explorer.exe","Win+Z - Snap layout")],
    (0x0008, 0x4B): [("Windows","explorer.exe","Win+K - Conectar")],
    (0x0008, 0x4D): [("Windows","explorer.exe","Win+M - Minimizar tudo")],
    (0x0008, 0x55): [("Windows","explorer.exe","Win+U - Acessibilidade")],
    (0x0000, 0x2C): [("Snipping Tool / OBS","SnippingTool.exe","PrtScn - Captura de tela")],
    # Snipping Tool
    (0x0008|0x0002, 0x53): [("Snipping Tool","SnippingTool.exe","Win+Ctrl+S - Recorte de tela")],
    # OneDrive
    (0x0000, 0x7B): [("OneDrive","OneDrive.exe","F12 - Frequentemente capturado pelo OneDrive")],
    # Teams / Skype
    (0x0002, 0x4D): [("Teams / Skype","Teams.exe","Ctrl+M - Mudo no Teams")],
    # Spotify
    (0x0002, 0xBB): [("Spotify","Spotify.exe","Ctrl+= - Volume no Spotify")],
    (0x0002, 0xBD): [("Spotify","Spotify.exe","Ctrl+- - Volume no Spotify")],
    # OBS Studio
    (0x0002, 0x31): [("OBS Studio","obs64.exe","Ctrl+1 - Cena 1 (padrão OBS)")],
    # Discord
    (0x0002, 0xBF): [("Discord","Discord.exe","Ctrl+/ - Atalhos do Discord")],
    # Lightshot / Screenshot tools
    (0x0000, 0x2C): [("Lightshot / Print Screen","Lightshot.exe","PrtScn - Captura")],
    # AutoHotkey
    (0x0001, 0x52): [("AutoHotkey","AutoHotkey.exe","Alt+R - Script AHK comum")],
    # RDP / Remote Desktop
    (0x0002, 0x0D): [("RDP","mstsc.exe","Ctrl+Enter - Tela cheia RDP")],
    # 1Password
    (0x0002|0x0001, 0x5C): [("1Password","1Password.exe","Ctrl+Alt+\\ - Abrir 1Password")],
    # Greenshot
    (0x0000, 0x2C): [("Greenshot","Greenshot.exe","PrtScn - Captura Greenshot")],
    # ShareX
    (0x0002, 0x2C): [("ShareX","ShareX.exe","Ctrl+PrtScn - Captura ShareX")],
}

def check_hotkey(mod_val, vk_val):
    uid = 8877
    r = user32.RegisterHotKey(None, uid, mod_val | MOD_NOREPEAT, vk_val)
    if r:
        user32.UnregisterHotKey(None, uid)
        return True
    return False

def get_windows_and_processes():
    """Enumera todas as janelas visíveis com seus processos."""
    results = []
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)

    def enum_callback(hwnd, lparam):
        if user32.IsWindowVisible(hwnd):
            buf = ctypes.create_unicode_buffer(512)
            user32.GetWindowTextW(hwnd, buf, 512)
            title = buf.value.strip()
            if not title:
                return True
            pid = wintypes.DWORD(0)
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            pid_val = pid.value
            proc_name = "?"
            proc_path = "?"
            if HAS_PSUTIL:
                try:
                    p = psutil.Process(pid_val)
                    proc_name = p.name()
                    proc_path = p.exe()
                except:
                    pass
            else:
                hproc = kernel32.OpenProcess(
                    PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, False, pid_val)
                if hproc:
                    buf2 = ctypes.create_unicode_buffer(1024)
                    ctypes.windll.psapi.GetModuleFileNameExW(hproc, None, buf2, 1024)
                    proc_path = buf2.value
                    proc_name = os.path.basename(proc_path) if proc_path else "?"
                    kernel32.CloseHandle(hproc)
            results.append({
                "hwnd": hwnd, "title": title,
                "pid": pid_val, "name": proc_name, "path": proc_path
            })
        return True

    user32.EnumWindows(WNDENUMPROC(enum_callback), 0)
    return results

def get_known_app_info(mod_val, vk_val):
    """Retorna lista de apps conhecidos para esse atalho."""
    key = (mod_val & ~MOD_NOREPEAT, vk_val)
    return KNOWN_HOTKEYS.get(key, [])

def _test_one_process(pid, mod_val, vk_val):
    """
    Retorna:
      True           → hotkey liberou (este processo é o culpado)
      False          → hotkey continua ocupada (não é este)
      "ACCESS_DENIED"→ não conseguiu suspender (suspeito! ex: Windows Store app)
      "NO_PROCESS"   → processo não existe mais
    """
    p = None
    try:
        p = psutil.Process(pid)
        p.suspend()
        time.sleep(0.2)
        avail = check_hotkey(mod_val, vk_val)
        return avail  # True = culpado, False = inocente
    except psutil.AccessDenied:
        return "ACCESS_DENIED"   # app protegido — suspeito!
    except psutil.NoSuchProcess:
        return "NO_PROCESS"
    except Exception:
        return "ACCESS_DENIED"   # qualquer outro erro = suspeito
    finally:
        if p is not None:
            try: p.resume()
            except: pass

def find_by_elimination(mod_val, vk_val, proc_list, progress_cb=None, stop_event=None):
    """
    Retorna dict com:
      "confirmed": proc_info | None   → processo confirmado como culpado
      "suspects":  [proc_info, ...]   → processos que não puderam ser suspensos (Access Denied)
      "stopped":   bool               → se foi interrompido pelo usuário
    """
    if not HAS_PSUTIL:
        return {"confirmed": None, "suspects": [], "stopped": False}

    TIMEOUT_PER_PROC = 3.0
    suspects = []
    stopped  = False

    for i, proc_info in enumerate(proc_list):
        if stop_event and stop_event.is_set():
            stopped = True
            break
        if progress_cb:
            progress_cb(i, len(proc_list), proc_info.get("name", "?"))

        pid = proc_info["pid"]
        try:
            with concurrent.futures.ThreadPoolExecutor(max_workers=1) as ex:
                future = ex.submit(_test_one_process, pid, mod_val, vk_val)
                try:
                    result = future.result(timeout=TIMEOUT_PER_PROC)
                    if result is True:
                        return {"confirmed": proc_info, "suspects": suspects, "stopped": False}
                    elif result == "ACCESS_DENIED":
                        suspects.append(proc_info)   # guarda como suspeito
                except concurrent.futures.TimeoutError:
                    suspects.append(proc_info)       # timeout = suspeito também
                    try: psutil.Process(pid).resume()
                    except: pass
        except Exception:
            pass

    return {"confirmed": None, "suspects": suspects, "stopped": stopped}

# ─── APP ───────────────────────────────────────────────────────────────────────
class HotkeyTrackerApp:
    COLORS = {
        "bg":       "#1e1e2e","surface":  "#2a2a3e","surface2": "#313145",
        "accent":   "#7c6af7","accent2":  "#5b4fd4","green":    "#50fa7b",
        "red":      "#ff5555","yellow":   "#f1fa8c","orange":   "#ffb86c",
        "text":     "#cdd6f4","subtext":  "#a6adc8","border":   "#45475a",
    }

    def __init__(self, root):
        self.root = root
        self.root.title("🔑 HotkeyTracker Pro")
        self.root.geometry("1100x720")
        self.root.configure(bg=self.COLORS["bg"])
        self.root.resizable(True, True)
        self.scan_results   = []
        self.scanning       = False
        self.all_windows    = []
        self._build_styles()
        self._build_ui()
        if not HAS_PSUTIL:
            self.root.after(800, lambda: messagebox.showwarning(
                "psutil não encontrado",
                "psutil não foi encontrado neste ambiente Python.\n\n"
                "Instale com:\n    pip install psutil\n\n"
                "Depois feche e abra o app novamente.\n"
                "Algumas funções ficarão limitadas sem ele."))

    def _build_styles(self):
        style = ttk.Style(); style.theme_use("clam"); C = self.COLORS
        style.configure("TFrame",          background=C["bg"])
        style.configure("TLabel",          background=C["bg"], foreground=C["text"], font=("Segoe UI",10))
        style.configure("Card.TLabel",     background=C["surface"], foreground=C["text"], font=("Segoe UI",10))
        style.configure("Accent.TButton",  background=C["accent"],  foreground="white", font=("Segoe UI",10,"bold"), relief="flat", padding=(14,6))
        style.map("Accent.TButton",        background=[("active",C["accent2"])])
        style.configure("Danger.TButton",  background=C["red"],     foreground="white", font=("Segoe UI",10,"bold"), relief="flat", padding=(14,6))
        style.configure("Success.TButton", background=C["green"],   foreground="#1e1e2e",font=("Segoe UI",10,"bold"), relief="flat", padding=(14,6))
        style.configure("Warn.TButton",    background=C["orange"],  foreground="#1e1e2e",font=("Segoe UI",10,"bold"), relief="flat", padding=(14,6))
        style.configure("TCombobox",       fieldbackground=C["surface2"], background=C["surface2"],
                            foreground=C["text"], selectforeground=C["text"], selectbackground=C["accent"])
        style.map("TCombobox",             fieldbackground=[("readonly", C["surface2"])],
                                           foreground=[("readonly", C["text"]), ("disabled", C["subtext"])],
                                           selectforeground=[("readonly", C["text"])],
                                           selectbackground=[("readonly", C["accent"])])
        style.configure("Treeview",        background=C["surface"], fieldbackground=C["surface"], foreground=C["text"], rowheight=26, font=("Segoe UI",10))
        style.configure("Treeview.Heading",background=C["surface2"],foreground=C["accent"], font=("Segoe UI",10,"bold"))
        style.map("Treeview",              background=[("selected",C["accent2"])], foreground=[("selected","white")])
        style.configure("TProgressbar",    troughcolor=C["surface2"], background=C["accent"], thickness=8)
        style.configure("TEntry",          fieldbackground=C["surface2"], foreground=C["text"], insertcolor=C["text"])
        style.configure("TNotebook",       background=C["bg"])
        style.configure("TNotebook.Tab",   background=C["surface"], foreground=C["subtext"], font=("Segoe UI",10), padding=[16,6])
        style.map("TNotebook.Tab",         background=[("selected",C["accent"])], foreground=[("selected","white")])

    def _build_ui(self):
        C = self.COLORS
        hdr = tk.Frame(self.root, bg=C["surface"], height=58)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="🔑 HotkeyTracker Pro", bg=C["surface"], fg=C["text"],
                 font=("Segoe UI",16,"bold")).pack(side="left", padx=20, pady=10)
        tk.Label(hdr, text="Detecta atalhos globais e descobre qual app está usando",
                 bg=C["surface"], fg=C["subtext"], font=("Segoe UI",9)).pack(side="left", padx=4)
        if not HAS_PSUTIL:
            tk.Label(hdr, text="⚠ pip install psutil  →  reinicie o app",
                     bg="#ff5555", fg="white", font=("Segoe UI",9,"bold")).pack(side="right", padx=16)
        else:
            tk.Label(hdr, text="✅ psutil ativo",
                     bg="#313145", fg="#50fa7b", font=("Segoe UI",9,"bold")).pack(side="right", padx=16)

        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=10, pady=6)

        t1=ttk.Frame(nb); nb.add(t1, text="  🔍 Verificador  ")
        t2=ttk.Frame(nb); nb.add(t2, text="  📊 Scanner  ")
        t3=ttk.Frame(nb); nb.add(t3, text="  🖥 Processos / Apps  ")
        t4=ttk.Frame(nb); nb.add(t4, text="  📋 Resultados  ")
        t5=ttk.Frame(nb); nb.add(t5, text="  📚 Base de Atalhos  ")

        self._build_checker(t1)
        self._build_scanner(t2)
        self._build_processes(t3)
        self._build_results(t4)
        self._build_known(t5)

    # ──────────────────────────────────────────────────────────────────────────
    # ABA 1 – VERIFICADOR RÁPIDO
    # ──────────────────────────────────────────────────────────────────────────
    def _build_checker(self, parent):
        C = self.COLORS
        main = tk.Frame(parent, bg=C["bg"]); main.pack(fill="both", expand=True, padx=20, pady=14)

        tk.Label(main, text="Verificar Tecla Específica", bg=C["bg"], fg=C["text"],
                 font=("Segoe UI",13,"bold")).pack(anchor="w")
        tk.Label(main, text="Descobre se está ocupada E qual app provável está usando",
                 bg=C["bg"], fg=C["subtext"], font=("Segoe UI",9)).pack(anchor="w", pady=(2,14))

        row = tk.Frame(main, bg=C["bg"]); row.pack(fill="x", pady=6)
        tk.Label(row, text="Modificador:", bg=C["bg"], fg=C["subtext"]).pack(side="left")
        self.mod_var = tk.StringVar(value="Alt")
        ttk.Combobox(row, textvariable=self.mod_var, values=list(MODIFIERS.keys()),
                     width=14, state="readonly").pack(side="left", padx=(6,20))
        tk.Label(row, text="Tecla:", bg=C["bg"], fg=C["subtext"]).pack(side="left")
        self.key_var = tk.StringVar(value="A")
        ttk.Combobox(row, textvariable=self.key_var, values=sorted(VK_CODES.keys()),
                     width=12, state="readonly").pack(side="left", padx=(6,20))
        ttk.Button(row, text="✔ Verificar", style="Accent.TButton",
                   command=self._quick_check).pack(side="left", padx=4)

        # Resultado principal
        res = tk.Frame(main, bg=C["surface"], bd=0, highlightthickness=1,
                       highlightbackground=C["border"])
        res.pack(fill="x", pady=12, ipady=10)
        self.res_icon  = tk.Label(res, text="—", bg=C["surface"], fg=C["subtext"], font=("Segoe UI",30))
        self.res_icon.pack(pady=(8,2))
        self.res_label = tk.Label(res, text="Selecione um atalho e clique em Verificar",
                                   bg=C["surface"], fg=C["subtext"], font=("Segoe UI",12))
        self.res_label.pack()

        # Bloco de "quem usa"
        who_frame = tk.Frame(main, bg=C["surface2"], bd=0, highlightthickness=1,
                              highlightbackground=C["border"])
        who_frame.pack(fill="x", pady=(0,10), ipady=8)
        tk.Label(who_frame, text="🔎 Provável responsável:", bg=C["surface2"],
                 fg=C["yellow"], font=("Segoe UI",10,"bold")).pack(anchor="w", padx=12, pady=(6,2))
        self.who_label = tk.Label(who_frame, text="—", bg=C["surface2"],
                                   fg=C["text"], font=("Segoe UI",10), wraplength=700, justify="left")
        self.who_label.pack(anchor="w", padx=24, pady=(0,6))

        # Dica método de descoberta
        tip = tk.Frame(main, bg=C["surface2"], bd=0, highlightthickness=1,
                       highlightbackground=C["border"])
        tip.pack(fill="x", pady=(0,10), ipady=6)
        tk.Label(tip, text="💡 Como identificar o responsável com certeza:",
                 bg=C["surface2"], fg=C["accent"], font=("Segoe UI",10,"bold")).pack(anchor="w", padx=12, pady=(6,2))
        self.elim_info = tk.Label(tip,
            text="1. Use a aba 'Processos / Apps' para ver todos os apps rodando\n"
                 "2. Use o botão '⚗ Teste por Eliminação' para suspender processos um a um e ver qual libera o atalho\n"
                 "3. Verifique a aba 'Base de Atalhos' para atalhos conhecidos de apps populares",
            bg=C["surface2"], fg=C["subtext"], font=("Segoe UI",9), justify="left")
        self.elim_info.pack(anchor="w", padx=24, pady=(0,4))

        # Histórico
        tk.Label(main, text="Histórico", bg=C["bg"], fg=C["text"],
                 font=("Segoe UI",11,"bold")).pack(anchor="w", pady=(8,4))
        hf = tk.Frame(main, bg=C["surface"]); hf.pack(fill="both", expand=True)
        cols = ("Atalho","Status","Provável App","Horário")
        self.hist_tree = ttk.Treeview(hf, columns=cols, show="headings", height=7)
        for c,w in zip(cols,[240,120,280,120]):
            self.hist_tree.heading(c,text=c); self.hist_tree.column(c,width=w,anchor="center")
        self.hist_tree.tag_configure("livre",   foreground=C["green"])
        self.hist_tree.tag_configure("ocupada", foreground=C["red"])
        sb=ttk.Scrollbar(hf,orient="vertical",command=self.hist_tree.yview)
        self.hist_tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right",fill="y"); self.hist_tree.pack(fill="both",expand=True)

    def _quick_check(self):
        C = self.COLORS
        mod_name = self.mod_var.get(); key_name = self.key_var.get()
        mod_val  = MODIFIERS.get(mod_name, 0); vk_val = VK_CODES.get(key_name, 0)
        label    = f"{mod_name} + {key_name}" if mod_name != "None" else key_name
        avail    = check_hotkey(mod_val, vk_val)
        hora     = datetime.now().strftime("%H:%M:%S")

        if avail:
            self.res_icon.config(text="✅", fg=C["green"])
            self.res_label.config(text=f"  {label}  está LIVRE!", fg=C["green"])
            self.who_label.config(text="Nenhum app está usando este atalho.", fg=C["green"])
            tag = "livre"; status = "✅ Livre"; who_str = "—"
        else:
            self.res_icon.config(text="🚫", fg=C["red"])
            self.res_label.config(text=f"  {label}  está OCUPADA por outro app", fg=C["red"])
            tag = "ocupada"; status = "🚫 Ocupada"
            known = get_known_app_info(mod_val, vk_val)
            if known:
                lines = []
                for app, exe, desc in known:
                    lines.append(f"📌 {app}  ({exe})  —  {desc}")
                who_str = known[0][0]
                self.who_label.config(text="\n".join(lines), fg=C["yellow"])
            else:
                who_str = "App desconhecido"
                self.who_label.config(
                    text="Não identificado na base de dados.\n"
                         "➜ Use a aba 'Processos / Apps' e o Teste por Eliminação para descobrir.",
                    fg=C["orange"])
        self.hist_tree.insert("", 0, values=(label, status, who_str, hora), tags=(tag,))

    # ──────────────────────────────────────────────────────────────────────────
    # ABA 2 – SCANNER
    # ──────────────────────────────────────────────────────────────────────────
    def _build_scanner(self, parent):
        C = self.COLORS
        main = tk.Frame(parent, bg=C["bg"]); main.pack(fill="both", expand=True, padx=20, pady=14)
        tk.Label(main, text="Scanner em Massa", bg=C["bg"], fg=C["text"],
                 font=("Segoe UI",13,"bold")).pack(anchor="w")
        tk.Label(main, text="Varre combinações e identifica quais estão ocupadas (com provável dono)",
                 bg=C["bg"], fg=C["subtext"], font=("Segoe UI",9)).pack(anchor="w", pady=(2,14))

        cfg = tk.Frame(main, bg=C["bg"]); cfg.pack(fill="x", pady=4)

        card1 = tk.LabelFrame(cfg, text=" Modificadores ", bg=C["surface"], fg=C["accent"],
                               font=("Segoe UI",9,"bold"), bd=1, relief="solid")
        card1.pack(side="left", padx=(0,12), pady=4, fill="y")
        self.mod_checks = {}
        for m in ["None","Alt","Ctrl","Shift","Win","Ctrl+Alt","Ctrl+Shift","Alt+Shift"]:
            v = tk.BooleanVar(value=(m in ["Alt","Ctrl","Ctrl+Alt"]))
            tk.Checkbutton(card1, text=m, variable=v, bg=C["surface"], fg=C["text"],
                           selectcolor=C["surface2"], activebackground=C["surface"],
                           font=("Segoe UI",9)).pack(anchor="w", padx=10, pady=1)
            self.mod_checks[m] = v

        card2 = tk.LabelFrame(cfg, text=" Grupos de Teclas ", bg=C["surface"], fg=C["accent"],
                               font=("Segoe UI",9,"bold"), bd=1, relief="solid")
        card2.pack(side="left", padx=(0,12), pady=4, fill="y")
        self.key_groups = {
            "Letras A-Z":  tk.BooleanVar(value=True),
            "Números 0-9": tk.BooleanVar(value=True),
            "F1-F12":      tk.BooleanVar(value=True),
            "Especiais":   tk.BooleanVar(value=True),
            "Numpad":      tk.BooleanVar(value=False),
        }
        for name, var in self.key_groups.items():
            tk.Checkbutton(card2, text=name, variable=var, bg=C["surface"], fg=C["text"],
                           selectcolor=C["surface2"], activebackground=C["surface"],
                           font=("Segoe UI",9)).pack(anchor="w", padx=10, pady=1)

        card3 = tk.LabelFrame(cfg, text=" Estatísticas ", bg=C["surface"], fg=C["accent"],
                               font=("Segoe UI",9,"bold"), bd=1, relief="solid")
        card3.pack(side="left", padx=(0,12), pady=4, fill="y", expand=True)
        self.stat_total = self._stat_lbl(card3,"Total","0")
        self.stat_livre = self._stat_lbl(card3,"Livres","0",C["green"])
        self.stat_ocup  = self._stat_lbl(card3,"Ocupadas","0",C["red"])
        self.stat_known = self._stat_lbl(card3,"Identificadas","0",C["yellow"])

        pf = tk.Frame(main, bg=C["bg"]); pf.pack(fill="x", pady=(12,4))
        self.prog_label = tk.Label(pf, text="Aguardando...", bg=C["bg"], fg=C["subtext"],
                                    font=("Segoe UI",9))
        self.prog_label.pack(anchor="w")
        self.progress = ttk.Progressbar(pf, mode="determinate")
        self.progress.pack(fill="x", pady=4)

        bf = tk.Frame(main, bg=C["bg"]); bf.pack(fill="x", pady=4)
        self.scan_btn = ttk.Button(bf, text="▶ Iniciar", style="Accent.TButton", command=self._start_scan)
        self.scan_btn.pack(side="left", padx=(0,8))
        self.stop_btn = ttk.Button(bf, text="■ Parar", style="Danger.TButton",
                                    command=self._stop_scan, state="disabled")
        self.stop_btn.pack(side="left", padx=(0,8))
        ttk.Button(bf, text="💾 Exportar CSV", style="Success.TButton",
                   command=self._export_csv).pack(side="left")

        tk.Label(main, text="Atalhos Ocupados", bg=C["bg"], fg=C["text"],
                 font=("Segoe UI",11,"bold")).pack(anchor="w", pady=(12,4))
        of = tk.Frame(main, bg=C["surface"]); of.pack(fill="both", expand=True)
        cols = ("Modificador","Tecla","Atalho","Provável App","VK Code")
        self.occ_tree = ttk.Treeview(of, columns=cols, show="headings", height=7)
        for c,w in zip(cols,[130,80,180,260,90]):
            self.occ_tree.heading(c,text=c); self.occ_tree.column(c,width=w,anchor="center")
        self.occ_tree.tag_configure("row",     foreground=C["red"])
        self.occ_tree.tag_configure("known",   foreground=C["yellow"])
        sb2=ttk.Scrollbar(of,orient="vertical",command=self.occ_tree.yview)
        self.occ_tree.configure(yscrollcommand=sb2.set)
        sb2.pack(side="right",fill="y"); self.occ_tree.pack(fill="both",expand=True)

    def _stat_lbl(self, parent, title, value, color=None):
        color = color or self.COLORS["text"]
        f = tk.Frame(parent, bg=self.COLORS["surface"]); f.pack(fill="x", padx=10, pady=3)
        tk.Label(f, text=title+":", bg=self.COLORS["surface"], fg=self.COLORS["subtext"],
                 font=("Segoe UI",9)).pack(side="left")
        lbl = tk.Label(f, text=value, bg=self.COLORS["surface"], fg=color,
                       font=("Segoe UI",10,"bold"))
        lbl.pack(side="right"); return lbl

    def _get_scan_keys(self):
        keys = {}
        if self.key_groups["Letras A-Z"].get():
            keys.update({k:v for k,v in VK_CODES.items() if len(k)==1 and k.isalpha()})
        if self.key_groups["Números 0-9"].get():
            keys.update({k:v for k,v in VK_CODES.items() if len(k)==1 and k.isdigit()})
        if self.key_groups["F1-F12"].get():
            keys.update({k:v for k,v in VK_CODES.items() if k.startswith("F") and k[1:].isdigit()})
        if self.key_groups["Especiais"].get():
            specials=["PrtScn","Ins","Del","Home","End","PgUp","PgDn","Esc","Tab",
                      "Space","Enter","Backspace","Left","Up","Right","Down",
                      ";","=",",","-",".","/","`"]
            keys.update({k:VK_CODES[k] for k in specials if k in VK_CODES})
        if self.key_groups["Numpad"].get():
            keys.update({k:v for k,v in VK_CODES.items() if k.startswith("Num")})
        return keys

    def _start_scan(self):
        if self.scanning: return
        mods  = {n:MODIFIERS[n] for n,v in self.mod_checks.items() if v.get()}
        keys  = self._get_scan_keys()
        if not mods or not keys:
            messagebox.showwarning("Aviso","Selecione ao menos um modificador e grupo."); return
        self.scanning = True; self.scan_results.clear()
        self.scan_btn.config(state="disabled"); self.stop_btn.config(state="normal")
        for t in self.occ_tree.get_children(): self.occ_tree.delete(t)
        self.stat_total.config(text="0"); self.stat_livre.config(text="0")
        self.stat_ocup.config(text="0");  self.stat_known.config(text="0")
        total = len(mods)*len(keys)
        self.progress["maximum"]=total; self.progress["value"]=0
        threading.Thread(target=self._scan_worker, args=(mods,keys,total), daemon=True).start()

    def _scan_worker(self, mods, keys, total):
        done=livre=ocup=known=0
        for mod_name, mod_val in mods.items():
            if not self.scanning: break
            for key_name, vk_val in keys.items():
                if not self.scanning: break
                label = f"{mod_name} + {key_name}" if mod_name!="None" else key_name
                avail = check_hotkey(mod_val, vk_val)
                done+=1
                if avail:
                    livre+=1
                    kinfo=[]
                else:
                    ocup+=1
                    kinfo = get_known_app_info(mod_val, vk_val)
                    if kinfo: known+=1
                rec={"mod":mod_name,"key":key_name,"label":label,
                     "avail":avail,"vk":hex(vk_val),"known":kinfo}
                self.scan_results.append(rec)
                self.root.after(0, self._upd_scan, rec, done, livre, ocup, known, total)
        self.root.after(0, self._scan_done)

    def _upd_scan(self, rec, done, livre, ocup, known, total):
        self.progress["value"]=done
        self.prog_label.config(text=f"Verificando... {done}/{total}  |  🚫 {ocup} ocupadas  |  📌 {known} identificadas")
        self.stat_total.config(text=str(done)); self.stat_livre.config(text=str(livre))
        self.stat_ocup.config(text=str(ocup));  self.stat_known.config(text=str(known))
        if not rec["avail"]:
            app_str = rec["known"][0][0] if rec["known"] else "Desconhecido"
            tag = "known" if rec["known"] else "row"
            self.occ_tree.insert("","end",
                values=(rec["mod"],rec["key"],rec["label"],app_str,rec["vk"]),
                tags=(tag,))

    def _scan_done(self):
        self.scanning=False
        self.scan_btn.config(state="normal"); self.stop_btn.config(state="disabled")
        self.prog_label.config(text=f"✅ Concluído! {len(self.scan_results)} combinações verificadas.")

    def _stop_scan(self):
        self.scanning=False
        self.stop_btn.config(state="disabled"); self.scan_btn.config(state="normal")
        self.prog_label.config(text="⚠ Interrompido.")

    # ──────────────────────────────────────────────────────────────────────────
    # ABA 3 – PROCESSOS / APPS
    # ──────────────────────────────────────────────────────────────────────────
    def _build_processes(self, parent):
        C = self.COLORS

        # ── Estrutura: painel superior fixo (tabela) + painel inferior scrollável (ferramentas) ──
        outer = tk.Frame(parent, bg=C["bg"]); outer.pack(fill="both", expand=True)

        # ── Topo fixo: título + botão atualizar + tabela de processos ────────
        top_fixed = tk.Frame(outer, bg=C["bg"]); top_fixed.pack(fill="x", padx=20, pady=(12,4))
        tk.Label(top_fixed, text="Processos com Janelas Abertas", bg=C["bg"], fg=C["text"],
                 font=("Segoe UI",13,"bold")).pack(anchor="w")
        tk.Label(top_fixed, text="Lista todos os apps ativos que podem estar registrando atalhos globais",
                 bg=C["bg"], fg=C["subtext"], font=("Segoe UI",9)).pack(anchor="w", pady=(2,6))
        ttk.Button(top_fixed, text="🔄 Atualizar Lista", style="Accent.TButton",
                   command=self._refresh_procs).pack(anchor="w", pady=(0,6))

        # Tabela de processos (altura fixa, sempre visível)
        pf = tk.Frame(outer, bg=C["surface"]); pf.pack(fill="x", padx=20, pady=(0,6))
        cols = ("PID","Nome do Processo","Janela Ativa","Caminho do Executável")
        self.proc_tree = ttk.Treeview(pf, columns=cols, show="headings", height=8)
        for c,w in zip(cols,[70,180,260,420]):
            self.proc_tree.heading(c, text=c); self.proc_tree.column(c, width=w, anchor="w")
        self.proc_tree.tag_configure("known",   foreground=C["yellow"])
        self.proc_tree.tag_configure("suspect", foreground=C["accent"])
        sb_proc = ttk.Scrollbar(pf, orient="vertical",   command=self.proc_tree.yview)
        sb_proc_h = ttk.Scrollbar(pf, orient="horizontal", command=self.proc_tree.xview)
        self.proc_tree.configure(yscrollcommand=sb_proc.set, xscrollcommand=sb_proc_h.set)
        sb_proc.pack(side="right", fill="y")
        sb_proc_h.pack(side="bottom", fill="x")
        self.proc_tree.pack(fill="x", expand=False)
        tk.Label(outer, text="💛 Amarelo = na base de atalhos conhecidos   🟣 Roxo = suspeito identificado",
                 bg=C["bg"], fg=C["yellow"], font=("Segoe UI",8)).pack(anchor="w", padx=20)

        # ── Separador ────────────────────────────────────────────────────────
        tk.Frame(outer, bg=C["border"], height=1).pack(fill="x", padx=12, pady=6)
        tk.Label(outer, text="🛠  Ferramentas de Identificação  (role para ver todas)",
                 bg=C["bg"], fg=C["accent"], font=("Segoe UI",10,"bold")).pack(anchor="w", padx=20, pady=(0,4))

        # ── Área inferior scrollável com as ferramentas ───────────────────────
        scroll_outer = tk.Frame(outer, bg=C["bg"]); scroll_outer.pack(fill="both", expand=True, padx=12)
        canvas = tk.Canvas(scroll_outer, bg=C["bg"], highlightthickness=0)
        vsb    = ttk.Scrollbar(scroll_outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        tools = tk.Frame(canvas, bg=C["bg"]); tools.pack(fill="x")
        canvas_win = canvas.create_window((0, 0), window=tools, anchor="nw")

        def _on_tools_resize(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(canvas_win, width=canvas.winfo_width())
        tools.bind("<Configure>", _on_tools_resize)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_win, width=e.width))

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # ── Ferramenta 1: Teste por Eliminação (suspensão) ───────────────────
        elim_frame = tk.LabelFrame(tools, text=" ⚗ Técnica 1 — Teste por Eliminação (Suspensão) ",
                                    bg=C["surface"], fg=C["yellow"],
                                    font=("Segoe UI",10,"bold"), bd=1, relief="solid")
        elim_frame.pack(fill="x", padx=8, pady=(8,6), ipady=4)
        tk.Label(elim_frame,
            text="Suspende cada processo temporariamente e verifica se o atalho fica disponível.\n"
                 "⚠ Processos críticos do Windows são ignorados automaticamente.",
            bg=C["surface"], fg=C["subtext"], font=("Segoe UI",9)).pack(anchor="w", padx=12, pady=(4,4))
        row2 = tk.Frame(elim_frame, bg=C["surface"]); row2.pack(fill="x", padx=12, pady=4)
        tk.Label(row2, text="Modificador:", bg=C["surface"], fg=C["subtext"]).pack(side="left")
        self.elim_mod = tk.StringVar(value="Alt")
        ttk.Combobox(row2, textvariable=self.elim_mod, values=list(MODIFIERS.keys()),
                     width=12, state="readonly").pack(side="left", padx=(6,16))
        tk.Label(row2, text="Tecla:", bg=C["surface"], fg=C["subtext"]).pack(side="left")
        self.elim_key = tk.StringVar(value="A")
        ttk.Combobox(row2, textvariable=self.elim_key, values=sorted(VK_CODES.keys()),
                     width=10, state="readonly").pack(side="left", padx=(6,16))
        self.elim_btn = ttk.Button(row2, text="⚗ Iniciar Teste", style="Warn.TButton",
                                    command=self._start_elimination)
        self.elim_btn.pack(side="left", padx=4)
        self._elim_stop_btn = ttk.Button(row2, text="⛔ Parar", style="Danger.TButton",
                                          command=self._elim_stop_test, state="disabled")
        self._elim_stop_btn.pack(side="left", padx=4)
        self.elim_result = tk.Label(elim_frame, text="", bg=C["surface"],
                                     fg=C["text"], font=("Segoe UI",10,"bold"), wraplength=850, justify="left")
        self.elim_result.pack(anchor="w", padx=12, pady=(0,4))
        self.elim_prog = ttk.Progressbar(elim_frame, mode="determinate")
        self.elim_prog.pack(fill="x", padx=12, pady=(0,6))

        # ── Ferramenta 2: Sondagem por Janela (hwnd) ─────────────────────────
        probe_frame = tk.LabelFrame(tools, text=" 🔬 Técnica 2 — Sondagem por Janela (hwnd) ",
                                     bg=C["surface"], fg=C["accent"],
                                     font=("Segoe UI",10,"bold"), bd=1, relief="solid")
        probe_frame.pack(fill="x", padx=8, pady=(0,6), ipady=4)
        tk.Label(probe_frame,
            text="Tenta registrar o atalho em cada janela aberta e observa qual rejeita.\n"
                 "Não suspende processos — funciona mesmo com apps da Store e processos protegidos.",
            bg=C["surface"], fg=C["subtext"], font=("Segoe UI",9)).pack(anchor="w", padx=12, pady=(4,4))
        row_p = tk.Frame(probe_frame, bg=C["surface"]); row_p.pack(fill="x", padx=12, pady=4)
        tk.Label(row_p, text="Modificador:", bg=C["surface"], fg=C["subtext"]).pack(side="left")
        self.probe_mod = tk.StringVar(value="Alt")
        ttk.Combobox(row_p, textvariable=self.probe_mod, values=list(MODIFIERS.keys()),
                     width=12, state="readonly").pack(side="left", padx=(6,16))
        tk.Label(row_p, text="Tecla:", bg=C["surface"], fg=C["subtext"]).pack(side="left")
        self.probe_key = tk.StringVar(value="A")
        ttk.Combobox(row_p, textvariable=self.probe_key, values=sorted(VK_CODES.keys()),
                     width=10, state="readonly").pack(side="left", padx=(6,16))
        self.probe_use_focus = tk.BooleanVar(value=False)
        self.probe_use_wm    = tk.BooleanVar(value=False)
        chk_f = tk.Frame(row_p, bg=C["surface"]); chk_f.pack(side="left", padx=(0,12))
        tk.Checkbutton(chk_f, text="Técnica B (foco)", variable=self.probe_use_focus,
                       bg=C["surface"], fg=C["subtext"], selectcolor=C["surface2"],
                       activebackground=C["surface"], font=("Segoe UI",8)).pack(anchor="w")
        tk.Checkbutton(chk_f, text="Técnica C (WM_HOTKEY)", variable=self.probe_use_wm,
                       bg=C["surface"], fg=C["subtext"], selectcolor=C["surface2"],
                       activebackground=C["surface"], font=("Segoe UI",8)).pack(anchor="w")
        self.probe_btn = ttk.Button(row_p, text="🔬 Sondar", style="Accent.TButton",
                                     command=self._start_probe)
        self.probe_btn.pack(side="left", padx=4)
        self._probe_stop_btn = ttk.Button(row_p, text="⛔ Parar", style="Danger.TButton",
                                           command=self._probe_stop, state="disabled")
        self._probe_stop_btn.pack(side="left", padx=4)
        tk.Label(probe_frame,
            text="  A: RegisterHotKey por hwnd   B: Libera ao receber foco ⚠janela muda   C: WM_HOTKEY simulado ⚠heurístico",
            bg=C["surface"], fg=C["subtext"], font=("Segoe UI",8), justify="left"
        ).pack(anchor="w", padx=24, pady=(0,2))
        self.probe_result = tk.Label(probe_frame, text="", bg=C["surface"],
                                      fg=C["text"], font=("Segoe UI",10,"bold"),
                                      wraplength=850, justify="left")
        self.probe_result.pack(anchor="w", padx=12, pady=(0,4))
        self.probe_prog = ttk.Progressbar(probe_frame, mode="determinate")
        self.probe_prog.pack(fill="x", padx=12, pady=(0,6))

        # ── Ferramenta 3: Técnica D — Fechar e Reabrir ───────────────────────
        kill_frame = tk.LabelFrame(tools, text=" 💀 Técnica D — Fechar e Reabrir (Modo Agressivo) ",
                                    bg=C["surface"], fg=C["red"],
                                    font=("Segoe UI",10,"bold"), bd=1, relief="solid")
        kill_frame.pack(fill="x", padx=8, pady=(0,12), ipady=4)

        # Aviso de risco
        risk_box = tk.Frame(kill_frame, bg="#3a1a1a", bd=1,
                             highlightthickness=1, highlightbackground=C["red"])
        risk_box.pack(fill="x", padx=12, pady=(6,4))
        tk.Label(risk_box, text="⚠  ATENÇÃO — LEIA ANTES DE USAR",
                 bg="#3a1a1a", fg=C["red"], font=("Segoe UI",9,"bold")).pack(anchor="w", padx=10, pady=(6,2))
        tk.Label(risk_box,
            text="Esta técnica ENCERRA processos forçadamente (TerminateProcess) e tenta reabri-los.\n"
                 "Riscos conhecidos:\n"
                 "  • Perda de dados não salvos no app encerrado\n"
                 "  • Alguns apps não sobem corretamente ao serem reabertos desta forma\n"
                 "  • Apps da Store (UWP) podem não reabrir via caminho de executável direto\n"
                 "  • Apps com múltiplas instâncias ou argumentos especiais podem se comportar de forma inesperada\n"
                 "  • Processos protegidos (antivírus, DRM) serão ignorados automaticamente\n\n"
                 "Use apenas como último recurso, quando as Técnicas 1 e 2 não identificarem o responsável.",
            bg="#3a1a1a", fg="#ffb86c", font=("Segoe UI",8), justify="left"
        ).pack(anchor="w", padx=10, pady=(0,8))

        row_k = tk.Frame(kill_frame, bg=C["surface"]); row_k.pack(fill="x", padx=12, pady=4)
        tk.Label(row_k, text="Modificador:", bg=C["surface"], fg=C["subtext"]).pack(side="left")
        self.kill_mod = tk.StringVar(value="Alt")
        ttk.Combobox(row_k, textvariable=self.kill_mod, values=list(MODIFIERS.keys()),
                     width=12, state="readonly").pack(side="left", padx=(6,16))
        tk.Label(row_k, text="Tecla:", bg=C["surface"], fg=C["subtext"]).pack(side="left")
        self.kill_key = tk.StringVar(value="A")
        ttk.Combobox(row_k, textvariable=self.kill_key, values=sorted(VK_CODES.keys()),
                     width=10, state="readonly").pack(side="left", padx=(6,16))

        # Checkbox de confirmação obrigatória
        self.kill_confirmed = tk.BooleanVar(value=False)
        tk.Checkbutton(row_k,
                       text="Entendo os riscos e quero prosseguir",
                       variable=self.kill_confirmed,
                       bg=C["surface"], fg=C["red"], selectcolor=C["surface2"],
                       activebackground=C["surface"],
                       font=("Segoe UI",9,"bold")).pack(side="left", padx=(0,12))

        self.kill_btn = ttk.Button(row_k, text="💀 Iniciar", style="Danger.TButton",
                                    command=self._start_kill_probe)
        self.kill_btn.pack(side="left", padx=4)
        self._kill_stop_btn = ttk.Button(row_k, text="⛔ Parar", style="Warn.TButton",
                                          command=self._kill_probe_stop, state="disabled")
        self._kill_stop_btn.pack(side="left", padx=4)

        self.kill_result = tk.Label(kill_frame, text="", bg=C["surface"],
                                     fg=C["text"], font=("Segoe UI",10,"bold"),
                                     wraplength=850, justify="left")
        self.kill_result.pack(anchor="w", padx=12, pady=(4,4))
        self.kill_prog = ttk.Progressbar(kill_frame, mode="determinate")
        self.kill_prog.pack(fill="x", padx=12, pady=(0,6))

    def _refresh_procs(self):
        for t in self.proc_tree.get_children(): self.proc_tree.delete(t)
        self.all_windows = []
        threading.Thread(target=self._load_procs, daemon=True).start()

    def _load_procs(self):
        wins = get_windows_and_processes()
        seen_pids = set()
        unique = []
        for w in wins:
            if w["pid"] not in seen_pids and w["pid"] != OWN_PID:
                seen_pids.add(w["pid"])
                unique.append(w)
        self.all_windows = unique
        self.root.after(0, self._populate_procs, unique)

    def _populate_procs(self, procs):
        known_exes = {info[1].lower() for lst in KNOWN_HOTKEYS.values() for info in lst}
        for p in procs:
            is_known = p["name"].lower() in known_exes
            tag = ("known",) if is_known else ()
            self.proc_tree.insert("","end",
                values=(p["pid"], p["name"], p["title"][:60], p["path"]),
                tags=tag)

    def _start_elimination(self):
        if not HAS_PSUTIL:
            messagebox.showerror("Erro","psutil é necessário.\n\npip install psutil"); return
        if not self.all_windows:
            messagebox.showinfo("Info","Clique em 'Atualizar Lista' primeiro."); return
        mod_val = MODIFIERS.get(self.elim_mod.get(), 0)
        vk_val  = VK_CODES.get(self.elim_key.get(), 0)
        label   = f"{self.elim_mod.get()} + {self.elim_key.get()}"

        # Filtra processos críticos do Windows
        CRITICAL = {"system","idle","lsass.exe","winlogon.exe","csrss.exe","smss.exe",
                    "services.exe","svchost.exe","explorer.exe","dwm.exe","wininit.exe",
                    "fontdrvhost.exe","spoolsv.exe","taskhostw.exe","sihost.exe",
                    "ctfmon.exe","runtimebroker.exe","searchhost.exe","startmenuexperiencehost.exe",
                    "python.exe","pythonw.exe"}  # evita suspender a si mesmo

        proc_list = [w for w in self.all_windows
                     if w["name"].lower() not in CRITICAL
                     and w["pid"] != OWN_PID]

        self._elim_stop = threading.Event()
        self.elim_btn.config(state="disabled")
        if hasattr(self, "_elim_stop_btn"):
            self._elim_stop_btn.config(state="normal")
        self.elim_result.config(
            text=f"🔄 Testando '{label}' — {len(proc_list)} processos (3s timeout cada)...",
            fg=self.COLORS["subtext"])
        self.elim_prog["maximum"] = len(proc_list)
        self.elim_prog["value"]   = 0

        def run():
            def cb(i, total, name):
                self.root.after(0, self._elim_progress, i, total, name)
            result = find_by_elimination(mod_val, vk_val, proc_list, cb, self._elim_stop)
            self.root.after(0, self._elim_done, result, label)

        threading.Thread(target=run, daemon=True).start()

    def _elim_progress(self, i, total, name):
        self.elim_prog["value"] = i
        self.elim_result.config(
            text=f"🔄 Testando processo {i+1}/{total}: {name}",
            fg=self.COLORS["subtext"])

    def _elim_stop_test(self):
        if hasattr(self, "_elim_stop"):
            self._elim_stop.set()
        self.elim_result.config(text="⛔ Teste interrompido pelo usuário.", fg=self.COLORS["yellow"])
        self.elim_btn.config(state="normal")
        if hasattr(self, "_elim_stop_btn"):
            self._elim_stop_btn.config(state="disabled")

    # ──────────────────────────────────────────────────────────────────────────
    # SONDAGEM POR JANELA (hwnd_probe)
    # ──────────────────────────────────────────────────────────────────────────
    def _start_probe(self):
        mod_val = MODIFIERS.get(self.probe_mod.get(), 0)
        vk_val  = VK_CODES.get(self.probe_key.get(), 0)
        label   = f"{self.probe_mod.get()} + {self.probe_key.get()}"

        self._probe_stop_event = threading.Event()
        self.probe_btn.config(state="disabled")
        self._probe_stop_btn.config(state="normal")
        self.probe_result.config(
            text=f"🔬 Sondando janelas para '{label}'...",
            fg=self.COLORS["subtext"])
        self.probe_prog["value"] = 0
        self.probe_prog["maximum"] = 100

        use_focus = self.probe_use_focus.get()
        use_wm    = self.probe_use_wm.get()

        def run():
            def cb(current, total, name):
                self.root.after(0, self._probe_progress, current, total, name)

            result = probe_hotkey_owner(
                mod_val, vk_val,
                use_focus_technique=use_focus,
                use_wm_technique=use_wm,
                progress_cb=cb,
                stop_event=self._probe_stop_event
            )
            self.root.after(0, self._probe_done, result, label)

        threading.Thread(target=run, daemon=True).start()

    def _probe_progress(self, current, total, name):
        if total > 0:
            self.probe_prog["maximum"] = total
            self.probe_prog["value"]   = current
        self.probe_result.config(
            text=f"🔬 Sondando {current+1}/{total}: {name[:60]}",
            fg=self.COLORS["subtext"])

    def _probe_stop(self):
        if hasattr(self, "_probe_stop_event"):
            self._probe_stop_event.set()
        self.probe_result.config(text="⛔ Sondagem interrompida.", fg=self.COLORS["yellow"])
        self.probe_btn.config(state="normal")
        self._probe_stop_btn.config(state="disabled")

    def _probe_done(self, result: "ProbeResult", label: str):
        C = self.COLORS
        self.probe_btn.config(state="normal")
        self._probe_stop_btn.config(state="disabled")
        self.probe_prog["value"] = self.probe_prog["maximum"]

        METHOD_LABELS = {
            "hwnd_register": "RegisterHotKey por hwnd",
            "focus_release":  "Liberação ao receber foco",
            "wm_hotkey":      "Reação a WM_HOTKEY",
            "none":           "—",
        }
        CONF_COLORS = {
            "high":   C["green"],
            "medium": C["yellow"],
            "low":    C["orange"],
        }

        if not result.found:
            notes_str = "\n   ".join(result.notes) if result.notes else ""
            self.probe_result.config(
                text=f"❓ Nenhuma janela identificada como responsável por '{label}'.\n"
                     f"   {notes_str}",
                fg=C["orange"])
            return

        method_str = METHOD_LABELS.get(result.method, result.method)
        conf_color = CONF_COLORS.get(result.confidence, C["text"])
        conf_label = {"high":"ALTA","medium":"MÉDIA","low":"BAIXA"}.get(result.confidence, result.confidence)
        notes_str  = "  |  ".join(result.notes) if result.notes else ""

        self.probe_result.config(
            text=(
                f"{'✅' if result.confidence=='high' else '⚠'} "
                f"Responsável: '{result.name}'  (PID {result.pid})\n"
                f"   Janela  : {result.title[:70]}\n"
                f"   Caminho : {result.path}\n"
                f"   Método  : {method_str}   |   Confiança: {conf_label}\n"
                f"   {notes_str}"
            ),
            fg=conf_color)

        # Destaca o processo na tabela se estiver listado
        for item in self.proc_tree.get_children():
            vals = self.proc_tree.item(item, "values")
            try:
                if int(vals[0]) == result.pid:
                    self.proc_tree.item(item, tags=("suspect",))
                    self.proc_tree.see(item)
            except Exception:
                pass
        self.proc_tree.tag_configure("suspect", foreground=C["accent"])

    # ──────────────────────────────────────────────────────────────────────────
    # TÉCNICA D — FECHAR E REABRIR
    # ──────────────────────────────────────────────────────────────────────────
    def _start_kill_probe(self):
        if not self.kill_confirmed.get():
            messagebox.showwarning(
                "Confirmação necessária",
                "Marque a caixa 'Entendo os riscos e quero prosseguir' antes de iniciar.")
            return
        if not self.all_windows:
            messagebox.showinfo("Info", "Clique em 'Atualizar Lista' primeiro."); return

        mod_val = MODIFIERS.get(self.kill_mod.get(), 0)
        vk_val  = VK_CODES.get(self.kill_key.get(), 0)
        label   = f"{self.kill_mod.get()} + {self.kill_key.get()}"

        proc_list = [w for w in self.all_windows
                     if w["name"].lower() not in KILL_BLACKLIST
                     and w["pid"] != OWN_PID]

        # Confirmação final com lista do que será testado
        names_preview = "\n".join(f"  • {p['name']} (PID {p['pid']})"
                                   for p in proc_list[:12])
        extra = f"\n  ... e mais {len(proc_list)-12} outros" if len(proc_list) > 12 else ""
        confirm = messagebox.askyesno(
            "⚠ Confirmar Técnica D",
            f"Serão encerrados e reabertos até {len(proc_list)} processos "
            f"para identificar o dono de  '{label}'.\n\n"
            f"Processos que serão testados:\n{names_preview}{extra}\n\n"
            f"O teste para quando o responsável for encontrado.\n"
            f"Salve seus trabalhos abertos antes de continuar.\n\n"
            f"Deseja realmente prosseguir?",
            icon="warning")
        if not confirm:
            return

        self._kill_stop_event = threading.Event()
        self.kill_btn.config(state="disabled")
        self._kill_stop_btn.config(state="normal")
        self.kill_result.config(
            text=f"💀 Testando '{label}' — {len(proc_list)} processos...",
            fg=self.COLORS["subtext"])
        self.kill_prog["maximum"] = len(proc_list)
        self.kill_prog["value"]   = 0

        def run():
            def cb(current, total, name):
                self.root.after(0, self._kill_probe_progress, current, total, name)
            result = probe_by_kill_reopen(
                mod_val, vk_val, proc_list,
                reopen=True,
                progress_cb=cb,
                stop_event=self._kill_stop_event
            )
            self.root.after(0, self._kill_probe_done, result, label)

        threading.Thread(target=run, daemon=True).start()

    def _kill_probe_progress(self, current, total, name):
        self.kill_prog["value"] = current
        self.kill_result.config(
            text=f"💀 Encerrando e verificando {current+1}/{total}: {name[:60]}",
            fg=self.COLORS["subtext"])

    def _kill_probe_stop(self):
        if hasattr(self, "_kill_stop_event"):
            self._kill_stop_event.set()
        self.kill_result.config(text="⛔ Técnica D interrompida.", fg=self.COLORS["yellow"])
        self.kill_btn.config(state="normal")
        self._kill_stop_btn.config(state="disabled")

    def _kill_probe_done(self, result: "ProbeResult", label: str):
        C = self.COLORS
        self.kill_btn.config(state="normal")
        self._kill_stop_btn.config(state="disabled")
        self.kill_prog["value"] = self.kill_prog["maximum"]

        if not result.found:
            notes_str = "\n   ".join(result.notes) if result.notes else ""
            self.kill_result.config(
                text=f"❓ Técnica D não identificou o responsável por '{label}'.\n   {notes_str}",
                fg=C["orange"])
            return

        notes_str = "\n   ".join(result.notes) if result.notes else ""
        self.kill_result.config(
            text=(
                f"✅ IDENTIFICADO via Técnica D!\n"
                f"   Processo: '{result.name}'  (PID {result.pid})\n"
                f"   Caminho : {result.path}\n"
                f"   {notes_str}"
            ),
            fg=C["green"])

        # Destaca na tabela de processos (PID pode ter mudado se reabriu)
        for item in self.proc_tree.get_children():
            vals = self.proc_tree.item(item, "values")
            try:
                if int(vals[0]) == result.pid or vals[1] == result.name:
                    self.proc_tree.item(item, tags=("suspect",))
                    self.proc_tree.see(item)
            except Exception:
                pass

    def _show_suspects_popup(self, suspects: list):
        """Abre janela popup com tabela legível dos processos suspeitos."""
        C = self.COLORS
        popup = tk.Toplevel(self.root)
        popup.title(f"⚠ Suspeitos — Access Denied ({len(suspects)} processos)")
        popup.geometry("780x420")
        popup.configure(bg=C["bg"])
        popup.resizable(True, True)
        popup.transient(self.root)
        popup.grab_set()

        # Cabeçalho
        hdr = tk.Frame(popup, bg=C["surface"]); hdr.pack(fill="x")
        tk.Label(hdr,
            text=f"  ⚠  {len(suspects)} processo(s) não puderam ser suspensos",
            bg=C["surface"], fg=C["yellow"],
            font=("Segoe UI", 12, "bold")).pack(side="left", padx=16, pady=10)
        tk.Label(hdr,
            text="Access Denied — possivelmente apps da Store ou processos protegidos",
            bg=C["surface"], fg=C["subtext"],
            font=("Segoe UI", 9)).pack(side="left", padx=4)

        # Tabela
        frm = tk.Frame(popup, bg=C["bg"]); frm.pack(fill="both", expand=True, padx=12, pady=8)
        cols = ("PID", "Nome", "Janela", "Caminho")
        tree = ttk.Treeview(frm, columns=cols, show="headings")
        for col, w in zip(cols, [70, 160, 240, 340]):
            tree.heading(col, text=col)
            tree.column(col, width=w, anchor="w")
        tree.tag_configure("row", foreground=C["orange"])

        vsb = ttk.Scrollbar(frm, orient="vertical",   command=tree.yview)
        hsb = ttk.Scrollbar(frm, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right",  fill="y")
        hsb.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)

        for s in suspects:
            tree.insert("", "end", tags=("row",), values=(
                s.get("pid",  "?"),
                s.get("name", "?"),
                s.get("title","?")[:60],
                s.get("path", "?"),
            ))

        # Rodapé
        foot = tk.Frame(popup, bg=C["bg"]); foot.pack(fill="x", padx=12, pady=(0,10))
        tk.Label(foot,
            text="💡 Para testar estes apps, execute o HotkeyTracker como Administrador.",
            bg=C["bg"], fg=C["subtext"], font=("Segoe UI", 9)).pack(side="left")
        ttk.Button(foot, text="Fechar", style="Accent.TButton",
                   command=popup.destroy).pack(side="right")

    def _elim_done(self, result, label):
        C = self.COLORS
        self.elim_btn.config(state="normal")
        if hasattr(self, "_elim_stop_btn"):
            self._elim_stop_btn.config(state="disabled")

        confirmed = result.get("confirmed")
        suspects  = result.get("suspects", [])
        stopped   = result.get("stopped", False)

        if stopped:
            self.elim_result.config(text="⛔ Teste interrompido.", fg=C["yellow"])
            return

        if confirmed:
            name = confirmed.get("name","?")
            pid  = confirmed.get("pid","?")
            path = confirmed.get("path","?")
            self.elim_result.config(
                text=f"✅ CONFIRMADO!  '{name}'  (PID {pid})\n📁 {path}",
                fg=C["green"])
            return

        if suspects:
            count = len(suspects)
            self.elim_result.config(
                text=f"⚠ {count} processo(s) não puderam ser testados (Access Denied).\n"
                     f"   Clique em '📋 Ver Suspeitos' para a lista completa.",
                fg=C["yellow"])
            # Destaca suspeitos na tabela principal
            suspect_pids = {s["pid"] for s in suspects}
            for item in self.proc_tree.get_children():
                vals = self.proc_tree.item(item, "values")
                try:
                    if int(vals[0]) in suspect_pids:
                        self.proc_tree.item(item, tags=("suspect",))
                except Exception:
                    pass
            self.proc_tree.tag_configure("suspect", foreground=C["orange"])

            # Botão para abrir popup com a lista completa
            if hasattr(self, "_suspects_btn"):
                self._suspects_btn.destroy()
            self._suspects_btn = ttk.Button(
                self.elim_result.master,
                text="📋 Ver Suspeitos",
                style="Warn.TButton",
                command=lambda s=suspects: self._show_suspects_popup(s))
            self._suspects_btn.pack(anchor="w", padx=12, pady=(0,4))
        else:
            self.elim_result.config(
                text="❓ Nenhum processo identificado. Pode ser serviço do sistema ou driver.",
                fg=C["orange"])

    # ──────────────────────────────────────────────────────────────────────────
    # ABA 4 – RESULTADOS
    # ──────────────────────────────────────────────────────────────────────────
    def _build_results(self, parent):
        C = self.COLORS
        main = tk.Frame(parent, bg=C["bg"]); main.pack(fill="both", expand=True, padx=12, pady=12)
        top = tk.Frame(main, bg=C["bg"]); top.pack(fill="x", pady=(0,8))
        tk.Label(top, text="Resultados do Último Scan", bg=C["bg"], fg=C["text"],
                 font=("Segoe UI",13,"bold")).pack(side="left")
        fil = tk.Frame(top, bg=C["bg"]); fil.pack(side="right")
        tk.Label(fil, text="Filtrar:", bg=C["bg"], fg=C["subtext"]).pack(side="left")
        self.filter_var = tk.StringVar(value="Todos")
        ttk.Combobox(fil, textvariable=self.filter_var,
                     values=["Todos","Livre","Ocupada","Identificada"], width=11,
                     state="readonly").pack(side="left", padx=6)
        self.filter_var.trace("w", lambda *a: self._apply_filter())
        tk.Label(fil, text="Buscar:", bg=C["bg"], fg=C["subtext"]).pack(side="left", padx=(12,0))
        self.search_var = tk.StringVar()
        self.search_var.trace("w", lambda *a: self._apply_filter())
        ttk.Entry(fil, textvariable=self.search_var, width=14).pack(side="left", padx=6)
        ttk.Button(fil, text="💾 Exportar", style="Success.TButton",
                   command=self._export_csv).pack(side="left", padx=6)

        cols = ("Modificador","Tecla","Atalho","Status","Provável App","VK (hex)")
        self.all_tree = ttk.Treeview(main, columns=cols, show="headings")
        for c,w in zip(cols,[120,80,180,110,230,90]):
            self.all_tree.heading(c,text=c); self.all_tree.column(c,width=w,anchor="center")
        self.all_tree.tag_configure("livre",      foreground=C["green"])
        self.all_tree.tag_configure("ocupada",    foreground=C["red"])
        self.all_tree.tag_configure("identified", foreground=C["yellow"])
        vsb=ttk.Scrollbar(main,orient="vertical",command=self.all_tree.yview)
        hsb=ttk.Scrollbar(main,orient="horizontal",command=self.all_tree.xview)
        self.all_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right",fill="y"); hsb.pack(side="bottom",fill="x")
        self.all_tree.pack(fill="both",expand=True)

    def _apply_filter(self, *args):
        f = self.filter_var.get(); q = self.search_var.get().lower()
        for t in self.all_tree.get_children(): self.all_tree.delete(t)
        for rec in self.scan_results:
            status = "✅ Livre" if rec["avail"] else "🚫 Ocupada"
            app_str = rec["known"][0][0] if rec.get("known") else ("—" if rec["avail"] else "Desconhecido")
            if rec["avail"]:   tag="livre"
            elif rec["known"]: tag="identified"
            else:              tag="ocupada"
            if f=="Livre"       and not rec["avail"]: continue
            if f=="Ocupada"     and rec["avail"]:     continue
            if f=="Identificada"and (rec["avail"] or not rec.get("known")): continue
            if q and q not in rec["label"].lower() and q not in app_str.lower(): continue
            self.all_tree.insert("","end",
                values=(rec["mod"],rec["key"],rec["label"],status,app_str,rec["vk"]),
                tags=(tag,))

    # ──────────────────────────────────────────────────────────────────────────
    # ABA 5 – BASE DE ATALHOS CONHECIDOS
    # ──────────────────────────────────────────────────────────────────────────
    def _build_known(self, parent):
        C = self.COLORS
        main = tk.Frame(parent, bg=C["bg"]); main.pack(fill="both", expand=True, padx=20, pady=14)
        tk.Label(main, text="Base de Atalhos Conhecidos", bg=C["bg"], fg=C["text"],
                 font=("Segoe UI",13,"bold")).pack(anchor="w")
        tk.Label(main, text="Lista de atalhos globais registrados por aplicativos populares",
                 bg=C["bg"], fg=C["subtext"], font=("Segoe UI",9)).pack(anchor="w", pady=(2,12))

        sf = tk.Frame(main, bg=C["bg"]); sf.pack(fill="x", pady=4)
        tk.Label(sf, text="Buscar:", bg=C["bg"], fg=C["subtext"]).pack(side="left")
        self.kb_search = tk.StringVar()
        self.kb_search.trace("w", lambda *a: self._filter_known())
        ttk.Entry(sf, textvariable=self.kb_search, width=20).pack(side="left", padx=6)

        kf = tk.Frame(main, bg=C["surface"]); kf.pack(fill="both", expand=True)
        cols = ("Atalho","Aplicativo","Executável","Descrição")
        self.kb_tree = ttk.Treeview(kf, columns=cols, show="headings")
        for c,w in zip(cols,[160,160,180,380]):
            self.kb_tree.heading(c,text=c); self.kb_tree.column(c,width=w,anchor="w")
        sb=ttk.Scrollbar(kf,orient="vertical",command=self.kb_tree.yview)
        self.kb_tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right",fill="y"); self.kb_tree.pack(fill="both",expand=True)
        self._populate_known()

    def _hotkey_label(self, mod_val, vk_val):
        mod_name = next((n for n,v in MODIFIERS.items() if v==mod_val and n!="None"), None)
        key_name = next((n for n,v in VK_CODES.items() if v==vk_val), hex(vk_val))
        return f"{mod_name} + {key_name}" if mod_name else key_name

    def _populate_known(self, q=""):
        for t in self.kb_tree.get_children(): self.kb_tree.delete(t)
        for (mod_val, vk_val), apps in sorted(KNOWN_HOTKEYS.items()):
            lbl = self._hotkey_label(mod_val & ~MOD_NOREPEAT, vk_val)
            for app, exe, desc in apps:
                if q and q not in lbl.lower() and q not in app.lower() and q not in desc.lower():
                    continue
                self.kb_tree.insert("","end", values=(lbl, app, exe, desc))

    def _filter_known(self):
        self._populate_known(self.kb_search.get().lower())

    # ──────────────────────────────────────────────────────────────────────────
    # EXPORTAR
    # ──────────────────────────────────────────────────────────────────────────
    def _export_csv(self):
        if not self.scan_results:
            messagebox.showinfo("Info","Execute o scanner primeiro."); return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV","*.csv"),("JSON","*.json")],
            initialfile=f"hotkeys_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        if not path: return
        export = [{**r, "known_apps": ", ".join(k[0] for k in r.get("known",[]))} for r in self.scan_results]
        if path.endswith(".json"):
            with open(path,"w",encoding="utf-8") as f: json.dump(export,f,ensure_ascii=False,indent=2)
        else:
            with open(path,"w",newline="",encoding="utf-8") as f:
                w=csv.DictWriter(f,fieldnames=["mod","key","label","avail","vk","known_apps"])
                w.writeheader(); w.writerows(export)
        messagebox.showinfo("✅ Exportado", f"Salvo em:\n{path}")

# ─── MAIN ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    HotkeyTrackerApp(root)
    root.mainloop()
