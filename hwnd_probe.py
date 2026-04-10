"""
hwnd_probe.py — Descoberta de dono de hotkey por sondagem de janelas (hwnd)

Estratégia:
  O Windows permite registrar um hotkey vinculado a um hwnd específico
  (RegisterHotKey(hwnd, id, mod, vk)), e não apenas global (hwnd=None).
  Alguns apps registram atalhos globais sem hwnd, mas muitos registram
  vinculados à janela principal. Além disso, quando passamos um hwnd já
  "dono" do atalho, o RegisterHotKey falha de forma diferente.

  Aqui usamos três técnicas em sequência, da mais rápida à mais lenta:

  Técnica A — Sondagem por hwnd (rápida):
    Para cada janela top-level visível, tentamos RegisterHotKey(hwnd, ...)
    Um app que registrou o atalho vinculado ao seu hwnd irá causar falha
    quando tentarmos registrar com aquele mesmo hwnd.
    Obs: isso NÃO detecta atalhos registrados com hwnd=None por outro processo.

  Técnica B — Envio de WM_HOTKEY (rápida, read-only):
    Enviamos PostMessage(hwnd, WM_HOTKEY, id, lparam) para cada janela
    e observamos se a janela fica ativa/responde — apps que "escutam"
    WM_HOTKEY vão processar a mensagem.
    Limitação: não é um teste definitivo, mas serve de sinal secundário.

  Técnica C — Bloqueio temporário de foco (lenta, invasiva):
    Trazemos cada janela para foreground e tentamos RegisterHotKey(None,...).
    Se liberar (retornar True) depois de trazer aquela janela para frente,
    pode indicar que o app libera o atalho ao perder foco — padrão de
    alguns apps como Discord, Teams, Slack.

Resultado final:
  ProbeResult com:
    - method   : qual técnica identificou
    - hwnd     : handle da janela suspeita/confirmada
    - pid      : PID do processo
    - name     : nome do executável
    - title    : título da janela
    - path     : caminho do executável
    - confidence: "high" | "medium" | "low"
"""

import ctypes
import ctypes.wintypes as wintypes
import time
import logging
import os
from dataclasses import dataclass, field
from typing import Optional, List, Callable

try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

import os as _os
_OWN_PID = _os.getpid()

log = logging.getLogger("hotkey.hwnd_probe")

user32   = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# Win32 constants
WM_HOTKEY        = 0x0312
MOD_NOREPEAT     = 0x4000
PROBE_HOTKEY_ID  = 0xF002   # ID reservado para nossa sondagem
PROCESS_QUERY_INFORMATION = 0x0400
PROCESS_VM_READ           = 0x0010

# ─── Estrutura de resultado ───────────────────────────────────────────────────

@dataclass
class ProbeResult:
    found:      bool         = False
    method:     str          = ""       # "hwnd_register" | "focus_release" | "wm_hotkey" | "none"
    hwnd:       int          = 0
    pid:        int          = 0
    name:       str          = ""
    title:      str          = ""
    path:       str          = ""
    confidence: str          = "low"   # "high" | "medium" | "low"
    notes:      List[str]    = field(default_factory=list)


# ─── Helpers de Win32 ─────────────────────────────────────────────────────────

def _get_pid_from_hwnd(hwnd: int) -> int:
    pid = wintypes.DWORD(0)
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value


def _get_proc_info(pid: int) -> dict:
    """Retorna name/path do processo. Usa psutil se disponível."""
    info = {"name": "?", "path": "?"}
    if HAS_PSUTIL:
        try:
            p = psutil.Process(pid)
            info["name"] = p.name()
            info["path"] = p.exe()
        except Exception:
            pass
    else:
        hproc = kernel32.OpenProcess(
            PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, False, pid)
        if hproc:
            buf = ctypes.create_unicode_buffer(1024)
            ctypes.windll.psapi.GetModuleFileNameExW(hproc, None, buf, 1024)
            path = buf.value
            info["path"] = path
            info["name"] = os.path.basename(path) if path else "?"
            kernel32.CloseHandle(hproc)
    return info


def _get_window_title(hwnd: int) -> str:
    buf = ctypes.create_unicode_buffer(512)
    user32.GetWindowTextW(hwnd, buf, 512)
    return buf.value.strip()


def _enum_top_windows() -> List[int]:
    """Retorna todos os hwnds de janelas top-level visíveis com título."""
    hwnds = []
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)

    def cb(hwnd, _):
        if user32.IsWindowVisible(hwnd):
            buf = ctypes.create_unicode_buffer(256)
            user32.GetWindowTextW(hwnd, buf, 256)
            if buf.value.strip():
                hwnds.append(hwnd)
        return True

    user32.EnumWindows(WNDENUMPROC(cb), 0)
    return hwnds


def _try_register_global(mod_val: int, vk_val: int) -> bool:
    """Tenta registrar o hotkey globalmente (hwnd=None). True = disponível."""
    ok = bool(user32.RegisterHotKey(None, PROBE_HOTKEY_ID, mod_val | MOD_NOREPEAT, vk_val))
    if ok:
        user32.UnregisterHotKey(None, PROBE_HOTKEY_ID)
    return ok


def _try_register_on_hwnd(hwnd: int, mod_val: int, vk_val: int) -> bool:
    """
    Tenta registrar o hotkey vinculado ao hwnd.
    Retorna True se CONSEGUIU registrar (o atalho NÃO estava preso neste hwnd).
    Retorna False se FALHOU (o atalho pode estar preso neste hwnd).
    """
    ok = bool(user32.RegisterHotKey(hwnd, PROBE_HOTKEY_ID, mod_val | MOD_NOREPEAT, vk_val))
    if ok:
        user32.UnregisterHotKey(hwnd, PROBE_HOTKEY_ID)
    return ok


# ─── Técnica A: Sondagem por hwnd ────────────────────────────────────────────

def _probe_by_hwnd_register(
    mod_val: int, vk_val: int,
    hwnds: List[int],
    progress_cb: Optional[Callable] = None,
    stop_event=None
) -> Optional[ProbeResult]:
    """
    Para cada hwnd, tenta RegisterHotKey(hwnd, ...).
    Se falhar especificamente naquele hwnd enquanto falhou globalmente,
    é sinal de que o atalho está registrado nesse hwnd.
    """
    log.debug(f"[Técnica A] Sondando {len(hwnds)} janelas via hwnd_register...")
    suspects = []

    for i, hwnd in enumerate(hwnds):
        if stop_event and stop_event.is_set():
            break
        if progress_cb:
            progress_cb(i, len(hwnds), _get_window_title(hwnd) or f"hwnd={hwnd:#x}")

        # Ignora janelas do próprio processo
        if _get_pid_from_hwnd(hwnd) == _OWN_PID:
            continue

        ok_on_hwnd = _try_register_on_hwnd(hwnd, mod_val, vk_val)

        if not ok_on_hwnd:
            # Falhou neste hwnd — pode estar registrado aqui
            pid   = _get_pid_from_hwnd(hwnd)
            title = _get_window_title(hwnd)
            info  = _get_proc_info(pid)
            log.info(f"[Técnica A] ❗ Falhou no hwnd {hwnd:#x} | '{title}' | {info['name']} (PID {pid})")
            suspects.append((hwnd, pid, title, info))

    if len(suspects) == 1:
        hwnd, pid, title, info = suspects[0]
        return ProbeResult(
            found=True, method="hwnd_register",
            hwnd=hwnd, pid=pid, name=info["name"],
            title=title, path=info["path"],
            confidence="high",
            notes=["Única janela que recusou RegisterHotKey com hwnd próprio."]
        )
    elif len(suspects) > 1:
        # Múltiplos suspeitos — retorna o primeiro com confiança média
        hwnd, pid, title, info = suspects[0]
        names = [f"{s[3]['name']} ({s[2][:30]})" for s in suspects]
        return ProbeResult(
            found=True, method="hwnd_register",
            hwnd=hwnd, pid=pid, name=info["name"],
            title=title, path=info["path"],
            confidence="medium",
            notes=[f"Múltiplos suspeitos ({len(suspects)}): " + "; ".join(names)]
        )
    return None


# ─── Técnica B: Liberação ao mudar foco ──────────────────────────────────────

def _probe_by_focus_release(
    mod_val: int, vk_val: int,
    hwnds: List[int],
    progress_cb: Optional[Callable] = None,
    stop_event=None
) -> Optional[ProbeResult]:
    """
    Traz cada janela para o foreground e verifica se o atalho global
    fica disponível. Alguns apps liberam o atalho ao perder o foco
    (ex: Discord, Teams). Invasivo — muda o foco da tela.
    """
    log.debug(f"[Técnica B] Sondando {len(hwnds)} janelas via focus_release...")

    original_fg = user32.GetForegroundWindow()

    for i, hwnd in enumerate(hwnds):
        if stop_event and stop_event.is_set():
            break
        title = _get_window_title(hwnd)
        if progress_cb:
            progress_cb(i, len(hwnds), title or f"hwnd={hwnd:#x}")

        if _get_pid_from_hwnd(hwnd) == _OWN_PID:
            continue

        # Traz janela para frente
        user32.SetForegroundWindow(hwnd)
        time.sleep(0.15)

        now_avail = _try_register_global(mod_val, vk_val)
        if now_avail:
            pid  = _get_pid_from_hwnd(hwnd)
            info = _get_proc_info(pid)
            log.info(f"[Técnica B] ✅ Atalho LIBEROU ao focar '{title}' | {info['name']} (PID {pid})")
            # Restaura foco
            try:
                user32.SetForegroundWindow(original_fg)
            except Exception:
                pass
            return ProbeResult(
                found=True, method="focus_release",
                hwnd=hwnd, pid=pid, name=info["name"],
                title=title, path=info["path"],
                confidence="medium",
                notes=["Atalho ficou disponível ao trazer esta janela para o foreground.",
                       "Indica app que segura o atalho apenas quando em background."]
            )

    # Restaura foco
    try:
        user32.SetForegroundWindow(original_fg)
    except Exception:
        pass
    return None


# ─── Técnica C: WM_HOTKEY broadcast ──────────────────────────────────────────

def _probe_by_wm_hotkey(
    mod_val: int, vk_val: int,
    hwnds: List[int],
    progress_cb: Optional[Callable] = None,
    stop_event=None
) -> Optional[ProbeResult]:
    """
    Envia WM_HOTKEY para cada janela e observa qual "processa" a mensagem
    verificando mudança de estado da janela (ativação, resposta).
    Técnica heurística — confiança baixa, útil como último recurso.
    """
    log.debug(f"[Técnica C] Sondando {len(hwnds)} janelas via WM_HOTKEY post...")

    # lparam: bits 0-15 = mod flags, bits 16-31 = vk
    lparam = (vk_val << 16) | (mod_val & 0xFFFF)
    reactions = []

    for i, hwnd in enumerate(hwnds):
        if stop_event and stop_event.is_set():
            break
        title = _get_window_title(hwnd)
        if progress_cb:
            progress_cb(i, len(hwnds), title or f"hwnd={hwnd:#x}")

        if _get_pid_from_hwnd(hwnd) == _OWN_PID:
            continue

        # Captura estado antes
        placement_before = ctypes.create_string_buffer(44)
        user32.GetWindowPlacement(hwnd, placement_before)

        # Envia mensagem simulando WM_HOTKEY
        user32.PostMessageW(hwnd, WM_HOTKEY, PROBE_HOTKEY_ID, lparam)
        time.sleep(0.08)

        # Verifica se houve mudança (janela ficou ativa, ou mudou placement)
        placement_after = ctypes.create_string_buffer(44)
        user32.GetWindowPlacement(hwnd, placement_after)

        changed = (placement_before.raw != placement_after.raw)
        is_fg   = (user32.GetForegroundWindow() == hwnd)

        if changed or is_fg:
            pid  = _get_pid_from_hwnd(hwnd)
            info = _get_proc_info(pid)
            log.info(f"[Técnica C] ⚡ Reação de '{title}' | {info['name']} (PID {pid}) | changed={changed} fg={is_fg}")
            reactions.append((hwnd, pid, title, info, changed, is_fg))

    if reactions:
        # Prioriza quem ficou em foreground
        reactions.sort(key=lambda x: (x[5], x[4]), reverse=True)
        hwnd, pid, title, info, changed, is_fg = reactions[0]
        return ProbeResult(
            found=True, method="wm_hotkey",
            hwnd=hwnd, pid=pid, name=info["name"],
            title=title, path=info["path"],
            confidence="low",
            notes=[
                "Janela reagiu ao receber WM_HOTKEY simulado.",
                f"Ficou em foreground: {is_fg} | Placement mudou: {changed}",
                "⚠ Técnica heurística — pode ter falsos positivos."
            ]
        )
    return None


# ─── Técnica D: Fechar e Reabrir ─────────────────────────────────────────────

# Processos que NUNCA devem ser encerrados
KILL_BLACKLIST = {
    "system", "idle", "lsass.exe", "winlogon.exe", "csrss.exe", "smss.exe",
    "services.exe", "svchost.exe", "explorer.exe", "dwm.exe", "wininit.exe",
    "fontdrvhost.exe", "spoolsv.exe", "taskhostw.exe", "sihost.exe",
    "ctfmon.exe", "runtimebroker.exe", "searchhost.exe",
    "startmenuexperiencehost.exe", "shellexperiencehost.exe",
    "securityhealthservice.exe", "mssense.exe", "antimalware",
    "python.exe", "pythonw.exe",   # evita matar a si mesmo
}

PROCESS_TERMINATE = 0x0001


def _is_store_app(path: str) -> bool:
    """Detecta se é um app UWP/Store pelo caminho."""
    p = path.lower()
    return "windowsapps" in p or "program files\\windowsapps" in p


def _kill_and_check(
    pid: int, proc_name: str, proc_path: str,
    mod_val: int, vk_val: int,
    reopen: bool = True
) -> dict:
    """
    Encerra o processo, verifica se o atalho ficou livre e tenta reabrir.
    Retorna dict com: freed (bool), reopened (bool), error (str|None)
    """
    result = {"freed": False, "reopened": False, "error": None, "is_store": False}

    name_lower = proc_name.lower()
    if name_lower in KILL_BLACKLIST:
        result["error"] = "BLACKLISTED"
        return result

    if _is_store_app(proc_path):
        result["is_store"] = True
        # Para Store apps tentamos via taskkill /PID ao invés de TerminateProcess
        import subprocess
        try:
            subprocess.run(["taskkill", "/PID", str(pid), "/F"],
                           capture_output=True, timeout=5)
        except Exception as e:
            result["error"] = f"taskkill falhou: {e}"
            return result
    else:
        hproc = kernel32.OpenProcess(PROCESS_TERMINATE, False, pid)
        if not hproc:
            result["error"] = "ACCESS_DENIED — não foi possível abrir o processo para encerrar"
            return result
        killed = bool(kernel32.TerminateProcess(hproc, 0))
        kernel32.CloseHandle(hproc)
        if not killed:
            result["error"] = f"TerminateProcess retornou False (errno={ctypes.get_last_error()})"
            return result

    time.sleep(0.4)  # aguarda kernel liberar os handles do processo

    result["freed"] = _try_register_global(mod_val, vk_val)

    if reopen and proc_path and proc_path != "?" and not result["is_store"]:
        try:
            import subprocess
            subprocess.Popen([proc_path], shell=False,
                             creationflags=subprocess.DETACHED_PROCESS |
                                           subprocess.CREATE_NEW_PROCESS_GROUP)
            result["reopened"] = True
            log.info(f"[Técnica D] Reaberto: {proc_path}")
        except Exception as e:
            log.warning(f"[Técnica D] Falha ao reabrir '{proc_path}': {e}")
            result["reopened"] = False

    return result


def probe_by_kill_reopen(
    mod_val: int, vk_val: int,
    proc_list: list,
    reopen: bool = True,
    progress_cb=None,
    stop_event=None
) -> ProbeResult:
    """
    Técnica D: encerra processos um a um e verifica se o atalho libera.
    proc_list: lista de dicts com pid, name, path, title (mesmo formato de get_windows_and_processes).
    """
    log.debug(f"[Técnica D] Iniciando kill_reopen em {len(proc_list)} processos...")

    for i, proc_info in enumerate(proc_list):
        if stop_event and stop_event.is_set():
            return ProbeResult(found=False, method="none", notes=["Interrompido pelo usuário."])

        pid   = proc_info.get("pid", 0)
        name  = proc_info.get("name", "?")
        path  = proc_info.get("path", "?")
        title = proc_info.get("title", "")

        if progress_cb:
            progress_cb(i, len(proc_list), name)

        log.debug(f"[Técnica D] Testando PID {pid} — {name}")
        r = _kill_and_check(pid, name, path, mod_val, vk_val, reopen=reopen)

        if r.get("error") == "BLACKLISTED":
            log.debug(f"[Técnica D] Ignorado (blacklist): {name}")
            continue

        if r.get("error"):
            log.warning(f"[Técnica D] Erro em {name} (PID {pid}): {r['error']}")
            continue

        if r["freed"]:
            notes = [f"Atalho ficou livre após encerrar '{name}'."]
            if r["reopened"]:
                notes.append("App reaberto automaticamente.")
            elif r["is_store"]:
                notes.append("⚠ App da Store — reabra manualmente pela lista de apps.")
            else:
                notes.append("⚠ Não foi possível reabrir automaticamente — verifique manualmente.")
            log.info(f"[Técnica D] ✅ Culpado: {name} (PID {pid}) | reaberto={r['reopened']}")
            return ProbeResult(
                found=True, method="kill_reopen",
                hwnd=0, pid=pid, name=name,
                title=title, path=path,
                confidence="high",
                notes=notes
            )

    return ProbeResult(
        found=False, method="none",
        notes=[
            "Nenhum processo liberou o atalho ao ser encerrado.",
            "Pode ser um driver de kernel, serviço sem janela, ou o próprio Windows."
        ]
    )

def probe_hotkey_owner(
    mod_val: int,
    vk_val:  int,
    use_focus_technique: bool = False,   # Técnica B — invasiva (muda foco)
    use_wm_technique:    bool = False,   # Técnica C — heurística
    progress_cb: Optional[Callable] = None,
    stop_event=None
) -> ProbeResult:
    """
    Tenta descobrir o dono de um hotkey global usando sondagem de janelas.

    Parâmetros:
      mod_val             : valor do modificador (ex: 0x0002 para Ctrl)
      vk_val              : código virtual da tecla
      use_focus_technique : ativa Técnica B (muda foco da janela — visível ao usuário)
      use_wm_technique    : ativa Técnica C (envia WM_HOTKEY — heurístico)
      progress_cb         : callable(current, total, label) para atualizar UI
      stop_event          : threading.Event para interrupção

    Retorna ProbeResult com found=False se nenhuma técnica identificou.
    """

    # Verifica se o atalho realmente está ocupado
    if _try_register_global(mod_val, vk_val):
        log.info("probe_hotkey_owner: atalho está LIVRE — nada a sondar.")
        return ProbeResult(
            found=False, method="none",
            notes=["Atalho não está ocupado — RegisterHotKey global teve sucesso."]
        )

    hwnds = _enum_top_windows()
    log.debug(f"probe_hotkey_owner: {len(hwnds)} janelas encontradas para sondar.")

    # ── Técnica A (sempre ativa) ──────────────────────────────────────────────
    result = _probe_by_hwnd_register(mod_val, vk_val, hwnds, progress_cb, stop_event)
    if result and result.confidence == "high":
        log.info(f"[A] Identificado com alta confiança: {result.name} (PID {result.pid})")
        return result

    # ── Técnica B (opcional — muda foco) ─────────────────────────────────────
    if use_focus_technique and (not result or result.confidence != "high"):
        if stop_event and stop_event.is_set():
            return result or ProbeResult(found=False, method="none", notes=["Interrompido."])
        result_b = _probe_by_focus_release(mod_val, vk_val, hwnds, progress_cb, stop_event)
        if result_b:
            # B tem prioridade sobre A com média confiança
            return result_b

    # ── Técnica C (opcional — heurística) ────────────────────────────────────
    if use_wm_technique and (not result):
        if stop_event and stop_event.is_set():
            return ProbeResult(found=False, method="none", notes=["Interrompido."])
        result_c = _probe_by_wm_hotkey(mod_val, vk_val, hwnds, progress_cb, stop_event)
        if result_c:
            return result_c

    # Retorna A (confiança média) ou "não encontrado"
    if result:
        return result

    return ProbeResult(
        found=False, method="none",
        notes=[
            "Nenhuma técnica identificou o responsável.",
            "Possíveis causas: serviço do sistema sem janela, driver de kernel,",
            "ou app que registrou o atalho em thread sem hwnd visível.",
            "Tente o Teste por Eliminação (suspensão de processos) na aba acima."
        ]
    )
