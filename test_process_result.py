import psutil
import ctypes
import time
import logging
import traceback
from enum import Enum

# ─────────────────────────────────────────────────────
# Resultado tipado — sem ambiguidade
# ─────────────────────────────────────────────────────
class TestResult(Enum):
    GUILTY       = "GUILTY"        # ficou disponível ao suspender → culpado
    INNOCENT     = "INNOCENT"      # ainda bloqueado após suspender → inocente
    ACCESS_DENIED = "ACCESS_DENIED" # sem permissão para suspender → suspeito
    NO_PROCESS   = "NO_PROCESS"    # processo sumiu durante o teste
    TIMEOUT      = "TIMEOUT"       # suspend travou (processo não respondeu)
    ERROR        = "ERROR"         # erro inesperado (com detalhe no log)


# ─────────────────────────────────────────────────────
# Logger dedicado ao teste
# ─────────────────────────────────────────────────────
log = logging.getLogger("hotkey.suspension_test")

def _setup_log(level=logging.DEBUG):
    """Chame uma vez no startup do app para ativar o log."""
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
        datefmt="%H:%M:%S",
    )

# ─────────────────────────────────────────────────────
# Função de teste com log detalhado
# ─────────────────────────────────────────────────────
def test_process_suspension(pid: int, proc_name: str,
                             mod_val: int, vk_val: int,
                             timeout: float = 3.0) -> TestResult:
    """
    Suspende o processo `pid` e verifica se o atalho (mod_val+vk_val) fica livre.
    Retorna um TestResult com log detalhado de cada etapa.
    """
    log.debug(f"[{proc_name} | PID {pid}] Iniciando teste de suspensão...")

    # ── 1. Obter handle do processo ──────────────────
    try:
        p = psutil.Process(pid)
    except psutil.NoSuchProcess:
        log.warning(f"[{proc_name} | PID {pid}] Processo não existe mais (NoSuchProcess).")
        return TestResult.NO_PROCESS
    except Exception as e:
        log.error(f"[{proc_name} | PID {pid}] Erro ao obter processo: {type(e).__name__}: {e}")
        return TestResult.ERROR

    # ── 2. Tentar suspender ──────────────────────────
    try:
        import threading, concurrent.futures
        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as ex:
            future = ex.submit(p.suspend)
            try:
                future.result(timeout=timeout)
                log.debug(f"[{proc_name} | PID {pid}] Suspenso com sucesso.")
            except concurrent.futures.TimeoutError:
                log.warning(f"[{proc_name} | PID {pid}] Timeout ao suspender ({timeout}s). Ignorando.")
                return TestResult.TIMEOUT

    except psutil.AccessDenied as e:
        log.warning(
            f"[{proc_name} | PID {pid}] AccessDenied ao suspender. "
            f"Detalhes: {e} | Admin necessário para apps da Store."
        )
        return TestResult.ACCESS_DENIED

    except psutil.NoSuchProcess:
        log.warning(f"[{proc_name} | PID {pid}] Processo encerrou antes de ser suspenso.")
        return TestResult.NO_PROCESS

    except PermissionError as e:
        # PermissionError é diferente de AccessDenied — processo protegido pelo sistema
        log.warning(f"[{proc_name} | PID {pid}] PermissionError (processo do sistema?): {e}")
        return TestResult.ACCESS_DENIED

    except OSError as e:
        log.error(f"[{proc_name} | PID {pid}] OSError ao suspender: errno={e.errno} | {e.strerror}")
        return TestResult.ERROR

    except Exception as e:
        # Captura qualquer outro erro COM stack trace completo
        log.error(
            f"[{proc_name} | PID {pid}] Erro inesperado ao suspender.\n"
            f"  Tipo : {type(e).__name__}\n"
            f"  Msg  : {e}\n"
            f"  Stack: {traceback.format_exc(limit=3)}"
        )
        return TestResult.ERROR

    # ── 3. Verificar se o atalho ficou livre ────────
    try:
        time.sleep(0.2)
        avail = _check_hotkey_available(mod_val, vk_val)
        log.debug(f"[{proc_name} | PID {pid}] Atalho disponível após suspensão: {avail}")
    except Exception as e:
        log.error(f"[{proc_name} | PID {pid}] Erro ao checar hotkey: {type(e).__name__}: {e}")
        avail = False
    finally:
        # ── 4. Retomar SEMPRE no finally ────────────
        try:
            p.resume()
            log.debug(f"[{proc_name} | PID {pid}] Retomado com sucesso.")
        except psutil.NoSuchProcess:
            log.warning(f"[{proc_name} | PID {pid}] Processo encerrou durante o teste (NoSuchProcess no resume).")
        except psutil.AccessDenied:
            log.error(f"[{proc_name} | PID {pid}] ⚠ NÃO FOI POSSÍVEL RETOMAR! AccessDenied no resume.")
        except Exception as e:
            log.error(f"[{proc_name} | PID {pid}] ⚠ Erro inesperado no resume: {type(e).__name__}: {e}")

    return TestResult.GUILTY if avail else TestResult.INNOCENT


# ─────────────────────────────────────────────────────
# Helper: checar hotkey via RegisterHotKey Win32
# ─────────────────────────────────────────────────────
def _check_hotkey_available(mod_flag: int, vk: int) -> bool:
    TEST_ID = 0xF001
    user32  = ctypes.windll.user32
    ok = bool(user32.RegisterHotKey(None, TEST_ID, mod_flag, vk))
    if ok:
        user32.UnregisterHotKey(None, TEST_ID)
    return ok
