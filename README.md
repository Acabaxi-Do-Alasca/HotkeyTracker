# 🔍 HotkeyTracker Pro

Uma ferramenta desktop para **Windows** que detecta atalhos de teclado globais em uso e descobre **qual aplicativo está bloqueando** cada combinação de teclas.

Ideal para desenvolvedores, gamers e usuários avançados que precisam diagnosticar conflitos de hotkeys entre aplicativos como Discord, Teams, OBS, Spotify, AutoHotkey, entre outros.

---

## ✨ Funcionalidades

- **Verificador Rápido** — Verifica se uma combinação de teclas específica está livre ou ocupada e exibe o provável responsável com histórico
- **Scanner em Massa** — Varre combinações de modificadores (Alt, Ctrl, Shift, Win, etc.) e grupos de teclas (letras, números, F1-F12, Numpad...) e lista todos os atalhos ocupados
- **Processos & Apps** — Lista todos os processos com janelas abertas que podem estar registrando atalhos globais
- **Identificação por múltiplas técnicas:**
  - **Técnica A** — Sondagem por `hwnd` (RegisterHotKey por janela) — rápida e não invasiva
  - **Técnica B** — Liberação ao mudar foco — detecta apps como Discord e Teams
  - **Técnica C** — Envio de `WM_HOTKEY` simulado — heurística como último recurso
  - **Técnica D** — Kill & Reopen — encerra processos um a um e verifica se o atalho libera (requer confirmação)
  - **Teste por Eliminação** — Suspende processos temporariamente para identificar o culpado sem encerrar nada
- **Base de Atalhos Conhecidos** — Banco de dados com atalhos registrados por apps populares (Windows, Teams, Spotify, OBS, Discord, Lightshot, ShareX, 1Password, etc.)
- **Exportação** — Exporta resultados do scan em `.csv` ou `.json`

---

## 🖥️ Requisitos

- **Windows 10/11** (usa Win32 API nativa)
- **Python 3.8+**
- [`psutil`](https://pypi.org/project/psutil/) — instalado automaticamente se ausente

### Instalar dependências manualmente

```bash
pip install psutil
```

---

## 🚀 Como usar

1. Clone o repositório:

```bash
git clone https://github.com/seu-usuario/HotkeyTracker.git
cd HotkeyTracker
```

2. Execute o aplicativo:

```bash
python HotkeyTracker.py
```

> **Recomendado:** Execute como **Administrador** para habilitar o Teste por Eliminação (suspensão de processos) e a Técnica D (Kill & Reopen).

---

## 📁 Estrutura dos arquivos

| Arquivo | Descrição |
|---|---|
| `HotkeyTracker.py` | Aplicação principal com interface gráfica (Tkinter) |
| `hwnd_probe.py` | Módulo de sondagem por janelas (Técnicas A, B, C e D) |
| `test_process_result.py` | Módulo de suspensão de processos para teste por eliminação |

---

## ⚙️ Detalhes técnicos

O projeto utiliza **Win32 API** via `ctypes` para:

- `RegisterHotKey` / `UnregisterHotKey` — verificar disponibilidade de atalhos
- `EnumWindows` / `GetWindowThreadProcessId` — enumerar janelas e seus processos
- `NtSuspendProcess` / `NtResumeProcess` (via psutil) — suspender processos temporariamente
- `PostMessage(WM_HOTKEY)` — sondagem heurística por mensagem
- `SetForegroundWindow` — técnica de troca de foco

O resultado de cada sondagem é encapsulado em um `ProbeResult` com nível de confiança (`high`, `medium`, `low`) e notas explicativas.

---

## ⚠️ Aviso

A **Técnica D** encerra processos para verificar se o atalho é liberado. Use com cautela:
- Salve seu trabalho antes de utilizá-la
- O app tenta reabrir os processos automaticamente após o teste
- Apps da Microsoft Store podem precisar ser reabertos manualmente

---

## 📸 Interface

A interface conta com **5 abas**:

1. **Verificador** — Teste rápido de uma combinação específica
2. **Scanner** — Varredura em massa configurável
3. **Processos & Apps** — Lista de janelas abertas + ferramentas de diagnóstico
4. **Resultados** — Resultados detalhados com filtros e busca
5. **Base de Atalhos** — Banco de dados de atalhos conhecidos pesquisável

---

## 📄 Licença

MIT License — veja o arquivo `LICENSE` para detalhes.
