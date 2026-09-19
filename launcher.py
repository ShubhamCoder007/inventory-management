"""Streamlit Windows Launcher script.

Launches the Streamlit application as a background server tied to the
lifetime of the launcher window/process, automatically picking an open port
and opening the browser when ready.
"""

import ctypes
from ctypes import wintypes
import os
from pathlib import Path
import socket
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser

# ---------------------------------------------------------------------------
# Configuration & Constants
# ---------------------------------------------------------------------------
APP_FILE = r"C:\Users\shubh\Desktop\softwares\bms\inventory_manager_v3.py"
PREFERRED_PORT = 8501
STARTUP_TIMEOUT = 90  # seconds to wait for the server to accept connections

JOB_OBJECT_EXTENDED_LIMIT_INFORMATION = 9
JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x00002000
PROCESS_TERMINATE = 0x0001
PROCESS_SET_QUOTA = 0x0100
SYNCHRONIZE = 0x00100000
INFINITE = 0xFFFFFFFF

# Holds the job handle open for as long as this process lives. Windows destroys
# the job (and everything in it) once the last handle closes, which is what
# stops the server from outliving the launcher.
_job_handle = None


# ---------------------------------------------------------------------------
# Ctypes Structures for Windows Job Object Binding
# ---------------------------------------------------------------------------
class _BASIC_LIMITS(ctypes.Structure):
    _fields_ = [
        ("PerProcessUserTimeLimit", ctypes.c_int64),
        ("PerJobUserTimeLimit", ctypes.c_int64),
        ("LimitFlags", wintypes.DWORD),
        ("MinimumWorkingSetSize", ctypes.c_size_t),
        ("MaximumWorkingSetSize", ctypes.c_size_t),
        ("ActiveProcessLimit", wintypes.DWORD),
        ("Affinity", ctypes.c_size_t),
        ("PriorityClass", wintypes.DWORD),
        ("SchedulingClass", wintypes.DWORD),
    ]


class _IO_COUNTERS(ctypes.Structure):
    _fields_ = [
        ("ReadOperationCount", ctypes.c_uint64),
        ("WriteOperationCount", ctypes.c_uint64),
        ("OtherOperationCount", ctypes.c_uint64),
        ("ReadTransferCount", ctypes.c_uint64),
        ("WriteTransferCount", ctypes.c_uint64),
        ("OtherTransferCount", ctypes.c_uint64),
    ]


class _EXTENDED_LIMITS(ctypes.Structure):
    _fields_ = [
        ("BasicLimitInformation", _BASIC_LIMITS),
        ("IoInfo", _IO_COUNTERS),
        ("ProcessMemoryLimit", ctypes.c_size_t),
        ("JobMemoryLimit", ctypes.c_size_t),
        ("PeakProcessMemoryUsed", ctypes.c_size_t),
        ("PeakJobMemoryUsed", ctypes.c_size_t),
    ]


# ---------------------------------------------------------------------------
# Lifetime Management Functions
# ---------------------------------------------------------------------------
def bind_to_lifetime(pid: int) -> bool:
    """Tie the server process to this launcher's lifetime.

    Without this, killing the launcher (closing the window, End task) would
    leave the Streamlit server running and holding the port.
    """
    global _job_handle
    try:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    except (AttributeError, OSError):
        return False

    kernel32.CreateJobObjectW.restype = wintypes.HANDLE
    kernel32.OpenProcess.restype = wintypes.HANDLE

    job = kernel32.CreateJobObjectW(None, None)
    if not job:
        return False

    limits = _EXTENDED_LIMITS()
    limits.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    if not kernel32.SetInformationJobObject(
        job,
        JOB_OBJECT_EXTENDED_LIMIT_INFORMATION,
        ctypes.byref(limits),
        ctypes.sizeof(limits),
    ):
        kernel32.CloseHandle(job)
        return False

    process = kernel32.OpenProcess(PROCESS_SET_QUOTA | PROCESS_TERMINATE, False, pid)
    if not process:
        kernel32.CloseHandle(job)
        return False

    assigned = bool(kernel32.AssignProcessToJobObject(job, process))
    kernel32.CloseHandle(process)
    if not assigned:
        kernel32.CloseHandle(job)
        return False

    _job_handle = job
    return True


def exit_with_bootloader(server: subprocess.Popen) -> None:
    """Exit if the PyInstaller bootloader that wrapped us is killed.

    A one-file build runs this code in a child of the bootloader process. Only
    the bootloader is visible as "Inventory Manager.exe", so killing it would
    otherwise leave this process, its job, and the server running.
    """
    if not getattr(sys, "frozen", False):
        return

    try:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.OpenProcess.restype = wintypes.HANDLE
        parent = kernel32.OpenProcess(SYNCHRONIZE, False, os.getppid())
    except (AttributeError, OSError):
        return
    if not parent:
        return

    def wait_then_quit() -> None:
        kernel32.WaitForSingleObject(parent, INFINITE)
        if server.poll() is None:
            server.kill()
        os._exit(1)

    threading.Thread(target=wait_then_quit, daemon=True).start()


# ---------------------------------------------------------------------------
# Utility & Environment Helpers
# ---------------------------------------------------------------------------
def app_dir() -> Path:
    """The folder holding app.py: next to the exe, or next to this script."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def fail(message: str) -> None:
    """Print an error message banner and pause before exiting."""
    print("\n" + "=" * 70)
    print(message)
    print("=" * 70)
    input("\nPress Enter to close this window... ")
    sys.exit(1)


def has_streamlit(python_exe: str) -> bool:
    """Check whether a given Python environment has Streamlit installed."""
    try:
        return (
            subprocess.run(
                [python_exe, "-c", "import streamlit"],
                capture_output=True,
                timeout=60,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            ).returncode
            == 0
        )
    except (OSError, subprocess.SubprocessError):
        return False


def find_python() -> str:
    """Locate a Python executable that has Streamlit installed."""
    if has_streamlit(sys.executable):
        return sys.executable

    for cmd in ["python", "python3", "py"]:
        try:
            path = subprocess.check_output(
                [cmd, "-c", "import sys; print(sys.executable)"],
                text=True,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            ).strip()
            if path and has_streamlit(path):
                return path
        except (OSError, subprocess.SubprocessError):
            continue

    return ""


def pick_port() -> int:
    """Find an available TCP port starting from PREFERRED_PORT."""
    for port in range(PREFERRED_PORT, PREFERRED_PORT + 100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            if sock.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return PREFERRED_PORT


def wait_for_server(port: int) -> bool:
    """Wait until the Streamlit server accepts HTTP connections."""
    url = f"http://127.0.0.1:{port}"
    start = time.time()
    while time.time() - start < STARTUP_TIMEOUT:
        try:
            with urllib.request.urlopen(url, timeout=1) as resp:
                if resp.status == 200:
                    return True
        except Exception:
            pass
        time.sleep(0.5)
    return False


# ---------------------------------------------------------------------------
# Main Launcher Workflow
# ---------------------------------------------------------------------------
def main() -> None:
    """Launch the Streamlit server and open the browser interface."""
    target_dir = app_dir()
    app_path = target_dir / APP_FILE

    if not app_path.exists():
        fail(f"Could not find {APP_FILE} in {target_dir}")

    python_exe = find_python()
    if not python_exe:
        fail(
            "Could not find a Python environment with Streamlit installed.\n"
            "Please ensure Streamlit is installed: pip install streamlit"
        )

    port = pick_port()
    cmd = [
        python_exe,
        "-m",
        "streamlit",
        "run",
        str(app_path),
        "--server.port",
        str(port),
        "--server.headless",
        "true",
        "--global.developmentMode",
        "false",
    ]

    try:
        server = subprocess.Popen(
            cmd,
            cwd=str(target_dir),
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception as exc:
        fail(f"Failed to start Streamlit server: {exc}")

    bind_to_lifetime(server.pid)
    exit_with_bootloader(server)

    url = f"http://127.0.0.1:{port}"
    print(f"Starting Streamlit server on {url}...")

    if wait_for_server(port):
        webbrowser.open(url)
        print("Server is running. Press Ctrl+C in this console to stop.")
        try:
            server.wait()
        except KeyboardInterrupt:
            server.terminate()
    else:
        server.terminate()
        fail("Streamlit server failed to start within the timeout period.")


if __name__ == "__main__":
    main()
