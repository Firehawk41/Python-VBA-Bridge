"""soffice subprocess lifecycle: launch, port/profile allocation, shutdown.

soffice's launcher (oosplash) forks the real soffice.bin off as a detached
child and then exits on its own once startup/handoff completes -- it is NOT
a stand-in for soffice.bin's lifetime. That means the PID subprocess.Popen
hands back cannot be trusted for either liveness checks or termination once
more than a couple seconds have passed (confirmed empirically: oosplash was
already gone, `poll()`-based liveness silently false, by the time a session
had created a document and run one macro). Instead we identify the real
process by a marker unique to this instance -- the profile directory baked
into its command line via -env:UserInstallation= -- scanned directly out of
/proc, and track/signal by that.
"""

import os
import shutil
import signal
import socket
import subprocess
import tempfile
import time


def find_free_port() -> int:
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("localhost", 0))
        return s.getsockname()[1]
    finally:
        s.close()


def _pids_matching(marker: bytes) -> list:
    pids = []
    try:
        entries = os.listdir("/proc")
    except OSError:
        return pids
    for entry in entries:
        if not entry.isdigit():
            continue
        try:
            with open(f"/proc/{entry}/cmdline", "rb") as f:
                cmdline = f.read()
        except OSError:
            continue
        if marker in cmdline:
            pids.append(int(entry))
    return pids


def _pid_alive(pid: int) -> bool:
    try:
        os.kill(pid, 0)
        return True
    except ProcessLookupError:
        return False
    except PermissionError:
        return True  # exists, just not ours to signal-check further


class SofficeProcess:
    def __init__(self, soffice_binary: str = "soffice"):
        self.soffice_binary = soffice_binary
        self.port = None
        self.profile_dir = None
        self._proc = None

    def launch(self) -> int:
        self.port = find_free_port()
        self.profile_dir = tempfile.mkdtemp(prefix="vba_bridge_profile_")
        cmd = [
            self.soffice_binary,
            "--headless",
            "--invisible",
            "--nocrashreport",
            "--nodefault",
            "--norestore",
            "--nologo",
            "--nofirststartwizard",
            f"--accept=socket,host=localhost,port={self.port};urp;",
            f"-env:UserInstallation=file://{self.profile_dir}",
        ]
        self._proc = subprocess.Popen(
            cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        return self.port

    def _marker(self) -> bytes:
        return f"UserInstallation=file://{self.profile_dir}".encode()

    def _matching_pids(self) -> list:
        return _pids_matching(self._marker())

    @property
    def is_running(self) -> bool:
        if self._proc is None:
            return False
        # Don't trust poll() alone: oosplash (the tracked PID) legitimately
        # exits on its own after handing off to soffice.bin. A live process
        # matching our profile-dir marker means the real instance is up,
        # even once the originally-launched PID is gone.
        if self._proc.poll() is None:
            return True
        return bool(self._matching_pids())

    def terminate(self, timeout: float = 5.0) -> None:
        if self._proc is None:
            return

        def alive_targets():
            return set(self._matching_pids()) | (
                {self._proc.pid} if _pid_alive(self._proc.pid) else set()
            )

        targets = alive_targets()
        for pid in targets:
            self._signal(pid, signal.SIGTERM)

        deadline = time.monotonic() + timeout
        remaining = {p for p in targets if _pid_alive(p)}
        while remaining and time.monotonic() < deadline:
            time.sleep(0.2)
            remaining = {p for p in remaining if _pid_alive(p)}

        if remaining:
            for pid in remaining:
                self._signal(pid, signal.SIGKILL)
            deadline = time.monotonic() + timeout
            while remaining and time.monotonic() < deadline:
                time.sleep(0.2)
                remaining = {p for p in remaining if _pid_alive(p)}

        try:
            self._proc.wait(timeout=1.0)
        except subprocess.TimeoutExpired:
            pass

        if self.profile_dir:
            shutil.rmtree(self.profile_dir, ignore_errors=True)
        self._proc = None

    @staticmethod
    def _signal(pid: int, sig: int) -> None:
        try:
            os.kill(pid, sig)
        except ProcessLookupError:
            pass
