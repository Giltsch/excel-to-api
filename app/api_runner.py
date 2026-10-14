"""Manages a FastAPI subprocess for live preview."""

import os
import signal
import socket
import subprocess
import sys
import time
from typing import Optional

import requests


class APIRunner:
    """Manages a FastAPI subprocess for live preview."""

    def __init__(self, project_dir: str, port: int = 8100):
        self.project_dir = project_dir
        self.port = port
        self.process: Optional[subprocess.Popen] = None

    def start(self) -> bool:
        """Start the FastAPI server. Returns True if server is healthy."""
        self.stop()

        if not self._port_available():
            self._kill_port_occupant()

        # Install dependencies
        req_file = os.path.join(self.project_dir, "requirements.txt")
        if os.path.exists(req_file):
            subprocess.run(
                [sys.executable, "-m", "pip", "install", "-q", "-r", req_file],
                cwd=self.project_dir,
                timeout=120,
                capture_output=True,
            )

        # Seed the database
        seed_script = os.path.join(self.project_dir, "seed_data.py")
        if os.path.exists(seed_script):
            subprocess.run(
                [sys.executable, seed_script],
                cwd=self.project_dir,
                timeout=30,
                capture_output=True,
            )

        # Launch uvicorn
        self.process = subprocess.Popen(
            [
                sys.executable,
                "-m",
                "uvicorn",
                "main:app",
                "--host",
                "0.0.0.0",
                "--port",
                str(self.port),
                "--log-level",
                "warning",
            ],
            cwd=self.project_dir,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )

        return self._wait_for_healthy(timeout=15)

    def stop(self) -> None:
        """Stop the running FastAPI subprocess."""
        if self.process is not None:
            self.process.terminate()
            try:
                self.process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                self.process.kill()
                self.process.wait()
            self.process = None

    def is_running(self) -> bool:
        if self.process is None:
            return False
        return self.process.poll() is None

    def get_url(self) -> str:
        return f"http://127.0.0.1:{self.port}"

    def get_docs_url(self) -> str:
        return f"{self.get_url()}/docs"

    def get_stderr(self) -> str:
        if self.process and self.process.stderr:
            try:
                return self.process.stderr.read().decode("utf-8", errors="replace")
            except Exception:
                return ""
        return ""

    def _wait_for_healthy(self, timeout: int = 15) -> bool:
        start = time.time()
        while time.time() - start < timeout:
            try:
                resp = requests.get(f"{self.get_url()}/docs", timeout=1)
                if resp.status_code == 200:
                    return True
            except requests.ConnectionError:
                pass
            if self.process and self.process.poll() is not None:
                return False
            time.sleep(0.5)
        return False

    def _port_available(self) -> bool:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            return s.connect_ex(("127.0.0.1", self.port)) != 0

    def _kill_port_occupant(self) -> None:
        try:
            result = subprocess.run(
                ["lsof", "-ti", f":{self.port}"],
                capture_output=True,
                text=True,
            )
            if result.stdout.strip():
                for pid in result.stdout.strip().split("\n"):
                    os.kill(int(pid), signal.SIGTERM)
                time.sleep(1)
        except Exception:
            pass
