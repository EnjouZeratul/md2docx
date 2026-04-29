"""Mermaid diagram renderer using mermaid.ink API."""

import base64
import json
import zlib
import time
from typing import Optional

import requests


class MermaidRenderError(Exception):
    """Mermaid rendering error."""
    pass


class MermaidRenderer:
    """Renders Mermaid diagrams via mermaid.ink API."""

    API_BASE = "https://mermaid.ink/img/pako:"

    def __init__(self, timeout: int = 25):
        """Initialize renderer.

        Args:
            timeout: Request timeout in seconds
        """
        self.timeout = timeout

    def _encode_mermaid(self, code: str) -> str:
        """Encode mermaid code for API URL.

        Uses pako compression + base64 encoding.

        Args:
            code: Mermaid diagram code

        Returns:
            URL-safe base64 encoded string
        """
        # Create JSON config
        config = {
            "code": code,
            "mermaid": {"theme": "default"}
        }
        json_str = json.dumps(config, separators=(',', ':'))

        # Compress with zlib
        compressed = zlib.compress(json_str.encode('utf-8'), level=9)

        # Base64 encode (URL-safe)
        encoded = base64.urlsafe_b64encode(compressed).decode('ascii')

        return encoded

    def render(self, code: str, retries: int = 1) -> Optional[bytes]:
        """Render mermaid diagram to PNG.

        Args:
            code: Mermaid diagram code
            retries: Number of retries on failure

        Returns:
            PNG bytes if successful, None if failed
        """
        encoded = self._encode_mermaid(code)
        url = self.API_BASE + encoded

        for attempt in range(retries + 1):
            try:
                response = requests.get(url, timeout=self.timeout)

                if response.status_code == 200:
                    # Check content size (should be > 100 bytes for valid PNG)
                    if len(response.content) > 100:
                        return response.content
                    else:
                        # Too small, likely an error
                        continue

                elif response.status_code == 429:
                    # Rate limited - wait and retry
                    if attempt < retries:
                        time.sleep(2)
                        continue
                    return None

                elif 400 <= response.status_code < 500:
                    # Client error - don't retry
                    return None

                elif response.status_code >= 500:
                    # Server error - retry
                    if attempt < retries:
                        continue
                    return None

            except requests.exceptions.Timeout:
                if attempt < retries:
                    continue
                return None

            except requests.exceptions.RequestException:
                if attempt < retries:
                    continue
                return None

        return None

    def is_valid(self, code: str) -> bool:
        """Check if mermaid code appears valid.

        Does a basic syntax check.

        Args:
            code: Mermaid diagram code

        Returns:
            True if code appears valid
        """
        if not code or not code.strip():
            return False

        # Check for common mermaid diagram types
        valid_starts = [
            'flowchart', 'graph', 'sequenceDiagram', 'classDiagram',
            'stateDiagram', 'erDiagram', 'gantt', 'pie', 'gitgraph',
            'mindmap', 'timeline', 'quadrantChart'
        ]

        first_line = code.strip().split('\n')[0].strip()
        return any(first_line.startswith(start) for start in valid_starts)
