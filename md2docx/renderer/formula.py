"""LaTeX formula renderer using matplotlib and CodeCogs API."""

import re
import io
from typing import Optional, Dict

import requests


def is_complex_formula(latex: str) -> bool:
    """Check if formula uses complex environments."""
    complex_envs = [
        r'\begin{cases}',
        r'\begin{align}',
        r'\begin{split}',
        r'\begin{gather}',
        r'\begin{eqnarray}',
        r'\begin{matrix}',
        r'\begin{array}'
    ]
    return any(env in latex for env in complex_envs)


class FormulaRenderer:
    """Renders LaTeX formulas to PNG images."""

    CODECOGS_API = "https://latex.codecogs.com/png.latex"

    def __init__(self, timeout: int = 25, chinese_replacements: Dict[str, str] = None):
        self.timeout = timeout
        self.chinese_replacements = chinese_replacements or {}

    def _remove_tag(self, latex: str) -> str:
        """Remove \\tag{} from formula."""
        return re.sub(r'\\tag\{[^}]*\}', '', latex).strip()

    def _replace_chinese(self, latex: str) -> str:
        """Replace Chinese characters in formula."""
        for ch, en in self.chinese_replacements.items():
            latex = latex.replace(f"\\text{{{ch}}}", f"\\mathrm{{{en}}}")
            if ch in latex and '\\text{' not in latex:
                latex = latex.replace(ch, en)
        return latex

    def render(self, latex: str) -> Optional[bytes]:
        """Render formula to PNG."""
        latex = self._remove_tag(latex)
        latex = self._replace_chinese(latex)

        if is_complex_formula(latex):
            return self.render_codecogs(latex)
        else:
            result = self.render_matplotlib(latex)
            if result is None:
                return self.render_codecogs(latex)
            return result

    def render_matplotlib(self, latex: str) -> Optional[bytes]:
        """Render formula using matplotlib."""
        try:
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt

            fig = plt.figure(figsize=(10, 2))
            ax = fig.add_subplot(111)
            ax.axis('off')

            if not latex.strip().startswith('$'):
                latex = f"${latex}$"

            ax.text(0.5, 0.5, latex, ha='center', va='center',
                   fontsize=14, transform=ax.transAxes)

            buf = io.BytesIO()
            fig.savefig(buf, format='png', dpi=150,
                       bbox_inches='tight', pad_inches=0.1,
                       transparent=False, facecolor='white')
            plt.close(fig)

            buf.seek(0)
            return buf.read()

        except Exception:
            return None

    def render_codecogs(self, latex: str) -> Optional[bytes]:
        """Render formula using CodeCogs API."""
        try:
            params = {r'\inline': '', 'latex': latex}
            response = requests.get(
                self.CODECOGS_API,
                params=params,
                timeout=self.timeout
            )

            if response.status_code == 200 and len(response.content) > 100:
                return response.content

        except requests.exceptions.RequestException:
            pass

        return None
