"""Configuration management for md2docx."""

from pathlib import Path
from dataclasses import dataclass, field
from typing import Dict, Optional

import yaml
from docx.shared import Cm, Pt


class ConfigError(Exception):
    """Configuration error."""
    pass


def parse_unit(value: str):
    """Parse unit string to docx unit object.

    Args:
        value: Unit string like "3.0cm" or "12pt"

    Returns:
        Cm or Pt object
    """
    if isinstance(value, (int, float)):
        return Pt(value)

    value = str(value).strip()
    if value.endswith("cm"):
        return Cm(float(value[:-2]))
    elif value.endswith("pt"):
        return Pt(float(value[:-2]))
    else:
        raise ConfigError(f"Invalid unit format: {value}. Use 'cm' or 'pt' suffix.")


@dataclass
class MarginsConfig:
    left: str = "3.0cm"  # cm as string
    right: str = "2.5cm"
    top: str = "2.54cm"
    bottom: str = "2.54cm"


@dataclass
class PageConfig:
    size: str = "A4"
    margins: MarginsConfig = field(default_factory=MarginsConfig)


@dataclass
class FontsConfig:
    default: str = "宋体"
    heading: str = "宋体"
    code: str = "Consolas"
    formula: str = "Times New Roman"


@dataclass
class HeadingStyleConfig:
    size: int = 12
    bold: bool = True
    align: str = "left"
    indent: bool = False


@dataclass
class HeadingsConfig:
    chapter: HeadingStyleConfig = field(default_factory=lambda: HeadingStyleConfig(size=18, bold=True, align="center"))
    section: HeadingStyleConfig = field(default_factory=lambda: HeadingStyleConfig(size=14, bold=True, align="left"))
    subsection: HeadingStyleConfig = field(default_factory=lambda: HeadingStyleConfig(size=12, bold=True, indent=True))


@dataclass
class BodyConfig:
    size: int = 12
    line_spacing: float = 1.5
    first_line_indent: str = "0.74cm"  # cm as string


@dataclass
class FiguresConfig:
    max_width: str = "14cm"  # cm as string
    caption_size: int = 9
    caption_font: str = "宋体"


@dataclass
class TocConfig:
    enabled: bool = True
    levels: str = "1-3"


@dataclass
class CoverConfig:
    enabled: bool = False
    template: str = ""


@dataclass
class AbstractConfig:
    auto_extract: bool = True
    title_size: int = 18
    title_font: str = "宋体"
    body_size: int = 12
    body_font: str = "宋体"


@dataclass
class HeaderFooterConfig:
    header_text: str = ""
    header_font_size: float = 10.5
    page_number_start: int = 1


@dataclass
class FormulaConfig:
    chinese_replacements: Dict[str, str] = field(default_factory=dict)


@dataclass
class TableConfig:
    border_top_weight: float = 1.5
    border_header_weight: float = 1.0
    border_bottom_weight: float = 1.5
    border_color: str = "000000"


@dataclass
class DocumentConfig:
    title_zh: str = ""
    title_en: str = ""
    author: str = ""
    date: str = ""


@dataclass
class Config:
    """Main configuration class."""
    document: DocumentConfig = field(default_factory=DocumentConfig)
    page: PageConfig = field(default_factory=PageConfig)
    fonts: FontsConfig = field(default_factory=FontsConfig)
    headings: HeadingsConfig = field(default_factory=HeadingsConfig)
    body: BodyConfig = field(default_factory=BodyConfig)
    figures: FiguresConfig = field(default_factory=FiguresConfig)
    toc: TocConfig = field(default_factory=TocConfig)
    cover: CoverConfig = field(default_factory=CoverConfig)
    abstract: AbstractConfig = field(default_factory=AbstractConfig)
    header_footer: HeaderFooterConfig = field(default_factory=HeaderFooterConfig)
    formula: FormulaConfig = field(default_factory=FormulaConfig)
    table: TableConfig = field(default_factory=TableConfig)


def _dict_to_config(d: dict) -> Config:
    """Convert dictionary to Config dataclass with proper defaults."""
    # Start with a default config
    config = Config()

    def update_nested(obj, data: dict):
        """Update object attributes from data dict."""
        if data is None:
            return
        for key, value in data.items():
            if hasattr(obj, key):
                current = getattr(obj, key)
                if hasattr(current, '__dataclass_fields__') and isinstance(value, dict):
                    # Recursively update nested dataclass
                    update_nested(current, value)
                else:
                    # Direct assignment for simple values
                    setattr(obj, key, value)

    update_nested(config, d)
    return config


def load_config(config_path: Optional[str] = None) -> Config:
    """Load configuration from file or use default.

    Args:
        config_path: Path to YAML config file. If None, uses internal default.

    Returns:
        Config object

    Raises:
        ConfigError: If config file is invalid
    """
    # Get default config path
    if config_path is None:
        default_path = Path(__file__).parent / "templates" / "default.yaml"
        config_path = str(default_path)

    # Load YAML
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            data = yaml.safe_load(f)
    except yaml.YAMLError as e:
        raise ConfigError(f"YAML parsing error: {e}")
    except FileNotFoundError:
        raise ConfigError(f"Config file not found: {config_path}")

    if data is None:
        data = {}

    # Convert to Config object with defaults
    config = _dict_to_config(data)

    # Validate required fields
    _validate_config(config)

    return config


def _validate_config(config: Config) -> None:
    """Validate configuration values."""
    # Page size must be valid
    valid_sizes = ["A4", "A3", "Letter", "Legal"]
    if config.page.size not in valid_sizes:
        raise ConfigError(f"Invalid page size: {config.page.size}. Must be one of {valid_sizes}")

    # Margins must be positive (parse string values like "3.0cm")
    for field_name in ['left', 'right', 'top', 'bottom']:
        value = getattr(config.page.margins, field_name)
        if isinstance(value, str):
            try:
                value = float(value.replace('cm', '').replace('pt', ''))
            except ValueError:
                raise ConfigError(f"Invalid margin value: {getattr(config.page.margins, field_name)}")
        if value < 0:
            raise ConfigError(f"Page margin {field_name} must be non-negative")

    # Font size must be positive
    if config.body.size <= 0:
        raise ConfigError("Body font size must be positive")
