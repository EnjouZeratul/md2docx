# tests/test_config.py
import pytest
from pathlib import Path
from md2docx.config import Config, load_config, ConfigError


class TestLoadConfig:
    def test_load_default_config(self):
        """Test loading the default configuration."""
        config = load_config()
        assert config.page.size == "A4"
        assert config.fonts.default == "宋体"
        assert config.body.size == 12

    def test_load_config_from_file(self, tmp_path):
        """Test loading configuration from a YAML file."""
        yaml_content = """
page:
  size: A4
  margins:
    left: 3.0cm
    right: 2.5cm
    top: 2.54cm
    bottom: 2.54cm
fonts:
  default: 宋体
  heading: 宋体
  code: Consolas
  formula: Times New Roman
body:
  size: 12
  line_spacing: 1.5
"""
        config_file = tmp_path / "test.yaml"
        config_file.write_text(yaml_content, encoding="utf-8")

        config = load_config(str(config_file))
        assert config.page.size == "A4"


class TestUnitParsing:
    def test_parse_cm_unit(self):
        from md2docx.config import parse_unit
        result = parse_unit("3.0cm")
        assert result.cm == 3.0

    def test_invalid_unit_raises_error(self):
        from md2docx.config import parse_unit
        with pytest.raises(ConfigError):
            parse_unit("invalid")
