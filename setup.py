from setuptools import setup, find_packages

setup(
    name="md2docx",
    version="0.1.0",
    description="Markdown thesis to DOCX converter with Mermaid and LaTeX support",
    packages=find_packages(),
    install_requires=[
        "python-docx>=0.8.11",
        "PyYAML>=6.0",
        "requests>=2.28.0",
        "matplotlib>=3.6.0",
    ],
    entry_points={
        "console_scripts": [
            "md2docx=md2docx.cli:main",
        ],
    },
    python_requires=">=3.8",
)
