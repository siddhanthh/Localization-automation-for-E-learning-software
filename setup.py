from setuptools import setup, find_packages

setup(
    name="docx-translator",
    version="1.0.0",
    description="AI-powered DOCX document translation and layout preservation tool",
    author="Localization Specialist",
    packages=find_packages(),
    install_requires=[
        "python-docx>=1.1.2",
        "deep-translator>=1.11.4",
        "requests>=2.31.0",
    ],
    entry_points={
        "console_scripts": [
            "docx-translate=docx_translator.cli:main",
        ],
    },
    python_requires=">=3.8",
)
