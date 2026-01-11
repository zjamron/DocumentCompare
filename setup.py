"""Setup script for DocumentCompare package."""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="document-compare",
    version="1.0.0",
    author="Valon Technologies",
    description="Document comparison (redlining) library for Word and PDF documents",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/zjamron/DocumentCompare",
    packages=find_packages(),
    py_modules=[
        "document_compare",
        "pdf_support",
        "compare_preserve_formatting"
    ],
    package_dir={"": "samples"},
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Legal Industry",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.13",
        "Topic :: Office/Business",
        "Topic :: Text Processing :: General",
    ],
    python_requires=">=3.9",
    install_requires=[
        "python-docx>=0.8.11",
        "pymupdf>=1.23.0",
        "reportlab>=4.0.0",
        "lxml>=4.9.0",
    ],
    extras_require={
        "dev": [
            "pytest>=7.0.0",
            "black>=23.0.0",
            "flake8>=6.0.0",
        ],
    },
    entry_points={
        "console_scripts": [
            "document-compare=document_compare:main",
            "doccompare=document_compare:main",
        ],
    },
)
