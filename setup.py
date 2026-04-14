"""
Setup script for packaging application
"""
from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="oracle-rac-report-generator",
    version="1.0.0",
    author="Your Name",
    author_email="your.email@example.com",
    description="Oracle RAC Report Generator - Desktop Application",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/oracle-rac-report-generator",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Development Status :: 3 - Alpha",
        "Intended Audience :: System Administrators",
        "Topic :: System :: Monitoring",
    ],
    python_requires=">=3.8",
    install_requires=[
        "PyQt5>=5.15.9",
        "beautifulsoup4>=4.12.2",
        "lxml>=4.9.3",
        "python-docx>=0.8.11",
        "reportlab>=4.0.7",
        "Pillow>=10.0.0",
        "openpyxl>=3.1.2",
    ],
    entry_points={
        "console_scripts": [
            "oracle-rac-report-generator=main:main",
        ],
    },
)
