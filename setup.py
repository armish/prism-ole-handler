"""
Setup script for prism-ole-handler package.
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read the README file
this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text()

setup(
    name="prism-ole-handler",
    version="0.1.0",
    author="B. Arman Aksoy",
    author_email="arman@aksoy.org",
    description="Extract and insert GraphPad PRISM objects from Microsoft Office documents (PowerPoint, Word, Excel) on macOS",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/armish/prism-ole-handler",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Science/Research",
        "License :: OSI Approved :: MIT License",
        "Operating System :: MacOS",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Scientific/Engineering",
        "Topic :: Office/Business :: Office Suites",
    ],
    python_requires=">=3.8",
    install_requires=[
        "olefile>=0.46",
    ],
    extras_require={
        "dev": [
            "pytest>=6.0",
            "black",
            "flake8",
            "mypy",
        ],
    },
    entry_points={
        "console_scripts": [
            "prism-extract=prism_ole_handler.cli.extract:main",
            "prism-insert=prism_ole_handler.cli.insert:main",
        ],
    },
    include_package_data=True,
    package_data={
        "prism_ole_handler": [
            "test/*.pptx",
            "test/*.prism",
            "test/*.pzfx",
        ],
    },
    keywords="prism ole handler microsoft office powerpoint word excel macos graphpad",
    project_urls={
        "Bug Reports": "https://github.com/armish/prism-ole-handler/issues",
        "Source": "https://github.com/armish/prism-ole-handler",
    },
)