"""
PRISM OLE Handler

A Python package for extracting and inserting GraphPad PRISM objects 
from Microsoft Office documents (PowerPoint, Word, Excel) on macOS.
"""

__version__ = "0.1.0"
__author__ = "B. Arman Aksoy"
__email__ = "arman@aksoy.org"

from .core.extractor import PrismExtractor
from .core.inserter import PrismInserter

__all__ = ['PrismExtractor', 'PrismInserter']