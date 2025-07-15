"""
Pytest configuration and fixtures for prism-ole-handler tests.
"""

import pytest
import tempfile
import shutil
from pathlib import Path
import zipfile


@pytest.fixture
def temp_dir():
    """Create a temporary directory for tests."""
    temp_dir = Path(tempfile.mkdtemp())
    yield temp_dir
    shutil.rmtree(temp_dir)


@pytest.fixture
def mock_pptx_file(temp_dir):
    """Create a mock PPTX file for testing."""
    pptx_file = temp_dir / "test.pptx"
    with zipfile.ZipFile(pptx_file, 'w') as zf:
        zf.writestr("_rels/.rels", '''<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''')
        zf.writestr("ppt/presentation.xml", '''<?xml version="1.0"?>
<presentation xmlns="http://schemas.openxmlformats.org/presentationml/2006/main">
</presentation>''')
        zf.writestr("ppt/slides/slide1.xml", '''<?xml version="1.0"?>
<slide xmlns="http://schemas.openxmlformats.org/presentationml/2006/main">
</slide>''')
        zf.writestr("ppt/slides/_rels/slide1.xml.rels", '''<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''')
    return pptx_file


@pytest.fixture
def mock_prism_file(temp_dir):
    """Create a mock PRISM (.pzfx) file for testing."""
    prism_file = temp_dir / "test.pzfx"
    with zipfile.ZipFile(prism_file, 'w') as zf:
        zf.writestr("prism.xml", '''<?xml version="1.0"?>
<prism>
    <data>
        <graph>
            <title>Test Graph</title>
        </graph>
    </data>
</prism>''')
    return prism_file


@pytest.fixture
def mock_empty_pptx(temp_dir):
    """Create a mock empty PPTX file for testing."""
    pptx_file = temp_dir / "empty.pptx"
    with zipfile.ZipFile(pptx_file, 'w') as zf:
        zf.writestr("_rels/.rels", '''<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''')
        zf.writestr("ppt/presentation.xml", '''<?xml version="1.0"?>
<presentation xmlns="http://schemas.openxmlformats.org/presentationml/2006/main">
</presentation>''')
    return pptx_file


@pytest.fixture
def output_dir(temp_dir):
    """Create an output directory for tests."""
    output_dir = temp_dir / "output"
    output_dir.mkdir()
    return output_dir