"""
Tests for PrismExtractor class.
"""

import pytest
import tempfile
import shutil
from pathlib import Path
from unittest.mock import Mock, patch, mock_open
import zipfile

from prism_ole_handler.core.extractor import PrismExtractor


class TestPrismExtractor:
    """Test cases for PrismExtractor class."""

    def setup_method(self):
        """Set up test environment."""
        self.test_dir = Path(tempfile.mkdtemp())
        self.test_pptx = self.test_dir / "test.pptx"
        self.output_dir = self.test_dir / "output"
        self.output_dir.mkdir()
        
        # Create a mock PPTX file
        self._create_mock_pptx()
        
    def teardown_method(self):
        """Clean up test environment."""
        shutil.rmtree(self.test_dir)
        
    def _create_mock_pptx(self):
        """Create a mock PPTX file for testing."""
        with zipfile.ZipFile(self.test_pptx, 'w') as zf:
            # Add basic PPTX structure
            zf.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')
            zf.writestr("ppt/presentation.xml", '<?xml version="1.0"?><presentation/>')
            zf.writestr("ppt/slides/slide1.xml", '<?xml version="1.0"?><slide/>')
            zf.writestr("ppt/slides/_rels/slide1.xml.rels", '<?xml version="1.0"?><Relationships/>')
            
    def test_initialization(self):
        """Test PrismExtractor initialization."""
        extractor = PrismExtractor(str(self.test_pptx))
        assert extractor.pptx_path == str(self.test_pptx)
        assert extractor.pptx_path == str(self.test_pptx)
        
    def test_initialization_with_nonexistent_file(self):
        """Test initialization with non-existent file."""
        with pytest.raises(FileNotFoundError):
            PrismExtractor("nonexistent.pptx")
            
    def test_get_slide_count(self):
        """Test getting slide count from PPTX."""
        extractor = PrismExtractor(str(self.test_pptx))
        slide_count = extractor.get_slide_count()
        assert slide_count == 1
        
    def test_get_slide_count_empty_pptx(self):
        """Test getting slide count from empty PPTX."""
        empty_pptx = self.test_dir / "empty.pptx"
        with zipfile.ZipFile(empty_pptx, 'w') as zf:
            zf.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')
            zf.writestr("ppt/presentation.xml", '<?xml version="1.0"?><presentation/>')
            
        extractor = PrismExtractor(str(empty_pptx))
        slide_count = extractor.get_slide_count()
        assert slide_count == 0
        
    @patch('prism_ole_handler.core.extractor.olefile.isOleFile')
    @patch('prism_ole_handler.core.extractor.olefile.OleFileIO')
    def test_process_ole_file_success(self, mock_ole_io, mock_is_ole):
        """Test successful OLE file processing."""
        mock_is_ole.return_value = True
        mock_ole = Mock()
        mock_ole_io.return_value = mock_ole
        mock_ole.openstream.return_value = Mock()
        mock_ole.openstream.return_value.read.return_value = b'mock_prism_data'
        
        extractor = PrismExtractor(str(self.test_pptx))
        result = extractor._process_ole_file(b'mock_ole_data', 'test_object')
        
        assert result is not None
        mock_is_ole.assert_called_once()
        mock_ole_io.assert_called_once()
        
    @patch('prism_ole_handler.core.extractor.olefile.isOleFile')
    def test_process_ole_file_not_ole(self, mock_is_ole):
        """Test processing non-OLE file."""
        mock_is_ole.return_value = False
        
        extractor = PrismExtractor(str(self.test_pptx))
        result = extractor._process_ole_file(b'not_ole_data', 'test_object')
        
        assert result is None
        
    def test_extract_prism_objects_no_embeddings(self):
        """Test extraction when no embeddings exist."""
        extractor = PrismExtractor(str(self.test_pptx))
        extracted = extractor.extract_prism_objects(str(self.output_dir))
        
        assert extracted == []
        
    def test_extract_prism_objects_with_slide_selection(self):
        """Test extraction with specific slide selection."""
        extractor = PrismExtractor(str(self.test_pptx))
        extracted = extractor.extract_prism_objects(
            str(self.output_dir), 
            selected_slides=[1]
        )
        
        assert isinstance(extracted, list)
        
    def test_extract_prism_objects_invalid_slide_selection(self):
        """Test extraction with invalid slide selection."""
        extractor = PrismExtractor(str(self.test_pptx))
        extracted = extractor.extract_prism_objects(
            str(self.output_dir), 
            selected_slides=[999]
        )
        
        assert extracted == []
        
    def test_extract_prism_objects_creates_output_dir(self):
        """Test that extraction creates output directory if it doesn't exist."""
        new_output = self.test_dir / "new_output"
        extractor = PrismExtractor(str(self.test_pptx))
        extractor.extract_prism_objects(str(new_output))
        
        assert new_output.exists()
        assert new_output.is_dir()