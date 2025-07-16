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
        with zipfile.ZipFile(self.test_pptx, "w") as zf:
            # Add basic PPTX structure
            zf.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')
            zf.writestr("ppt/presentation.xml", '<?xml version="1.0"?><presentation/>')
            zf.writestr("ppt/slides/slide1.xml", '<?xml version="1.0"?><slide/>')
            zf.writestr(
                "ppt/slides/_rels/slide1.xml.rels",
                '<?xml version="1.0"?><Relationships/>',
            )

    def test_initialization(self):
        """Test PrismExtractor initialization."""
        extractor = PrismExtractor(str(self.test_pptx))
        assert extractor.pptx_path == self.test_pptx
        assert extractor.temp_dir == Path("temp_pptx_extract")

    def test_initialization_with_nonexistent_file(self):
        """Test initialization with non-existent file (should not raise error)."""
        # The current implementation doesn't validate file existence in __init__
        extractor = PrismExtractor("nonexistent.pptx")
        assert extractor.pptx_path == Path("nonexistent.pptx")

    def test_extract_pptx(self):
        """Test extracting PPTX file."""
        extractor = PrismExtractor(str(self.test_pptx))
        extractor.extract_pptx()

        # Check that temp directory was created
        assert extractor.temp_dir.exists()

        # Check that basic PPTX structure was extracted
        assert (extractor.temp_dir / "ppt" / "presentation.xml").exists()
        assert (extractor.temp_dir / "ppt" / "slides" / "slide1.xml").exists()

        # Cleanup
        shutil.rmtree(extractor.temp_dir)

    def test_is_prism_xml_positive(self):
        """Test PRISM XML detection with valid data."""
        extractor = PrismExtractor(str(self.test_pptx))
        prism_xml = (
            b'<?xml version="1.0"?><GraphPadPrismFile><data></data></GraphPadPrismFile>'
        )

        assert extractor.is_prism_xml(prism_xml) is True

    def test_is_prism_xml_negative(self):
        """Test PRISM XML detection with invalid data."""
        extractor = PrismExtractor(str(self.test_pptx))
        non_prism_xml = b'<?xml version="1.0"?><regular><data></data></regular>'

        assert extractor.is_prism_xml(non_prism_xml) is False

    def test_is_prism_data_zip_format(self):
        """Test PRISM data detection with ZIP format."""
        extractor = PrismExtractor(str(self.test_pptx))
        zip_data = b"PK\x03\x04\x14\x00\x00\x00"  # ZIP header

        assert extractor.is_prism_data(zip_data) is True

    def test_is_prism_data_with_prefix(self):
        """Test PRISM data detection with 4-byte prefix."""
        extractor = PrismExtractor(str(self.test_pptx))
        prefixed_zip = b"\x04\x2c\x00\x00PK\x03\x04"  # 4-byte prefix + ZIP header

        assert extractor.is_prism_data(prefixed_zip) is True

    def test_extract_prism_objects_no_embeddings(self):
        """Test extraction when no embeddings exist."""
        extractor = PrismExtractor(str(self.test_pptx))

        # Mock the methods to return empty results
        with patch.object(extractor, "find_embedded_objects", return_value=([], [])):
            with patch("builtins.print"):  # Suppress print output
                result = extractor.extract_prism_objects(str(self.output_dir))

        # The method doesn't return anything, so we just verify it doesn't crash
        assert result is None

    def test_extract_prism_objects_creates_output_dir(self):
        """Test that extraction creates output directory if it doesn't exist."""
        new_output = self.test_dir / "new_output"
        extractor = PrismExtractor(str(self.test_pptx))

        with patch.object(extractor, "find_embedded_objects", return_value=([], [])):
            with patch("builtins.print"):  # Suppress print output
                extractor.extract_prism_objects(str(new_output))

        assert new_output.exists()
        assert new_output.is_dir()

    def test_analyze_slide_relationships_empty(self):
        """Test analyzing slide relationships with no relationships."""
        extractor = PrismExtractor(str(self.test_pptx))
        extractor.extract_pptx()

        result = extractor.analyze_slide_relationships()

        # Should return empty list for slides with no embedded objects
        assert result == []

        # Cleanup
        shutil.rmtree(extractor.temp_dir)

    def test_find_embedded_objects_no_embeddings(self):
        """Test finding embedded objects when none exist."""
        extractor = PrismExtractor(str(self.test_pptx))
        extractor.extract_pptx()

        objects, slides_info = extractor.find_embedded_objects()

        # Should return empty lists
        assert objects == []
        assert slides_info == []

        # Cleanup
        shutil.rmtree(extractor.temp_dir)

    @patch("prism_ole_handler.core.extractor.HAS_OLEFILE", False)
    def test_extract_from_ole_no_olefile(self):
        """Test OLE extraction when olefile is not available."""
        extractor = PrismExtractor(str(self.test_pptx))

        result = extractor.extract_from_ole(b"dummy_data", "test")

        assert result == []

    @patch("prism_ole_handler.core.extractor.HAS_OLEFILE", True)
    @patch("prism_ole_handler.core.extractor.olefile.OleFileIO")
    def test_extract_from_ole_success(self, mock_ole_io):
        """Test successful OLE extraction."""
        mock_ole = Mock()
        mock_ole_io.return_value = mock_ole
        mock_ole.listdir.return_value = [["CONTENTS"]]

        # Mock stream data
        mock_stream = Mock()
        mock_stream.read.return_value = b"PK\x03\x04test_prism_data"
        mock_ole.openstream.return_value = mock_stream

        extractor = PrismExtractor(str(self.test_pptx))
        result = extractor.extract_from_ole(b"dummy_ole_data", "test")

        assert len(result) > 0
        mock_ole.close.assert_called_once()

    @patch("prism_ole_handler.core.extractor.HAS_OLEFILE", True)
    @patch("prism_ole_handler.core.extractor.olefile.OleFileIO")
    def test_extract_from_ole_exception(self, mock_ole_io):
        """Test OLE extraction with exception."""
        mock_ole_io.side_effect = Exception("OLE error")

        extractor = PrismExtractor(str(self.test_pptx))

        with patch("builtins.print"):  # Suppress error output
            result = extractor.extract_from_ole(b"dummy_ole_data", "test")

        assert result == []
