"""
Tests for PrismInserter class.
"""

import pytest
import tempfile
import shutil
from pathlib import Path
from unittest.mock import Mock, patch, mock_open
import zipfile

from prism_ole_handler.core.inserter import PrismInserter


class TestPrismInserter:
    """Test cases for PrismInserter class."""

    def setup_method(self):
        """Set up test environment."""
        self.test_dir = Path(tempfile.mkdtemp())
        self.test_pptx = self.test_dir / "test.pptx"
        self.test_prism = self.test_dir / "test.pzfx"

        # Create mock files
        self._create_mock_pptx()
        self._create_mock_prism()

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
                """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>""",
            )

    def _create_mock_prism(self):
        """Create a mock PRISM file for testing."""
        with zipfile.ZipFile(self.test_prism, "w") as zf:
            zf.writestr(
                "prism.xml", '<?xml version="1.0"?><prism><data>test</data></prism>'
            )

    def test_initialization(self):
        """Test PrismInserter initialization."""
        inserter = PrismInserter(str(self.test_pptx))
        assert inserter.pptx_path == self.test_pptx
        assert inserter.temp_dir == Path("temp_pptx_insert")
        assert inserter.backup_path is None

    def test_initialization_with_nonexistent_file(self):
        """Test initialization with non-existent file (should not raise error)."""
        # The current implementation doesn't validate file existence in __init__
        inserter = PrismInserter("nonexistent.pptx")
        assert inserter.pptx_path == Path("nonexistent.pptx")

    def test_create_backup(self):
        """Test creating backup of original file."""
        inserter = PrismInserter(str(self.test_pptx))
        inserter.create_backup()

        expected_backup = self.test_pptx.parent / (
            self.test_pptx.stem + "_backup" + self.test_pptx.suffix
        )
        assert expected_backup.exists()
        assert inserter.backup_path == expected_backup

    def test_extract_pptx(self):
        """Test extracting PPTX file."""
        inserter = PrismInserter(str(self.test_pptx))
        inserter.extract_pptx()

        # Check that temp directory was created
        assert inserter.temp_dir.exists()

        # Check that basic PPTX structure was extracted
        assert (inserter.temp_dir / "ppt" / "presentation.xml").exists()
        assert (inserter.temp_dir / "ppt" / "slides" / "slide1.xml").exists()

        # Cleanup
        shutil.rmtree(inserter.temp_dir)

    def test_get_slide_count(self):
        """Test getting slide count."""
        inserter = PrismInserter(str(self.test_pptx))
        inserter.extract_pptx()

        slide_count = inserter.get_slide_count()
        assert slide_count == 1

        # Cleanup
        shutil.rmtree(inserter.temp_dir)

    def test_get_slide_count_no_slides(self):
        """Test getting slide count when no slides directory exists."""
        inserter = PrismInserter(str(self.test_pptx))
        # Don't extract, so slides directory won't exist

        slide_count = inserter.get_slide_count()
        assert slide_count == 0

    def test_slide_exists(self):
        """Test checking if slide exists."""
        inserter = PrismInserter(str(self.test_pptx))
        inserter.extract_pptx()

        assert inserter.slide_exists(1) is True
        assert inserter.slide_exists(2) is False

        # Cleanup
        shutil.rmtree(inserter.temp_dir)

    def test_slide_has_embeddings_false(self):
        """Test checking for embeddings when none exist."""
        inserter = PrismInserter(str(self.test_pptx))
        inserter.extract_pptx()

        has_embeddings = inserter.slide_has_embeddings(1)
        assert has_embeddings is False

        # Cleanup
        shutil.rmtree(inserter.temp_dir)

    def test_find_embedding_for_slide_none(self):
        """Test finding embedding when none exists."""
        inserter = PrismInserter(str(self.test_pptx))
        inserter.extract_pptx()

        embedding = inserter.find_embedding_for_slide(1)
        assert embedding is None

        # Cleanup
        shutil.rmtree(inserter.temp_dir)

    def test_insert_prism_object_prism_extension_error(self):
        """Test insertion with .prism file (should fail)."""
        prism_file = self.test_dir / "test.prism"
        prism_file.write_text("dummy content")

        inserter = PrismInserter(str(self.test_pptx))
        inserter.extract_pptx()

        with patch("builtins.print"):  # Suppress error output
            result = inserter.insert_prism_object(1, str(prism_file))

        assert result is False

        # Cleanup
        shutil.rmtree(inserter.temp_dir)

    def test_insert_prism_object_nonexistent_file(self):
        """Test insertion with non-existent PRISM file."""
        inserter = PrismInserter(str(self.test_pptx))
        inserter.extract_pptx()

        with patch("builtins.print"):  # Suppress error output
            result = inserter.insert_prism_object(1, "nonexistent.pzfx")

        assert result is False

        # Cleanup
        shutil.rmtree(inserter.temp_dir)

    def test_insert_prism_object_nonexistent_slide_no_create(self):
        """Test insertion into non-existent slide without create flag."""
        inserter = PrismInserter(str(self.test_pptx))
        inserter.extract_pptx()

        with patch("builtins.print"):  # Suppress error output
            result = inserter.insert_prism_object(99, str(self.test_prism))

        assert result is False

        # Cleanup
        shutil.rmtree(inserter.temp_dir)

    def test_create_new_slide(self):
        """Test creating a new slide."""
        inserter = PrismInserter(str(self.test_pptx))
        inserter.extract_pptx()

        with patch("builtins.print"):  # Suppress output
            result = inserter.create_new_slide(2)

        assert result is True
        assert inserter.slide_exists(2) is True

        # Cleanup
        shutil.rmtree(inserter.temp_dir)

    def test_batch_insert_success(self):
        """Test batch insertion with successful operations."""
        updates = [(1, str(self.test_prism))]

        inserter = PrismInserter(str(self.test_pptx))

        with patch("builtins.print"):  # Suppress output
            with patch.object(inserter, "insert_prism_object", return_value=True):
                result = inserter.batch_insert(updates)

        assert result == 1

    def test_batch_insert_failure(self):
        """Test batch insertion with failed operations."""
        updates = [(1, str(self.test_prism))]

        inserter = PrismInserter(str(self.test_pptx))

        with patch("builtins.print"):  # Suppress output
            with patch.object(inserter, "insert_prism_object", return_value=False):
                result = inserter.batch_insert(updates)

        assert result == 0

    @patch("prism_ole_handler.core.inserter.update_ole_file")
    def test_update_ole_contents(self, mock_update_ole):
        """Test updating OLE contents."""
        mock_update_ole.return_value = b"updated_ole_data"

        # Create a mock OLE file
        ole_file = self.test_dir / "test.bin"
        ole_file.write_bytes(b"original_ole_data")

        inserter = PrismInserter(str(self.test_pptx))
        inserter.update_ole_contents(ole_file, b"new_prism_data")

        mock_update_ole.assert_called_once_with(b"original_ole_data", b"new_prism_data")

        # Check that file was updated
        assert ole_file.read_bytes() == b"updated_ole_data"

    def test_extract_template_ole_no_file(self):
        """Test extracting template OLE when template file doesn't exist."""
        inserter = PrismInserter(str(self.test_pptx))

        with patch("builtins.print"):  # Suppress output
            result = inserter.extract_template_ole(Path("nonexistent.pptx"))

        assert result is None

    def test_insert_into_empty_slide_success(self):
        """Test inserting into empty slide."""
        inserter = PrismInserter(str(self.test_pptx))
        inserter.extract_pptx()

        with patch("builtins.print"):  # Suppress output
            with patch.object(inserter, "create_ole_file"):
                with patch.object(inserter, "add_embedded_object_to_slide"):
                    result = inserter.insert_into_empty_slide(1, str(self.test_prism))

        assert result is True

        # Cleanup
        shutil.rmtree(inserter.temp_dir)
