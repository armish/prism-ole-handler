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
        with zipfile.ZipFile(self.test_pptx, 'w') as zf:
            # Add basic PPTX structure
            zf.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')
            zf.writestr("ppt/presentation.xml", '<?xml version="1.0"?><presentation/>')
            zf.writestr("ppt/slides/slide1.xml", '<?xml version="1.0"?><slide/>')
            zf.writestr("ppt/slides/_rels/slide1.xml.rels", '''<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''')
            
    def _create_mock_prism(self):
        """Create a mock PRISM file for testing."""
        with zipfile.ZipFile(self.test_prism, 'w') as zf:
            zf.writestr("prism.xml", '<?xml version="1.0"?><prism><data>test</data></prism>')
            
    def test_initialization(self):
        """Test PrismInserter initialization."""
        inserter = PrismInserter(str(self.test_pptx))
        assert inserter.pptx_path == str(self.test_pptx)
        
    def test_initialization_with_nonexistent_file(self):
        """Test initialization with non-existent file."""
        with pytest.raises(FileNotFoundError):
            PrismInserter("nonexistent.pptx")
            
    def test_initialization_with_output_path(self):
        """Test initialization with output path."""
        output_path = str(self.test_dir / "output.pptx")
        inserter = PrismInserter(str(self.test_pptx), output_path=output_path)
        assert inserter.output_path == output_path
        
    def test_validate_prism_file_valid(self):
        """Test validation of valid PRISM file."""
        inserter = PrismInserter(str(self.test_pptx))
        # Should not raise exception
        inserter._validate_prism_file(str(self.test_prism))
        
    def test_validate_prism_file_invalid_extension(self):
        """Test validation with invalid file extension."""
        invalid_file = self.test_dir / "test.prism"
        invalid_file.write_text("test")
        
        inserter = PrismInserter(str(self.test_pptx))
        with pytest.raises(ValueError, match="Only .pzfx files are supported"):
            inserter._validate_prism_file(str(invalid_file))
            
    def test_validate_prism_file_nonexistent(self):
        """Test validation with non-existent file."""
        inserter = PrismInserter(str(self.test_pptx))
        with pytest.raises(FileNotFoundError):
            inserter._validate_prism_file("nonexistent.pzfx")
            
    def test_get_slide_count(self):
        """Test getting slide count."""
        inserter = PrismInserter(str(self.test_pptx))
        slide_count = inserter.get_slide_count()
        assert slide_count == 1
        
    def test_create_new_slide(self):
        """Test creating a new slide."""
        inserter = PrismInserter(str(self.test_pptx))
        inserter._create_new_slide(2)
        
        # Verify slide was created
        slide_count = inserter.get_slide_count()
        assert slide_count == 2
        
    def test_create_backup(self):
        """Test creating backup of original file."""
        inserter = PrismInserter(str(self.test_pptx))
        backup_path = inserter._create_backup()
        
        assert Path(backup_path).exists()
        assert backup_path.endswith(".backup")
        
    def test_has_existing_embeddings_false(self):
        """Test checking for existing embeddings when none exist."""
        inserter = PrismInserter(str(self.test_pptx))
        has_embeddings = inserter._has_existing_embeddings(1)
        assert has_embeddings is False
        
    @patch('prism_ole_handler.core.inserter.update_ole_file')
    def test_insert_prism_object_success(self, mock_update_ole):
        """Test successful PRISM object insertion."""
        mock_update_ole.return_value = b'mock_ole_data'
        
        inserter = PrismInserter(str(self.test_pptx))
        result = inserter.insert_prism_object(1, str(self.test_prism))
        
        assert result is True
        mock_update_ole.assert_called_once()
        
    def test_insert_prism_object_invalid_slide(self):
        """Test insertion with invalid slide number."""
        inserter = PrismInserter(str(self.test_pptx))
        with pytest.raises(ValueError, match="Invalid slide number"):
            inserter.insert_prism_object(999, str(self.test_prism))
            
    def test_insert_prism_object_create_new_slide(self):
        """Test insertion with new slide creation."""
        inserter = PrismInserter(str(self.test_pptx))
        
        with patch.object(inserter, '_create_new_slide'):
            with patch.object(inserter, '_add_prism_to_slide'):
                result = inserter.insert_prism_object(
                    2, str(self.test_prism), create_new=True
                )
                
        assert result is True
        
    def test_insert_prism_object_existing_embeddings_no_force(self):
        """Test insertion when embeddings exist without force flag."""
        inserter = PrismInserter(str(self.test_pptx))
        
        with patch.object(inserter, '_has_existing_embeddings', return_value=True):
            with pytest.raises(ValueError, match="already has embedded objects"):
                inserter.insert_prism_object(1, str(self.test_prism))
                
    def test_insert_prism_object_existing_embeddings_with_force(self):
        """Test insertion when embeddings exist with force flag."""
        inserter = PrismInserter(str(self.test_pptx))
        
        with patch.object(inserter, '_has_existing_embeddings', return_value=True):
            with patch.object(inserter, '_add_prism_to_slide'):
                result = inserter.insert_prism_object(
                    1, str(self.test_prism), force_insert=True
                )
                
        assert result is True