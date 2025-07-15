"""
Tests for CLI commands.
"""

import pytest
import tempfile
import shutil
from pathlib import Path
from unittest.mock import Mock, patch, call
import sys
import zipfile

from prism_ole_handler.cli import extract, insert


class TestExtractCLI:
    """Test cases for extract CLI command."""

    def setup_method(self):
        """Set up test environment."""
        self.test_dir = Path(tempfile.mkdtemp())
        self.test_pptx = self.test_dir / "test.pptx"
        self.output_dir = self.test_dir / "output"
        
        # Create mock PPTX
        self._create_mock_pptx()
        
    def teardown_method(self):
        """Clean up test environment."""
        shutil.rmtree(self.test_dir)
        
    def _create_mock_pptx(self):
        """Create a mock PPTX file for testing."""
        with zipfile.ZipFile(self.test_pptx, 'w') as zf:
            zf.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')
            zf.writestr("ppt/presentation.xml", '<?xml version="1.0"?><presentation/>')
            zf.writestr("ppt/slides/slide1.xml", '<?xml version="1.0"?><slide/>')
            
    @patch('prism_ole_handler.cli.extract.PrismExtractor')
    def test_main_basic_extraction(self, mock_extractor_class):
        """Test basic extraction command."""
        mock_extractor = Mock()
        mock_extractor_class.return_value = mock_extractor
        mock_extractor.extract_prism_objects.return_value = ['extracted_file.pzfx']
        
        test_args = ['prism-extract', str(self.test_pptx), '-o', str(self.output_dir)]
        
        with patch.object(sys, 'argv', test_args):
            extract.main()
            
        mock_extractor_class.assert_called_once_with(str(self.test_pptx))
        mock_extractor.extract_prism_objects.assert_called_once_with(str(self.output_dir), selected_slides=None)
        
    @patch('prism_ole_handler.cli.extract.PrismExtractor')
    def test_main_with_slide_selection(self, mock_extractor_class):
        """Test extraction with slide selection."""
        mock_extractor = Mock()
        mock_extractor_class.return_value = mock_extractor
        mock_extractor.extract_prism_objects.return_value = ['extracted_file.pzfx']
        
        test_args = ['prism-extract', str(self.test_pptx), '-o', str(self.output_dir), '--slide', '2', '--slide', '3']
        
        with patch.object(sys, 'argv', test_args):
            extract.main()
            
        mock_extractor.extract_prism_objects.assert_called_once_with(str(self.output_dir), selected_slides=[2, 3])
        
    @patch('prism_ole_handler.cli.extract.PrismExtractor')
    def test_main_with_slides_csv(self, mock_extractor_class):
        """Test extraction with comma-separated slides."""
        mock_extractor = Mock()
        mock_extractor_class.return_value = mock_extractor
        mock_extractor.extract_prism_objects.return_value = ['extracted_file.pzfx']
        
        test_args = ['prism-extract', str(self.test_pptx), '-o', str(self.output_dir), '--slides', '1,2,3']
        
        with patch.object(sys, 'argv', test_args):
            extract.main()
            
        mock_extractor.extract_prism_objects.assert_called_once_with(str(self.output_dir), selected_slides=[1, 2, 3])
        
    @patch('prism_ole_handler.cli.extract.PrismExtractor')
    def test_main_no_objects_found(self, mock_extractor_class):
        """Test extraction when no objects are found."""
        mock_extractor = Mock()
        mock_extractor_class.return_value = mock_extractor
        mock_extractor.extract_prism_objects.return_value = []
        
        test_args = ['prism-extract', str(self.test_pptx), '-o', str(self.output_dir)]
        
        with patch.object(sys, 'argv', test_args):
            extract.main()
            
        mock_extractor.extract_prism_objects.assert_called_once()
        
    @patch('prism_ole_handler.cli.extract.PrismExtractor')
    def test_main_file_not_found(self, mock_extractor_class):
        """Test extraction with non-existent file."""
        mock_extractor_class.side_effect = FileNotFoundError("File not found")
        
        test_args = ['prism-extract', 'nonexistent.pptx', '-o', str(self.output_dir)]
        
        with patch.object(sys, 'argv', test_args):
            with pytest.raises(SystemExit):
                extract.main()


class TestInsertCLI:
    """Test cases for insert CLI command."""

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
            zf.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')
            zf.writestr("ppt/presentation.xml", '<?xml version="1.0"?><presentation/>')
            zf.writestr("ppt/slides/slide1.xml", '<?xml version="1.0"?><slide/>')
            
    def _create_mock_prism(self):
        """Create a mock PRISM file for testing."""
        with zipfile.ZipFile(self.test_prism, 'w') as zf:
            zf.writestr("prism.xml", '<?xml version="1.0"?><prism><data>test</data></prism>')
            
    @patch('prism_ole_handler.cli.insert.PrismInserter')
    def test_main_basic_insertion(self, mock_inserter_class):
        """Test basic insertion command."""
        mock_inserter = Mock()
        mock_inserter_class.return_value = mock_inserter
        mock_inserter.insert_prism_object.return_value = True
        
        test_args = ['prism-insert', str(self.test_pptx), '--slide', '1', '--prism', str(self.test_prism)]
        
        with patch.object(sys, 'argv', test_args):
            insert.main()
            
        mock_inserter_class.assert_called_once_with(str(self.test_pptx), output_path=None)
        mock_inserter.insert_prism_object.assert_called_once_with(1, str(self.test_prism), create_new=False, force_insert=False)
        
    @patch('prism_ole_handler.cli.insert.PrismInserter')
    def test_main_with_create_new(self, mock_inserter_class):
        """Test insertion with create new slide option."""
        mock_inserter = Mock()
        mock_inserter_class.return_value = mock_inserter
        mock_inserter.insert_prism_object.return_value = True
        
        test_args = ['prism-insert', str(self.test_pptx), '--slide', '2', '--prism', str(self.test_prism), '--create-new']
        
        with patch.object(sys, 'argv', test_args):
            insert.main()
            
        mock_inserter.insert_prism_object.assert_called_once_with(2, str(self.test_prism), create_new=True, force_insert=False)
        
    @patch('prism_ole_handler.cli.insert.PrismInserter')
    def test_main_with_force_insert(self, mock_inserter_class):
        """Test insertion with force insert option."""
        mock_inserter = Mock()
        mock_inserter_class.return_value = mock_inserter
        mock_inserter.insert_prism_object.return_value = True
        
        test_args = ['prism-insert', str(self.test_pptx), '--slide', '1', '--prism', str(self.test_prism), '--force-insert']
        
        with patch.object(sys, 'argv', test_args):
            insert.main()
            
        mock_inserter.insert_prism_object.assert_called_once_with(1, str(self.test_prism), create_new=False, force_insert=True)
        
    @patch('prism_ole_handler.cli.insert.PrismInserter')
    def test_main_with_output_path(self, mock_inserter_class):
        """Test insertion with output path."""
        mock_inserter = Mock()
        mock_inserter_class.return_value = mock_inserter
        mock_inserter.insert_prism_object.return_value = True
        
        output_path = str(self.test_dir / "output.pptx")
        test_args = ['prism-insert', str(self.test_pptx), '--slide', '1', '--prism', str(self.test_prism), '-o', output_path]
        
        with patch.object(sys, 'argv', test_args):
            insert.main()
            
        mock_inserter_class.assert_called_once_with(str(self.test_pptx), output_path=output_path)
        
    @patch('prism_ole_handler.cli.insert.PrismInserter')
    def test_main_insertion_failure(self, mock_inserter_class):
        """Test insertion when it fails."""
        mock_inserter = Mock()
        mock_inserter_class.return_value = mock_inserter
        mock_inserter.insert_prism_object.return_value = False
        
        test_args = ['prism-insert', str(self.test_pptx), '--slide', '1', '--prism', str(self.test_prism)]
        
        with patch.object(sys, 'argv', test_args):
            with pytest.raises(SystemExit):
                insert.main()
                
    @patch('prism_ole_handler.cli.insert.PrismInserter')
    def test_main_file_not_found(self, mock_inserter_class):
        """Test insertion with non-existent file."""
        mock_inserter_class.side_effect = FileNotFoundError("File not found")
        
        test_args = ['prism-insert', 'nonexistent.pptx', '--slide', '1', '--prism', str(self.test_prism)]
        
        with patch.object(sys, 'argv', test_args):
            with pytest.raises(SystemExit):
                insert.main()