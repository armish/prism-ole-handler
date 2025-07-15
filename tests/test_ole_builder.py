"""
Tests for OLE builder utilities.
"""

import pytest
import tempfile
import shutil
from pathlib import Path
from unittest.mock import Mock, patch, mock_open
import struct

from prism_ole_handler.utils.ole_builder import OLEBuilder, update_ole_file


class TestOLEBuilder:
    """Test cases for OLEBuilder class."""

    def setup_method(self):
        """Set up test environment."""
        self.test_dir = Path(tempfile.mkdtemp())
        self.test_ole_data = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' + b'\x00' * 500  # Mock OLE header + data
        self.test_prism_data = b'<?xml version="1.0"?><prism><data>test</data></prism>'
        
    def teardown_method(self):
        """Clean up test environment."""
        shutil.rmtree(self.test_dir)
        
    @patch('prism_ole_handler.utils.ole_builder.olefile.OleFileIO')
    def test_initialization(self, mock_ole_io):
        """Test OLEBuilder initialization."""
        builder = OLEBuilder(self.test_ole_data)
        mock_ole_io.assert_called_once_with(self.test_ole_data)
        assert builder.original_ole == mock_ole_io.return_value
        
    @patch('prism_ole_handler.utils.ole_builder.olefile.OleFileIO')
    def test_close_method(self, mock_ole_io):
        """Test the close method."""
        mock_ole = Mock()
        mock_ole_io.return_value = mock_ole
        
        builder = OLEBuilder(self.test_ole_data)
        builder.close()
        
        mock_ole.close.assert_called_once()
        
    @patch('prism_ole_handler.utils.ole_builder.olefile.OleFileIO')
    def test_build_updated_ole_with_contents_stream(self, mock_ole_io):
        """Test building updated OLE with CONTENTS stream."""
        mock_ole = Mock()
        mock_ole_io.return_value = mock_ole
        mock_ole.listdir.return_value = [['CONTENTS'], ['OTHER_STREAM']]
        
        # Mock the CONTENTS stream
        mock_contents_stream = Mock()
        mock_contents_stream.read.return_value = b'\x04\x2c\x00\x00PK' + b'original_prism_data'
        
        # Mock the OTHER_STREAM
        mock_other_stream = Mock()
        mock_other_stream.read.return_value = b'other_data'
        
        def mock_openstream(stream_path):
            if stream_path == ['CONTENTS']:
                return mock_contents_stream
            else:
                return mock_other_stream
                
        mock_ole.openstream.side_effect = mock_openstream
        
        # Mock the file pointer operations
        mock_ole.fp.seek.return_value = None
        mock_ole.fp.read.return_value = b'\x04\x2c\x00\x00PK' + b'original_prism_data' + b'\x00' * 400
        
        builder = OLEBuilder(self.test_ole_data)
        result = builder.build_updated_ole(self.test_prism_data)
        
        assert isinstance(result, bytes)
        assert len(result) > 0
        
    @patch('prism_ole_handler.utils.ole_builder.olefile.OleFileIO')
    def test_build_updated_ole_no_contents_stream(self, mock_ole_io):
        """Test building updated OLE without CONTENTS stream."""
        mock_ole = Mock()
        mock_ole_io.return_value = mock_ole
        mock_ole.listdir.return_value = [['OTHER_STREAM']]
        mock_ole.openstream.return_value.read.return_value = b'other_data'
        
        # Mock the file pointer operations for build_ole_compound_file
        mock_ole.fp.seek.return_value = None
        mock_ole.fp.read.return_value = b'\x04\x2c\x00\x00PK' + b'dummy_data' + b'\x00' * 400
        
        # Mock openstream to return CONTENTS stream data when called
        def mock_openstream_side_effect(stream_path):
            if stream_path == ['CONTENTS']:
                mock_stream = Mock()
                mock_stream.read.return_value = b'\x04\x2c\x00\x00PK' + b'dummy_data'
                return mock_stream
            else:
                mock_stream = Mock()
                mock_stream.read.return_value = b'other_data'
                return mock_stream
                
        mock_ole.openstream.side_effect = mock_openstream_side_effect
        
        builder = OLEBuilder(self.test_ole_data)
        result = builder.build_updated_ole(self.test_prism_data)
        
        assert isinstance(result, bytes)
        assert len(result) > 0
        
    @patch('prism_ole_handler.utils.ole_builder.olefile.OleFileIO')
    def test_build_ole_compound_file_error_handling(self, mock_ole_io):
        """Test error handling in build_ole_compound_file."""
        mock_ole = Mock()
        mock_ole_io.return_value = mock_ole
        mock_ole.listdir.return_value = [['CONTENTS']]
        
        # Mock openstream to return data without PK header
        mock_stream = Mock()
        mock_stream.read.return_value = b'no_pk_header'
        mock_ole.openstream.return_value = mock_stream
        
        mock_ole.fp.seek.return_value = None
        mock_ole.fp.read.return_value = b'no_pk_header' + b'\x00' * 400
        
        builder = OLEBuilder(self.test_ole_data)
        
        with pytest.raises(ValueError, match="Could not find CONTENTS stream"):
            builder.build_updated_ole(self.test_prism_data)
            
    @patch('prism_ole_handler.utils.ole_builder.olefile.OleFileIO')
    def test_build_updated_ole_exception_handling(self, mock_ole_io):
        """Test exception handling in build_updated_ole."""
        mock_ole = Mock()
        mock_ole_io.return_value = mock_ole
        mock_ole.listdir.side_effect = Exception("Test error")
        
        builder = OLEBuilder(self.test_ole_data)
        
        with pytest.raises(Exception, match="Test error"):
            builder.build_updated_ole(self.test_prism_data)


class TestUpdateOleFile:
    """Test cases for update_ole_file function."""
    
    def setup_method(self):
        """Set up test environment."""
        self.test_ole_data = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' + b'\x00' * 500
        self.test_prism_data = b'<?xml version="1.0"?><prism><data>test</data></prism>'
        
    @patch('prism_ole_handler.utils.ole_builder.OLEBuilder')
    def test_update_ole_file_success(self, mock_ole_builder_class):
        """Test successful OLE file update."""
        mock_builder = Mock()
        mock_ole_builder_class.return_value = mock_builder
        mock_builder.build_updated_ole.return_value = b'updated_ole_data'
        
        result = update_ole_file(self.test_ole_data, self.test_prism_data)
        
        mock_ole_builder_class.assert_called_once_with(self.test_ole_data)
        mock_builder.build_updated_ole.assert_called_once_with(self.test_prism_data)
        mock_builder.close.assert_called_once()
        assert result == b'updated_ole_data'
        
    @patch('prism_ole_handler.utils.ole_builder.OLEBuilder')
    def test_update_ole_file_closes_on_exception(self, mock_ole_builder_class):
        """Test that update_ole_file properly closes builder on exception."""
        mock_builder = Mock()
        mock_ole_builder_class.return_value = mock_builder
        mock_builder.build_updated_ole.side_effect = Exception("Test error")
        
        with pytest.raises(Exception, match="Test error"):
            update_ole_file(self.test_ole_data, self.test_prism_data)
            
        mock_builder.close.assert_called_once()
        
    @patch('prism_ole_handler.utils.ole_builder.OLEBuilder')
    def test_update_ole_file_with_empty_data(self, mock_ole_builder_class):
        """Test update_ole_file with empty PRISM data."""
        mock_builder = Mock()
        mock_ole_builder_class.return_value = mock_builder
        mock_builder.build_updated_ole.return_value = b'updated_ole_data'
        
        result = update_ole_file(self.test_ole_data, b'')
        
        mock_builder.build_updated_ole.assert_called_once_with(b'')
        assert result == b'updated_ole_data'
        
    @patch('prism_ole_handler.utils.ole_builder.OLEBuilder')
    def test_update_ole_file_with_large_data(self, mock_ole_builder_class):
        """Test update_ole_file with large PRISM data."""
        mock_builder = Mock()
        mock_ole_builder_class.return_value = mock_builder
        mock_builder.build_updated_ole.return_value = b'updated_ole_data'
        
        large_data = b'x' * 10000  # 10KB of data
        result = update_ole_file(self.test_ole_data, large_data)
        
        mock_builder.build_updated_ole.assert_called_once_with(large_data)
        assert result == b'updated_ole_data'
        
    @patch('prism_ole_handler.utils.ole_builder.OLEBuilder')
    def test_update_ole_file_builder_initialization_error(self, mock_ole_builder_class):
        """Test update_ole_file when builder initialization fails."""
        mock_ole_builder_class.side_effect = Exception("Initialization error")
        
        with pytest.raises(Exception, match="Initialization error"):
            update_ole_file(self.test_ole_data, self.test_prism_data)