#!/usr/bin/env python3
"""
Build new OLE files with updated PRISM data
"""
import struct
import io
import olefile
import tempfile
from pathlib import Path


class OLEBuilder:
    """Build new OLE files with updated PRISM data while preserving structure"""

    def __init__(self, original_ole_data):
        self.original_ole = olefile.OleFileIO(original_ole_data)

    def build_updated_ole(self, new_prism_data):
        """Build a new OLE file with updated PRISM data"""
        # Create a temporary file to build the new OLE
        with tempfile.NamedTemporaryFile() as temp_file:
            # Write OLE header
            self._write_ole_header(temp_file)

            # Copy all streams except CONTENTS
            streams_to_copy = []
            contents_stream = None

            for stream_path in self.original_ole.listdir():
                stream_name = "/".join(stream_path)
                if stream_name == "CONTENTS":
                    contents_stream = stream_path
                else:
                    stream_data = self.original_ole.openstream(stream_path).read()
                    streams_to_copy.append((stream_path, stream_data))

            # Create new CONTENTS stream
            # Determine the correct header by looking at the original
            if contents_stream:
                original_contents = self.original_ole.openstream(contents_stream).read()
                header = original_contents[:4]
            else:
                header = b"\x04\x2c\x00\x00"  # Default header

            new_contents_data = header + new_prism_data

            # Build the new OLE file using a simpler approach
            return self._build_ole_compound_file(streams_to_copy, new_contents_data)

    def _build_ole_compound_file(self, other_streams, contents_data):
        """Build OLE compound file with the given streams"""
        # For now, we'll use a workaround by modifying the existing file
        # This is a simplified approach that works for most cases

        # Read the original file data
        self.original_ole.fp.seek(0)
        original_data = bytearray(self.original_ole.fp.read())

        # Find the CONTENTS stream position
        contents_stream_data = self.original_ole.openstream(["CONTENTS"]).read()
        pattern = contents_stream_data[:4] + b"PK"

        pos = original_data.find(pattern)
        if pos == -1:
            raise ValueError("Could not find CONTENTS stream")

        # Calculate the space needed
        original_size = len(contents_stream_data)
        new_size = len(contents_data)
        size_diff = new_size - original_size

        if size_diff > 0:
            # We need more space - expand the file
            # Insert the additional bytes at the end of the CONTENTS stream
            insert_pos = pos + original_size
            original_data[insert_pos:insert_pos] = b"\x00" * size_diff

        # Replace the CONTENTS stream
        original_data[pos : pos + len(contents_data)] = contents_data

        # If we made the file smaller, we might need to adjust
        if size_diff < 0:
            # Fill the remaining space with zeros
            fill_start = pos + len(contents_data)
            fill_end = fill_start + abs(size_diff)
            for i in range(fill_start, fill_end):
                if i < len(original_data):
                    original_data[i] = 0

        return bytes(original_data)

    def _write_ole_header(self, file_obj):
        """Write OLE compound file header"""
        # OLE compound file signature
        file_obj.write(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")
        # This is a simplified header - in practice, you'd need to calculate
        # proper directory entries, sector chains, etc.

    def close(self):
        """Close the original OLE file"""
        self.original_ole.close()


def update_ole_file(original_ole_data, new_prism_data):
    """Update an OLE file with new PRISM data"""
    builder = OLEBuilder(original_ole_data)
    try:
        updated_ole = builder.build_updated_ole(new_prism_data)
        return updated_ole
    finally:
        builder.close()
