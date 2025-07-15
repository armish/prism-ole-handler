#!/usr/bin/env python3
import zipfile
import os
import sys
import argparse
import shutil
from pathlib import Path
import xml.etree.ElementTree as ET
try:
    import olefile
    HAS_OLEFILE = True
except ImportError:
    HAS_OLEFILE = False
    print("Warning: olefile not installed. OLE extraction will be limited.")
    print("Install with: pip install olefile")

class AdvancedPrismExtractor:
    def __init__(self, pptx_path):
        self.pptx_path = Path(pptx_path)
        self.temp_dir = Path("temp_pptx_extract")
        
    def extract_pptx(self):
        """Extract PPTX file to temporary directory"""
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir)
        
        with zipfile.ZipFile(self.pptx_path, 'r') as zip_ref:
            zip_ref.extractall(self.temp_dir)
    
    def analyze_slide_relationships(self):
        """Analyze slide relationships to understand embedded objects"""
        slides_info = []
        slides_dir = self.temp_dir / "ppt" / "slides"
        rels_dir = self.temp_dir / "ppt" / "slides" / "_rels"
        
        if slides_dir.exists():
            for slide_file in sorted(slides_dir.glob("slide*.xml")):
                slide_num = slide_file.stem.replace("slide", "")
                
                # Check relationships
                rel_file = rels_dir / f"{slide_file.name}.rels"
                embedded_refs = []
                
                if rel_file.exists():
                    tree = ET.parse(rel_file)
                    root = tree.getroot()
                    
                    for rel in root.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                        rel_type = rel.get("Type", "")
                        target = rel.get("Target", "")
                        
                        if "oleObject" in rel_type or "package" in rel_type:
                            embedded_refs.append({
                                "type": rel_type,
                                "target": target,
                                "id": rel.get("Id", "")
                            })
                
                if embedded_refs:
                    slides_info.append({
                        "slide": int(slide_num),
                        "file": slide_file.name,
                        "embedded_objects": embedded_refs
                    })
        
        return slides_info
            
    def find_embedded_objects(self):
        """Find all embedded objects in the PPTX structure with metadata"""
        embeddings_dir = self.temp_dir / "ppt" / "embeddings"
        embedded_objects = []
        
        # Get slide relationship information
        slides_info = self.analyze_slide_relationships()
        
        if embeddings_dir.exists():
            for file in embeddings_dir.iterdir():
                if file.suffix == ".bin":
                    # Try to find which slide this belongs to
                    slide_num = None
                    for slide in slides_info:
                        for ref in slide["embedded_objects"]:
                            if file.name in ref["target"]:
                                slide_num = slide["slide"]
                                break
                    
                    embedded_objects.append({
                        "path": file,
                        "slide": slide_num
                    })
                    
        return embedded_objects, slides_info
    
    def extract_from_ole(self, ole_data, base_name):
        """Extract PRISM data from OLE file using olefile"""
        if not HAS_OLEFILE:
            return []
        
        extracted = []
        
        try:
            ole = olefile.OleFileIO(ole_data)
            
            # List all streams in the OLE file
            streams = ole.listdir()
            print(f"  OLE streams found: {streams}")
            
            for stream_path in streams:
                stream_name = "/".join(stream_path)
                
                # Common locations for embedded content
                if any(keyword in stream_name.lower() for keyword in ['package', 'contents', 'prism', 'graphpad']):
                    try:
                        stream_data = ole.openstream(stream_path).read()
                        
                        # Check if it's PRISM data
                        if self.is_prism_data(stream_data):
                            # Skip the first 4 bytes which appear to be a length header
                            if stream_data[:2] == b'\x04\x30':
                                stream_data = stream_data[4:]
                            extracted.append((f"{base_name}_prism.pzfx", stream_data))
                        else:
                            # Save for investigation
                            safe_name = stream_name.replace("/", "_").replace("\\", "_")
                            extracted.append((f"{base_name}_stream_{safe_name}.bin", stream_data))
                    except Exception as e:
                        print(f"  Error reading stream {stream_name}: {e}")
            
            ole.close()
            
        except Exception as e:
            print(f"  Error processing OLE file: {e}")
        
        return extracted
    
    def extract_prism_objects(self, output_dir):
        """Extract PRISM objects from PPTX file"""
        output_path = Path(output_dir)
        output_path.mkdir(exist_ok=True)
        
        print(f"Extracting from: {self.pptx_path}")
        self.extract_pptx()
        
        embedded_objects, slides_info = self.find_embedded_objects()
        print(f"\nFound {len(embedded_objects)} embedded objects")
        
        if slides_info:
            print("\nEmbedded objects by slide:")
            for slide in slides_info:
                print(f"  Slide {slide['slide']}: {len(slide['embedded_objects'])} objects")
        
        prism_count = 0
        
        for idx, obj_info in enumerate(embedded_objects):
            obj_path = obj_info["path"]
            slide_num = obj_info["slide"]
            
            # Create descriptive name
            if slide_num:
                base_name = f"slide{slide_num}_object{idx + 1}"
            else:
                base_name = f"object{idx + 1}"
            
            print(f"\nProcessing: {obj_path.name} ({base_name})")
            
            # Read the embedded object
            with open(obj_path, 'rb') as f:
                data = f.read()
            
            # Check if it's an OLE file
            if data[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
                print("  Detected OLE compound file")
                extracted = self.extract_from_ole(data, base_name)
                
                for name, content in extracted:
                    output_file = output_path / name
                    with open(output_file, 'wb') as f:
                        f.write(content)
                    
                    if name.endswith('.pzfx'):
                        print(f"  ✓ Extracted PRISM object: {output_file}")
                        prism_count += 1
                    else:
                        print(f"  - Saved stream: {output_file}")
                
                if not extracted:
                    # Save the OLE file for manual investigation
                    output_file = output_path / f"{base_name}.ole"
                    with open(output_file, 'wb') as f:
                        f.write(data)
                    print(f"  - Saved OLE file: {output_file}")
                    
            elif self.is_prism_xml(data):
                # Direct PRISM XML file
                output_file = output_path / f"{base_name}.pzfx"
                with open(output_file, 'wb') as f:
                    f.write(data)
                print(f"  ✓ Extracted PRISM object: {output_file}")
                prism_count += 1
                
            else:
                # Unknown format
                output_file = output_path / f"{base_name}.bin"
                with open(output_file, 'wb') as f:
                    f.write(data)
                print(f"  - Saved unknown format: {output_file}")
                
                # Try to identify the format
                print(f"  - First 16 bytes: {data[:16].hex()}")
        
        # Cleanup
        shutil.rmtree(self.temp_dir)
        
        print(f"\n{'='*50}")
        print(f"Extraction complete!")
        print(f"Total PRISM objects extracted: {prism_count}")
        print(f"Output directory: {output_path}")
        
    def is_prism_xml(self, data):
        """Check if data is PRISM XML format"""
        if data.startswith(b'<?xml'):
            try:
                xml_str = data.decode('utf-8', errors='ignore')
                if 'GraphPadPrismFile' in xml_str or 'PrismFile' in xml_str:
                    return True
            except:
                pass
        return False
    
    def is_prism_data(self, data):
        """Check if data is PRISM format (ZIP or XML)"""
        # Check for ZIP header (with possible 4-byte prefix)
        if len(data) > 6:
            if data[:2] == b'PK' or data[4:6] == b'PK':
                return True
        # Check for XML
        return self.is_prism_xml(data)

def main():
    parser = argparse.ArgumentParser(
        description='Extract PRISM objects from PowerPoint files',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s presentation.pptx
  %(prog)s presentation.pptx -o extracted_prism_files
  
The tool will extract all PRISM objects found in the PowerPoint file,
organizing them by slide number when possible.
        """
    )
    
    parser.add_argument('pptx_file', help='Path to the PowerPoint file')
    parser.add_argument('-o', '--output', default='extracted_prism_objects', 
                        help='Output directory for extracted objects (default: extracted_prism_objects)')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.pptx_file):
        print(f"Error: File not found: {args.pptx_file}")
        sys.exit(1)
    
    extractor = AdvancedPrismExtractor(args.pptx_file)
    extractor.extract_prism_objects(args.output)

if __name__ == "__main__":
    main()