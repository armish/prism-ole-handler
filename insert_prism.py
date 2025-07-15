#!/usr/bin/env python3
import zipfile
import os
import sys
import argparse
import shutil
from pathlib import Path
import xml.etree.ElementTree as ET
import json
import tempfile
from ole_builder import update_ole_file

class PrismInserter:
    def __init__(self, pptx_path):
        self.pptx_path = Path(pptx_path)
        self.temp_dir = Path("temp_pptx_insert")
        self.backup_path = None
        
    def create_backup(self):
        """Create a backup of the original PPTX file"""
        backup_name = self.pptx_path.stem + "_backup" + self.pptx_path.suffix
        self.backup_path = self.pptx_path.parent / backup_name
        shutil.copy2(self.pptx_path, self.backup_path)
        print(f"Created backup: {self.backup_path}")
        
    def extract_pptx(self):
        """Extract PPTX file to temporary directory"""
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir)
        
        with zipfile.ZipFile(self.pptx_path, 'r') as zip_ref:
            zip_ref.extractall(self.temp_dir)
    
    def repack_pptx(self, output_path=None):
        """Repack the modified PPTX file"""
        if output_path is None:
            output_path = self.pptx_path
            
        # Create new PPTX
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(self.temp_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(self.temp_dir)
                    zipf.write(file_path, arcname)
        
        # Cleanup
        shutil.rmtree(self.temp_dir)
        print(f"Updated PPTX saved: {output_path}")
    
    def find_embedding_for_slide(self, slide_num):
        """Find the embedding file for a specific slide"""
        slides_dir = self.temp_dir / "ppt" / "slides"
        rels_dir = self.temp_dir / "ppt" / "slides" / "_rels"
        
        slide_file = slides_dir / f"slide{slide_num}.xml"
        rel_file = rels_dir / f"slide{slide_num}.xml.rels"
        
        if not rel_file.exists():
            return None
            
        # Parse relationships
        tree = ET.parse(rel_file)
        root = tree.getroot()
        
        for rel in root.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
            rel_type = rel.get("Type", "")
            target = rel.get("Target", "")
            
            if "oleObject" in rel_type:
                # Extract the embedding filename
                if "../embeddings/" in target:
                    embedding_name = target.replace("../embeddings/", "")
                    return embedding_name
        
        return None
    
    def update_ole_contents(self, ole_path, new_prism_data):
        """Update the CONTENTS stream in an OLE file with new PRISM data"""
        # Read the original OLE file
        with open(ole_path, 'rb') as f:
            ole_data = f.read()
        
        # Update the CONTENTS stream using the builder
        updated_ole = update_ole_file(ole_data, new_prism_data)
        
        # Save the updated OLE file
        with open(ole_path, 'wb') as f:
            f.write(updated_ole)
    
    def insert_prism_object(self, slide_num, prism_file_path):
        """Insert/update a PRISM object for a specific slide"""
        prism_path = Path(prism_file_path)
        
        if not prism_path.exists():
            print(f"Error: PRISM file not found: {prism_path}")
            return False
            
        # Find the embedding file for this slide
        embedding_name = self.find_embedding_for_slide(slide_num)
        
        if not embedding_name:
            print(f"Error: No embedded object found for slide {slide_num}")
            return False
            
        print(f"Found embedding for slide {slide_num}: {embedding_name}")
        
        # Read the new PRISM data
        with open(prism_path, 'rb') as f:
            prism_data = f.read()
        
        # Update the OLE file
        ole_path = self.temp_dir / "ppt" / "embeddings" / embedding_name
        
        if not ole_path.exists():
            print(f"Error: Embedding file not found: {ole_path}")
            return False
            
        try:
            self.update_ole_contents(ole_path, prism_data)
            print(f"✓ Updated PRISM object in slide {slide_num}")
            return True
        except Exception as e:
            print(f"Error updating OLE file: {e}")
            return False
    
    def batch_insert(self, updates):
        """Insert multiple PRISM objects at once"""
        # Create backup
        self.create_backup()
        
        # Extract PPTX
        print(f"\nExtracting: {self.pptx_path}")
        self.extract_pptx()
        
        # Perform updates
        success_count = 0
        for slide_num, prism_file in updates:
            print(f"\nUpdating slide {slide_num} with {prism_file}")
            if self.insert_prism_object(slide_num, prism_file):
                success_count += 1
        
        # Repack PPTX
        if success_count > 0:
            self.repack_pptx()
            print(f"\n{'='*50}")
            print(f"Successfully updated {success_count} PRISM objects")
        else:
            print("\nNo updates were successful. Original file unchanged.")
            shutil.rmtree(self.temp_dir)
            
        return success_count

def main():
    parser = argparse.ArgumentParser(
        description='Insert PRISM objects back into PowerPoint files',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Update a single slide
  %(prog)s presentation.pptx --slide 2 --prism updated_graph.pzfx
  
  # Update multiple slides
  %(prog)s presentation.pptx --slide 2 --prism graph1.pzfx --slide 3 --prism graph2.pzfx
  
  # Use a mapping file
  %(prog)s presentation.pptx --mapping updates.json
  
Mapping file format (updates.json):
{
  "updates": [
    {"slide": 2, "prism": "slide2_updated.pzfx"},
    {"slide": 3, "prism": "slide3_updated.pzfx"}
  ]
}
        """
    )
    
    parser.add_argument('pptx_file', help='Path to the PowerPoint file')
    parser.add_argument('--slide', '-s', type=int, action='append', 
                        help='Slide number to update')
    parser.add_argument('--prism', '-p', action='append',
                        help='PRISM file to insert')
    parser.add_argument('--mapping', '-m', 
                        help='JSON file with slide-to-prism mappings')
    parser.add_argument('--output', '-o',
                        help='Output PPTX file (default: overwrite original)')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.pptx_file):
        print(f"Error: File not found: {args.pptx_file}")
        sys.exit(1)
    
    # Build update list
    updates = []
    
    if args.mapping:
        # Load from mapping file
        with open(args.mapping, 'r') as f:
            mapping_data = json.load(f)
            for update in mapping_data.get('updates', []):
                updates.append((update['slide'], update['prism']))
    elif args.slide and args.prism:
        # Use command line arguments
        if len(args.slide) != len(args.prism):
            print("Error: Number of --slide and --prism arguments must match")
            sys.exit(1)
        updates = list(zip(args.slide, args.prism))
    else:
        print("Error: Specify either --slide/--prism pairs or --mapping file")
        sys.exit(1)
    
    # Perform insertion
    inserter = PrismInserter(args.pptx_file)
    if args.output and args.output != args.pptx_file:
        # If different output specified, don't create backup
        inserter.backup_path = None
        inserter.batch_insert(updates)
        # Save to different file
        output_path = Path(args.output)
        shutil.move(args.pptx_file, output_path)
        print(f"Saved to: {output_path}")
    else:
        inserter.batch_insert(updates)

if __name__ == "__main__":
    main()