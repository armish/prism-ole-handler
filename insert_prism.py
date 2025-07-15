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
    
    def get_slide_count(self):
        """Get the total number of slides in the presentation"""
        slides_dir = self.temp_dir / "ppt" / "slides"
        if not slides_dir.exists():
            return 0
        slide_files = list(slides_dir.glob("slide*.xml"))
        return len(slide_files)
    
    def slide_exists(self, slide_num):
        """Check if a slide exists"""
        slides_dir = self.temp_dir / "ppt" / "slides"
        slide_file = slides_dir / f"slide{slide_num}.xml"
        return slide_file.exists()
    
    def slide_has_embeddings(self, slide_num):
        """Check if a slide has existing embeddings"""
        return self.find_embedding_for_slide(slide_num) is not None
    
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
    
    def create_new_slide(self, slide_num):
        """Create a new slide with the specified number"""
        slides_dir = self.temp_dir / "ppt" / "slides"
        rels_dir = self.temp_dir / "ppt" / "slides" / "_rels"
        
        # Create directories if they don't exist
        slides_dir.mkdir(parents=True, exist_ok=True)
        rels_dir.mkdir(parents=True, exist_ok=True)
        
        # Create slide XML with basic structure
        slide_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name=""/>
                <p:cNvGrpSpPr/>
                <p:nvPr/>
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="0" cy="0"/>
                    <a:chOff x="0" y="0"/>
                    <a:chExt cx="0" cy="0"/>
                </a:xfrm>
            </p:grpSpPr>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping/>
    </p:clrMapOvr>
</p:sld>"""
        
        slide_file = slides_dir / f"slide{slide_num}.xml"
        with open(slide_file, 'w', encoding='utf-8') as f:
            f.write(slide_xml)
        
        # Update presentation.xml to include the new slide
        self.update_presentation_xml(slide_num)
        
        print(f"Created new slide {slide_num}")
        return True
    
    def update_presentation_xml(self, slide_num):
        """Update presentation.xml to include the new slide"""
        pres_file = self.temp_dir / "ppt" / "presentation.xml"
        
        if not pres_file.exists():
            print("Warning: presentation.xml not found, slide may not appear correctly")
            return
        
        # Parse and update presentation.xml
        tree = ET.parse(pres_file)
        root = tree.getroot()
        
        # Find the slide ID list
        slide_id_list = root.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst")
        
        if slide_id_list is not None:
            # Find the highest existing slide ID
            max_id = 256  # Start with a reasonable base
            for slide_id in slide_id_list.findall(".//{http://schemas.openxmlformats.org/presentationml/2006/main}sldId"):
                current_id = int(slide_id.get("id", "0"))
                if current_id > max_id:
                    max_id = current_id
            
            # Create new slide ID element
            new_slide_id = ET.SubElement(slide_id_list, "{http://schemas.openxmlformats.org/presentationml/2006/main}sldId")
            new_slide_id.set("id", str(max_id + 1))
            new_slide_id.set("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", f"rId{slide_num}")
            
            # Save the updated presentation.xml
            tree.write(pres_file, encoding='utf-8', xml_declaration=True)
    
    def insert_into_empty_slide(self, slide_num, prism_file_path):
        """Insert PRISM object into an empty slide"""
        prism_path = Path(prism_file_path)
        
        if not prism_path.exists():
            print(f"Error: PRISM file not found: {prism_path}")
            return False
        
        # Read the PRISM data
        with open(prism_path, 'rb') as f:
            prism_data = f.read()
        
        # Create embedding file
        embeddings_dir = self.temp_dir / "ppt" / "embeddings"
        embeddings_dir.mkdir(parents=True, exist_ok=True)
        
        # Find next available embedding filename
        existing_embeddings = list(embeddings_dir.glob("oleObject*.bin"))
        if existing_embeddings:
            nums = [int(f.stem.replace("oleObject", "")) for f in existing_embeddings if f.stem.replace("oleObject", "").isdigit()]
            next_num = max(nums) + 1 if nums else 1
        else:
            next_num = 1
        
        embedding_filename = f"oleObject{next_num}.bin"
        embedding_path = embeddings_dir / embedding_filename
        
        # Create OLE file with PRISM data
        self.create_ole_file(embedding_path, prism_data)
        
        # Update slide XML to include the embedded object
        self.add_embedded_object_to_slide(slide_num, embedding_filename, next_num)
        
        print(f"✓ Inserted PRISM object into slide {slide_num}")
        return True
    
    def create_ole_file(self, ole_path, prism_data):
        """Create a new OLE file with PRISM data"""
        # Use an existing OLE file as a template if available
        embeddings_dir = ole_path.parent
        existing_ole_files = list(embeddings_dir.glob("oleObject*.bin"))
        
        if existing_ole_files:
            # Use existing OLE file as template
            template_path = existing_ole_files[0]
            with open(template_path, 'rb') as f:
                template_data = f.read()
            
            # Update the template with new PRISM data
            updated_ole = update_ole_file(template_data, prism_data)
            
            with open(ole_path, 'wb') as f:
                f.write(updated_ole)
        else:
            # Try to use a template from the test files
            test_template = Path(__file__).parent / "test" / "test_01.pptx"
            if test_template.exists():
                print(f"Using template from {test_template}")
                # Extract a template OLE file from the test presentation
                template_ole = self.extract_template_ole(test_template)
                if template_ole:
                    updated_ole = update_ole_file(template_ole, prism_data)
                    with open(ole_path, 'wb') as f:
                        f.write(updated_ole)
                    return
            
            # Fallback: Create a minimal OLE structure
            print("Warning: Creating minimal OLE structure - may not work properly")
            ole_header = b'\x04\x2c\x00\x00'  # Common header pattern
            ole_data = ole_header + prism_data
            
            with open(ole_path, 'wb') as f:
                f.write(ole_data)
    
    def extract_template_ole(self, template_pptx):
        """Extract an OLE file from a template PPTX for use as a template"""
        try:
            import tempfile
            with tempfile.TemporaryDirectory() as temp_dir:
                with zipfile.ZipFile(template_pptx, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                embeddings_dir = Path(temp_dir) / "ppt" / "embeddings"
                if embeddings_dir.exists():
                    ole_files = list(embeddings_dir.glob("oleObject*.bin"))
                    if ole_files:
                        with open(ole_files[0], 'rb') as f:
                            return f.read()
        except Exception as e:
            print(f"Could not extract template OLE: {e}")
        return None
    
    def add_embedded_object_to_slide(self, slide_num, embedding_filename, object_id):
        """Add embedded object reference to slide XML"""
        slides_dir = self.temp_dir / "ppt" / "slides"
        rels_dir = self.temp_dir / "ppt" / "slides" / "_rels"
        
        slide_file = slides_dir / f"slide{slide_num}.xml"
        rel_file = rels_dir / f"slide{slide_num}.xml.rels"
        
        # Create relationships file if it doesn't exist
        if not rel_file.exists():
            # Create the _rels directory if it doesn't exist
            rel_file.parent.mkdir(parents=True, exist_ok=True)
            rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""
            with open(rel_file, 'w', encoding='utf-8') as f:
                f.write(rels_xml)
        
        # Add relationship for the embedded object
        tree = ET.parse(rel_file)
        root = tree.getroot()
        
        # Create new relationship
        rel_element = ET.SubElement(root, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
        rel_element.set("Id", f"rId{object_id}")
        rel_element.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject")
        rel_element.set("Target", f"../embeddings/{embedding_filename}")
        
        # Save relationships file
        tree.write(rel_file, encoding='utf-8', xml_declaration=True)
        
        # Update slide XML to include the embedded object
        self.add_object_to_slide_xml(slide_num, object_id)
    
    def add_object_to_slide_xml(self, slide_num, object_id):
        """Add embedded object to slide XML content"""
        slides_dir = self.temp_dir / "ppt" / "slides"
        slide_file = slides_dir / f"slide{slide_num}.xml"
        
        tree = ET.parse(slide_file)
        root = tree.getroot()
        
        # Find the shape tree
        sp_tree = root.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}spTree")
        
        if sp_tree is not None:
            # Create embedded object shape
            ole_shape_xml = f"""<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <p:nvSpPr>
                    <p:cNvPr id="{object_id + 1}" name="PRISM Object {object_id}"/>
                    <p:cNvSpPr/>
                    <p:nvPr>
                        <p:extLst>
                            <p:ext uri="{{D42A27DB-BD31-4B8C-83A1-26A3C78FFF5A}}">
                                <p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1"/>
                            </p:ext>
                        </p:extLst>
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="1270000" y="1270000"/>
                        <a:ext cx="7620000" cy="5715000"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                        <a:avLst/>
                    </a:prstGeom>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr rtlCol="0" anchor="ctr"/>
                    <a:lstStyle/>
                    <a:p>
                        <a:pPr algn="ctr"/>
                    </a:p>
                </p:txBody>
            </p:sp>"""
            
            # Parse and add the OLE shape
            ole_element = ET.fromstring(ole_shape_xml)
            sp_tree.append(ole_element)
            
            # Save the updated slide
            tree.write(slide_file, encoding='utf-8', xml_declaration=True)
    
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
    
    def insert_prism_object(self, slide_num, prism_file_path, create_new=False, force_insert=False):
        """Insert/update a PRISM object for a specific slide"""
        prism_path = Path(prism_file_path)
        
        if not prism_path.exists():
            print(f"Error: PRISM file not found: {prism_path}")
            return False
        
        # Check if the file is a .prism file (need to convert to .pzfx format)
        if prism_path.suffix.lower() == '.prism':
            print(f"Error: Cannot insert .prism files directly. Please use .pzfx files.")
            print(f"Tip: Open {prism_path.name} in PRISM and save/export as .pzfx format")
            return False
        
        # Check if slide exists
        if not self.slide_exists(slide_num):
            if create_new:
                print(f"Creating new slide {slide_num}")
                self.create_new_slide(slide_num)
            else:
                print(f"Error: Slide {slide_num} does not exist. Use --create-new to create it.")
                return False
        
        # Check if slide has existing embeddings
        has_embeddings = self.slide_has_embeddings(slide_num)
        
        if has_embeddings and not force_insert:
            # Replace existing embedding
            embedding_name = self.find_embedding_for_slide(slide_num)
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
        else:
            # Insert into empty slide or create new
            if has_embeddings:
                print(f"Warning: Slide {slide_num} already has embeddings. Adding new object.")
            
            return self.insert_into_empty_slide(slide_num, prism_file_path)
    
    def batch_insert(self, updates, create_new=False, force_insert=False):
        """Insert multiple PRISM objects at once"""
        # Create backup
        self.create_backup()
        
        # Extract PPTX
        print(f"\nExtracting: {self.pptx_path}")
        self.extract_pptx()
        
        # Show current slide information
        slide_count = self.get_slide_count()
        print(f"Current presentation has {slide_count} slides")
        
        # Perform updates
        success_count = 0
        for slide_num, prism_file in updates:
            print(f"\nProcessing slide {slide_num} with {prism_file}")
            if self.insert_prism_object(slide_num, prism_file, create_new, force_insert):
                success_count += 1
        
        # Repack PPTX
        if success_count > 0:
            self.repack_pptx()
            print(f"\n{'='*50}")
            print(f"Successfully processed {success_count} PRISM objects")
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
  # Update an existing slide
  %(prog)s presentation.pptx --slide 2 --prism updated_graph.pzfx
  
  # Insert into empty slide
  %(prog)s presentation.pptx --slide 3 --prism new_graph.pzfx
  
  # Create new slide with PRISM object
  %(prog)s presentation.pptx --slide 10 --prism graph.pzfx --create-new
  
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
                        help='PRISM file to insert (.pzfx format)')
    parser.add_argument('--mapping', '-m', 
                        help='JSON file with slide-to-prism mappings')
    parser.add_argument('--output', '-o',
                        help='Output PPTX file (default: overwrite original)')
    parser.add_argument('--create-new', action='store_true',
                        help='Create new slides if they don\'t exist')
    parser.add_argument('--force-insert', action='store_true',
                        help='Insert into slides even if they already have embeddings')
    
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
        inserter.batch_insert(updates, args.create_new, args.force_insert)
        # Save to different file
        output_path = Path(args.output)
        shutil.move(args.pptx_file, output_path)
        print(f"Saved to: {output_path}")
    else:
        inserter.batch_insert(updates, args.create_new, args.force_insert)

if __name__ == "__main__":
    main()