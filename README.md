# PRISM Object Extractor for PowerPoint

Extract GraphPad PRISM objects from PowerPoint presentations on macOS.

## Background

PowerPoint for Mac doesn't support direct editing of embedded PRISM objects (OLE), unlike Windows. This tool extracts embedded PRISM objects from .pptx files so they can be edited in PRISM and re-embedded.

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### Extraction:
```bash
python extract_prism.py presentation.pptx -o output_folder
```

This provides:
- Better OLE compound document parsing
- Slide number tracking for each object
- More detailed extraction information

## Output

Extracted files will be saved with descriptive names:
- `slide1_object1.pzfx` - PRISM file from slide 1
- `slide2_object1.ole` - Unprocessed OLE file for investigation
- `object1_stream_Package.bin` - Extracted OLE stream

## How it works

1. PPTX files are ZIP archives containing XML and embedded objects
2. Embedded objects are stored in `ppt/embeddings/` as .bin files
3. These .bin files are often OLE compound documents containing PRISM data
4. The tool extracts and identifies PRISM XML data from these containers

## Insertion

To re-insert updated PRISM objects back into PowerPoint:

```bash
# Update a single slide
python insert_prism.py presentation.pptx --slide 2 --prism updated_graph.pzfx

# Update multiple slides
python insert_prism.py presentation.pptx --slide 2 --prism graph1.pzfx --slide 3 --prism graph2.pzfx
```

The insertion tool:
- Creates a backup of the original file
- Updates the OLE containers with new PRISM data
- Preserves the visual representation streams
- Maintains all PowerPoint relationships

## Complete Workflow

1. **Extract PRISM objects from PowerPoint:**
   ```bash
   python extract_prism.py presentation.pptx -o extracted_files
   ```

2. **Edit the extracted .pzfx files in GraphPad PRISM**

3. **Insert the updated files back into PowerPoint:**
   ```bash
   python insert_prism.py presentation.pptx --slide 2 --prism extracted_files/slide2_updated.pzfx
   ```

## Limitations

- OLE files have size constraints - very large PRISM files may not fit
- Complex OLE structures may require manual investigation
- Visual representation is preserved but may not reflect all changes