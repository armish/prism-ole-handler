# Development with Claude Code

This document describes the development process of the PRISM OLE Handler package, which was built collaboratively with [Claude Code](https://claude.ai/code).

## Project Overview

**Goal**: Create a command-line tool and Python package to extract and insert GraphPad PRISM objects from Microsoft Office documents (PowerPoint, Word, Excel) on macOS.

**Problem**: Microsoft Office for Mac doesn't support direct editing of embedded PRISM objects (OLE), unlike Windows. Users needed a programmatic solution to extract PRISM objects, edit them, and re-embed them.

## Development Timeline

### Phase 1: Initial Extraction Tool (Session 1)
- **Challenge**: Understanding PowerPoint's internal structure and OLE compound documents
- **Solution**: Built `extract_prism.py` with ZIP file extraction and OLE parsing using `olefile`
- **Success**: Successfully extracted 4 PRISM objects from test PowerPoint file
- **Key Learning**: PPTX files are ZIP archives containing XML and embedded `.bin` files

### Phase 2: Insertion Tool Development (Session 1)
- **Challenge**: Reverse-engineering PowerPoint's OLE embedding format
- **Solution**: Created `insert_prism.py` with OLE file reconstruction capabilities
- **Obstacles**: PowerPoint corruption issues, missing `_rels` directory structure
- **Resolution**: Proper PowerPoint relationship handling and file validation

### Phase 3: Package Structure Refactoring (Session 2)
- **Motivation**: Make the tool installable and extensible for future Word/Excel support
- **Changes**:
  - Renamed from `prism-powerpoint` to `prism-ole-handler`
  - Created proper Python package structure
  - Added `setup.py` and `pyproject.toml` for pip installation
  - Converted scripts to console entry points (`prism-extract`, `prism-insert`)
  - Added MIT License and proper metadata

### Phase 4: Testing Infrastructure (Session 3)
- **Goal**: Add comprehensive unit tests and CI/CD
- **Implementation**:
  - Created 58 unit tests across 4 test modules
  - Added GitHub Actions CI for multi-platform testing
  - Configured code quality tools (flake8, black, mypy)
  - Set up package building and validation

### Phase 5: Test Fixes and CI Optimization (Session 3)
- **Problem**: Tests failed in CI due to mocking assumptions about internal APIs
- **Solution**: Rewrote tests to match actual implementation behavior
- **CI Improvements**:
  - Updated deprecated GitHub Actions (v3→v4, v4→v5)
  - Removed Python 3.8 support, focused on 3.9+
  - Streamlined CI matrix for faster builds

## Architecture

### Core Components

1. **`PrismExtractor`** (`prism_ole_handler/core/extractor.py`)
   - Extracts PPTX files as ZIP archives
   - Parses slide relationships to find embedded objects
   - Uses `olefile` to extract PRISM data from OLE compound documents
   - Supports selective extraction by slide number

2. **`PrismInserter`** (`prism_ole_handler/core/inserter.py`)
   - Inserts/updates PRISM objects in PowerPoint slides
   - Creates new slides when needed
   - Handles OLE file reconstruction with updated PRISM data
   - Maintains PowerPoint relationships and structure

3. **`OLEBuilder`** (`prism_ole_handler/utils/ole_builder.py`)
   - Rebuilds OLE compound documents with new PRISM data
   - Preserves existing streams while updating CONTENTS stream
   - Handles size constraints and file structure integrity

4. **CLI Tools** (`prism_ole_handler/cli/`)
   - `prism-extract`: Command-line extraction tool
   - `prism-insert`: Command-line insertion tool
   - Support for batch operations and JSON mapping files

### File Format Support

- **Extraction**: Creates `.pzfx` files compatible with GraphPad PRISM
- **Insertion**: Requires `.pzfx` files (not `.prism` files)
- **PowerPoint**: Full support for `.pptx` files
- **Future**: Designed for extensibility to Word (`.docx`) and Excel (`.xlsx`)

## Technical Challenges & Solutions

### 1. PowerPoint Structure Understanding
- **Challenge**: PPTX files are complex ZIP archives with XML relationships
- **Solution**: Systematic analysis of PowerPoint's internal structure
- **Tools**: ZIP extraction, XML parsing, relationship mapping

### 2. OLE Compound Document Manipulation
- **Challenge**: OLE files have complex binary structure with sectors and streams
- **Solution**: Used `olefile` library for parsing, custom builder for reconstruction
- **Key Insight**: PRISM data is stored in CONTENTS stream with 4-byte header

### 3. File Size Constraints
- **Challenge**: OLE files have size limitations that can cause corruption
- **Solution**: Implemented file size validation and rebuilding strategies
- **Approach**: Modify existing files in-place when possible, rebuild when necessary

### 4. Cross-Platform Compatibility
- **Challenge**: Different behavior on macOS, Windows, and Linux
- **Solution**: Comprehensive CI testing across all platforms
- **Focus**: Primary target is macOS (where the problem exists)

## Testing Strategy

### Unit Tests (58 tests)
- **`test_extractor.py`**: PrismExtractor functionality and edge cases
- **`test_inserter.py`**: PrismInserter operations and error handling
- **`test_ole_builder.py`**: OLE file manipulation and reconstruction
- **`test_cli.py`**: Command-line interface and argument parsing

### Integration Tests
- **Real-world scenarios**: Extraction and insertion with actual PowerPoint files
- **End-to-end workflows**: Complete extract-edit-insert cycles
- **Error handling**: File corruption, missing files, invalid formats

### CI/CD Pipeline
- **Multi-platform**: Ubuntu, Windows, macOS
- **Multi-Python**: Python 3.9, 3.10, 3.11, 3.12
- **Code quality**: Linting, formatting, type checking
- **Package validation**: Building and distribution testing

## Key Learnings

### 1. File Format Complexity
- Modern office documents are sophisticated archives with intricate relationships
- Binary formats like OLE require specialized libraries and careful handling
- PowerPoint's embedding system preserves both data and visual representations

### 2. Testing Challenges
- Mocking complex file operations requires understanding actual implementation
- Integration tests are crucial for file manipulation tools
- CI environments can behave differently than local development

### 3. Package Design
- Extensible naming and structure enable future feature additions
- Console entry points provide professional CLI experience
- Proper metadata and documentation enhance usability

### 4. Error Handling
- File corruption is a primary concern with binary manipulation
- Backup strategies are essential for data safety
- Clear error messages improve user experience

## Future Enhancements

### Planned Features
1. **Microsoft Word Support** (`.docx` files)
2. **Microsoft Excel Support** (`.xlsx` files)
3. **Batch processing improvements**
4. **GUI interface** for non-technical users
5. **Plugin system** for other scientific software

### Technical Improvements
1. **Performance optimization** for large files
2. **Memory efficiency** for batch operations
3. **Better error recovery** mechanisms
4. **Enhanced OLE rebuilding** algorithms

## Usage Examples

### Basic Extraction
```bash
# Extract all PRISM objects
prism-extract presentation.pptx -o extracted_files

# Extract from specific slides
prism-extract presentation.pptx --slides 2,3,5 -o extracted_files
```

### Basic Insertion
```bash
# Update existing embedding
prism-insert presentation.pptx --slide 2 --prism updated_graph.pzfx

# Create new slide with PRISM object
prism-insert presentation.pptx --slide 10 --prism new_graph.pzfx --create-new
```

### Python API
```python
from prism_ole_handler import PrismExtractor, PrismInserter

# Extract PRISM objects
extractor = PrismExtractor("presentation.pptx")
extractor.extract_prism_objects("output_folder", selected_slides=[2, 3])

# Insert PRISM objects
inserter = PrismInserter("presentation.pptx")
inserter.insert_prism_object(slide_num=2, prism_file_path="graph.pzfx")
```

## Acknowledgments

This project was developed collaboratively with Claude Code, Anthropic's AI assistant for software development. The iterative development process, from initial proof-of-concept to production-ready package, demonstrates the effectiveness of AI-assisted development for complex technical challenges.

## Resources

- **Repository**: https://github.com/armish/prism-ole-handler
- **PyPI Package**: https://pypi.org/project/prism-ole-handler/
- **Documentation**: See README.md for detailed usage instructions
- **Issue Tracker**: https://github.com/armish/prism-ole-handler/issues

---

*This document serves as both a development log and a guide for future contributors to understand the project's evolution and technical decisions.*