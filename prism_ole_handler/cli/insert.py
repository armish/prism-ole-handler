"""
Command-line interface for PRISM insertion.
"""

import argparse
import sys
import os
import json
import shutil
from pathlib import Path
from ..core.inserter import PrismInserter


def main():
    """Main entry point for insertion CLI."""
    parser = argparse.ArgumentParser(
        description="Insert PRISM objects back into PowerPoint files",
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
        """,
    )

    parser.add_argument("pptx_file", help="Path to the PowerPoint file")
    parser.add_argument(
        "--slide", "-s", type=int, action="append", help="Slide number to update"
    )
    parser.add_argument(
        "--prism", "-p", action="append", help="PRISM file to insert (.pzfx format)"
    )
    parser.add_argument(
        "--mapping", "-m", help="JSON file with slide-to-prism mappings"
    )
    parser.add_argument(
        "--output", "-o", help="Output PPTX file (default: overwrite original)"
    )
    parser.add_argument(
        "--create-new",
        action="store_true",
        help="Create new slides if they don't exist",
    )
    parser.add_argument(
        "--force-insert",
        action="store_true",
        help="Insert into slides even if they already have embeddings",
    )

    args = parser.parse_args()

    if not os.path.exists(args.pptx_file):
        print(f"Error: File not found: {args.pptx_file}")
        sys.exit(1)

    # Build update list
    updates = []

    if args.mapping:
        # Load from mapping file
        with open(args.mapping, "r") as f:
            mapping_data = json.load(f)
            for update in mapping_data.get("updates", []):
                updates.append((update["slide"], update["prism"]))
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
