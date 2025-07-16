"""
Command-line interface for PRISM extraction.
"""

import argparse
import sys
import os
from ..core.extractor import PrismExtractor

try:
    import olefile

    HAS_OLEFILE = True
except ImportError:
    HAS_OLEFILE = False


def main():
    """Main entry point for extraction CLI."""
    if not HAS_OLEFILE:
        print("Warning: olefile not installed. OLE extraction will be limited.")
        print("Install with: pip install olefile")

    parser = argparse.ArgumentParser(
        description="Extract PRISM objects from PowerPoint files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s presentation.pptx
  %(prog)s presentation.pptx -o extracted_prism_files
  %(prog)s presentation.pptx --slide 2
  %(prog)s presentation.pptx --slide 2 --slide 3 --slide 5
  %(prog)s presentation.pptx --slides 2,3,5
  
The tool will extract PRISM objects from the specified slides, or all slides if none specified.
        """,
    )

    parser.add_argument("pptx_file", help="Path to the PowerPoint file")
    parser.add_argument(
        "-o",
        "--output",
        default="extracted_prism_objects",
        help="Output directory for extracted objects (default: extracted_prism_objects)",
    )
    parser.add_argument(
        "--slide",
        "-s",
        type=int,
        action="append",
        dest="individual_slides",
        help="Specific slide number to extract (can be used multiple times)",
    )
    parser.add_argument(
        "--slides",
        type=str,
        help='Comma-separated list of slide numbers to extract (e.g., "2,3,5")',
    )
    parser.add_argument(
        "--padding",
        "-p",
        type=int,
        default=3,
        help="Number of digits for zero-padding in filenames (default: 3, e.g., slide001_object001)",
    )

    args = parser.parse_args()

    if not os.path.exists(args.pptx_file):
        print(f"Error: File not found: {args.pptx_file}")
        sys.exit(1)

    # Parse slide numbers
    selected_slides = []
    if args.slides:
        # Handle comma-separated list
        try:
            selected_slides.extend([int(s.strip()) for s in args.slides.split(",")])
        except ValueError:
            print("Error: Invalid slide number format in --slides argument")
            sys.exit(1)

    if args.individual_slides:
        # Handle individual --slide arguments
        selected_slides.extend(args.individual_slides)

    # Remove duplicates and sort
    if selected_slides:
        selected_slides = sorted(list(set(selected_slides)))
        if any(s <= 0 for s in selected_slides):
            print("Error: Slide numbers must be positive integers")
            sys.exit(1)

    # Validate padding argument
    if args.padding < 1:
        print("Error: Padding must be at least 1 digit")
        sys.exit(1)

    extractor = PrismExtractor(args.pptx_file)
    extractor.extract_prism_objects(args.output, selected_slides, padding=args.padding)


if __name__ == "__main__":
    main()
