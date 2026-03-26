#!/usr/bin/env python3
"""Generate the DSO DIIM Presentation (38 slides)."""
import os
import sys

# Ensure we can import sibling modules
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pptx import Presentation
from design_constants import SLIDE_WIDTH, SLIDE_HEIGHT
from content_data import build_all_slides

def main():
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    build_all_slides(prs)

    # Output path
    base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    output = os.path.join(base, "DSO_DIIM_Presentation.pptx")
    prs.save(output)
    print(f"Presentation saved to: {output}")
    print(f"Total slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
