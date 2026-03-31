"""
annotate_timecard.py

Inserts a correction annotation below every "Approved by:" label in a PDF.

Usage:
    python3 annotate_timecard.py \
        --input  "Input/Complete_with_Docusign_EDITEDdocx.pdf" \
        --output "Output/Complete_with_Docusign_EDITEDdocx_corrected.pdf" \
        --adjustment 2.0 \
        --corrected_total 42.3

The text inserted reads:
    Hours correction = <adjustment> | New total hours = <corrected_total> | Reason for correction = Unrecorded PTO
"""

import argparse
import os
import fitz  # PyMuPDF


ANNOTATION_TEMPLATE = (
    "Hours correction = {adjustment} | "
    "New total hours = {corrected_total} | "
    "Reason for correction = Unrecorded PTO"
)

# Vertical gap between bottom of "Approved by:" text and top of annotation
Y_OFFSET = 4  # points


def find_approved_by_instances(page):
    """
    Return one rect per distinct (panel, timesheet-row) combination.

    The PDF has two side-by-side panels per page.  Within each panel there may
    be multiple "Approved by:" labels at nearly the same Y (DocuSign artefact);
    those should collapse to one annotation.  But the left panel and right panel
    are independent and each need their own annotation even if their Y values are
    close.

    Strategy
    --------
    1. Split rects into LEFT panel (x0 < PAGE_MID) and RIGHT panel (x0 >= PAGE_MID).
    2. Within each panel, cluster by Y proximity and keep the topmost rect per cluster.
    3. Return all surviving rects from both panels.
    """
    all_rects = page.search_for("Approved by:")
    if not all_rects:
        return []

    page_mid = page.rect.width / 2
    Y_CLUSTER_TOLERANCE = 10  # points

    def cluster_by_y(rects):
        rects = sorted(rects, key=lambda r: r.y0)
        groups = []
        for rect in rects:
            placed = False
            for group in groups:
                if abs(rect.y0 - group[0].y0) <= Y_CLUSTER_TOLERANCE:
                    group.append(rect)
                    placed = True
                    break
            if not placed:
                groups.append([rect])
        # Keep the topmost (lowest y0) rect from each cluster
        return [min(g, key=lambda r: r.y0) for g in groups]

    left_rects  = [r for r in all_rects if r.x0 <  page_mid]
    right_rects = [r for r in all_rects if r.x0 >= page_mid]

    return cluster_by_y(left_rects) + cluster_by_y(right_rects)


def insert_annotation(page, rect, annotation_text, font_size):
    """Insert annotation text just below the given rect."""
    x = rect.x0
    y = rect.y1 + Y_OFFSET + font_size  # fitz inserts text at baseline

    page.insert_text(
        (x, y),
        annotation_text,
        fontsize=font_size,
        color=(0, 0, 0),
    )


def annotate_pdf(input_path, output_path, adjustment, corrected_total):
    doc = fitz.open(input_path)

    annotation_text = ANNOTATION_TEMPLATE.format(
        adjustment=adjustment,
        corrected_total=corrected_total,
    )

    total_insertions = 0

    for page_num, page in enumerate(doc, start=1):
        instances = find_approved_by_instances(page)
        if not instances:
            continue

        for rect in instances:
            # Measure the font size of the "Approved by:" text by inspecting
            # the text blocks on this page and matching by bbox proximity.
            font_size = get_font_size_near_rect(page, rect)
            insert_annotation(page, rect, annotation_text, font_size)
            total_insertions += 1
            print(
                f"  Page {page_num}: inserted below 'Approved by:' at y={rect.y1:.1f} "
                f"(font size {font_size:.1f}pt)"
            )

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    doc.save(output_path)
    doc.close()

    print(f"\nDone. {total_insertions} annotation(s) inserted.")
    print(f"Output saved to: {output_path}")


def get_font_size_near_rect(page, target_rect, tolerance=5):
    """
    Walk the page's text blocks and find the font size of the span
    whose bbox most closely matches target_rect.
    Falls back to 8pt if no match found.
    """
    blocks = page.get_text("rawdict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
    best_size = 8.0
    best_dist = float("inf")

    for block in blocks:
        if block.get("type") != 0:  # 0 = text block
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                span_rect = fitz.Rect(span["bbox"])
                # Check horizontal and vertical overlap with target_rect
                dx = abs(span_rect.x0 - target_rect.x0)
                dy = abs(span_rect.y0 - target_rect.y0)
                dist = (dx**2 + dy**2) ** 0.5
                if dist < best_dist and dist < 20:
                    best_dist = dist
                    best_size = span["size"]

    return best_size


def main():
    parser = argparse.ArgumentParser(
        description="Insert hours-correction annotation below every 'Approved by:' in a timecard PDF."
    )
    parser.add_argument("--input", required=True, help="Path to the input PDF")
    parser.add_argument("--output", required=True, help="Path for the output PDF")
    parser.add_argument(
        "--adjustment",
        required=True,
        help="Hours adjustment value (e.g. 2.0 or -1.5)",
    )
    parser.add_argument(
        "--corrected_total",
        required=True,
        help="New total hours after correction (e.g. 44.3)",
    )
    args = parser.parse_args()

    print(f"Input:  {args.input}")
    print(f"Output: {args.output}")
    print(f"Adjustment:      {args.adjustment}")
    print(f"Corrected total: {args.corrected_total}")
    print()

    annotate_pdf(args.input, args.output, args.adjustment, args.corrected_total)


if __name__ == "__main__":
    main()
