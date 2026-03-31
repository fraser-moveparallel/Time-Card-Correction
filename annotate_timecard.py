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


ANNOTATION_LINE1 = "Hours correction = {adjustment} | New total hours = {corrected_total}"
ANNOTATION_LINE2 = "Reason for correction = Unrecorded PTO"

# Vertical gap between bottom of "Approved by:" text and top of annotation
Y_OFFSET = 4  # points


def find_approved_by_instances(page):
    """
    Return one rect per distinct (panel, timesheet-row) combination.

    Strategy
    --------
    1. Find the natural X gap between the left and right panels by looking at
       the largest gap in sorted x0 values.  This handles pages where the right
       panel starts before the nominal page midpoint.
    2. Split rects into LEFT / RIGHT panels at that gap.
    3. Within each panel, cluster by Y proximity (DocuSign duplicate artefacts)
       and keep the topmost rect per cluster.
    4. Return one representative rect per surviving cluster.
    """
    all_rects = page.search_for("Approved by:")
    if not all_rects:
        return []

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
        return [min(g, key=lambda r: r.y0) for g in groups]

    # Find the panel split by locating the largest gap in x0 values
    sorted_x = sorted(set(round(r.x0) for r in all_rects))
    if len(sorted_x) > 1:
        gaps = [(sorted_x[i+1] - sorted_x[i], sorted_x[i], sorted_x[i+1])
                for i in range(len(sorted_x) - 1)]
        largest_gap = max(gaps, key=lambda g: g[0])
        x_split = (largest_gap[1] + largest_gap[2]) / 2
    else:
        x_split = page.rect.width / 2

    left_rects  = [r for r in all_rects if r.x0 <= x_split]
    right_rects = [r for r in all_rects if r.x0 >  x_split]

    return cluster_by_y(left_rects) + cluster_by_y(right_rects)


def insert_annotation(page, rect, line1, line2, font_size):
    """Insert two-line annotation just below the given rect."""
    x = rect.x0
    line_height = font_size * 1.2  # standard leading

    # First line
    y1 = rect.y1 + Y_OFFSET + font_size
    page.insert_text((x, y1), line1, fontsize=font_size, color=(0, 0, 0))

    # Second line
    y2 = y1 + line_height
    page.insert_text((x, y2), line2, fontsize=font_size, color=(0, 0, 0))


def annotate_pdf(input_path, output_path, adjustment, corrected_total):
    doc = fitz.open(input_path)

    line1 = ANNOTATION_LINE1.format(adjustment=adjustment, corrected_total=corrected_total)
    line2 = ANNOTATION_LINE2

    total_insertions = 0

    for page_num, page in enumerate(doc, start=1):
        instances = find_approved_by_instances(page)
        if not instances:
            continue

        for rect in instances:
            font_size = get_font_size_near_rect(page, rect)
            insert_annotation(page, rect, line1, line2, font_size)
            total_insertions += 1
            print(
                f"  Page {page_num}: inserted below 'Approved by:' at y={rect.y1:.1f} "
                f"x={rect.x0:.1f} (font size {font_size:.1f}pt)"
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
