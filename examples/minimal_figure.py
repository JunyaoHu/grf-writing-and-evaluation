"""Create an editable, fictional proposal diagram; no PDF export or real evidence."""
from pathlib import Path
import argparse
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from pptx_figure_kit import new_slide, round_rect, label, arrow_right, M1f, M2f, M3f
from pptx.util import Inches

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    target = args.output.resolve()
    if target.suffix.lower() != ".pptx":
        parser.error("--output must end in .pptx")
    if target.exists():
        parser.error("output already exists; choose another path")
    prs, slide = new_slide(width_in=10, height_in=3)
    for i, (title, fill) in enumerate(zip(
        ("Research question", "Proposed mechanism", "Planned validation"),
        (M1f, M2f, M3f),
    )):
        x = Inches(0.4 + i * 3.2)
        round_rect(slide, x, Inches(0.75), Inches(2.8), Inches(1.3), fill)
        label(slide, x + Inches(0.1), Inches(1.18), Inches(2.6), Inches(0.45), title, size=15)
        if i < 2:
            arrow_right(slide, x + Inches(2.85), Inches(1.28), Inches(0.3), Inches(0.2))
    label(slide, Inches(0.4), Inches(2.45), Inches(9.2), Inches(0.3),
          "Fictional layout example - no experimental results", size=10)
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open("xb") as output:
        prs.save(output)
    print(target)

if __name__ == "__main__":
    main()
