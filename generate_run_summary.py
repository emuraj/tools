from reportlab.pdfgen import canvas
from pathlib import Path
import hashlib, csv

def build_run_summary(pdf_path: Path,
                      manifest_path: Path,
                      metrics: dict):
    c = canvas.Canvas(str(pdf_path))
    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, 800, "Run Summary")
    c.setFont("Helvetica", 9)
    y = 770
    for k, v in metrics.items():
        c.drawString(40, y, f"{k}: {v}")
        y -= 12
    # embed manifest hash
    sha = hashlib.sha256(manifest_path.read_bytes()).hexdigest()
    y -= 24
    c.drawString(40, y, f"SHA‑256 of manifest CSV: {sha}")
    c.save()