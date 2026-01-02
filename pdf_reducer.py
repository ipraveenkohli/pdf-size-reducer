import os
import sys
from tkinter import Tk, filedialog

try:
    import fitz  # PyMuPDF
except ImportError:
    print("ERROR: PyMuPDF (fitz) is not installed.")
    print("Install it with: pip install pymupdf")
    input("Press Enter to exit...")
    sys.exit(1)

try:
    from PIL import Image
except ImportError:
    print("ERROR: Pillow (PIL) is not installed.")
    print("Install it with: pip install pillow")
    input("Press Enter to exit...")
    sys.exit(1)

from io import BytesIO


def pixmap_to_jpeg_bytes(pix, quality=75):
    """
    Convert a PyMuPDF Pixmap to JPEG bytes using Pillow, with given quality.
    """
    if pix.n >= 4:  # has alpha
        mode = "RGBA"
    else:
        mode = "RGB"

    img = Image.frombytes(mode, (pix.width, pix.height), pix.samples)

    if mode == "RGBA":
        img = img.convert("RGB")

    buf = BytesIO()
    img.save(buf, format="JPEG", quality=quality, optimize=True)
    return buf.getvalue()


def build_flattened_pdf(doc, quality, dpi):
    """
    Render each page to a JPEG image (at given DPI & quality)
    and rebuild a new image-only PDF in memory.
    Returns: bytes of the new PDF.
    """
    out = fitz.open()  # new PDF

    for page in doc:
        pix = page.get_pixmap(dpi=dpi)
        img_bytes = pixmap_to_jpeg_bytes(pix, quality=quality)

        width, height = pix.width, pix.height
        rect = fitz.Rect(0, 0, width, height)
        pdf_page = out.new_page(width=width, height=height)
        pdf_page.insert_image(rect, stream=img_bytes)

    pdf_bytes = out.tobytes()
    out.close()
    return pdf_bytes


def compress_pdf_to_size(input_pdf, target_kb, dpi=96, max_iters=8, output_dir=None):
    """
    Compress a PDF to be as close as possible to target_kb using
    page rasterization + JPEG quality binary search.

    - target_kb: desired size in KB
    - dpi: rendering DPI for pages
    - output_dir: if not None, save compressed file there with same name.
                  if None, overwrite the original file.
    """
    target_bytes = int(target_kb * 1024)
    original_size = os.path.getsize(input_pdf)
    original_kb = original_size / 1024

    print(f"\nProcessing: {os.path.basename(input_pdf)}")
    print(f"  Original size: {original_kb:.1f} KB")

    doc = fitz.open(input_pdf)

    low_q, high_q = 10, 95
    iterations = 0

    best_under = None
    best_under_size = 0

    best_over = None
    best_over_size = None

    while low_q <= high_q and iterations < max_iters:
        q = (low_q + high_q) // 2
        iterations += 1

        print(f"  ▶ Trying quality {q} ... ", end="", flush=True)
        pdf_bytes = build_flattened_pdf(doc, quality=q, dpi=dpi)
        size = len(pdf_bytes)
        size_kb = size / 1024
        print(f"done ({size_kb:.1f} KB)")

        if size <= target_bytes:
            if size > best_under_size:
                best_under = pdf_bytes
                best_under_size = size
            low_q = q + 1
        else:
            if best_over is None or size < best_over_size:
                best_over = pdf_bytes
                best_over_size = size
            high_q = q - 1

    doc.close()

    chosen_bytes = None
    chosen_size = None

    if best_under is not None:
        chosen_bytes = best_under
        chosen_size = best_under_size
    elif best_over is not None:
        chosen_bytes = best_over
        chosen_size = best_over_size

    # If no candidate or candidate is larger than original, keep original
    if chosen_bytes is None or chosen_size >= original_size:
        print("  ⚠ Could not make a file smaller than the original with this method.")
        print("  ➜ Keeping original file as the smallest version.")
        return

    # Decide output path
    filename = os.path.basename(input_pdf)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, filename)
    else:
        # overwrite original
        output_path = input_pdf

    out_kb = chosen_size / 1024

    with open(output_path, "wb") as f:
        f.write(chosen_bytes)

    status = "✅ within target" if chosen_size <= target_bytes else "⚠ above target (best possible)"
    print(f"  ➜ Saved: {output_path}")
    print(f"     Size: {out_kb:.1f} KB (target {target_kb:.1f} KB) {status}")


def main():
    print("=== PDF Size Reducer (Pi7-style, image-based) ===")

    while True:
        try:
            target_kb = float(input("Enter desired maximum PDF size (in KB): ").strip())
            if target_kb <= 0:
                print("Please enter a positive number.")
                continue
            break
        except ValueError:
            print("Please enter a valid number (e.g., 500).")

    dpi_input = input("Enter DPI for pages (default 96, lower = smaller file): ").strip()
    if dpi_input:
        try:
            dpi = int(dpi_input)
        except ValueError:
            print("Invalid DPI, using default 96.")
            dpi = 96
    else:
        dpi = 96

    # One Tk root for both file and folder dialogs
    root = Tk()
    root.withdraw()
    root.update()

    file_paths = filedialog.askopenfilenames(
        title="Select PDF files to compress",
        filetypes=[("PDF files", "*.pdf")]
    )

    if not file_paths:
        root.destroy()
        print("No files selected. Exiting.")
        input("Press Enter to close...")
        return

    print("\nSelect a folder to save compressed PDFs.")
    print("If you click 'Cancel', originals will be OVERWRITTEN with compressed versions.\n")

    output_dir = filedialog.askdirectory(
        title="Select output folder (Cancel = overwrite originals)"
    )

    root.destroy()

    if output_dir:
        print(f"Compressed PDFs will be saved in: {output_dir}")
    else:
        print("Compressed PDFs will overwrite the original files.")

    print(f"\nSelected {len(file_paths)} file(s). Starting compression...")

    for path in file_paths:
        compress_pdf_to_size(path, target_kb, dpi=dpi, output_dir=output_dir)

    print("\nAll done!")
    input("Press Enter to close...")


if __name__ == "__main__":
    main()
