import pikepdf
import os

def compress_pdf_pikepdf(input_path, output_path):
    try:
        # Open the PDF
        with pikepdf.open(input_path) as pdf:
            # Save the PDF with optimization enabled
            # linearize=True makes it "Fast Web View" enabled
            pdf.save(output_path, compress_streams=True)
            
        original_size = os.path.getsize(input_path)
        new_size = os.path.getsize(output_path)
        print(f"Compressed from {original_size / 1024:.2f} KB to {new_size / 1024:.2f} KB")
        
    except Exception as e:
        print(f"Error: {e}")

# Usage
compress_pdf_pikepdf('input.pdf', 'compressed_pike.pdf')
