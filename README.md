PDF File Size Reducer (Python)

A Python-based utility to reduce the size of multiple PDF files in bulk by adjusting DPI and target file size.
Designed for DevOps workflows, batch processing, and server-side automation.

⸻

🚀 Features
	•	📂 Process multiple PDF files at once
	•	🎯 Reduce PDFs based on target file size
	•	🖨 Control output quality using DPI settings
	•	🔁 Batch-friendly & automation-ready
	•	⚙️ Suitable for CI/CD pipelines, cron jobs, and containerized environments
	•	🐍 Pure Python interface (easy to extend)

⸻

📦 Use Cases
	•	Compress PDFs before uploading to cloud storage
	•	Optimize artifacts in CI/CD pipelines
	•	Reduce document sizes for email or archival
	•	Pre-process PDFs in DevOps automation workflows

⸻

🛠 Requirements
	•	Python 3.8+
	•	System dependencies (example):
	•	ghostscript
	•	poppler-utils (if image-based processing is used)

Python dependencies (example):

pip install -r requirements.txt


⸻

▶️ Usage

Basic Example

python reducer.py \
  --input ./input_pdfs \
  --output ./output_pdfs \
  --dpi 150 \
  --target-size 500

Arguments

Argument	Description
--input	Directory containing source PDF files
--output	Directory for reduced PDFs
--dpi	DPI value for PDF rendering (lower = smaller file)
--target-size	Target file size in KB
--recursive	(Optional) Process PDFs recursively


⸻



Output files retain the same structure inside the output directory.



⸻

📈 Performance Notes
	•	Lower DPI = smaller file size, lower visual quality
	•	Image-heavy PDFs benefit most
	•	Text-based PDFs may see minimal reduction


# second version w'll updated soon with ghostreducer

verson 1
