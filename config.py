from pathlib import Path
import os
import tempfile
import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# These are placeholders; app.py will manage the actual temporary directory
TARGET_EXCEL_FILENAME = Path("preprocessed_data.xlsx")
OUTPUT_DIRECTORY = Path("output")

# Path to your PowerPoint template
# os.path.dirname(__file__) gets the directory of the current config.py file
# Assumes FINAL_PowerPoint_Template.pptx is in the same directory as config.py
PPT_TEMPLATE_FILENAME = os.path.join(os.path.dirname(__file__), "FINAL_PowerPoint_Template.pptx")
logging.info(f"Configured PPT_TEMPLATE_FILENAME: {PPT_TEMPLATE_FILENAME}")

# All Azure Blob Storage specific configurations are removed as they are no longer needed.