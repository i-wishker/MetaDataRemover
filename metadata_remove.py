import os
import sys
from docx import Document
from datetime import datetime

def remove_docx_metadata(input_path, output_path):
    """Remove metadata from a Word (.docx) file and save a new copy."""
    try:
        print(f"📂 Opening file: {input_path}")
        doc = Document(input_path)
        props = doc.core_properties

        # Remove metadata fields
        props.author = ""
        props.title = ""
        props.subject = ""
        props.keywords = ""
        props.comments = ""
        props.last_modified_by = ""
        props.created = datetime.min  # Set to earliest possible date
        props.modified = datetime.min  # Set to earliest possible date

        doc.save(output_path)
        print(f"✅ DOCX metadata removed: {output_path}")
    except Exception as e:
        print(f"❌ Error processing DOCX: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("⚠️ Usage: python remove_docx_metadata.py <file_path>")
    else:
        file_path = sys.argv[1]
        print(f"🔍 File received: {file_path}")

        if not os.path.exists(file_path):
            print("❌ Error: File not found! Check the file path.")
        else:
            filename, ext = os.path.splitext(file_path)

            if ext.lower() != ".docx":
                print("❌ Unsupported file type. Please provide a .docx file.")
            else:
                output_path = f"{filename}_clean.docx"
                remove_docx_metadata(file_path, output_path)
