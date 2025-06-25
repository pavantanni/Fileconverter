import os
import sys

from file_converter import to_pdf, to_docx, to_pptx, to_jpg, compress_file

def normalize_command(cmd):
    return cmd.strip().lower().replace(" ", "")

if len(sys.argv) != 3:
    print("Usage: python launcher.py <filename> <command>")
    input("Press Enter to exit...")
    sys.exit()

file = sys.argv[1]
command = normalize_command(sys.argv[2])

if not os.path.exists(file):
    print(f"❌ File not found: {file}")
    input("Press Enter to exit...")
    sys.exit()

if command == "topdf":
    to_pdf(file)
elif command == "todocx":
    to_docx(file)
elif command == "topptx":
    to_pptx(file)
elif command == "tojpg":
    to_jpg(file)
elif command == "compress":
    compress_file(file)
else:
    print(f"❌ Unsupported command: {command}")

input("Press Enter to exit...")
