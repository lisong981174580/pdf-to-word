"""Run each conversion in a separate process for reliable timeout cleanup."""
import sys
from pdf2docx import Converter

if __name__ == '__main__':
    converter = Converter(sys.argv[1])
    try:
        converter.convert(sys.argv[2], multi_processing=False)
    finally:
        converter.close()
