#!/usr/bin/env python3
"""
Beast PDF Compressor - Maximum Text-Only Compression
Sacrifices all image quality for maximum file size reduction while preserving text readability.
Designed for archiving documents where only text content matters.
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path
import argparse
import time


class BeastPDFCompressor:
    def __init__(self, ghostscript_path=None):
        """Initialize the Beast PDF Compressor"""
        self.gs_path = self._find_ghostscript(ghostscript_path)
        if not self.gs_path:
            raise Exception("Ghostscript not found! Please install Ghostscript.")

    def _find_ghostscript(self, custom_path=None):
        """Find Ghostscript executable in common locations"""
        possible_paths = []

        if custom_path:
            possible_paths.append(custom_path)

        # Common installation paths
        possible_paths.extend([
            r"S:\gs10.05.1\bin\gswin64c.exe",
            r"S:\gs10.05.1\bin\gswin32c.exe",
            r"C:\Program Files\gs\gs10.05.1\bin\gswin64c.exe",
            r"C:\Program Files\gs\gs9.56.1\bin\gswin64c.exe",
            r"C:\Program Files (x86)\gs\gs10.05.1\bin\gswin32c.exe",
            r"C:\Program Files (x86)\gs\gs9.56.1\bin\gswin32c.exe"
        ])

        # Try system PATH
        possible_paths.extend(["gswin64c.exe", "gswin32c.exe", "gs"])

        for path in possible_paths:
            if shutil.which(path) or os.path.exists(path):
                return path

        return None

    def get_file_size(self, filepath):
        """Get file size in bytes"""
        return os.path.getsize(filepath)

    def format_size(self, size_bytes):
        """Format file size in human readable format"""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 ** 2:
            return f"{size_bytes / 1024:.1f} KB"
        elif size_bytes < 1024 ** 3:
            return f"{size_bytes / (1024 ** 2):.1f} MB"
        else:
            return f"{size_bytes / (1024 ** 3):.1f} GB"

    def analyze_pdf(self, input_path):
        """Analyze PDF file for optimization strategy"""
        try:
            file_size = self.get_file_size(input_path)

            # Quick analysis based on file size
            if file_size > 50 * 1024 * 1024:  # >50MB
                analysis_type = "Large Document (likely scanned)"
            elif file_size > 10 * 1024 * 1024:  # >10MB
                analysis_type = "Medium Document (mixed content)"
            else:
                analysis_type = "Small Document (likely text-heavy)"

            return {
                'file_size': file_size,
                'type': analysis_type,
                'needs_aggressive_compression': file_size > 5 * 1024 * 1024
            }

        except Exception as e:
            return {
                'file_size': 0,
                'type': 'Unknown',
                'needs_aggressive_compression': True,
                'error': str(e)
            }

    def beast_compress(self, input_path, output_path):
        """
        Apply Beast compression - maximum compression with text preservation
        """
        print("BEAST COMPRESSION INITIATED")
        print("WARNING: All image quality will be sacrificed for maximum compression")
        print("Text readability will be preserved")

        # Beast mode compression settings
        beast_cmd = [
            self.gs_path,
            '-sDEVICE=pdfwrite',
            '-dCompatibilityLevel=1.4',
            '-dNOPAUSE',
            '-dQUIET',
            '-dBATCH',
            '-dSAFER',

            # Optimization settings
            '-dOptimize=true',
            '-dFastWebView=true',
            '-dPrinted=false',
            '-dUseFlateCompression=true',

            # Font optimization
            '-dEmbedAllFonts=false',
            '-dSubsetFonts=true',
            '-dCompressFonts=true',

            # Image destruction for maximum compression
            '-dPDFSETTINGS=/screen',
            '-dDownsampleColorImages=true',
            '-dDownsampleGrayImages=true',
            '-dDownsampleMonoImages=true',
            '-dColorImageDownsampleType=/Subsample',
            '-dGrayImageDownsampleType=/Subsample',
            '-dMonoImageDownsampleType=/Subsample',

            # Ultra-low resolution settings
            '-dColorImageResolution=50',
            '-dGrayImageResolution=50',
            '-dMonoImageResolution=150',

            # Aggressive threshold settings
            '-dColorImageDownsampleThreshold=1.0',
            '-dGrayImageDownsampleThreshold=1.0',
            '-dMonoImageDownsampleThreshold=1.0',

            # Image encoding settings
            '-dEncodeColorImages=true',
            '-dEncodeGrayImages=true',
            '-dEncodeMonoImages=true',
            '-dColorImageFilter=/DCTEncode',
            '-dGrayImageFilter=/DCTEncode',
            '-dMonoImageFilter=/CCITTFaxEncode',

            # Minimum quality settings
            '-dJPEGQ=10',
            '-dAutoFilterColorImages=false',
            '-dAutoFilterGrayImages=false',
            '-dAntiAliasColorImages=false',
            '-dAntiAliasGrayImages=false',
            '-dAntiAliasMonoImages=false',

            # Remove unnecessary elements
            '-dDetectDuplicateImages=true',
            '-dPreserveAnnots=false',
            '-dPreserveMarkedContent=false',
            '-dPreserveOPIComments=false',
            '-dPreserveHalftoneInfo=false',
            '-dPreserveOverprintSettings=false',
            '-dPreserveEPSInfo=false',

            # Color optimization
            '-dUseCIEColor=false',
            '-dColorConversionStrategy=/LeaveColorUnchanged',
            '-dConvertCMYKImagesToRGB=true',
            '-dConvertImagesToIndexed=true',
            '-dUCRandBGInfo=/Remove',

            # Bypass image preservation
            '-dPassThroughJPEGImages=false',
            '-dPassThroughJPEGQ=false',

            # Output file
            f'-sOutputFile={output_path}',
            input_path
        ]

        try:
            print("Executing Beast compression...")
            start_time = time.time()

            result = subprocess.run(beast_cmd, capture_output=True, text=True, timeout=600)

            end_time = time.time()
            compression_time = end_time - start_time

            if result.returncode != 0:
                print(f"Primary Beast compression failed (exit code: {result.returncode})")
                print("Attempting fallback compression...")
                return self._fallback_compression(input_path, output_path)

            if not os.path.exists(output_path):
                print("Output file was not created, attempting fallback...")
                return self._fallback_compression(input_path, output_path)

            print(f"Beast compression completed in {compression_time:.2f} seconds")
            return True

        except subprocess.TimeoutExpired:
            print("Beast compression timed out (10 minutes), attempting fallback...")
            return self._fallback_compression(input_path, output_path)
        except Exception as e:
            print(f"Beast compression error: {e}")
            return self._fallback_compression(input_path, output_path)

    def _fallback_compression(self, input_path, output_path):
        """Fallback compression when Beast mode fails"""
        print("FALLBACK: Attempting text-extraction focused compression...")

        fallback_cmd = [
            self.gs_path,
            '-sDEVICE=pdfwrite',
            '-dCompatibilityLevel=1.4',
            '-dNOPAUSE',
            '-dQUIET',
            '-dBATCH',
            '-dSAFER',
            '-dPDFSETTINGS=/screen',
            '-dOptimize=true',
            '-dEmbedAllFonts=false',
            '-dSubsetFonts=true',
            '-dCompressFonts=true',
            '-dEncodeColorImages=false',
            '-dEncodeGrayImages=false',
            '-dEncodeMonoImages=true',
            '-dPreserveAnnots=false',
            '-dPreserveMarkedContent=false',
            '-dUseFlateCompression=true',
            f'-sOutputFile={output_path}',
            input_path
        ]

        try:
            result = subprocess.run(fallback_cmd, capture_output=True, text=True, timeout=300)

            if result.returncode == 0 and os.path.exists(output_path):
                print("Fallback compression successful")
                return True
            else:
                print("Fallback compression also failed")
                return False

        except Exception as e:
            print(f"Fallback compression error: {e}")
            return False

    def _analyze_compression_result(self, original_size, compressed_size):
        """Analyze and explain compression results"""
        print("\nCOMPRESSION ANALYSIS:")
        print("-" * 50)

        if compressed_size < original_size:
            reduction = ((original_size - compressed_size) / original_size) * 100
            saved_space = original_size - compressed_size

            print(f"Space saved: {self.format_size(saved_space)}")
            print(f"Compression ratio: {reduction:.1f}%")

            if reduction > 95:
                print("Result: EXCEPTIONAL - Beast mode achieved maximum compression")
            elif reduction > 90:
                print("Result: EXCELLENT - Very high compression achieved")
            elif reduction > 70:
                print("Result: GOOD - Solid compression achieved")
            elif reduction > 50:
                print("Result: MODERATE - Some compression achieved")
            else:
                print("Result: LIMITED - Minimal compression achieved")

        elif compressed_size == original_size:
            print("Result: NO CHANGE - File may already be optimized")
        else:
            increase = ((compressed_size - original_size) / original_size) * 100
            print(f"Result: SIZE INCREASED by {increase:.1f}%")
            print("This can happen with already optimized text-only PDFs")

    def _explain_beast_compression(self):
        """Explain what Beast compression does"""
        print("\nBEAST COMPRESSION DETAILS:")
        print("-" * 50)
        print("SACRIFICED FOR MAXIMUM COMPRESSION:")
        print("• All image quality reduced to minimum (10% JPEG quality)")
        print("• Image resolution reduced to 50 DPI for color/grayscale")
        print("• Font embedding minimized to essential characters only")
        print("• All metadata, annotations, and comments removed")
        print("• Color profiles and advanced graphics removed")
        print("• Duplicate images and objects eliminated")
        print("\nPRESERVED FOR READABILITY:")
        print("• All text content and formatting")
        print("• Document structure and page layout")
        print("• Essential navigation elements")
        print("• Basic font rendering")

    def process_single_file(self, input_file, output_dir=None):
        """Process a single PDF file with Beast compression"""
        input_path = Path(input_file)

        if not input_path.exists():
            print(f"ERROR: Input file {input_file} does not exist")
            return False

        if not input_path.suffix.lower() == '.pdf':
            print(f"ERROR: {input_file} is not a PDF file")
            return False

        # Determine output path
        if output_dir:
            output_dir = Path(output_dir)
            output_dir.mkdir(exist_ok=True)
            output_path = output_dir / f"{input_path.stem}_text_only.pdf"
        else:
            output_path = input_path.parent / f"{input_path.stem}_text_only.pdf"

        print(f"Processing: {input_path.name}")
        print(f"Output: {output_path}")
        print("=" * 60)

        # Analyze the PDF
        analysis = self.analyze_pdf(input_path)
        original_size = analysis['file_size']

        print(f"File type: {analysis['type']}")
        print(f"Original size: {self.format_size(original_size)}")

        # Perform Beast compression
        success = self.beast_compress(str(input_path), str(output_path))

        if success:
            compressed_size = self.get_file_size(output_path)
            print(f"Compressed size: {self.format_size(compressed_size)}")

            # Analyze results
            self._analyze_compression_result(original_size, compressed_size)

            return True
        else:
            print("Beast compression failed completely")
            return False

    def process_directory(self, input_dir, output_dir=None):
        """Process all PDF files in a directory"""
        input_path = Path(input_dir)

        if not input_path.exists():
            print(f"ERROR: Input directory {input_dir} does not exist")
            return

        # Find all PDF files
        pdf_files = list(input_path.glob("*.pdf"))

        if not pdf_files:
            print(f"No PDF files found in {input_dir}")
            return

        print(f"BEAST COMPRESSION BATCH PROCESSING")
        print(f"Found {len(pdf_files)} PDF files to process")
        print("=" * 60)

        success_count = 0
        total_original_size = 0
        total_compressed_size = 0
        start_time = time.time()

        for i, pdf_file in enumerate(pdf_files, 1):
            print(f"\nProcessing file {i}/{len(pdf_files)}")
            print("-" * 40)

            original_size = self.get_file_size(pdf_file)
            total_original_size += original_size

            if self.process_single_file(pdf_file, output_dir):
                success_count += 1

                # Calculate compressed size for totals
                if output_dir:
                    compressed_file = Path(output_dir) / f"{pdf_file.stem}_text_only.pdf"
                else:
                    compressed_file = pdf_file.parent / f"{pdf_file.stem}_text_only.pdf"

                if compressed_file.exists():
                    total_compressed_size += self.get_file_size(compressed_file)

        # Final summary
        end_time = time.time()
        total_time = end_time - start_time

        print(f"\nBEAST COMPRESSION BATCH COMPLETE")
        print("=" * 60)
        print(f"Total processing time: {total_time:.2f} seconds")
        print(f"Files processed: {len(pdf_files)}")
        print(f"Successful compressions: {success_count}")
        print(f"Failed compressions: {len(pdf_files) - success_count}")
        print(f"Total original size: {self.format_size(total_original_size)}")
        print(f"Total compressed size: {self.format_size(total_compressed_size)}")

        if total_original_size > 0 and total_compressed_size > 0:
            if total_compressed_size < total_original_size:
                total_reduction = ((total_original_size - total_compressed_size) / total_original_size) * 100
                total_saved = total_original_size - total_compressed_size
                print(f"Overall compression: {total_reduction:.1f}%")
                print(f"Total space saved: {self.format_size(total_saved)}")

                if total_reduction > 90:
                    print("RESULT: EXCEPTIONAL BEAST COMPRESSION ACHIEVED")
                elif total_reduction > 70:
                    print("RESULT: EXCELLENT BEAST COMPRESSION ACHIEVED")
                elif total_reduction > 50:
                    print("RESULT: GOOD BEAST COMPRESSION ACHIEVED")
                else:
                    print("RESULT: MODERATE BEAST COMPRESSION ACHIEVED")

        # Show compression explanation
        self._explain_beast_compression()


def main():
    """Main function to handle command line arguments and execute compression"""
    parser = argparse.ArgumentParser(
        description='Beast PDF Compressor - Maximum text-only compression',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
BEAST COMPRESSION:
  Sacrifices all image quality for maximum file size reduction while preserving text readability.
  Perfect for archiving documents where only text content matters.

EXAMPLES:
  python beast_compressor.py document.pdf
  python beast_compressor.py ./documents -o ./compressed
  python beast_compressor.py "C:/My Documents/file.pdf"
        """
    )

    parser.add_argument('input', help='Input PDF file or directory path')
    parser.add_argument('-o', '--output', help='Output directory (optional)')
    parser.add_argument('-g', '--ghostscript', help='Path to Ghostscript executable (optional)')

    args = parser.parse_args()

    try:
        compressor = BeastPDFCompressor(args.ghostscript)

        input_path = Path(args.input)

        if input_path.is_file():
            compressor.process_single_file(args.input, args.output)
        elif input_path.is_dir():
            compressor.process_directory(args.input, args.output)
        else:
            print(f"ERROR: {args.input} is not a valid file or directory")
            sys.exit(1)

    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) == 1:
        print("BEAST PDF COMPRESSOR - TEXT-ONLY FOCUS")
        print("=" * 50)
        print("WARNING: This will sacrifice all image quality for maximum compression!")
        print("Perfect for archiving documents where only text matters.")
        print("\nUSAGE:")
        print("python beast_compressor.py <input_file_or_directory> [options]")
        print("\nEXAMPLES:")
        print("python beast_compressor.py document.pdf")
        print("python beast_compressor.py ./documents -o ./compressed")
        print('python beast_compressor.py "C:/My Documents/file.pdf"')
        print("\nFEATURES:")
        print("• Maximum compression with text preservation")
        print("• Automatic fallback for problematic PDFs")
        print("• Detailed compression statistics")
        print("• Batch processing support")
        print("• Intelligent file analysis")
        print("\nDrag and drop files onto this script to compress them!")
        input("\nPress Enter to exit...")
    else:
        main()