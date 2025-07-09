#!/usr/bin/env python3
"""
Interactive PDF Compressor using Ghostscript
Provides four compression modes with smart analysis and user-friendly interface
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path
import argparse
import time


class InteractivePDFCompressor:
    def __init__(self, ghostscript_path=None):
        self.gs_path = self._find_ghostscript(ghostscript_path)
        if not self.gs_path:
            raise Exception("Ghostscript not found! Please install Ghostscript or provide correct path.")

        self.compression_modes = {
            '1': 'conservative',
            '2': 'balanced',
            '3': 'aggressive',
            '4': 'nuclear'
        }

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

    def analyze_pdf_content(self, input_path):
        """Analyze PDF to understand its content type and recommend compression"""
        try:
            # Basic analysis using file size and Ghostscript info
            file_size = self.get_file_size(input_path)

            # Try to get basic PDF info
            cmd = [
                self.gs_path,
                '-sDEVICE=bbox',
                '-dNOPAUSE',
                '-dBATCH',
                '-dQUIET',
                input_path
            ]

            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)

            analysis = {
                'file_size': file_size,
                'size_mb': file_size / (1024 * 1024),
                'likely_scanned': file_size > 10 * 1024 * 1024,
                'likely_text_heavy': file_size < 2 * 1024 * 1024,
                'has_errors': result.returncode != 0,
                'recommended_mode': self._recommend_compression_mode(file_size)
            }

            return analysis

        except Exception as e:
            return {
                'file_size': self.get_file_size(input_path),
                'size_mb': self.get_file_size(input_path) / (1024 * 1024),
                'error': str(e),
                'recommended_mode': 'balanced'
            }

    def _recommend_compression_mode(self, file_size):
        """Recommend compression mode based on file size"""
        size_mb = file_size / (1024 * 1024)

        if size_mb > 50:
            return 'nuclear'
        elif size_mb > 10:
            return 'aggressive'
        elif size_mb > 2:
            return 'balanced'
        else:
            return 'conservative'

    def display_compression_menu(self):
        """Display compression mode selection menu"""
        print("\n" + "=" * 60)
        print("PDF COMPRESSION MODES")
        print("=" * 60)
        print("1. Conservative - High quality, minimal compression")
        print("   Best for: Documents with important images, presentations")
        print("   Compression: 10-30% reduction")
        print()
        print("2. Balanced - Medium quality, moderate compression")
        print("   Best for: General documents, mixed content")
        print("   Compression: 30-50% reduction")
        print()
        print("3. Aggressive - Lower quality, high compression")
        print("   Best for: Archival documents, size-critical files")
        print("   Compression: 50-70% reduction")
        print()
        print("4. Nuclear - Text-focused, maximum compression")
        print("   Best for: Text-heavy documents, extreme size reduction")
        print("   Compression: 70-90% reduction")
        print("=" * 60)

    def get_user_compression_choice(self, recommended_mode=None):
        """Get compression mode choice from user"""
        if recommended_mode:
            mode_names = {
                'conservative': '1 (Conservative)',
                'balanced': '2 (Balanced)',
                'aggressive': '3 (Aggressive)',
                'nuclear': '4 (Nuclear)'
            }
            print(f"\nRecommended mode based on file analysis: {mode_names[recommended_mode]}")

        while True:
            choice = input("\nSelect compression mode (1-4): ").strip()

            if choice in self.compression_modes:
                selected_mode = self.compression_modes[choice]
                print(f"\nSelected: {selected_mode.upper()} compression")
                return selected_mode
            else:
                print("Invalid choice. Please select 1, 2, 3, or 4.")

    def compress_pdf(self, input_path, output_path, compression_level):
        """Compress PDF with specified compression level"""

        # Base Ghostscript command
        base_cmd = [
            self.gs_path,
            '-sDEVICE=pdfwrite',
            '-dCompatibilityLevel=1.4',
            '-dNOPAUSE',
            '-dQUIET',
            '-dBATCH',
            '-dSAFER',
            '-dDetectDuplicateImages=true',
            '-dCompressFonts=true',
            '-dSubsetFonts=true',
            '-dOptimize=true',
            '-dUseFlateCompression=true',
            '-dFastWebView=true'
        ]

        # Mode-specific compression settings
        if compression_level == "conservative":
            compression_args = [
                '-dPDFSETTINGS=/printer',
                '-dDownsampleColorImages=true',
                '-dDownsampleGrayImages=true',
                '-dColorImageResolution=300',
                '-dGrayImageResolution=300',
                '-dMonoImageResolution=1200',
                '-dColorImageDownsampleType=/Bicubic',
                '-dGrayImageDownsampleType=/Bicubic',
                '-dJPEGQ=90',
                '-dEmbedAllFonts=true',
                '-dPreserveAnnots=true'
            ]

        elif compression_level == "balanced":
            compression_args = [
                '-dPDFSETTINGS=/ebook',
                '-dDownsampleColorImages=true',
                '-dDownsampleGrayImages=true',
                '-dColorImageResolution=200',
                '-dGrayImageResolution=200',
                '-dMonoImageResolution=600',
                '-dColorImageDownsampleType=/Bicubic',
                '-dGrayImageDownsampleType=/Bicubic',
                '-dJPEGQ=85',
                '-dEmbedAllFonts=true',
                '-dPreserveAnnots=false'
            ]

        elif compression_level == "aggressive":
            compression_args = [
                '-dPDFSETTINGS=/screen',
                '-dDownsampleColorImages=true',
                '-dDownsampleGrayImages=true',
                '-dDownsampleMonoImages=true',
                '-dColorImageDownsampleType=/Bicubic',
                '-dGrayImageDownsampleType=/Bicubic',
                '-dMonoImageDownsampleType=/Bicubic',
                '-dColorImageResolution=150',
                '-dGrayImageResolution=150',
                '-dMonoImageResolution=300',
                '-dColorImageDownsampleThreshold=1.5',
                '-dGrayImageDownsampleThreshold=1.5',
                '-dMonoImageDownsampleThreshold=1.5',
                '-dEncodeColorImages=true',
                '-dEncodeGrayImages=true',
                '-dEncodeMonoImages=true',
                '-dColorImageFilter=/DCTEncode',
                '-dGrayImageFilter=/DCTEncode',
                '-dMonoImageFilter=/CCITTFaxEncode',
                '-dJPEGQ=75',
                '-dEmbedAllFonts=false',
                '-dPreserveAnnots=false',
                '-dPreserveMarkedContent=false'
            ]

        else:  # nuclear
            compression_args = [
                '-dPDFSETTINGS=/screen',
                '-dDownsampleColorImages=true',
                '-dDownsampleGrayImages=true',
                '-dDownsampleMonoImages=true',
                '-dColorImageDownsampleType=/Subsample',
                '-dGrayImageDownsampleType=/Subsample',
                '-dMonoImageDownsampleType=/Subsample',
                '-dColorImageResolution=72',
                '-dGrayImageResolution=72',
                '-dMonoImageResolution=200',
                '-dColorImageDownsampleThreshold=1.0',
                '-dGrayImageDownsampleThreshold=1.0',
                '-dMonoImageDownsampleThreshold=1.0',
                '-dEncodeColorImages=true',
                '-dEncodeGrayImages=true',
                '-dEncodeMonoImages=true',
                '-dColorImageFilter=/DCTEncode',
                '-dGrayImageFilter=/DCTEncode',
                '-dMonoImageFilter=/CCITTFaxEncode',
                '-dJPEGQ=40',
                '-dEmbedAllFonts=false',
                '-dPreserveAnnots=false',
                '-dPreserveMarkedContent=false',
                '-dPassThroughJPEGImages=false',
                '-dConvertCMYKImagesToRGB=true',
                '-dConvertImagesToIndexed=true'
            ]

        # Complete command
        cmd = base_cmd + compression_args + [f'-sOutputFile={output_path}', input_path]

        try:
            # Run compression
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)

            if result.returncode != 0:
                print(f"Primary compression failed, trying fallback approach...")
                return self._try_fallback_compression(input_path, output_path, compression_level)

            # Check if output file was created
            if not os.path.exists(output_path):
                print(f"Output file was not created, trying fallback...")
                return self._try_fallback_compression(input_path, output_path, compression_level)

            return True

        except subprocess.TimeoutExpired:
            print("Compression timed out, trying fallback...")
            return self._try_fallback_compression(input_path, output_path, compression_level)
        except Exception as e:
            print(f"Error during compression: {e}")
            return self._try_fallback_compression(input_path, output_path, compression_level)

    def _try_fallback_compression(self, input_path, output_path, compression_level):
        """Try fallback compression with minimal settings"""
        print("Attempting fallback compression with minimal settings...")

        # Minimal compression command for problematic PDFs
        cmd = [
            self.gs_path,
            '-sDEVICE=pdfwrite',
            '-dCompatibilityLevel=1.4',
            '-dNOPAUSE',
            '-dQUIET',
            '-dBATCH',
            '-dSAFER',
            '-dOptimize=true',
            '-dCompressFonts=true',
            '-dSubsetFonts=true'
        ]

        # Add basic compression based on level
        if compression_level in ['aggressive', 'nuclear']:
            cmd.extend(['-dPDFSETTINGS=/screen'])
        else:
            cmd.extend(['-dPDFSETTINGS=/default'])

        cmd.extend([f'-sOutputFile={output_path}', input_path])

        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)

            if result.returncode == 0 and os.path.exists(output_path):
                print("Fallback compression successful")
                return True
            else:
                print("Fallback compression also failed")
                return False

        except Exception as e:
            print(f"Fallback compression failed: {e}")
            return False

    def process_single_file(self, input_file, output_dir=None, compression_level=None, interactive=True):
        """Process a single PDF file"""
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
            output_path = output_dir / f"{input_path.stem}_compressed.pdf"
        else:
            output_path = input_path.parent / f"{input_path.stem}_compressed.pdf"

        print(f"\nProcessing: {input_path}")
        print(f"Output: {output_path}")

        # Analyze PDF content
        print("\nAnalyzing PDF content...")
        analysis = self.analyze_pdf_content(input_path)

        print(f"File size: {self.format_size(analysis['file_size'])}")

        if analysis['size_mb'] > 50:
            print("Analysis: Large file detected - likely scanned document")
        elif analysis['size_mb'] > 10:
            print("Analysis: Medium file - mixed content likely")
        elif analysis['size_mb'] < 2:
            print("Analysis: Small file - likely text/vector content")
        else:
            print("Analysis: Standard document size")

        # Get compression mode
        if interactive and not compression_level:
            self.display_compression_menu()
            compression_level = self.get_user_compression_choice(analysis['recommended_mode'])
        elif not compression_level:
            compression_level = analysis['recommended_mode']
            print(f"Using recommended compression mode: {compression_level}")

        print(f"\nApplying {compression_level.upper()} compression...")
        start_time = time.time()

        # Compress the PDF
        success = self.compress_pdf(str(input_path), str(output_path), compression_level)

        if success:
            # Get compressed file size
            compressed_size = self.get_file_size(output_path)
            compression_time = time.time() - start_time

            print(f"\nCompression Results:")
            print(f"Original size: {self.format_size(analysis['file_size'])}")
            print(f"Compressed size: {self.format_size(compressed_size)}")
            print(f"Processing time: {compression_time:.1f} seconds")

            # Calculate compression ratio
            if analysis['file_size'] > 0:
                if compressed_size < analysis['file_size']:
                    ratio = ((analysis['file_size'] - compressed_size) / analysis['file_size']) * 100
                    space_saved = analysis['file_size'] - compressed_size
                    print(f"Compression achieved: {ratio:.1f}% reduction")
                    print(f"Space saved: {self.format_size(space_saved)}")
                    print("Status: SUCCESS")
                elif compressed_size == analysis['file_size']:
                    print("File size unchanged - PDF may already be optimized")
                    print("Status: NO CHANGE")
                else:
                    ratio = ((compressed_size - analysis['file_size']) / analysis['file_size']) * 100
                    print(f"File size increased by {ratio:.1f}%")
                    print("Status: SIZE INCREASED")
                    self._explain_compression_result(analysis['file_size'], compressed_size, input_path)

            return True
        else:
            print("Status: FAILED")
            return False

    def _explain_compression_result(self, original_size, compressed_size, input_path):
        """Explain why compression didn't work well"""
        print("\nCompression Analysis:")

        size_mb = original_size / (1024 * 1024)

        if size_mb < 0.5:
            print("Small file - likely contains efficient text/vector content")
        elif size_mb < 2:
            print("Medium file - may contain optimized images or vector graphics")
        else:
            print("Large file - may contain high-quality images or complex content")

        print("\nPossible reasons for poor compression:")
        print("- PDF is already optimized")
        print("- Contains mostly vector graphics or text")
        print("- Images are already highly compressed")
        print("- PDF has password protection or encryption")
        print("- Contains forms, annotations, or complex structures")

        print("\nSuggestions:")
        print("- Try a different compression mode")
        print("- Check if PDF has password protection")
        print("- Consider if current file size is acceptable")

    def process_directory(self, input_dir, output_dir=None, compression_level=None, interactive=True):
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

        print(f"\nFound {len(pdf_files)} PDF files to process")

        # Get compression mode for batch processing
        if interactive and not compression_level:
            self.display_compression_menu()
            compression_level = self.get_user_compression_choice()
            print(f"\nUsing {compression_level.upper()} compression for all files")
        elif not compression_level:
            compression_level = 'balanced'
            print(f"Using default compression mode: {compression_level}")

        success_count = 0
        total_original_size = 0
        total_compressed_size = 0

        print(f"\nProcessing {len(pdf_files)} files with {compression_level.upper()} compression...")
        print("=" * 60)

        for i, pdf_file in enumerate(pdf_files, 1):
            print(f"\nFile {i}/{len(pdf_files)}: {pdf_file.name}")
            print("-" * 40)

            original_size = self.get_file_size(pdf_file)
            total_original_size += original_size

            if self.process_single_file(pdf_file, output_dir, compression_level, interactive=False):
                success_count += 1

                # Calculate compressed size for totals
                if output_dir:
                    compressed_file = Path(output_dir) / f"{pdf_file.stem}_compressed.pdf"
                else:
                    compressed_file = pdf_file.parent / f"{pdf_file.stem}_compressed.pdf"

                if compressed_file.exists():
                    total_compressed_size += self.get_file_size(compressed_file)

        # Summary
        print(f"\n{'=' * 60}")
        print("BATCH COMPRESSION SUMMARY")
        print(f"{'=' * 60}")
        print(f"Files processed: {len(pdf_files)}")
        print(f"Successful compressions: {success_count}")
        print(f"Failed compressions: {len(pdf_files) - success_count}")
        print(f"Total original size: {self.format_size(total_original_size)}")
        print(f"Total compressed size: {self.format_size(total_compressed_size)}")

        if total_original_size > 0:
            if total_compressed_size < total_original_size:
                total_ratio = ((total_original_size - total_compressed_size) / total_original_size) * 100
                total_saved = total_original_size - total_compressed_size
                print(f"Overall compression: {total_ratio:.1f}% reduction")
                print(f"Total space saved: {self.format_size(total_saved)}")
            else:
                total_ratio = ((total_compressed_size - total_original_size) / total_original_size) * 100
                print(f"Overall size change: {total_ratio:.1f}% increase")


def main():
    parser = argparse.ArgumentParser(
        description='Interactive PDF Compressor with four compression modes',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
COMPRESSION MODES:
  conservative - High quality, minimal compression (10-30% reduction)
  balanced     - Medium quality, moderate compression (30-50% reduction)
  aggressive   - Lower quality, high compression (50-70% reduction)
  nuclear      - Text-focused, maximum compression (70-90% reduction)

EXAMPLES:
  python compressor.py document.pdf
  python compressor.py ./pdfs -o ./compressed
  python compressor.py document.pdf -c aggressive
        """
    )

    parser.add_argument('input', help='Input PDF file or directory')
    parser.add_argument('-o', '--output', help='Output directory (optional)')
    parser.add_argument('-c', '--compression',
                        choices=['conservative', 'balanced', 'aggressive', 'nuclear'],
                        help='Compression mode (skip interactive selection)')
    parser.add_argument('-g', '--ghostscript', help='Path to Ghostscript executable')
    parser.add_argument('--batch', action='store_true',
                        help='Non-interactive batch mode (uses recommended compression)')

    args = parser.parse_args()

    try:
        compressor = InteractivePDFCompressor(args.ghostscript)

        input_path = Path(args.input)
        interactive = not args.batch

        if input_path.is_file():
            compressor.process_single_file(args.input, args.output, args.compression, interactive)
        elif input_path.is_dir():
            compressor.process_directory(args.input, args.output, args.compression, interactive)
        else:
            print(f"ERROR: {args.input} is not a valid file or directory")
            sys.exit(1)

    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) == 1:
        print("Interactive PDF Compressor")
        print("=" * 50)
        print("Compress PDF files with four different quality levels")
        print("\nUSAGE:")
        print("python compressor.py <input_file_or_directory> [options]")
        print("\nEXAMPLES:")
        print("python compressor.py document.pdf")
        print("python compressor.py ./pdfs -o ./compressed")
        print("python compressor.py document.pdf -c aggressive")
        print("\nFor more options, use: python compressor.py --help")
        print("\nDrag and drop files to compress them interactively!")
        input("\nPress Enter to exit...")
    else:
        main()