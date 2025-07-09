# PDF Compression

## Ghost Compression Usage:

### 1. Interactive mode (prompts for compression choice)
```
python compressor.py document.pdf
```

### 2. Batch directory processing with interactive mode selection
```
python compressor.py ./pdfs -o ./compressed
```

### 3. Skip interactive selection, use specific mode
```
python compressor.py document.pdf -c aggressive
```

### 4. Non-interactive batch mode (uses recommended compression)
```
python compressor.py ./pdfs --batch -o ./output
```
## Beast Compression Usage:

### 1. Single file
```
python beast_compressor.py document.pdf
```

### 2. Directory with custom output
```
python beast_compressor.py ./documents -o ./compressed
```

### 3. With custom Ghostscript path
```
python beast_compressor.py file.pdf -g "D:\gs\bin\gswin64c.exe"
```
<hr>

© 2025 Fahad Elahi Khan · Licensed under [CC BY-NC-ND 4.0](https://creativecommons.org/licenses/by-nc-nd/4.0/)
