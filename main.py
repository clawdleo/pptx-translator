"""
PPTX Translator API
-------------------
FastAPI application for translating PowerPoint presentations.

Endpoints:
- GET /: Serve the web interface
- POST /api/translate: Upload and translate a PPTX file
- GET /health: Health check endpoint
"""

import os
import uuid
import shutil
import logging
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

from translator import Translator
from pptx_processor import PPTXProcessor

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Initialize FastAPI app
app = FastAPI(
    title="PPTX Translator",
    description="Translate PowerPoint presentations while preserving formatting",
    version="2.0.0"
)

# Temporary directory for file processing
TEMP_DIR = Path("/tmp/pptx-translator")
TEMP_DIR.mkdir(exist_ok=True)

# DeepL API key (optional, loaded from environment)
DEEPL_API_KEY = os.environ.get('DEEPL_API_KEY', 'e87352a7-9518-4019-bb38-73f09eb2581b:fx')

# Supported languages
SUPPORTED_LANGUAGES = ['slovenian', 'croatian', 'serbian', 'english', 'german', 'french', 'spanish', 'italian']


@app.get("/", response_class=HTMLResponse)
async def root():
    """Serve the web interface."""
    return get_html_page()


@app.get("/health")
async def health_check():
    """Health check endpoint for Render."""
    return {
        "status": "healthy",
        "service": "pptx-translator",
        "version": "2.0.0"
    }


@app.post("/api/translate")
async def translate_pptx(
    file: UploadFile = File(...),
    language: str = Form(default="slovenian"),
    use_deepl: bool = Form(default=True)
):
    """
    Translate a PowerPoint file.
    
    Args:
        file: The .pptx file to translate
        language: Target language (slovenian, croatian, serbian, etc.)
        use_deepl: Whether to use DeepL API (True) or googletrans (False)
    
    Returns:
        The translated .pptx file as a download
    """
    # Validate file type
    if not file.filename.lower().endswith('.pptx'):
        raise HTTPException(
            status_code=400,
            detail="Only .pptx files are supported"
        )
    
    # Validate language
    if language.lower() not in SUPPORTED_LANGUAGES:
        raise HTTPException(
            status_code=400,
            detail=f"Language must be one of: {', '.join(SUPPORTED_LANGUAGES)}"
        )
    
    # Generate unique file paths
    job_id = str(uuid.uuid4())[:8]
    input_path = TEMP_DIR / f"{job_id}_input.pptx"
    output_path = TEMP_DIR / f"{job_id}_output.pptx"
    
    try:
        # Save uploaded file
        logger.info(f"Processing file: {file.filename} -> {language}")
        with open(input_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
        
        # Initialize translator
        deepl_key = DEEPL_API_KEY if use_deepl else None
        translator = Translator(target_lang=language, deepl_api_key=deepl_key)
        
        # Process the file
        processor = PPTXProcessor(translator)
        stats = processor.process_file(str(input_path), str(output_path))
        
        logger.info(f"Translation complete: {stats}")
        
        # Generate output filename
        original_name = Path(file.filename).stem
        output_filename = f"{original_name}_{language}.pptx"
        
        # Return the translated file
        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            background=cleanup_task(input_path, output_path)
        )
        
    except Exception as e:
        logger.error(f"Translation failed: {e}")
        # Cleanup on error
        cleanup_files(input_path, output_path)
        raise HTTPException(
            status_code=500,
            detail=f"Translation failed: {str(e)}"
        )


def cleanup_files(*paths):
    """Remove temporary files."""
    for path in paths:
        try:
            if path.exists():
                path.unlink()
        except Exception as e:
            logger.warning(f"Failed to cleanup {path}: {e}")


async def cleanup_task(input_path, output_path):
    """Background task to cleanup files after response is sent."""
    import asyncio
    await asyncio.sleep(60)  # Wait 60 seconds before cleanup
    cleanup_files(input_path, output_path)


def get_html_page():
    """Return the HTML page for the web interface."""
    return '''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PPTX Translator</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        
        .container {
            background: rgba(255, 255, 255, 0.97);
            border-radius: 20px;
            padding: 40px;
            max-width: 520px;
            width: 100%;
            box-shadow: 0 25px 50px rgba(0, 0, 0, 0.3);
        }
        
        .logo { text-align: center; margin-bottom: 30px; }
        .logo h1 { font-size: 26px; color: #1a1a2e; margin-bottom: 5px; }
        .logo p { color: #666; font-size: 14px; }
        
        .upload-area {
            border: 3px dashed #ddd;
            border-radius: 15px;
            padding: 40px 20px;
            text-align: center;
            transition: all 0.3s ease;
            cursor: pointer;
            margin-bottom: 20px;
        }
        
        .upload-area:hover, .upload-area.dragover {
            border-color: #0f3460;
            background: rgba(15, 52, 96, 0.05);
        }
        
        .upload-area.has-file {
            border-color: #27ae60;
            background: rgba(39, 174, 96, 0.05);
        }
        
        .upload-icon { font-size: 48px; margin-bottom: 15px; }
        .upload-text { color: #666; font-size: 16px; }
        .file-name { color: #27ae60; font-weight: 600; margin-top: 10px; word-break: break-all; }
        
        #file-input { display: none; }
        
        .form-group { margin-bottom: 20px; }
        .form-group label { display: block; margin-bottom: 8px; color: #333; font-weight: 500; }
        
        .form-group select {
            width: 100%;
            padding: 12px 15px;
            border: 2px solid #ddd;
            border-radius: 10px;
            font-size: 16px;
            background: white;
            cursor: pointer;
        }
        
        .form-group select:focus { outline: none; border-color: #0f3460; }
        
        .translate-btn {
            width: 100%;
            padding: 15px;
            background: linear-gradient(135deg, #0f3460 0%, #1a1a2e 100%);
            color: white;
            border: none;
            border-radius: 10px;
            font-size: 18px;
            font-weight: 600;
            cursor: pointer;
            transition: transform 0.2s, box-shadow 0.2s;
        }
        
        .translate-btn:hover:not(:disabled) {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(15, 52, 96, 0.3);
        }
        
        .translate-btn:disabled { background: #ccc; cursor: not-allowed; }
        
        .progress { display: none; margin-top: 20px; text-align: center; }
        .progress.active { display: block; }
        
        .spinner {
            width: 40px; height: 40px;
            border: 4px solid #ddd;
            border-top-color: #0f3460;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin: 0 auto 15px;
        }
        
        @keyframes spin { to { transform: rotate(360deg); } }
        
        .progress-text { color: #666; }
        
        .error {
            background: #fee; border: 1px solid #fcc; color: #c00;
            padding: 15px; border-radius: 10px; margin-top: 20px;
            display: none;
        }
        .error.active { display: block; }
        
        .features {
            margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee;
        }
        .features h3 { color: #333; font-size: 14px; margin-bottom: 10px; }
        .features ul { list-style: none; color: #666; font-size: 13px; }
        .features li { padding: 5px 0; }
        .features li::before { content: "✓ "; color: #27ae60; }
    </style>
</head>
<body>
    <div class="container">
        <div class="logo">
            <h1>📊 PPTX Translator</h1>
            <p>Translate presentations with perfect formatting</p>
        </div>
        
        <div class="upload-area" id="upload-area">
            <div class="upload-icon">📁</div>
            <div class="upload-text">Drop your .pptx file here or click to browse</div>
            <div class="file-name" id="file-name"></div>
        </div>
        <input type="file" id="file-input" accept=".pptx">
        
        <div class="form-group">
            <label for="language">Translate to:</label>
            <select id="language">
                <option value="slovenian">🇸🇮 Slovenian</option>
                <option value="croatian">🇭🇷 Croatian</option>
                <option value="serbian">🇷🇸 Serbian</option>
                <option value="german">🇩🇪 German</option>
                <option value="french">🇫🇷 French</option>
                <option value="spanish">🇪🇸 Spanish</option>
                <option value="italian">🇮🇹 Italian</option>
            </select>
        </div>
        
        <button class="translate-btn" id="translate-btn" disabled>Translate Presentation</button>
        
        <div class="progress" id="progress">
            <div class="spinner"></div>
            <div class="progress-text">Translating... This may take a minute.</div>
        </div>
        
        <div class="error" id="error"></div>
        
        <div class="features">
            <h3>Features:</h3>
            <ul>
                <li>Preserves all formatting (fonts, colors, sizes)</li>
                <li>Handles grouped shapes and diagrams</li>
                <li>Translates tables and speaker notes</li>
                <li>Powered by DeepL for accuracy</li>
            </ul>
        </div>
    </div>
    
    <script>
        const uploadArea = document.getElementById('upload-area');
        const fileInput = document.getElementById('file-input');
        const fileName = document.getElementById('file-name');
        const translateBtn = document.getElementById('translate-btn');
        const progress = document.getElementById('progress');
        const error = document.getElementById('error');
        const languageSelect = document.getElementById('language');
        
        let selectedFile = null;
        
        uploadArea.addEventListener('click', () => fileInput.click());
        
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });
        
        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });
        
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            if (file && file.name.toLowerCase().endsWith('.pptx')) {
                handleFile(file);
            } else {
                showError('Please upload a .pptx file');
            }
        });
        
        fileInput.addEventListener('change', (e) => {
            if (e.target.files[0]) handleFile(e.target.files[0]);
        });
        
        function handleFile(file) {
            selectedFile = file;
            fileName.textContent = file.name;
            uploadArea.classList.add('has-file');
            translateBtn.disabled = false;
            error.classList.remove('active');
        }
        
        function showError(message) {
            error.textContent = message;
            error.classList.add('active');
        }
        
        translateBtn.addEventListener('click', async () => {
            if (!selectedFile) return;
            
            translateBtn.disabled = true;
            progress.classList.add('active');
            error.classList.remove('active');
            
            const formData = new FormData();
            formData.append('file', selectedFile);
            formData.append('language', languageSelect.value);
            formData.append('use_deepl', 'true');
            
            try {
                const response = await fetch('/api/translate', {
                    method: 'POST',
                    body: formData
                });
                
                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.detail || 'Translation failed');
                }
                
                const blob = await response.blob();
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = selectedFile.name.replace('.pptx', `_${languageSelect.value}.pptx`);
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                
                progress.classList.remove('active');
                translateBtn.disabled = false;
                
            } catch (err) {
                progress.classList.remove('active');
                translateBtn.disabled = false;
                showError(err.message);
            }
        });
    </script>
</body>
</html>'''


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
