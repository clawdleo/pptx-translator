"""
PPTX Translator API
-------------------
FastAPI application for translating PowerPoint presentations.
Includes aggressive error handling and logging for debugging.
"""

import os
import uuid
import shutil
import logging
import traceback
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

from translator import Translator
from pptx_processor import PPTXProcessor

# Configure logging - verbose for debugging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Initialize FastAPI app
app = FastAPI(
    title="PPTX Translator",
    description="Translate PowerPoint presentations while preserving formatting",
    version="2.1.0"
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Temporary directory for file processing
TEMP_DIR = Path("/tmp/pptx-translator")
TEMP_DIR.mkdir(exist_ok=True)

# DeepL API key
DEEPL_API_KEY = os.environ.get('DEEPL_API_KEY', 'e87352a7-9518-4019-bb38-73f09eb2581b:fx')

# Supported languages
SUPPORTED_LANGUAGES = ['slovenian', 'croatian', 'serbian', 'english', 'german', 'french', 'spanish', 'italian']


# Global exception handler
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    """Catch all unhandled exceptions and return proper JSON."""
    error_msg = str(exc)
    logger.error(f"Unhandled exception: {error_msg}")
    logger.error(traceback.format_exc())
    print(f"CRITICAL ERROR: {error_msg}")
    print(traceback.format_exc())
    return JSONResponse(
        status_code=500,
        content={"error": error_msg, "detail": "An unexpected error occurred. Check server logs."}
    )


@app.get("/", response_class=HTMLResponse)
async def root():
    """Serve the web interface."""
    try:
        return get_html_page()
    except Exception as e:
        logger.error(f"Error serving HTML: {e}")
        print(traceback.format_exc())
        return HTMLResponse(content=f"<h1>Error: {e}</h1>", status_code=500)


@app.get("/health")
async def health_check():
    """Health check endpoint for Render."""
    return {"status": "healthy", "service": "pptx-translator", "version": "2.1.0"}


@app.post("/api/translate")
async def translate_pptx(
    file: UploadFile = File(...),
    language: str = Form(default="slovenian"),
    use_deepl: bool = Form(default=True)
):
    """
    Translate a PowerPoint file with comprehensive error handling.
    """
    input_path = None
    output_path = None
    
    try:
        logger.info(f"=== NEW TRANSLATION REQUEST ===")
        logger.info(f"File: {file.filename}, Language: {language}")
        print(f"Processing: {file.filename} -> {language}")
        
        # Validate file type
        if not file.filename:
            return JSONResponse(
                status_code=400,
                content={"error": "No filename provided"}
            )
        
        if not file.filename.lower().endswith('.pptx'):
            return JSONResponse(
                status_code=400,
                content={"error": "Only .pptx files are supported"}
            )
        
        # Validate language
        lang_lower = language.lower()
        if lang_lower not in SUPPORTED_LANGUAGES:
            return JSONResponse(
                status_code=400,
                content={"error": f"Language must be one of: {', '.join(SUPPORTED_LANGUAGES)}"}
            )
        
        # Generate unique file paths
        job_id = str(uuid.uuid4())[:8]
        input_path = TEMP_DIR / f"{job_id}_input.pptx"
        output_path = TEMP_DIR / f"{job_id}_output.pptx"
        
        logger.info(f"Job ID: {job_id}")
        
        # Save uploaded file
        try:
            with open(input_path, "wb") as f:
                content = await file.read()
                f.write(content)
            logger.info(f"Saved input file: {input_path} ({len(content)} bytes)")
            print(f"Saved: {input_path} ({len(content)} bytes)")
        except Exception as e:
            logger.error(f"Failed to save uploaded file: {e}")
            print(traceback.format_exc())
            return JSONResponse(
                status_code=500,
                content={"error": f"Failed to save uploaded file: {str(e)}"}
            )
        
        # Initialize translator
        try:
            logger.info(f"Initializing translator with DeepL API")
            translator = Translator(target_lang=lang_lower, deepl_api_key=DEEPL_API_KEY)
            logger.info("Translator initialized successfully")
        except Exception as e:
            logger.error(f"Failed to initialize translator: {e}")
            print(traceback.format_exc())
            return JSONResponse(
                status_code=500,
                content={"error": f"Failed to initialize translation service: {str(e)}"}
            )
        
        # Process the file
        try:
            logger.info("Starting PPTX processing...")
            print("Starting PPTX processing...")
            processor = PPTXProcessor(translator)
            stats = processor.process_file(str(input_path), str(output_path))
            logger.info(f"Processing complete: {stats}")
            print(f"Processing complete: {stats}")
        except Exception as e:
            logger.error(f"Failed to process PPTX: {e}")
            print(f"PPTX PROCESSING ERROR: {e}")
            print(traceback.format_exc())
            
            # Check for specific translation errors
            error_msg = str(e).lower()
            if 'nonetype' in error_msg or 'group' in error_msg:
                return JSONResponse(
                    status_code=500,
                    content={"error": "Translation service failed. Please try again later or use a smaller file."}
                )
            
            return JSONResponse(
                status_code=500,
                content={"error": f"Failed to process presentation: {str(e)}"}
            )
        
        # Verify output file exists
        if not output_path.exists():
            logger.error("Output file was not created")
            return JSONResponse(
                status_code=500,
                content={"error": "Translation completed but output file was not created"}
            )
        
        # Generate output filename
        original_name = Path(file.filename).stem
        output_filename = f"{original_name}_{language}.pptx"
        
        logger.info(f"Returning translated file: {output_filename}")
        print(f"Success! Returning: {output_filename}")
        
        # Return the translated file
        return FileResponse(
            path=str(output_path),
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
        
    except Exception as e:
        # Catch-all for any unexpected errors
        logger.error(f"UNEXPECTED ERROR in translate_pptx: {e}")
        print(f"UNEXPECTED ERROR: {e}")
        print(traceback.format_exc())
        return JSONResponse(
            status_code=500,
            content={"error": f"Unexpected error: {str(e)}"}
        )
    
    finally:
        # Cleanup input file (output file will be cleaned up after response)
        try:
            if input_path and input_path.exists():
                input_path.unlink()
                logger.debug(f"Cleaned up input file: {input_path}")
        except Exception as e:
            logger.warning(f"Failed to cleanup input file: {e}")


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
            display: none; word-break: break-word;
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
        const errorDiv = document.getElementById('error');
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
            fileName.textContent = file.name + ' (' + (file.size / 1024 / 1024).toFixed(2) + ' MB)';
            uploadArea.classList.add('has-file');
            translateBtn.disabled = false;
            errorDiv.classList.remove('active');
        }
        
        function showError(message) {
            errorDiv.textContent = message;
            errorDiv.classList.add('active');
        }
        
        translateBtn.addEventListener('click', async () => {
            if (!selectedFile) return;
            
            translateBtn.disabled = true;
            progress.classList.add('active');
            errorDiv.classList.remove('active');
            
            const formData = new FormData();
            formData.append('file', selectedFile);
            formData.append('language', languageSelect.value);
            formData.append('use_deepl', 'true');
            
            try {
                const response = await fetch('/api/translate', {
                    method: 'POST',
                    body: formData
                });
                
                // Check content type to determine response type
                const contentType = response.headers.get('content-type') || '';
                
                if (!response.ok) {
                    // Try to parse error as JSON
                    let errorMsg = 'Translation failed';
                    try {
                        if (contentType.includes('application/json')) {
                            const err = await response.json();
                            errorMsg = err.error || err.detail || errorMsg;
                        } else {
                            errorMsg = await response.text() || errorMsg;
                        }
                    } catch (e) {
                        errorMsg = `Server error (${response.status})`;
                    }
                    throw new Error(errorMsg);
                }
                
                // Check if response is a file (success) or JSON (error)
                if (contentType.includes('application/json')) {
                    const data = await response.json();
                    if (data.error) {
                        throw new Error(data.error);
                    }
                }
                
                // Success - download the file
                const blob = await response.blob();
                if (blob.size === 0) {
                    throw new Error('Received empty file from server');
                }
                
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = selectedFile.name.replace('.pptx', '_' + languageSelect.value + '.pptx');
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                
                progress.classList.remove('active');
                translateBtn.disabled = false;
                
            } catch (err) {
                progress.classList.remove('active');
                translateBtn.disabled = false;
                showError(err.message || 'An unexpected error occurred');
            }
        });
    </script>
</body>
</html>'''


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    print(f"Starting PPTX Translator on port {port}")
    uvicorn.run(app, host="0.0.0.0", port=port)
