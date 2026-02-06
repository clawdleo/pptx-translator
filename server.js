const express = require('express');
const multer = require('multer');
const AdmZip = require('adm-zip');
const xml2js = require('xml2js');
const { v4: uuidv4 } = require('uuid');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// Storage for uploads
const upload = multer({ 
  dest: '/tmp/uploads/',
  limits: { fileSize: 50 * 1024 * 1024 } // 50MB max
});

app.use(express.static('public'));
app.use(express.json());

// Language codes for Google Translate
const LANG_CODES = {
  'slovenian': 'sl',
  'croatian': 'hr', 
  'serbian': 'sr',
  'english': 'en'
};

// Delay helper
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

// Free translation using Google Translate (unofficial API)
async function translateText(text, targetLang) {
  if (!text || text.trim().length === 0) return text;
  // Skip if just numbers, whitespace, or single chars
  if (/^[\d\s\.\,\-\+\%\€\$\£]+$/.test(text)) return text;
  if (text.trim().length < 2) return text;
  
  const langCode = LANG_CODES[targetLang] || targetLang;
  
  try {
    // Add small delay to avoid rate limiting
    await delay(100);
    
    const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=${langCode}&dt=t&q=${encodeURIComponent(text)}`;
    const response = await fetch(url);
    const responseText = await response.text();
    
    // Check if response is valid JSON
    if (responseText.startsWith('<!') || responseText.startsWith('<html')) {
      console.error('Got HTML response, rate limited. Returning original.');
      return text;
    }
    
    const data = JSON.parse(responseText);
    
    if (data && data[0]) {
      return data[0].map(item => item[0]).filter(Boolean).join('');
    }
    return text;
  } catch (error) {
    console.error('Translation error:', error.message);
    return text; // Return original on error
  }
}

// Recursively find and translate all text in XML object
async function translateXmlObject(obj, targetLang, stats) {
  if (!obj) return obj;
  
  if (Array.isArray(obj)) {
    for (let i = 0; i < obj.length; i++) {
      obj[i] = await translateXmlObject(obj[i], targetLang, stats);
    }
    return obj;
  }
  
  if (typeof obj === 'object') {
    // Handle text elements <a:t>
    if (obj['a:t']) {
      for (let i = 0; i < obj['a:t'].length; i++) {
        const text = obj['a:t'][i];
        if (typeof text === 'string' && text.trim().length > 0) {
          const translated = await translateText(text, targetLang);
          obj['a:t'][i] = translated;
          stats.translated++;
        }
      }
    }
    
    // Recurse into all properties
    for (const key of Object.keys(obj)) {
      obj[key] = await translateXmlObject(obj[key], targetLang, stats);
    }
  }
  
  return obj;
}

// Process PPTX file
async function processPptx(inputPath, targetLang) {
  const zip = new AdmZip(inputPath);
  const entries = zip.getEntries();
  const stats = { total: 0, translated: 0, files: 0 };
  
  const parser = new xml2js.Parser({ preserveChildrenOrder: true, explicitChildren: true });
  const builder = new xml2js.Builder({ headless: true, renderOpts: { pretty: false } });
  
  for (const entry of entries) {
    // Process slide XMLs, masters, layouts
    if (entry.entryName.match(/ppt\/(slides|slideMasters|slideLayouts|notesSlides)\/.*\.xml$/)) {
      stats.files++;
      
      try {
        const xmlContent = entry.getData().toString('utf8');
        const parsed = await parser.parseStringPromise(xmlContent);
        
        // Translate all text in the parsed XML
        await translateXmlObject(parsed, targetLang, stats);
        
        // Rebuild XML
        const newXml = builder.buildObject(parsed);
        zip.updateFile(entry.entryName, Buffer.from(newXml, 'utf8'));
        
      } catch (err) {
        console.error(`Error processing ${entry.entryName}:`, err.message);
      }
    }
  }
  
  const outputPath = inputPath.replace('.pptx', `_${targetLang}.pptx`);
  zip.writeZip(outputPath);
  
  return { outputPath, stats };
}

// Process DOCX file
async function processDocx(inputPath, targetLang) {
  const zip = new AdmZip(inputPath);
  const entries = zip.getEntries();
  const stats = { total: 0, translated: 0, files: 0 };
  
  const parser = new xml2js.Parser({ preserveChildrenOrder: true, explicitChildren: true });
  const builder = new xml2js.Builder({ headless: true, renderOpts: { pretty: false } });
  
  for (const entry of entries) {
    // Process Word document XMLs
    if (entry.entryName.match(/word\/(document|header[0-9]*|footer[0-9]*|comments|footnotes|endnotes)\.xml$/)) {
      stats.files++;
      
      try {
        const xmlContent = entry.getData().toString('utf8');
        const parsed = await parser.parseStringPromise(xmlContent);
        
        // Translate all text in the parsed XML (Word uses w:t for text)
        await translateDocxObject(parsed, targetLang, stats);
        
        // Rebuild XML
        const newXml = builder.buildObject(parsed);
        zip.updateFile(entry.entryName, Buffer.from(newXml, 'utf8'));
        
      } catch (err) {
        console.error(`Error processing ${entry.entryName}:`, err.message);
      }
    }
  }
  
  const outputPath = inputPath.replace('.docx', `_${targetLang}.docx`);
  zip.writeZip(outputPath);
  
  return { outputPath, stats };
}

// Recursively find and translate all text in DOCX XML object
async function translateDocxObject(obj, targetLang, stats) {
  if (!obj) return obj;
  
  if (Array.isArray(obj)) {
    for (let i = 0; i < obj.length; i++) {
      obj[i] = await translateDocxObject(obj[i], targetLang, stats);
    }
    return obj;
  }
  
  if (typeof obj === 'object') {
    // Handle Word text elements <w:t>
    if (obj['w:t']) {
      for (let i = 0; i < obj['w:t'].length; i++) {
        let text = obj['w:t'][i];
        // w:t can be string or object with _ property
        if (typeof text === 'object' && text._) {
          if (text._.trim().length > 0) {
            const translated = await translateText(text._, targetLang);
            obj['w:t'][i]._ = translated;
            stats.translated++;
            console.log(`Translated: "${text._}" → "${translated}"`);
          }
        } else if (typeof text === 'string' && text.trim().length > 0) {
          const translated = await translateText(text, targetLang);
          obj['w:t'][i] = translated;
          stats.translated++;
          console.log(`Translated: "${text}" → "${translated}"`);
        }
      }
    }
    
    // Recurse into all properties
    for (const key of Object.keys(obj)) {
      if (key !== 'w:t') { // Don't re-process w:t
        obj[key] = await translateDocxObject(obj[key], targetLang, stats);
      }
    }
  }
  
  return obj;
}

// API endpoint for translation
app.post('/api/translate', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    
    const targetLang = req.body.language || 'slovenian';
    const originalName = req.file.originalname.toLowerCase();
    const isPptx = originalName.endsWith('.pptx');
    const isDocx = originalName.endsWith('.docx');
    
    if (!isPptx && !isDocx) {
      return res.status(400).json({ error: 'Only .pptx and .docx files are supported' });
    }
    
    console.log(`Processing: ${req.file.originalname} → ${targetLang}`);
    
    // Rename for processing
    const ext = isPptx ? '.pptx' : '.docx';
    const inputPath = req.file.path + ext;
    fs.renameSync(req.file.path, inputPath);
    
    // Process based on file type
    const result = isPptx 
      ? await processPptx(inputPath, targetLang)
      : await processDocx(inputPath, targetLang);
    
    // Read the translated file
    const translatedFile = fs.readFileSync(result.outputPath);
    const outputName = req.file.originalname.replace(ext, `_${targetLang}${ext}`);
    
    // Cleanup
    fs.unlinkSync(inputPath);
    fs.unlinkSync(result.outputPath);
    
    const contentType = isPptx 
      ? 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
      : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    
    res.setHeader('Content-Disposition', `attachment; filename="${outputName}"`);
    res.setHeader('Content-Type', contentType);
    res.send(translatedFile);
    
    console.log(`Done: ${result.stats.translated} texts translated in ${result.stats.files} files`);
    
  } catch (error) {
    console.error('Processing error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Health check
app.get('/health', (req, res) => {
  res.json({ status: 'ok', service: 'pptx-translator' });
});

app.listen(PORT, () => {
  console.log(`PPTX Translator running on port ${PORT}`);
});
