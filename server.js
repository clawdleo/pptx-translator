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

// Free translation using Google Translate (unofficial API)
async function translateText(text, targetLang) {
  if (!text || text.trim().length === 0) return text;
  
  const langCode = LANG_CODES[targetLang] || targetLang;
  
  try {
    const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=${langCode}&dt=t&q=${encodeURIComponent(text)}`;
    const response = await fetch(url);
    const data = await response.json();
    
    if (data && data[0]) {
      return data[0].map(item => item[0]).join('');
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

// API endpoint for translation
app.post('/api/translate', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    
    const targetLang = req.body.language || 'slovenian';
    console.log(`Processing: ${req.file.originalname} → ${targetLang}`);
    
    // Rename to .pptx for processing
    const inputPath = req.file.path + '.pptx';
    fs.renameSync(req.file.path, inputPath);
    
    const result = await processPptx(inputPath, targetLang);
    
    // Read the translated file
    const translatedFile = fs.readFileSync(result.outputPath);
    const outputName = req.file.originalname.replace('.pptx', `_${targetLang}.pptx`);
    
    // Cleanup
    fs.unlinkSync(inputPath);
    fs.unlinkSync(result.outputPath);
    
    res.setHeader('Content-Disposition', `attachment; filename="${outputName}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
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
