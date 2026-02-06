"""
Translator Module
-----------------
Handles text translation with retry logic and multiple backend support.
Primary: googletrans (free)
Backup: DeepL API (requires API key)
"""

import time
import logging
from typing import Optional
from functools import lru_cache

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Language code mapping
LANGUAGE_CODES = {
    'slovenian': 'sl',
    'croatian': 'hr',
    'serbian': 'sr',
    'english': 'en',
    'german': 'de',
    'italian': 'it',
    'french': 'fr',
    'spanish': 'es',
}


class Translator:
    """
    Translation engine with retry mechanism and caching.
    Uses googletrans as primary, with DeepL as optional backup.
    """
    
    def __init__(self, target_lang: str, deepl_api_key: Optional[str] = None):
        """
        Initialize translator with target language.
        
        Args:
            target_lang: Target language (e.g., 'slovenian', 'croatian', 'sr')
            deepl_api_key: Optional DeepL API key for premium translation
        """
        self.target_lang = LANGUAGE_CODES.get(target_lang.lower(), target_lang.lower())
        self.deepl_api_key = deepl_api_key
        self._translator = None
        self._cache = {}
        self._init_translator()
    
    def _init_translator(self):
        """Initialize the translator."""
        self._translator = None
        logger.info(f"Initialized translator for target language: {self.target_lang}")
    
    def translate(self, text: str, max_retries: int = 3) -> str:
        """
        Translate text to target language with retry logic.
        
        Args:
            text: Source text to translate
            max_retries: Maximum retry attempts on failure
            
        Returns:
            Translated text, or original text on failure
        """
        # Skip empty or whitespace-only text
        if not text or not text.strip():
            return text
        
        # Skip if only numbers, punctuation, or very short
        stripped = text.strip()
        if len(stripped) < 2:
            return text
        if stripped.replace(' ', '').replace('\n', '').isnumeric():
            return text
        
        # Check cache first
        cache_key = f"{self.target_lang}:{text}"
        if cache_key in self._cache:
            return self._cache[cache_key]
        
        # Use DeepL API
        result = self._translate_deepl(text, max_retries)
        if result:
            self._cache[cache_key] = result
            return result
        
        # Return original on failure
        logger.warning(f"Translation failed for: {text[:50]}...")
        return text
    
    def _translate_deepl(self, text: str, max_retries: int) -> Optional[str]:
        """
        Translate using DeepL API with retry logic.
        """
        import httpx
        import traceback
        
        # DeepL uses uppercase language codes
        deepl_lang = self.target_lang.upper()
        
        for attempt in range(max_retries):
            try:
                logger.debug(f"DeepL attempt {attempt + 1}: translating {len(text)} chars to {deepl_lang}")
                
                response = httpx.post(
                    'https://api-free.deepl.com/v2/translate',
                    headers={
                        'Authorization': f'DeepL-Auth-Key {self.deepl_api_key}',
                        'Content-Type': 'application/json'
                    },
                    json={
                        'text': [text],
                        'target_lang': deepl_lang
                    },
                    timeout=30.0
                )
                
                if response.status_code == 200:
                    data = response.json()
                    if data.get('translations'):
                        translated = data['translations'][0]['text']
                        logger.debug(f"DeepL success: '{text[:30]}...' -> '{translated[:30]}...'")
                        return translated
                    else:
                        logger.warning(f"DeepL returned empty translations: {data}")
                elif response.status_code == 429:
                    logger.warning("DeepL rate limited, waiting...")
                    time.sleep(5)
                elif response.status_code == 456:
                    logger.error("DeepL quota exceeded")
                    return None
                else:
                    logger.warning(f"DeepL returned status {response.status_code}: {response.text[:200]}")
                    
            except httpx.TimeoutException:
                logger.warning(f"DeepL timeout on attempt {attempt + 1}")
                if attempt < max_retries - 1:
                    time.sleep(2 ** attempt)
            except Exception as e:
                logger.error(f"DeepL attempt {attempt + 1} failed: {e}")
                print(f"DeepL error: {e}")
                print(traceback.format_exc())
                if attempt < max_retries - 1:
                    time.sleep(2 ** attempt)
        
        logger.error(f"All DeepL attempts failed for text: {text[:50]}...")
        return None
    
    def get_stats(self) -> dict:
        """Return translation statistics."""
        return {
            'cached_translations': len(self._cache),
            'target_language': self.target_lang
        }
