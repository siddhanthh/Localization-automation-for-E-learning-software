import re
import time
import logging
from threading import Lock

logger = logging.getLogger("docx_translator.translator")

# Regex patterns for protected strings that should not be translated
PROTECTED_PATTERNS = [
    r'^[A-Z0-9_-]+$',  # IDs, SKUs, Part numbers like ID-001, ABC-12345
    r'^v?\d+(\.\d+)+$',  # Version numbers like v2.1.4, 1.0.0
    r'https?://[^\s]+',  # URLs
    r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$',  # Emails
    r'^[a-zA-Z0-9_-]+\.[a-zA-Z0-9]+$',  # File names like document.docx
    r'^[a-zA-Z]:\\.*$',  # Windows file paths
    r'^/.*$',  # Unix file paths
    r'^\{[a-zA-Z_0-9]+\}$',  # Placeholders like {username}
    r'^\{\d+\}$',  # Placeholders like {0}
    r'^%\w$',  # Placeholders like %s, %d
    r'^%\([a-zA-Z_0-9]+\)\w$',  # Placeholders like %(name)s
    r'^\{\{[a-zA-Z_0-9]+\}\}$',  # Placeholders like {{customer_name}}
    r'^<[A-Z0-9_]+>$',  # Tokens like <TOKEN>
    r'^\[[A-Z0-9_]+\]$',  # Variables like [VARIABLE]
]

# Unicode patterns for target languages to detect pre-translated text
SCRIPT_PATTERNS = {
    'th': re.compile(r'[\u0e00-\u0e7f]'),                        # Thai
    'ru': re.compile(r'[\u0400-\u04ff]'),                        # Russian / Cyrillic
    'zh': re.compile(r'[\u4e00-\u9fff]'),                        # Chinese (Hanzi)
    'ja': re.compile(r'[\u3040-\u30ff\u4e00-\u9fff]'),            # Japanese (Kana + Kanji)
    'ko': re.compile(r'[\uac00-\ud7af\u1100-\u11ff]'),            # Korean (Hangul)
    'ar': re.compile(r'[\u0600-\u06ff]'),                        # Arabic
    'he': re.compile(r'[\u0590-\u05ff]'),                        # Hebrew
    'el': re.compile(r'[\u0370-\u03ff]'),                        # Greek
    'hi': re.compile(r'[\u0900-\u097f]'),                        # Hindi / Devanagari
}

class RateLimiter:
    """Centralized rate controller to enforce requests-per-minute limits."""
    def __init__(self, requests_per_minute=60, backoff_factor=2.0, max_retries=5):
        self.requests_per_minute = requests_per_minute
        self.delay = 60.0 / requests_per_minute if requests_per_minute > 0 else 0
        self.backoff_factor = backoff_factor
        self.max_retries = max_retries
        self.last_request_time = 0.0
        self.lock = Lock()

    def wait(self):
        if self.delay <= 0:
            return
        with self.lock:
            now = time.time()
            elapsed = now - self.last_request_time
            if elapsed < self.delay:
                time.sleep(self.delay - elapsed)
            self.last_request_time = time.time()

def resolve_language_code(lang):
    """Resolves language names (e.g. 'spanish') to standard ISO codes (e.g. 'es')."""
    if not lang:
        return lang
    
    lang_lower = lang.lower().strip()
    
    try:
        from deep_translator import GoogleTranslator
        langs_dict = GoogleTranslator().get_supported_languages(as_dict=True)
        if lang_lower in langs_dict:
            return langs_dict[lang_lower]
        if lang_lower in langs_dict.values():
            return lang_lower
    except Exception:
        pass
        
    return lang_lower

class TranslatorBackend:
    """Wrapper to support multiple translation backends dynamically."""
    def __init__(self, config, target_lang="th", source_lang="en"):
        self.config = config
        self.target_lang = resolve_language_code(target_lang)
        self.source_lang = resolve_language_code(source_lang)
        self.backend_type = config.translation_backend
        
        # Initialize Rate Limiter
        rl_cfg = config.rate_limiting
        self.rate_limiter = RateLimiter(
            requests_per_minute=rl_cfg.get("max_requests_per_minute", 60),
            backoff_factor=rl_cfg.get("retry_backoff_factor", 2.0),
            max_retries=rl_cfg.get("max_retries", 5)
        )
        self.gemini_config = config.gemini

        # Initialize GoogleTranslator fallback or main engine
        self.google_translator = None
        try:
            from deep_translator import GoogleTranslator
            self.google_translator = GoogleTranslator(source=self.source_lang, target=self.target_lang)
            logger.info(f"Initialized GoogleTranslator fallback from {self.source_lang} to {self.target_lang}")
        except ImportError:
            logger.warning("deep_translator not installed. GoogleTranslator backend is disabled.")

    def should_skip_translation(self, text):
        """Checks if text contains target language characters or matches protected patterns."""
        if not text.strip():
            return True

        # Check script pattern for target language to prevent re-translation
        lang_base = self.target_lang.split('-')[0]
        script_pattern = SCRIPT_PATTERNS.get(lang_base) or SCRIPT_PATTERNS.get(self.target_lang)
        if script_pattern and script_pattern.search(text):
            return True

        # Check protected patterns
        for pattern in PROTECTED_PATTERNS:
            if re.match(pattern, text.strip()):
                return True
        return False

    def translate(self, text, context=None):
        """Translates text to the target language with retries and rate limiting."""
        if self.should_skip_translation(text):
            return text

        retries = 0
        delay = self.rate_limiter.delay
        
        while retries <= self.rate_limiter.max_retries:
            self.rate_limiter.wait()
            try:
                if self.backend_type == "gemini" and self.gemini_config.get("api_key"):
                    res = self._translate_gemini(text, context)
                else:
                    res = self._translate_google(text)
                return res if res is not None else text
            except Exception as e:
                retries += 1
                if retries > self.rate_limiter.max_retries:
                    logger.error(f"Translation failed after {retries} retries: {e}")
                    raise e
                sleep_time = delay * (self.rate_limiter.backoff_factor ** retries)
                logger.warning(f"Error during translation: {e}. Retrying in {sleep_time:.2f} seconds...")
                time.sleep(sleep_time)
        return text

    def _translate_google(self, text):
        if not self.google_translator:
            raise RuntimeError("GoogleTranslator backend is not available. Please install deep-translator.")
        return self.google_translator.translate(text)

    def _translate_gemini(self, text, context):
        import requests
        api_key = self.gemini_config.get("api_key")
        model = self.gemini_config.get("model", "gemini-1.5-flash")
        endpoint = self.gemini_config.get("endpoint", "")
        url = endpoint.format(model=model) + f"?key={api_key}"
        
        system_instruction = (
            f"You are a professional technical localization specialist. Translate the English text to natural, professional target language code: {self.target_lang}. "
            "Preserve any placeholders (like {username}, {0}, %s), variables, product codes, or brand names exactly as they are. "
            "If XML tags like <r0>, <r1> are present, translate only the text inside the tags, preserving the tag structure, tag names, and tag order exactly. "
            "Do not add any explanations, notes, or introductory text. Return only the translated text."
        )
        
        prompt = f"English text: {text}\n"
        if context:
            prompt += f"Context (Table Header/Row context): {context}\n"
        prompt += f"Translation to target language {self.target_lang}:"
        
        payload = {
            "contents": [{
                "parts": [{"text": prompt}]
            }],
            "systemInstruction": {
                "parts": [{"text": system_instruction}]
            },
            "generationConfig": {
                "temperature": 0.1
            }
        }
        
        headers = {"Content-Type": "application/json"}
        response = requests.post(url, headers=headers, json=payload, timeout=30)
        response.raise_for_status()
        
        result_json = response.json()
        translated_text = result_json['candidates'][0]['content']['parts'][0]['text'].strip()
        return translated_text
