"""Translation service orchestrating providers, caching and chunking."""
from __future__ import annotations

import re
import json
import threading
import time
from pathlib import Path
from typing import Dict, List, Optional

from .providers.base import TranslationProvider

_SENTENCE_SPLIT_PATTERN = re.compile(r"(?<=[.!?。！？])\s+")


class TranslationService:
    """Translate text using a configured provider with caching support."""

    def __init__(self, provider: TranslationProvider, *, max_chunk_size: int = 1000, cache_file: Optional[Path] = None) -> None:
        self.provider = provider
        self.max_chunk_size = max_chunk_size
        self._cache: Dict[str, str] = {}
        self._lock = threading.Lock()
        
        # Persistent cache setup
        self.cache_file = cache_file
        if self.cache_file is None:
             # Default to a file in the same directory as this script or a user-local dir
             # For this skill, we'll put it in the parent of the script dir to avoid cluttering src
             self.cache_file = Path(__file__).parent.parent / "translation_cache.json"
        
        self._load_cache()

    def _load_cache(self) -> None:
        """Load cache from disk."""
        if self.cache_file and self.cache_file.exists():
            try:
                with open(self.cache_file, "r", encoding="utf-8") as f:
                    self._cache = json.load(f)
            except Exception as e:
                print(f"Warning: Failed to load cache file: {e}")
                self._cache = {}

    def _save_cache(self) -> None:
        """Save cache to disk."""
        if self.cache_file:
            try:
                with self._lock: # Ensure thread safety during write
                    temp_file = self.cache_file.with_suffix(".tmp")
                    with open(temp_file, "w", encoding="utf-8") as f:
                        json.dump(self._cache, f, ensure_ascii=False, indent=2)
                    temp_file.replace(self.cache_file)
            except Exception as e:
                print(f"Warning: Failed to save cache file: {e}")

    def translate(self, text: str, source_lang: str, target_lang: str, is_tagged: bool = False) -> str:
        """Translate ``text`` and cache repeated requests.
        
        Args:
            text: Text to translate.
            source_lang: Source language code.
            target_lang: Target language code.
            is_tagged: Whether the text contains formatting tags (e.g. <r0>...</r0>).
                       If True, sends specific instructions to the provider.
        """
        if not text or text.isspace():
            return text

        # Cache key includes languages to avoid collisions
        cache_key = f"{source_lang}:{target_lang}:{text}"

        with self._lock:
            if cache_key in self._cache:
                return self._cache[cache_key]

        # For tagged text, we generally don't chunk because splitting tags is dangerous.
        # We rely on the provider's context window (which is large for modern models).
        if is_tagged:
            translated = self._translate_with_retry(text, source_lang, target_lang, is_tagged=True)
        else:
            chunks = self.chunk_text(text, self.max_chunk_size)
            translated_chunks: List[str] = []
            for chunk in chunks:
                stripped = chunk.strip()
                if not stripped:
                    translated_chunks.append(chunk)
                    continue
                
                # Check cache for individual chunks too
                chunk_key = f"{source_lang}:{target_lang}:{chunk}"
                with self._lock:
                    cached_chunk = self._cache.get(chunk_key)
                
                if cached_chunk:
                    translated_chunks.append(cached_chunk)
                else:
                    t_chunk = self._translate_with_retry(chunk, source_lang, target_lang, is_tagged=False)
                    translated_chunks.append(t_chunk.strip())
                    # Cache the chunk
                    with self._lock:
                        self._cache[chunk_key] = t_chunk.strip()

            translated = " ".join(part for part in translated_chunks if part)
            if not translated:
                translated = text

        with self._lock:
            self._cache[cache_key] = translated
        
        # Save cache after update
        self._save_cache()
            
        return translated

    def _translate_with_retry(self, text: str, source_lang: str, target_lang: str, is_tagged: bool = False, retries: int = 3) -> str:
        """Execute translation with retry logic."""
        attempt = 0
        while attempt < retries:
            try:
                # If tagged, prepend instructions to the text or rely on system prompt in provider.
                # Since we can't easily change the provider's system prompt per call here without changing the interface,
                # we'll prepend a user instruction for tagged text.
                input_text = text
                if is_tagged:
                    input_text = (
                        f"Translate the following text from {source_lang} to {target_lang}. "
                        "The text contains XML-like tags (e.g., <r0>...</r0>) marking formatting. "
                        "**RULES:**\n"
                        "1. Translate ONLY the content inside the tags.\n"
                        "2. PRESERVE all tags exactly as they are.\n"
                        "3. DO NOT change the order of the tags.\n"
                        "4. DO NOT translate the tag names.\n\n"
                        f"Text to translate:\n{text}"
                    )
                
                return self.provider.translate(input_text, source_lang, target_lang)
            except Exception as e:
                attempt += 1
                print(f"Translation failed (attempt {attempt}/{retries}): {e}")
                if attempt >= retries:
                    print(f"Giving up on text: {text[:50]}...")
                    # Return original text on failure to avoid data loss, or empty?
                    # Returning original allows the process to continue.
                    return text 
                time.sleep(1 * attempt) # Exponential backoff
        return text

    def translate_batch_json(self, texts: List[str], source_lang: str, target_lang: str) -> List[str]:
        """Translate a batch of texts using ID-based JSON objects for robust mapping."""
        if not texts:
            return []
            
        # Filter and prepare objects with IDs
        items_to_translate = []
        for i, t in enumerate(texts):
            if t and not t.isspace():
                items_to_translate.append({"id": i, "text": t})
        
        if not items_to_translate:
            return texts # All were empty
            
        print(f"[BatchTranslation] Sending {len(items_to_translate)} items. IDs: {[i['id'] for i in items_to_translate]}")
        
        try:
            translated_items = self._translate_batch_with_retry_objects(items_to_translate, source_lang, target_lang)
        except Exception as e:
            print(f"[BatchTranslation] CRITICAL ERROR: Batch request failed completely: {e}")
            return texts # Fallback to original
            
        # Reconstruct result using ID mapping
        result = list(texts) # Copy original
        
        success_count = 0
        failure_count = 0
        
        # Create a lookup map
        response_map = {}
        received_ids = []
        for item in translated_items:
            if "id" in item and "text" in item:
                try:
                    rid = int(item["id"])
                    response_map[rid] = item["text"]
                    received_ids.append(rid)
                except ValueError:
                    print(f"[BatchTranslation] Warn: Invalid ID format in response: {item['id']}")
        
        received_ids.sort()
        print(f"[BatchTranslation] Received {len(received_ids)} items. IDs: {received_ids}")

        first_id = items_to_translate[0]["id"]
        last_id = items_to_translate[-1]["id"]
        
        for item in items_to_translate:
            original_id = item["id"]
            if original_id in response_map:
                trans_text = response_map[original_id]
                
                # Debug log samples (First and Last)
                if original_id == first_id:
                    print(f"[BatchTranslation] Sample [First]: '{item['text'][:30]}...' -> '{trans_text[:30]}...'\n")
                elif original_id == last_id:
                    print(f"[BatchTranslation] Sample [Last]:  '{item['text'][:30]}...' -> '{trans_text[:30]}...'\n")
                
                result[original_id] = trans_text
                success_count += 1
            else:
                print(f"[BatchTranslation] WARNING: Missing ID {original_id}. Orig: '{item['text'][:50]}...' Keeping original.")
                failure_count += 1
                
        print(f"[BatchTranslation] Batch Summary -> Success: {success_count}, Missing: {failure_count}")
        return result

    def _translate_batch_with_retry_objects(self, items: List[Dict], source_lang: str, target_lang: str, retries: int = 3) -> List[Dict]:
        import json
        
        prompt = (
            f"You are a professional translator translating a PowerPoint presentation from {source_lang} to {target_lang}.\n"
            "INPUT: A JSON array of objects, each with 'id' and 'text'.\n"
            "TASK: Translate the 'text' field of each object.\n"
            "CRITICAL RULES:\n"
            "1. Output ONLY a valid JSON array of objects.\n"
            "2. Each object MUST have 'id' (integer, matching input) and 'text' (translated string).\n"
            "3. PRESERVE all <rN>...</rN> tags exactly. The tags mark formatting boundaries.\n"
            "4. TRANSLATE the content inside the tags. DO NOT leave it in {source_lang} unless it is a proper noun.\n"
            "5. Do not output markdown code blocks, just raw JSON.\n\n"
            "EXAMPLE:\n"
            'Input: [{"id": 1, "text": "<r0>Hello</r0> <r1>World</r1>"}]\n'
            f'Output: [{{"id": 1, "text": "<r0>你好</r0> <r1>世界</r1>"}}]\n'
            "(Note: The tags <r0>, <r1> are kept, but content 'Hello', 'World' is translated.)\n\n"
            f"Input JSON to translate:\n{json.dumps(items, ensure_ascii=False)}"
        )

        attempt = 0
        while attempt < retries:
            try:
                response_text = self.provider.translate(prompt, source_lang, target_lang)
                # Cleanup potential markdown formatting
                cleaned_resp = response_text.replace("```json", "").replace("```", "").strip()
                
                result = json.loads(cleaned_resp)
                
                if not isinstance(result, list):
                    print(f"[BatchTranslation] Attempt {attempt+1}: Response is not a list.")
                    raise ValueError("Response is not a JSON list")
                    
                return result
            except Exception as e:
                attempt += 1
                print(f"[BatchTranslation] Attempt {attempt+1}/{retries} failed: {e}")
                if attempt >= retries:
                    raise
                time.sleep(1 * attempt)
        return [] # Should not reach here due to raise

    @staticmethod
    def chunk_text(text: str, max_chunk_size: int = 1000) -> List[str]:
        """Split long text into smaller chunks preserving sentence boundaries."""
        if len(text) <= max_chunk_size:
            return [text]

        sentences = [segment.strip() for segment in _SENTENCE_SPLIT_PATTERN.split(text) if segment.strip()]
        if not sentences:
            sentences = [text]

        chunks: List[str] = []
        current: List[str] = []
        current_len = 0

        for sentence in sentences:
            sentence_len = len(sentence)
            if current and current_len + sentence_len + 1 > max_chunk_size:
                chunks.append(" ".join(current))
                current = []
                current_len = 0
            if sentence_len > max_chunk_size:
                if current:
                    chunks.append(" ".join(current))
                    current = []
                    current_len = 0
                chunks.extend(
                    [sentence[i : i + max_chunk_size] for i in range(0, sentence_len, max_chunk_size)]
                )
                continue
            current.append(sentence)
            current_len += sentence_len + 1

        if current:
            chunks.append(" ".join(current))

        if not chunks:
            return [text]
        return chunks

    def clear_cache(self) -> None:
        """Drop cached translations."""
        with self._lock:
            self._cache.clear()
        self._save_cache()

    def cache_size(self) -> int:
        """Return the number of cached entries."""
        with self._lock:
            return len(self._cache)
