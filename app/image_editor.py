import base64
import io
from typing import Optional
from urllib.request import urlopen

from openai import OpenAI
from PIL import Image


class ImageEditor:
    """Uses OpenAI image API to apply prompt-driven edits to an image."""

    def __init__(self, api_key: str):
        if not api_key:
            raise ValueError("OpenAI API key is required")
        self.client = OpenAI(api_key=api_key)
        self.model = "gpt-image-1"

    def _build_strict_edit_prompt(self, user_prompt: str) -> str:
        """Wrap user prompt with hard constraints to reduce non-target edits."""
        return (
            "Task: Edit this document or image with surgical precision.\n\n"
            "User requested changes:\n"
            f"{user_prompt.strip()}\n\n"
            "Hard constraints:\n"
            "1. Change only the explicitly requested fields or regions in the user request.\n"
            "2. Do not modify any other text, numbers, punctuation, spacing, alignment, font, color, border, logo, stamp, signature, watermark, background, or layout.\n"
            "3. Preserve all non-target content exactly, character-by-character and pixel-by-pixel as much as possible.\n"
            "4. Keep the same page structure, proportions, and visual design.\n"
            "5. Do not rewrite, reformat, beautify, or regenerate the entire page/image.\n"
            "6. If a requested target is unclear or missing, leave it unchanged instead of guessing.\n"
            "7. Return only the edited result.\n"
        )

    def edit_image(self, image: Image.Image, prompt: str, size: str = "1024x1024") -> bytes:
        """Return edited image as PNG bytes."""
        if not prompt or not prompt.strip():
            raise ValueError("Edit prompt is required")

        strict_prompt = self._build_strict_edit_prompt(prompt)

        image_buffer = io.BytesIO()
        image.convert("RGBA").save(image_buffer, format="PNG")
        image_buffer.name = "input.png"
        image_buffer.seek(0)

        response = self.client.images.edit(
            model=self.model,
            image=image_buffer,
            prompt=strict_prompt,
            size=size,
        )

        if not getattr(response, "data", None):
            raise ValueError("No image returned from OpenAI image edit API")

        first_item = response.data[0]
        b64_data: Optional[str] = getattr(first_item, "b64_json", None)
        image_url: Optional[str] = getattr(first_item, "url", None)

        if b64_data:
            return base64.b64decode(b64_data)

        if image_url:
            with urlopen(image_url) as remote_file:  # nosec - URL returned by trusted API response
                return remote_file.read()

        raise ValueError("Image edit response did not contain image data")
