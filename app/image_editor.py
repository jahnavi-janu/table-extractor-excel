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

    def edit_image(self, image: Image.Image, prompt: str, size: str = "1024x1024") -> bytes:
        """Return edited image as PNG bytes."""
        if not prompt or not prompt.strip():
            raise ValueError("Edit prompt is required")

        image_buffer = io.BytesIO()
        image.convert("RGBA").save(image_buffer, format="PNG")
        image_buffer.name = "input.png"
        image_buffer.seek(0)

        response = self.client.images.edit(
            model=self.model,
            image=image_buffer,
            prompt=prompt.strip(),
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
