import base64
import io
import json
import re
from typing import Any, Dict, List

import pandas as pd
from openai import OpenAI
from PIL import Image


class LLMExtractor:
    """Uses OpenAI Vision to extract table data from images."""

    def __init__(self, api_key: str):
        if not api_key:
            raise ValueError("OpenAI API key is required")
        self.client = OpenAI(api_key=api_key)
        self.model = "gpt-4.1-mini"

    def extract_tables_from_image(self, image: Image.Image) -> List[Dict[str, Any]]:
        image_data = self._image_to_base64(image)

        prompt = """Analyze this image and extract ALL tables you can find.
For each table found:
1. Extract the complete table data (headers and rows)
2. Return the data as a structured JSON format

Return your response ONLY as valid JSON (no markdown, no code blocks) with this structure:
{
  "tables": [
    {
      "title": "Table title or description if visible",
      "headers": ["column1", "column2", ...],
      "rows": [
        ["value1", "value2", ...],
        ["value1", "value2", ...]
      ],
      "description": "Any additional context about the table"
    }
  ],
  "table_count": number,
  "extraction_notes": "Any issues or notes about extraction"
}

If no tables are found, return:
{
  "tables": [],
  "table_count": 0,
  "extraction_notes": "No tables detected in image"
}

Important:
- Extract ALL text exactly as it appears in the table
- Preserve all numerical values precisely
- Include all headers and columns
- Make sure the structure is valid JSON"""

        response = self.client.responses.create(
            model=self.model,
            max_output_tokens=4000,
            input=[
                {
                    "role": "user",
                    "content": [
                        {"type": "input_text", "text": prompt},
                        {
                            "type": "input_image",
                            "image_url": f"data:image/png;base64,{image_data}",
                        },
                    ],
                }
            ],
        )

        response_text = getattr(response, "output_text", "") or ""
        if not response_text:
            response_text = json.dumps(
                {"tables": [], "table_count": 0, "extraction_notes": "No response text"}
            )

        return self._parse_response(response_text)

    def _image_to_base64(self, image: Image.Image) -> str:
        image = image.convert("RGB")
        buffer = io.BytesIO()
        image.save(buffer, format="PNG")
        image_data = buffer.getvalue()
        return base64.b64encode(image_data).decode("utf-8")

    def _parse_response(self, response_text: str) -> List[Dict[str, Any]]:
        try:
            response_text = response_text.strip()

            if response_text.startswith("```"):
                response_text = re.sub(r"^```(?:json)?", "", response_text).strip()
                response_text = re.sub(r"```$", "", response_text).strip()

            start = response_text.find("{")
            end = response_text.rfind("}")
            if start != -1 and end != -1 and end >= start:
                response_text = response_text[start : end + 1]

            data = json.loads(response_text)

            tables: List[Dict[str, Any]] = []
            for table_data in data.get("tables", []):
                headers = table_data.get("headers") or []
                rows = table_data.get("rows") or []
                if not headers and not rows:
                    continue

                if not headers and rows:
                    max_cols = max(len(row) for row in rows)
                    headers = [f"Column {idx + 1}" for idx in range(max_cols)]

                normalized_rows = []
                for row in rows:
                    row_values = list(row)
                    if len(row_values) < len(headers):
                        row_values.extend([""] * (len(headers) - len(row_values)))
                    normalized_rows.append(row_values[: len(headers)])

                df = pd.DataFrame(data=normalized_rows, columns=headers)

                tables.append(
                    {
                        "title": table_data.get("title", "Extracted Table"),
                        "data": df,
                        "description": table_data.get("description", ""),
                        "original": table_data,
                    }
                )

            return tables

        except json.JSONDecodeError as error:
            raise ValueError(
                f"Failed to parse table data: {str(error)}\\nResponse: {response_text}"
            )
