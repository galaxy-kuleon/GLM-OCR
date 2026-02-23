"""
title: PDF Translation (GLM-OCR)
description: Converts PDF to DOCX and translates using GLM-OCR pipeline (run.sh).
author: noel
version: 0.1.0
licence: MIT
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import shutil
import time
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import aiohttp
from pydantic import BaseModel, Field


class Pipe:
    """PDF Translation pipe for OpenWebUI using GLM-OCR.

    Receives a PDF upload, extracts the target language from the user prompt,
    runs `bash run.sh "<target_lang>" input/<file>.pdf` inside the GLM-OCR
    project, streams output in real-time, and returns download links for
    the converted DOCX and translated DOCX.
    """

    class Valves(BaseModel):
        PROJECT_DIR: str = Field(
            default="/Users/admin/Works/GLM-OCR/",
            description="Path to the GLM-OCR project directory",
        )
        DEBUG_MODE: bool = Field(
            default=True,
            description="Enable debug logging",
        )
        DEBUG_LOG_FILE: str = Field(
            default="/tmp/pdf_to_docx_v7_pipe.log",
            description="Path to debug log file",
        )
        LANG_DETECT_API_BASE: str = Field(
            default="http://localhost:1234/v1",
            description="OpenAI-compatible API base URL for language extraction",
        )
        LANG_DETECT_MODEL: str = Field(
            default="qwen/qwen3-4b-2507",
            description="Model for extracting target language from user prompt",
        )

    def __init__(self):
        self.valves = self.Valves()

    # -----------------------------------------------------------------
    # Debug logging
    # -----------------------------------------------------------------

    def _debug_log(self, title: str, data: Any) -> None:
        if not self.valves.DEBUG_MODE:
            return

        payload = {
            "ts": datetime.now().isoformat(),
            "title": title,
            "data": data,
        }

        try:
            text = json.dumps(payload, indent=2, ensure_ascii=False, default=str)
        except Exception:
            text = str(payload)

        try:
            with open(self.valves.DEBUG_LOG_FILE, "a", encoding="utf-8") as f:
                f.write(text + "\n")
        except Exception:
            pass

    # -----------------------------------------------------------------
    # Internal task handling
    # -----------------------------------------------------------------

    def _handle_internal_task(self, task: str, body: dict) -> str:
        if task == "title_generation":
            return "PDF Translation (GLM-OCR)"
        if task == "query_generation":
            return ""
        if task == "tags_generation":
            return ""
        if task == "emoji_generation":
            return ""
        return ""

    # -----------------------------------------------------------------
    # Extract user message from body
    # -----------------------------------------------------------------

    def _extract_user_message(self, body: dict) -> str:
        """Extract the last user message from the request body."""
        messages = body.get("messages", [])
        for msg in reversed(messages):
            if msg.get("role") == "user":
                content = msg.get("content", "")
                if isinstance(content, str):
                    return content
                if isinstance(content, list):
                    text_parts = []
                    for part in content:
                        if isinstance(part, dict) and part.get("type") == "text":
                            text_parts.append(part.get("text", ""))
                        elif isinstance(part, str):
                            text_parts.append(part)
                    return " ".join(text_parts)
        return ""

    # -----------------------------------------------------------------
    # Find PDF in __files__
    # -----------------------------------------------------------------

    def _find_pdf_file(self, files: List[dict]) -> Optional[dict]:
        if not files:
            return None
        for f in files:
            name = f.get("name", "") or ""
            file_type = f.get("type", "") or ""
            if name.lower().endswith(".pdf") or "pdf" in file_type.lower():
                return f
        return None

    # -----------------------------------------------------------------
    # Register output file into OpenWebUI
    # -----------------------------------------------------------------

    def _register_output_file(
        self, docx_path: str, display_name: str, user_id: str
    ) -> Optional[Any]:
        from open_webui.models.files import FileForm, Files
        from open_webui.storage.provider import Storage

        file_id = str(uuid.uuid4())
        storage_filename = f"{file_id}_{display_name}"

        with open(docx_path, "rb") as f:
            contents, file_path = Storage.upload_file(
                f,
                storage_filename,
                {
                    "OpenWebUI-User-Id": user_id,
                    "OpenWebUI-File-Id": file_id,
                },
            )

        file_item = Files.insert_new_file(
            user_id,
            FileForm(
                id=file_id,
                filename=display_name,
                path=file_path,
                data={},
                meta={
                    "name": display_name,
                    "content_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "size": len(contents),
                },
            ),
        )

        self._debug_log(
            "FILE_REGISTERED",
            {
                "file_id": file_id,
                "display_name": display_name,
                "file_path": file_path,
                "size": len(contents),
            },
        )

        return file_item

    # -----------------------------------------------------------------
    # Extract target language from user prompt via LM Studio
    # -----------------------------------------------------------------

    async def _extract_target_language(self, user_message: str, emit: Callable) -> str:
        """
        Use a small model (LM Studio) to extract the target language description
        from the user's message.

        Returns:
            A string describing the target language, e.g. "台灣風格的繁體中文".
            Falls back to "English" if extraction fails.
        """
        if not user_message.strip():
            return "English"

        system_prompt = """Analyze the user's message and extract the target language they want the PDF translated into.
Respond ONLY with valid JSON in this exact format:
{"target_lang": "the target language description"}

Examples:
- User says "翻譯成台灣風格的繁體中文" → {"target_lang": "台灣風格的繁體中文"}
- User says "translate to Japanese" → {"target_lang": "Japanese"}
- User says "translate this to simplified Chinese" → {"target_lang": "Simplified Chinese"}
- User says "轉成英文" → {"target_lang": "English"}
- User says "convert to Korean" → {"target_lang": "Korean"}

If the user doesn't specify a language, default to "English".
Do NOT output any other text, ONLY the JSON object."""

        try:
            await emit(
                {
                    "type": "status",
                    "data": {
                        "description": "Extracting target language...",
                        "done": False,
                    },
                }
            )

            async with aiohttp.ClientSession() as session:
                payload = {
                    "model": self.valves.LANG_DETECT_MODEL,
                    "messages": [
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": f"User message: {user_message}"},
                    ],
                    "temperature": 0.1,
                    "max_tokens": 100,
                }

                api_url = f"{self.valves.LANG_DETECT_API_BASE}/chat/completions"
                self._debug_log(
                    "LANG_EXTRACT_REQUEST", {"url": api_url, "payload": payload}
                )

                async with session.post(
                    api_url,
                    json=payload,
                    timeout=aiohttp.ClientTimeout(total=30),
                ) as response:
                    if response.status != 200:
                        error_text = await response.text()
                        self._debug_log(
                            "LANG_EXTRACT_API_ERROR",
                            {
                                "status": response.status,
                                "error": error_text,
                            },
                        )
                        raise Exception(f"API error: {response.status}")

                    result = await response.json()
                    content = result["choices"][0]["message"]["content"]
                    self._debug_log("LANG_EXTRACT_RESPONSE", {"content": content})

                    json_match = re.search(r"\{[^}]+\}", content)
                    if json_match:
                        parsed = json.loads(json_match.group())
                    else:
                        parsed = json.loads(content)

                    target_lang = parsed.get("target_lang", "English")

                    self._debug_log(
                        "LANG_EXTRACT_PARSED",
                        {
                            "target_lang": target_lang,
                        },
                    )

                    await emit(
                        {
                            "type": "status",
                            "data": {
                                "description": f"Target language: {target_lang}",
                                "done": False,
                            },
                        }
                    )

                    return target_lang

        except Exception as e:
            self._debug_log("LANG_EXTRACT_FAILED", {"error": str(e)})
            await emit(
                {
                    "type": "status",
                    "data": {
                        "description": "Language extraction failed, defaulting to English",
                        "done": False,
                    },
                }
            )
            return "English"

    # -----------------------------------------------------------------
    # Load .env from PROJECT_DIR
    # -----------------------------------------------------------------

    def _load_dotenv(self) -> dict:
        """Read {PROJECT_DIR}/.env and merge with current environment."""
        env = os.environ.copy()
        dotenv_path = os.path.join(self.valves.PROJECT_DIR, ".env")

        if not os.path.isfile(dotenv_path):
            self._debug_log("DOTENV_NOT_FOUND", {"path": dotenv_path})
            return env

        try:
            with open(dotenv_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#"):
                        continue
                    if "=" not in line:
                        continue
                    key, _, value = line.partition("=")
                    key = key.strip()
                    value = value.strip()
                    # Remove surrounding quotes if present
                    if (
                        len(value) >= 2
                        and value[0] == value[-1]
                        and value[0] in ('"', "'")
                    ):
                        value = value[1:-1]
                    env[key] = value

            self._debug_log("DOTENV_LOADED", {"path": dotenv_path})
        except Exception as e:
            self._debug_log("DOTENV_ERROR", {"path": dotenv_path, "error": str(e)})

        return env

    # -----------------------------------------------------------------
    # Find workspace directory in output/
    # -----------------------------------------------------------------

    def _find_workspace_dir(self, stem: str, start_time: float) -> Optional[str]:
        """
        Locate the GLM-OCR workspace directory under output/.

        Strategy:
        1. Try the expected path: output/{stem}-docx-workspace/
        2. Fallback: scan output/ for any *-docx-workspace directory
           created after start_time, preferring ones whose name contains
           the stem.
        """
        output_root = os.path.join(self.valves.PROJECT_DIR, "output")

        # --- Attempt 1: exact expected path ---
        expected = os.path.join(output_root, f"{stem}-docx-workspace")
        if os.path.isdir(expected):
            self._debug_log("WORKSPACE_FOUND_EXACT", {"path": expected})
            return expected

        # --- Attempt 2: scan for recently created workspace dirs ---
        if not os.path.isdir(output_root):
            self._debug_log("WORKSPACE_OUTPUT_ROOT_MISSING", {"path": output_root})
            return None

        candidates: list[Tuple[str, float]] = []
        try:
            for entry in os.listdir(output_root):
                if not entry.endswith("-docx-workspace"):
                    continue
                full = os.path.join(output_root, entry)
                if not os.path.isdir(full):
                    continue
                mtime = os.path.getmtime(full)
                if mtime >= start_time:
                    candidates.append((full, mtime))
        except Exception as e:
            self._debug_log("WORKSPACE_SCAN_ERROR", {"error": str(e)})
            return None

        if not candidates:
            self._debug_log("WORKSPACE_NO_CANDIDATES", {"output_root": output_root})
            return None

        # Prefer directory whose name contains the stem
        for path, _ in candidates:
            if stem in os.path.basename(path):
                self._debug_log("WORKSPACE_FOUND_STEM_MATCH", {"path": path})
                return path

        # Otherwise pick the most recently modified
        candidates.sort(key=lambda x: x[1], reverse=True)
        chosen = candidates[0][0]
        self._debug_log("WORKSPACE_FOUND_RECENT", {"path": chosen})
        return chosen

    # -----------------------------------------------------------------
    # Run GLM-OCR pipeline (bash run.sh)
    # -----------------------------------------------------------------

    async def _run_glm_ocr(
        self,
        target_lang: str,
        pdf_input_relative: str,
        emit: Callable,
    ) -> Tuple[int, str, str]:
        """
        Execute `bash run.sh "<target_lang>" <pdf_input_relative>` in the
        GLM-OCR project directory.

        Returns:
            (exit_code, stdout_text, stderr_text)
        """
        cmd = ["bash", "run.sh", target_lang, pdf_input_relative]

        env = self._load_dotenv()
        # Ensure we use GLM-OCR's virtualenv, not OpenWebUI's
        env["VIRTUAL_ENV"] = os.path.join(self.valves.PROJECT_DIR, ".venv")
        # Update PATH to prioritize GLM-OCR's virtualenv
        venv_bin = os.path.join(env["VIRTUAL_ENV"], "bin")
        env["PATH"] = venv_bin + ":" + env.get("PATH", "")

        self._debug_log(
            "GLM_OCR_COMMAND",
            {
                "cmd": cmd,
                "cwd": self.valves.PROJECT_DIR,
            },
        )

        await emit(
            {
                "type": "status",
                "data": {"description": "Starting GLM-OCR pipeline...", "done": False},
            }
        )

        try:
            process = await asyncio.create_subprocess_exec(
                *cmd,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
                cwd=self.valves.PROJECT_DIR,
                env=env,
            )

            stdout_lines: list[str] = []
            stderr_lines: list[str] = []

            async def read_stream(
                stream: asyncio.StreamReader,
                collect: list[str],
                is_stderr: bool = False,
            ):
                while True:
                    line_bytes = await stream.readline()
                    if not line_bytes:
                        break
                    line = line_bytes.decode("utf-8", errors="replace").rstrip()
                    collect.append(line)
                    self._debug_log(
                        "GLM_OCR_STREAM", {"stderr" if is_stderr else "stdout": line}
                    )
                    if line:
                        prefix = "[stderr] " if is_stderr else ""
                        await emit(
                            {
                                "type": "status",
                                "data": {
                                    "description": f"{prefix}{line[:80]}",
                                    "done": False,
                                },
                            }
                        )

            await asyncio.gather(
                read_stream(process.stdout, stdout_lines, False),
                read_stream(process.stderr, stderr_lines, True),
            )

            exit_code = await process.wait()

            self._debug_log(
                "GLM_OCR_COMPLETED",
                {
                    "exit_code": exit_code,
                    "stdout_lines": len(stdout_lines),
                    "stderr_lines": len(stderr_lines),
                },
            )

            return exit_code, "\n".join(stdout_lines), "\n".join(stderr_lines)

        except FileNotFoundError:
            error_msg = f"bash or run.sh not found in {self.valves.PROJECT_DIR}"
            self._debug_log("GLM_OCR_NOT_FOUND", {"error": error_msg})
            await emit(
                {
                    "type": "status",
                    "data": {"description": error_msg, "done": True},
                }
            )
            return 1, "", error_msg
        except Exception as e:
            self._debug_log("GLM_OCR_ERROR", {"error": str(e)})
            await emit(
                {
                    "type": "status",
                    "data": {"description": f"GLM-OCR error: {e}", "done": True},
                }
            )
            return 1, "", str(e)

    # -----------------------------------------------------------------
    # Main pipe entry point
    # -----------------------------------------------------------------

    async def pipe(
        self,
        body: dict,
        __user__: Optional[dict] = None,
        __request__: Optional[Any] = None,
        __event_emitter__: Optional[Callable] = None,
        __task__: Optional[str] = None,
        __files__: Optional[List[dict]] = None,
        __metadata__: Optional[dict] = None,
    ) -> str:

        # No-op emitter fallback
        async def noop_emitter(event: dict) -> None:
            pass

        emit = __event_emitter__ or noop_emitter

        self._debug_log(
            "PIPE_CALLED",
            {
                "task": __task__,
                "files": __files__,
                "user": __user__,
                "body_keys": list(body.keys()) if isinstance(body, dict) else None,
            },
        )

        # ----------------------------------------------------------
        # 1. Short-circuit internal tasks
        # ----------------------------------------------------------
        if __task__:
            return self._handle_internal_task(__task__, body)

        # ----------------------------------------------------------
        # 2. Find PDF in uploaded files
        # ----------------------------------------------------------
        pdf_file = self._find_pdf_file(__files__ or [])
        if not pdf_file:
            await emit(
                {
                    "type": "status",
                    "data": {"description": "No PDF file found", "done": True},
                }
            )
            return (
                "Please upload a `.pdf` file to translate. "
                "This pipe converts PDF documents to DOCX and translates them using GLM-OCR."
            )

        # ----------------------------------------------------------
        # 3. Resolve disk path via Files.get_file_by_id()
        # ----------------------------------------------------------
        from open_webui.models.files import Files
        from open_webui.storage.provider import Storage

        file_id = pdf_file.get("id")
        file_record = Files.get_file_by_id(file_id)
        if not file_record or not file_record.path:
            await emit(
                {
                    "type": "status",
                    "data": {
                        "description": "Failed to resolve uploaded file path",
                        "done": True,
                    },
                }
            )
            return "Error: Could not resolve the uploaded PDF file path."

        source_path = Storage.get_file(file_record.path)
        original_filename = (
            pdf_file.get("name") or file_record.filename or "document.pdf"
        )

        self._debug_log(
            "FILE_RESOLVED",
            {
                "file_id": file_id,
                "source_path": source_path,
                "original_filename": original_filename,
            },
        )

        if not os.path.isfile(source_path):
            await emit(
                {
                    "type": "status",
                    "data": {
                        "description": "Uploaded file not found on disk",
                        "done": True,
                    },
                }
            )
            return f"Error: File not found at `{source_path}`."

        # ----------------------------------------------------------
        # 4. Extract target language from user message
        # ----------------------------------------------------------
        user_message = self._extract_user_message(body)
        target_lang = await self._extract_target_language(user_message, emit)

        self._debug_log(
            "TARGET_LANGUAGE",
            {
                "user_message": user_message[:200] if user_message else "",
                "target_lang": target_lang,
            },
        )

        # ----------------------------------------------------------
        # 5. Copy PDF to {PROJECT_DIR}/input/
        # ----------------------------------------------------------
        input_dir = os.path.join(self.valves.PROJECT_DIR, "input")
        os.makedirs(input_dir, exist_ok=True)

        epoch = int(time.time())
        safe_filename = f"{epoch}_{original_filename}"
        dest_pdf_path = os.path.join(input_dir, safe_filename)

        try:
            shutil.copy2(source_path, dest_pdf_path)
            self._debug_log("FILE_COPIED", {"from": source_path, "to": dest_pdf_path})
        except Exception as e:
            self._debug_log("FILE_COPY_FAILED", {"error": str(e)})
            await emit(
                {
                    "type": "status",
                    "data": {"description": f"Failed to copy file: {e}", "done": True},
                }
            )
            return f"Error: Failed to copy PDF to GLM-OCR input directory: {e}"

        await emit(
            {
                "type": "status",
                "data": {
                    "description": f"Processing: {original_filename}",
                    "done": False,
                },
            }
        )

        # ----------------------------------------------------------
        # 6. Run bash run.sh "<target_lang>" input/<filename>
        # ----------------------------------------------------------
        pdf_input_relative = f"input/{safe_filename}"
        start_time = time.time()

        exit_code, stdout, stderr = await self._run_glm_ocr(
            target_lang, pdf_input_relative, emit
        )

        # ----------------------------------------------------------
        # 7. Locate output files
        # ----------------------------------------------------------
        stem = Path(safe_filename).stem  # e.g. "1706000000_document"
        workspace_dir = self._find_workspace_dir(stem, start_time)

        if workspace_dir:
            final_docx_path = os.path.join(workspace_dir, "final-output.docx")
            translated_docx_path = os.path.join(
                workspace_dir, "translation", "translated-output.docx"
            )
        else:
            final_docx_path = ""
            translated_docx_path = ""

        self._debug_log(
            "LOOKING_FOR_OUTPUTS",
            {
                "workspace_dir": workspace_dir,
                "final_docx_path": final_docx_path,
                "translated_docx_path": translated_docx_path,
            },
        )

        final_exists = final_docx_path and os.path.isfile(final_docx_path)
        translated_exists = translated_docx_path and os.path.isfile(
            translated_docx_path
        )

        if not final_exists and not translated_exists:
            # Non-zero exit and no outputs at all
            if exit_code != 0:
                error_msg = stderr or stdout or "Unknown error"
                await emit(
                    {
                        "type": "status",
                        "data": {
                            "description": "GLM-OCR pipeline failed",
                            "done": True,
                        },
                    }
                )
                return f"**GLM-OCR Pipeline Failed** (exit code {exit_code})\n\n```\n{error_msg}\n```"

            await emit(
                {
                    "type": "status",
                    "data": {"description": "Output files not found", "done": True},
                }
            )
            expected_dir = os.path.join(
                self.valves.PROJECT_DIR, "output", f"{stem}-docx-workspace"
            )
            return (
                f"**Error**: Pipeline completed but output files not found.\n\n"
                f"Workspace searched: `{workspace_dir or expected_dir}`\n"
                f"Expected:\n"
                f"- `final-output.docx`\n"
                f"- `translation/translated-output.docx`"
            )

        # ----------------------------------------------------------
        # 8. Register output files in OpenWebUI
        # ----------------------------------------------------------
        user_id = (__user__ or {}).get("id", "")
        original_stem = Path(original_filename).stem
        download_links = []

        if final_exists:
            final_display = f"{original_stem}_converted.docx"
            final_item = self._register_output_file(
                final_docx_path, final_display, user_id
            )
            if final_item:
                url = f"/api/v1/files/{final_item.id}/content"
                download_links.append(
                    f"- [Download {final_display}]({url}) (converted DOCX)"
                )
                self._debug_log("FINAL_DOCX_REGISTERED", {"file_id": final_item.id})

        if translated_exists:
            translated_display = f"{original_stem}_translated.docx"
            translated_item = self._register_output_file(
                translated_docx_path, translated_display, user_id
            )
            if translated_item:
                url = f"/api/v1/files/{translated_item.id}/content"
                download_links.append(
                    f"- [Download {translated_display}]({url}) (translated DOCX)"
                )
                self._debug_log(
                    "TRANSLATED_DOCX_REGISTERED", {"file_id": translated_item.id}
                )

        if not download_links:
            await emit(
                {
                    "type": "status",
                    "data": {
                        "description": "Failed to register output files",
                        "done": True,
                    },
                }
            )
            return "Error: Pipeline completed but failed to register output files."

        # ----------------------------------------------------------
        # 9. Return download links
        # ----------------------------------------------------------
        await emit(
            {
                "type": "status",
                "data": {"description": "Translation complete!", "done": True},
            }
        )

        links_text = "\n".join(download_links)
        result_message = (
            f"**PDF Translation Complete!**\n\n"
            f"Target language: **{target_lang}**\n\n"
            f"{links_text}\n\n"
            f"---\n"
            f"*Source: `{original_filename}`*"
        )

        self._debug_log(
            "SUCCESS",
            {
                "download_links": download_links,
                "target_lang": target_lang,
            },
        )

        # ----------------------------------------------------------
        # 10. Emit "replace" to persist message
        # ----------------------------------------------------------
        await emit(
            {
                "type": "replace",
                "data": {"content": result_message},
            }
        )

        return result_message
