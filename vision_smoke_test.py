#!/usr/bin/env python3
"""Vision-capability smoke test for the configured AI provider (GPT-5.5 via
Workbench, by default) -- confirms two independent things, cheaply, BEFORE
building a real contract-extraction pipeline on top of an unverified
assumption. Same "test the smallest possible thing first" discipline this
project already used for plain text connectivity (a standalone connectivity
script was built and used earlier for that, separate from this one):
  1. Can a PDF page actually be rasterized locally at all (pypdfium2)?
  2. Does the configured model deployment actually ACCEPT and correctly
     read image input? Some enterprise LLM gateways only expose a
     text-only deployment even when the underlying model family supports
     vision -- this is a real, unverified assumption until tested against
     the real endpoint. Nothing in this codebase has ever sent an image to
     the AI provider before this script.

Uses the EXACT SAME AIClient connection setup (config.yml credentials,
endpoint, workbench headers, retry/param-adjustment logic) as the rest of
this pipeline -- not a second, separately-maintained copy of the auth/
connection logic. get_response()'s `content` field is passed through to
the API as-is regardless of whether it's a plain string or a list of
OpenAI-spec multimodal content blocks, so no changes to fdd_utils/ai.py
were needed to support this.

Usage:
    python vision_smoke_test.py "path/to/a/contract.pdf"
        # rasterizes page 1, sends it to the model, asks it to describe
        # what it sees in one sentence and quote one specific piece of
        # text it can actually read -- if the response correctly
        # describes a lease contract (not a refusal / "I cannot see
        # images" / hallucinated unrelated content), vision input works.
    python vision_smoke_test.py "path/to/a/contract.pdf" --page 2
    python vision_smoke_test.py "path/to/an/image.jpg"
        # works directly on a raw image file too, no PDF rasterization needed
    python vision_smoke_test.py "path/to/a/contract.pdf" --model openai
        # test a different configured provider instead of the default workbench
"""
import argparse
import base64
import io
import sys
from pathlib import Path

from fdd_utils.ai import AIClient

try:
    import pypdfium2 as pdfium
except ImportError:
    pdfium = None


def _page_to_png_bytes(pdf_path: str, page_num: int, dpi: int = 200) -> bytes:
    if pdfium is None:
        raise RuntimeError("pypdfium2 not installed -- run `pip install pypdfium2`")
    pdf = pdfium.PdfDocument(pdf_path)
    if page_num < 1 or page_num > len(pdf):
        raise ValueError(f"Page {page_num} out of range -- this PDF has {len(pdf)} page(s).")
    page = pdf[page_num - 1]
    bitmap = page.render(scale=dpi / 72)
    pil_image = bitmap.to_pil()
    buf = io.BytesIO()
    pil_image.save(buf, format="PNG")
    return buf.getvalue()


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file_path", help="path to a .pdf or image file (.jpg/.png/etc.)")
    ap.add_argument("--page", type=int, default=1, help="for a PDF, which page to test (1-indexed, default 1)")
    ap.add_argument("--model", default="workbench", help="model_type to test (default: workbench)")
    args = ap.parse_args()

    path = Path(args.file_path)
    if not path.exists():
        print(f"❌ File not found: {path}")
        return 1

    print(f"Loading {path.name!r}...")
    ext = path.suffix.lower()
    if ext == ".pdf":
        try:
            image_bytes = _page_to_png_bytes(str(path), args.page)
        except Exception as exc:
            print(f"❌ Could not rasterize page {args.page}: {exc}")
            return 1
        print(f"✅ Rasterized page {args.page} to a PNG ({len(image_bytes) / 1024:.0f} KB).")
        mime = "image/png"
    else:
        image_bytes = path.read_bytes()
        print(f"✅ Read image file directly ({len(image_bytes) / 1024:.0f} KB).")
        mime = f"image/{ext.lstrip('.') or 'png'}"

    b64 = base64.b64encode(image_bytes).decode("ascii")
    data_url = f"data:{mime};base64,{b64}"

    print(f"\nConnecting via AIClient(model_type={args.model!r})...")
    try:
        client = AIClient(model_type=args.model, agent_name="subagent_1", language="Eng")
    except Exception as exc:
        print(f"❌ Could not initialize AIClient: {exc}")
        return 1

    vision_content = [
        {"type": "text", "text": (
            "Describe in one sentence what kind of document this image shows "
            "(e.g. a lease contract's cover page, a signature page, a rent "
            "schedule table, etc.) and then quote ONE specific piece of text "
            "you can actually read on it, to prove you are reading the real "
            "image content and not guessing from context."
        )},
        {"type": "image_url", "image_url": {"url": data_url}},
    ]

    print("Sending to the model (this is the untested part -- confirms whether "
          "this deployment accepts image input at all)...\n")
    try:
        result = client.get_response(user_prompt=vision_content, system_prompt="You are a helpful assistant.")
    except Exception as exc:
        print(f"❌ Call failed: {exc}")
        print(
            "\nIf this is an authentication/model-not-found/400-style error, the likely "
            "cause is that this specific deployment/model ID doesn't support image input "
            "at all (text-only), not a bug in this script -- check with whoever "
            "administers the gateway whether the configured model supports vision."
        )
        return 1

    print("=" * 78)
    print("RESPONSE:")
    print("=" * 78)
    print(result.get("content", "(no content field in response)"))
    print("\n" + "=" * 78)
    print(f"Duration: {result.get('duration', '?')}s")
    print(
        "\nIf the response above correctly describes a real lease-contract-looking page "
        "AND quotes something plausibly readable from it, vision input works on this "
        "gateway -- safe to build the full extraction pipeline on top of this. If it "
        "refuses, hallucinates unrelated content, or errors, paste this whole output "
        "back before any further contract-extraction work is built."
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
