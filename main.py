"""
다온 EPUB 변환기 (웹 버전) - 1단계: DOCX to EPUB
FastAPI 기반. python-docx로 원고를 읽고, ebooklib으로 EPUB을 생성합니다.
"""

import io
import re
import uuid
from pathlib import Path

from fastapi import FastAPI, Form, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

import docx
from ebooklib import epub

app = FastAPI(title="다온 EPUB 변환기")

TEMPLATES = {
    "essay": {
        "label": "에세이",
        "css": """
            body { font-family: 'Noto Serif KR', serif; line-height: 1.9; margin: 5% 6%; }
            h1 { font-size: 1.5em; margin-bottom: 1.2em; text-align: center; }
            p { margin: 0 0 1em 0; text-indent: 1em; }
        """,
    },
    "novel": {
        "label": "소설",
        "css": """
            body { font-family: 'Noto Serif KR', serif; line-height: 2.0; margin: 6% 7%; }
            h1 { font-size: 1.6em; margin-bottom: 1.5em; text-align: center; letter-spacing: 0.05em; }
            p { margin: 0 0 0.8em 0; text-indent: 1em; }
        """,
    },
    "practical": {
        "label": "실용/정보",
        "css": """
            body { font-family: 'Noto Sans KR', sans-serif; line-height: 1.7; margin: 5%; }
            h1 { font-size: 1.4em; margin-bottom: 1em; border-bottom: 2px solid #333; padding-bottom: 0.3em; }
            p { margin: 0 0 1em 0; }
        """,
    },
}


def get_paragraph_html(para) -> str:
    """문단 안의 텍스트를 가져오되, 수동 줄바꿈(Shift+Enter)은 <br/>로 보존합니다."""
    from docx.oxml.ns import qn

    parts = []
    for run in para.runs:
        for child in run._element:
            tag = child.tag.split("}")[-1]
            if tag == "t":
                parts.append(escape_html(child.text or ""))
            elif tag == "br":
                parts.append("<br/>")
            elif tag == "tab":
                parts.append("&#9;")
    return "".join(parts)


def build_heading_tree(document: "docx.Document"):
    """
    '제목 1' ~ '제목 6' 스타일을 몇 단계든 자동으로 인식해 트리 구조를 만듭니다.
    'Title' 스타일 문단은 표지 텍스트로 취급해 최상위(root) 문단에 포함시킵니다.
    """
    root = {"title": None, "level": 0, "paras": [], "children": []}
    stack = [root]

    for para in document.paragraphs:
        text = para.text.strip()
        style_name = (para.style.name or "").lower()

        if style_name == "title":
            html = get_paragraph_html(para)
            if html.strip():
                root["paras"].append(html)
            continue

        m = re.match(r"heading (\d+)", style_name)
        level = int(m.group(1)) if m else None

        if level is not None and text:
            while len(stack) > 1 and stack[-1]["level"] >= level:
                stack.pop()
            node = {"title": text, "level": level, "paras": [], "children": []}
            stack[-1]["children"].append(node)
            stack.append(node)
        else:
            html = get_paragraph_html(para)
            if html.strip():
                stack[-1]["paras"].append(html)

    return root


def escape_html(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def build_epub(
    docx_bytes: bytes,
    title: str,
    author: str,
    publisher: str,
    isbn: str,
    language: str,
    template_key: str,
) -> bytes:
    document = docx.Document(io.BytesIO(docx_bytes))
    tree = build_heading_tree(document)
    template = TEMPLATES.get(template_key, TEMPLATES["essay"])

    book = epub.EpubBook()
    book.set_identifier(isbn if isbn else str(uuid.uuid4()))
    book.set_title(title or "제목 없음")
    book.set_language(language or "ko")
    if author:
        book.add_author(author)
    if publisher:
        book.add_metadata("DC", "publisher", publisher)

    style_item = epub.EpubItem(
        uid="style_default",
        file_name="style/default.css",
        media_type="text/css",
        content=template["css"],
    )
    book.add_item(style_item)

    all_chapters_flat = []
    counter = 0

    def make_chapter(chap_title: str, paras: list) -> "epub.EpubHtml":
        nonlocal counter
        counter += 1
        file_name = f"chap_{counter:03d}.xhtml"
        body_html = f"<h1>{escape_html(chap_title)}</h1>\n" if chap_title else ""
        body_html += "\n".join(f"<p>{html}</p>" for html in paras)
        c = epub.EpubHtml(title=chap_title or f"챕터 {counter}", file_name=file_name, lang=language or "ko")
        c.content = f"<html><head></head><body>{body_html}</body></html>"
        c.add_item(style_item)
        book.add_item(c)
        all_chapters_flat.append(c)
        return c

    def process_node(node):
        """노드를 EPUB 챕터/섹션으로 변환. 반환값은 EpubHtml 또는 (Section, tuple) 입니다."""
        chap = make_chapter(node["title"], node["paras"])
        if node["children"]:
            child_entries = [process_node(c) for c in node["children"]]
            return (epub.Section(node["title"], href=chap.file_name), tuple(child_entries))
        return chap

    toc_entries = []

    # 표지/제목 등 root의 문단(첫 제목 앞의 내용)이 있으면 별도 챕터로
    if tree["paras"]:
        intro_chap = make_chapter(title or "시작", tree["paras"])
        toc_entries.append(intro_chap)

    for child in tree["children"]:
        toc_entries.append(process_node(child))

    if not all_chapters_flat:
        make_chapter("본문", ["(빈 문서입니다)"])
        toc_entries = list(all_chapters_flat)

    book.toc = tuple(toc_entries)
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())
    book.spine = ["nav"] + all_chapters_flat

    out = io.BytesIO()
    epub.write_epub(out, book)
    out.seek(0)
    return out.read()


@app.get("/", response_class=HTMLResponse)
async def index():
    options_html = "".join(
        f'<option value="{key}">{val["label"]}</option>'
        for key, val in TEMPLATES.items()
    )
    return f"""
    <!DOCTYPE html>
    <html lang="ko">
    <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>다온 EPUB 변환기 (웹)</title>
    <style>
      body {{ font-family: 'Pretendard', -apple-system, sans-serif; background:#101423; color:#F5F9FF; max-width:640px; margin:60px auto; padding:0 20px; }}
      h1 {{ font-size:24px; margin-bottom:8px; }}
      .notice {{ background:rgba(79,182,198,0.12); border:1px solid rgba(79,182,198,0.4); padding:12px 16px; border-radius:6px; font-size:13px; margin-bottom:28px; }}
      label {{ display:block; margin-top:16px; font-size:14px; color:#B9C3D9; }}
      input, select {{ width:100%; padding:10px; margin-top:6px; border-radius:4px; border:1px solid #333; background:#181D2E; color:#fff; box-sizing:border-box; }}
      button {{ margin-top:28px; width:100%; padding:14px; background:#4C8DFF; border:none; border-radius:4px; color:#fff; font-size:15px; font-weight:600; cursor:pointer; }}
      button:hover {{ background:#6C93EF; }}
    </style>
    </head>
    <body>
      <h1>다온 EPUB 변환기 (웹)</h1>
      <div class="notice">현재는 워드(.docx) 파일만 지원합니다. 한글(HWP) 파일 지원은 준비 중입니다.</div>
      <form action="/convert" method="post" enctype="multipart/form-data">
        <label>원고 파일 (.docx)</label>
        <input type="file" name="file" accept=".docx" required>

        <label>제목</label>
        <input type="text" name="title" required>

        <label>저자</label>
        <input type="text" name="author">

        <label>출판사</label>
        <input type="text" name="publisher">

        <label>ISBN (선택)</label>
        <input type="text" name="isbn">

        <label>디자인 템플릿</label>
        <select name="template_key">{options_html}</select>

        <button type="submit">EPUB으로 변환하기</button>
      </form>
    </body>
    </html>
    """


@app.post("/convert")
async def convert(
    file: UploadFile = File(...),
    title: str = Form(...),
    author: str = Form(""),
    publisher: str = Form(""),
    isbn: str = Form(""),
    template_key: str = Form("essay"),
):
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="현재는 .docx 파일만 지원합니다.")

    docx_bytes = await file.read()

    try:
        epub_bytes = build_epub(
            docx_bytes=docx_bytes,
            title=title,
            author=author,
            publisher=publisher,
            isbn=isbn,
            language="ko",
            template_key=template_key,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"변환 중 오류가 발생했습니다: {e}")

    from urllib.parse import quote

    safe_title = re.sub(r"[^\w\-가-힣]", "_", title) or "output"
    ascii_fallback = "book.epub"
    encoded_name = quote(f"{safe_title}.epub")
    return StreamingResponse(
        io.BytesIO(epub_bytes),
        media_type="application/epub+zip",
        headers={
            "Content-Disposition": (
                f'attachment; filename="{ascii_fallback}"; '
                f"filename*=UTF-8''{encoded_name}"
            )
        },
    )


@app.get("/health")
async def health():
    return {"status": "ok"}
