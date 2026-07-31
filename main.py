"""
다온 EPUB 변환기 (웹 버전) - 2단계: DOCX to EPUB
FastAPI 기반. python-docx로 원고를 읽고, ebooklib으로 EPUB을 생성합니다.

수정 내역 (오류 리포트 반영):
1. 제목 입력 필수 해제 -> 워드의 "Title" 스타일을 자동 추출해 진짜 제목으로 사용
2. 저자 미입력 시 "저자 미상"으로 기본값 채움 (PC 버전과 동일 정책)
3. 목차(nav)를 spine 맨 앞이 아니라 표지(1번 챕터) 다음으로 이동
4. 챕터 분할 기준을 Heading 1(부)·Heading 2(장)까지로 제한.
   Heading 3 이상은 새 파일을 만들지 않고, 같은 챕터 파일 안의 소제목(anchor)으로 삽입
5. 문단 앞뒤 공백(trim) 처리 추가 -> 저자명 등에 남아있던 긴 공백 제거
6. EPUB 2.0 / 3.0 선택 기능 추가
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

# ---------------------------------------------------------------------------
# 챕터 분할 기준: 이 레벨까지만 "새 파일(페이지)"을 만듭니다.
#   Heading 1 = 부(Part)
#   Heading 2 = 장(Chapter)
#   Heading 3 이상 = 장 안의 소제목 -> 새 파일 대신 같은 파일 안 앵커(#id)로 삽입
# 원고 구조가 다르면 이 숫자만 조정하면 됩니다.
# ---------------------------------------------------------------------------
MAX_SPLIT_LEVEL = 2

TEMPLATES = {
    "essay": {
        "label": "에세이",
        "css": """
            body { font-family: 'Noto Serif KR', serif; line-height: 1.9; margin: 5% 6%; }
            h1 { font-size: 1.5em; margin-bottom: 1.2em; text-align: center; }
            h2, h3, h4 { margin-top: 2em; margin-bottom: 0.8em; }
            p { margin: 0 0 1em 0; text-indent: 1em; }
            p.subtitle { text-align: center; text-indent: 0; font-style: italic; color: #555; }
        """,
    },
    "novel": {
        "label": "소설",
        "css": """
            body { font-family: 'Noto Serif KR', serif; line-height: 2.0; margin: 6% 7%; }
            h1 { font-size: 1.6em; margin-bottom: 1.5em; text-align: center; letter-spacing: 0.05em; }
            h2, h3, h4 { margin-top: 2em; margin-bottom: 0.8em; }
            p { margin: 0 0 0.8em 0; text-indent: 1em; }
            p.subtitle { text-align: center; text-indent: 0; font-style: italic; color: #555; }
        """,
    },
    "practical": {
        "label": "실용/정보",
        "css": """
            body { font-family: 'Noto Sans KR', sans-serif; line-height: 1.7; margin: 5%; }
            h1 { font-size: 1.4em; margin-bottom: 1em; border-bottom: 2px solid #333; padding-bottom: 0.3em; }
            h2, h3, h4 { margin-top: 1.6em; margin-bottom: 0.6em; }
            p { margin: 0 0 1em 0; }
            p.subtitle { text-align: center; font-style: italic; color: #555; }
        """,
    },
}


def escape_html(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def get_paragraph_html(para) -> str:
    """문단 안의 텍스트를 가져오되, 수동 줄바꿈(Shift+Enter)은 <br/>로 보존합니다.
    앞뒤 공백은 제거하되(트림), 문장 중간 공백은 그대로 둡니다."""
    parts = []
    for run in para.runs:
        for child in run._element:
            tag = child.tag.split("}")[-1]
            if tag == "t":
                parts.append(escape_html(child.text or ""))
            elif tag == "br":
                parts.append("<br/>")
            elif tag == "tab":
                parts.append(" ")  # 탭 문자는 공백 하나로 정리 (소제목에 탭이 그대로 남는 문제 방지)
    html = "".join(parts)
    # 문단 앞뒤의 공백만 제거 (예: 저자명 앞에 정렬용으로 넣은 스페이스 다수)
    return html.strip()


def build_heading_tree(document: "docx.Document"):
    """
    'Heading 1' ~ 'Heading 6'(제목 1~6) 스타일을 몇 단계든 자동으로 인식해 트리 구조를 만듭니다.
    'Title' 스타일 문단은 진짜 책 제목으로, 그 외 Heading 이전에 나오는 문단(부제 등)은
    표지 텍스트(root paras)로 취급합니다.
    """
    root = {"title_lines": [], "paras": [], "level": 0, "children": []}
    stack = [root]

    for para in document.paragraphs:
        style_name = (para.style.name or "").lower()
        html = get_paragraph_html(para)

        if style_name == "title":
            if html:
                root["title_lines"].append(html)
            continue

        m = re.match(r"(?:heading|제목)\s*(\d+)", style_name)
        level = int(m.group(1)) if m else None

        if level is not None and html:
            while len(stack) > 1 and stack[-1]["level"] >= level:
                stack.pop()
            node = {"title": html, "level": level, "paras": [], "children": []}
            stack[-1]["children"].append(node)
            stack.append(node)
        else:
            if html:
                css_class = "subtitle" if style_name == "subtitle" else None
                stack[-1]["paras"].append({"html": html, "class": css_class})

    return root


def build_epub(
    docx_bytes: bytes,
    title_override: str,
    author: str,
    publisher: str,
    isbn: str,
    language: str,
    template_key: str,
    epub_version: str,
) -> bytes:
    document = docx.Document(io.BytesIO(docx_bytes))
    tree = build_heading_tree(document)
    template = TEMPLATES.get(template_key, TEMPLATES["essay"])

    # 진짜 제목 결정: 워드의 Title 스타일 문단이 있으면 그걸 우선 사용,
    # 없을 경우에만 사용자가 입력한 title_override(선택 입력)를 사용
    extracted_title = " ".join(tree["title_lines"]).strip()
    book_title = extracted_title or title_override.strip() or "제목 없음"
    book_author = author.strip() if author and author.strip() else "저자 미상"

    book = epub.EpubBook()
    book.set_identifier(isbn if isbn else str(uuid.uuid4()))
    book.set_title(book_title)
    book.set_language(language or "ko")
    book.add_author(book_author)
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

    def render_paras(paras) -> str:
        out = []
        for p in paras:
            cls = f' class="{p["class"]}"' if p.get("class") else ""
            out.append(f'<p{cls}>{p["html"]}</p>')
        return "\n".join(out)

    def make_chapter(chap_title: str, paras: list) -> "epub.EpubHtml":
        nonlocal counter
        counter += 1
        file_name = f"chap_{counter:03d}.xhtml"
        body_html = f"<h1>{escape_html(chap_title)}</h1>\n" if chap_title else ""
        body_html += render_paras(paras)
        c = epub.EpubHtml(title=chap_title or f"챕터 {counter}", file_name=file_name, lang=language or "ko")
        c.content = f"<html><head></head><body>{body_html}</body></html>"
        c.add_item(style_item)
        book.add_item(c)
        all_chapters_flat.append(c)
        return c

    def append_anchor_to_chapter(chapter, node) -> str:
        """새 파일을 만들지 않고, 기존 챕터 파일 끝에 소제목(h3~h6)+본문을 이어붙입니다."""
        anchor_id = f"sec_{uuid.uuid4().hex[:8]}"
        heading_tag = f"h{min(node['level'] + 1, 6)}"  # 챕터 h1 다음 레벨이므로 +1
        anchor_html = f'<{heading_tag} id="{anchor_id}">{escape_html(node["title"])}</{heading_tag}>\n'
        anchor_html += render_paras(node["paras"])
        chapter.content = chapter.content.replace("</body>", anchor_html + "\n</body>")
        return anchor_id

    def process_node(node):
        """
        반환값: (toc_entry, 사용된_챕터파일)
        toc_entry는 epub.Link, epub.EpubHtml, 또는 (Section, tuple) 형태입니다.
        """
        if node["level"] <= MAX_SPLIT_LEVEL:
            chap = make_chapter(node["title"], node["paras"])
            child_entries = []
            for child in node["children"]:
                entry, _ = process_node_in_chapter(child, chap)
                child_entries.append(entry)
            if child_entries:
                return (epub.Section(node["title"], href=chap.file_name), tuple(child_entries)), chap
            return chap, chap
        else:
            # 이론상 여기 도달하면 최상위 노드가 MAX_SPLIT_LEVEL보다 깊은 경우인데,
            # 실제로는 process_node_in_chapter 쪽에서 처리되므로 안전장치로만 둡니다.
            return process_node_in_chapter(node, all_chapters_flat[-1] if all_chapters_flat else None)

    def process_node_in_chapter(node, current_chapter):
        """MAX_SPLIT_LEVEL을 넘는 소제목을 현재 챕터 파일 안에 앵커로 삽입."""
        if node["level"] <= MAX_SPLIT_LEVEL:
            # 혹시 트리 구조상 다시 상위 레벨이 나오면 새 챕터로 분리
            chap = make_chapter(node["title"], node["paras"])
            child_entries = []
            for child in node["children"]:
                entry, _ = process_node_in_chapter(child, chap)
                child_entries.append(entry)
            if child_entries:
                return (epub.Section(node["title"], href=chap.file_name), tuple(child_entries)), chap
            return chap, chap

        anchor_id = append_anchor_to_chapter(current_chapter, node)
        link = epub.Link(f"{current_chapter.file_name}#{anchor_id}", node["title"], anchor_id)
        child_entries = []
        for child in node["children"]:
            entry, _ = process_node_in_chapter(child, current_chapter)
            child_entries.append(entry)
        if child_entries:
            return (link, tuple(child_entries)), current_chapter
        return link, current_chapter

    toc_entries = []

    # 표지(제목) 챕터: 항상 1번 챕터로 생성. 부제/저자명 등은 이 챕터의 본문으로 들어감.
    intro_chap = make_chapter(book_title, tree["paras"])
    toc_entries.append(intro_chap)

    for child in tree["children"]:
        entry, _ = process_node(child)
        toc_entries.append(entry)

    book.toc = tuple(toc_entries)

    # --- EPUB 2.0 / 3.0 분기 ---
    # ebooklib은 호환성을 위해 기본적으로 NCX(2.0 표준 목차)를 항상 생성합니다.
    # nav.xhtml(EPUB3 표준 목차)은 3.0을 선택했을 때만 추가합니다.
    book.add_item(epub.EpubNcx())
    if epub_version == "3":
        book.add_item(epub.EpubNav())
        # 목차(nav)는 "제목 다음"에 오도록 표지 챕터 바로 뒤에 배치 (맨 앞 X)
        book.spine = [all_chapters_flat[0], "nav"] + all_chapters_flat[1:]
    else:
        book.spine = list(all_chapters_flat)

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
      button:disabled {{ opacity:0.6; cursor:not-allowed; }}
      .hint {{ font-size:12px; color:#7C8AA8; margin-top:4px; }}
      #status {{ margin-top:14px; font-size:13px; color:#B9C3D9; min-height:18px; }}
    </style>
    </head>
    <body>
      <h1>다온 EPUB 변환기 (웹)</h1>
      <div class="notice">현재는 워드(.docx) 파일만 지원합니다. 한글(HWP) 파일 지원은 준비 중입니다.</div>
      <form id="convertForm">
        <label>원고 파일 (.docx)</label>
        <input type="file" name="file" accept=".docx" required>

        <label>제목 <span class="hint">(비워두면 원고의 "Title" 스타일 문단을 자동으로 사용합니다)</span></label>
        <input type="text" name="title">

        <label>저자</label>
        <input type="text" name="author">

        <label>출판사</label>
        <input type="text" name="publisher">

        <label>ISBN (선택)</label>
        <input type="text" name="isbn">

        <label>디자인 템플릿</label>
        <select name="template_key">{options_html}</select>

        <label>EPUB 버전</label>
        <select name="epub_version">
          <option value="3">EPUB 3.0 (권장)</option>
          <option value="2">EPUB 2.0</option>
        </select>

        <button type="submit" id="submitBtn">EPUB으로 변환하기</button>
        <div id="status"></div>
      </form>

      <script>
        const form = document.getElementById('convertForm');
        const btn = document.getElementById('submitBtn');
        const status = document.getElementById('status');

        form.addEventListener('submit', async (e) => {{
          e.preventDefault();
          btn.disabled = true;
          status.textContent = '변환 중입니다...';

          const formData = new FormData(form);

          try {{
            const res = await fetch('/convert', {{ method: 'POST', body: formData }});
            if (!res.ok) {{
              const err = await res.json().catch(() => ({{}}));
              throw new Error(err.detail || '변환 중 오류가 발생했습니다.');
            }}
            const blob = await res.blob();

            // 파일명 추출 (서버가 내려준 Content-Disposition 참고, 실패 시 기본값)
            const disposition = res.headers.get('Content-Disposition') || '';
            const match = disposition.match(/filename\\*=UTF-8''([^;]+)/);
            const suggestedName = match ? decodeURIComponent(match[1]) : 'book.epub';

            // 저장 위치 선택 (크롬/엣지 등 지원 브라우저)
            if (window.showSaveFilePicker) {{
              try {{
                const handle = await window.showSaveFilePicker({{
                  suggestedName,
                  types: [{{ description: 'EPUB 파일', accept: {{ 'application/epub+zip': ['.epub'] }} }}],
                }});
                const writable = await handle.createWritable();
                await writable.write(blob);
                await writable.close();
                status.textContent = '저장 완료!';
              }} catch (pickerErr) {{
                if (pickerErr.name === 'AbortError') {{
                  status.textContent = '저장이 취소되었습니다.';
                }} else {{
                  throw pickerErr;
                }}
              }}
            }} else {{
              // 저장 위치 선택 미지원 브라우저(사파리/파이어폭스 등) -> 기본 다운로드 폴더로 저장
              const url = URL.createObjectURL(blob);
              const a = document.createElement('a');
              a.href = url;
              a.download = suggestedName;
              document.body.appendChild(a);
              a.click();
              a.remove();
              URL.revokeObjectURL(url);
              status.textContent = '다운로드 폴더에 저장되었습니다.';
            }}
          }} catch (err) {{
            status.textContent = '오류: ' + err.message;
          }} finally {{
            btn.disabled = false;
          }}
        }});
      </script>
    </body>
    </html>
    """


@app.post("/convert")
async def convert(
    file: UploadFile = File(...),
    title: str = Form(""),
    author: str = Form(""),
    publisher: str = Form(""),
    isbn: str = Form(""),
    template_key: str = Form("essay"),
    epub_version: str = Form("3"),
):
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="현재는 .docx 파일만 지원합니다.")

    if epub_version not in ("2", "3"):
        raise HTTPException(status_code=400, detail="epub_version은 '2' 또는 '3'만 지원합니다.")

    docx_bytes = await file.read()

    try:
        epub_bytes = build_epub(
            docx_bytes=docx_bytes,
            title_override=title,
            author=author,
            publisher=publisher,
            isbn=isbn,
            language="ko",
            template_key=template_key,
            epub_version=epub_version,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"변환 중 오류가 발생했습니다: {e}")

    from urllib.parse import quote

    # 파일명은 실제 추출된 제목을 쓰기 어려우므로(스트리밍 이후 알 수 없음),
    # 사용자가 입력한 title 또는 파일 원본 이름을 기준으로 안전하게 생성
    base_name = title.strip() or Path(file.filename).stem
    safe_title = re.sub(r"[^\w\-가-힣]", "_", base_name) or "output"
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
