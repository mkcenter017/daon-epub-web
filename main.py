"""
다온 EPUB 변환기 (웹 버전) - 3단계: DOCX to EPUB + 변환 후 수정 기능

흐름
----
1) POST /parse   : 워드 파일을 업로드하면, 실제 EPUB을 만들지 않고
                    "편집 가능한 구조(JSON)"만 만들어서 돌려줍니다.
                    (제목/저자/장별 제목/본문 텍스트/새 페이지 분리 여부)
2) 프론트엔드에서 그 구조를 화면에 폼으로 그려서, 사용자가 직접 수정합니다.
3) POST /build    : 수정된 구조(JSON)를 그대로 받아서, 그때 비로소 실제
                    EPUB 파일을 만들어 다운로드로 내려줍니다.

서버에는 아무 상태도 저장하지 않습니다(stateless). 수정된 데이터는 전부
브라우저가 들고 있다가 /build 호출 때 통째로 다시 보내주는 구조라서,
Render 같은 무료 서버가 재시작되거나 여러 인스턴스로 떠도 문제가 없습니다.
"""

import io
import json
import re
import uuid
from typing import List, Optional

from fastapi import FastAPI, Form, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from pydantic import BaseModel

import docx
from ebooklib import epub

app = FastAPI(title="다온 EPUB 변환기")

# 이 레벨(포함)까지는 기본적으로 "새 페이지"로 분리합니다. (수정 화면에서 개별적으로 덮어쓸 수 있음)
DEFAULT_SPLIT_LEVEL = 2

TEMPLATES = {
    "essay": {
        "label": "에세이",
        "css": """
            body { font-family: 'Noto Serif KR', serif; line-height: 1.9; margin: 5% 6%; }
            h1 { font-size: 1.5em; margin-bottom: 1.2em; text-align: center; }
            h2, h3, h4, h5, h6 { margin-top: 2em; margin-bottom: 0.8em; }
            p { margin: 0 0 1em 0; text-indent: 1em; }
        """,
    },
    "novel": {
        "label": "소설",
        "css": """
            body { font-family: 'Noto Serif KR', serif; line-height: 2.0; margin: 6% 7%; }
            h1 { font-size: 1.6em; margin-bottom: 1.5em; text-align: center; letter-spacing: 0.05em; }
            h2, h3, h4, h5, h6 { margin-top: 2em; margin-bottom: 0.8em; }
            p { margin: 0 0 0.8em 0; text-indent: 1em; }
        """,
    },
    "practical": {
        "label": "실용/정보",
        "css": """
            body { font-family: 'Noto Sans KR', sans-serif; line-height: 1.7; margin: 5%; }
            h1 { font-size: 1.4em; margin-bottom: 1em; border-bottom: 2px solid #333; padding-bottom: 0.3em; }
            h2, h3, h4, h5, h6 { margin-top: 1.6em; margin-bottom: 0.6em; }
            p { margin: 0 0 1em 0; }
        """,
    },
}


# ---------------------------------------------------------------------------
# 공통 유틸
# ---------------------------------------------------------------------------

def escape_html(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def plain_to_html(text: str) -> str:
    """편집 화면에서 받은 순수 텍스트(줄바꿈 포함)를 안전한 HTML로 변환.
    한 줄바꿈(\\n)은 <br/>로, 문단 사이 빈 줄은 호출하는 쪽에서 이미 별도 문단으로 나눠 넘겨줌."""
    return escape_html(text).replace("\n", "<br/>")


def get_paragraph_plain_text(para) -> str:
    """워드 문단에서 순수 텍스트를 뽑아옵니다. Shift+Enter 줄바꿈은 \\n으로,
    탭은 공백 하나로 바꾸고, 앞뒤 공백은 제거합니다."""
    parts = []
    for run in para.runs:
        for child in run._element:
            tag = child.tag.split("}")[-1]
            if tag == "t":
                parts.append(child.text or "")
            elif tag == "br":
                parts.append("\n")
            elif tag == "tab":
                parts.append(" ")
    return "".join(parts).strip()


# ---------------------------------------------------------------------------
# 1단계: 워드 -> 편집 가능한 구조(JSON)
# ---------------------------------------------------------------------------

def build_heading_tree(document: "docx.Document") -> dict:
    """
    'Heading 1'~'Heading 6'(제목 1~6) 스타일을 자동으로 인식해 트리 구조를 만듭니다.
    'Title' 스타일 문단은 표지 제목(root)에 포함시킵니다.
    """
    root = {"title_lines": [], "paras": [], "level": 0, "children": []}
    stack = [root]

    for para in document.paragraphs:
        style_name = (para.style.name or "").lower()
        text = get_paragraph_plain_text(para)

        if style_name == "title":
            if text:
                root["title_lines"].append(text)
            continue

        m = re.match(r"(?:heading|제목)\s*(\d+)", style_name)
        level = int(m.group(1)) if m else None

        if level is not None and text:
            while len(stack) > 1 and stack[-1]["level"] >= level:
                stack.pop()
            node = {"title": text, "level": level, "paras": [], "children": []}
            stack[-1]["children"].append(node)
            stack.append(node)
        else:
            if text:
                stack[-1]["paras"].append(text)

    return root


_id_counter = 0


def _next_id() -> str:
    global _id_counter
    _id_counter += 1
    return f"n_{_id_counter}"


def tree_to_editable(node: dict, is_root: bool = False) -> dict:
    """내부 트리 구조를 프론트엔드가 그대로 렌더링할 수 있는 JSON 구조로 변환."""
    if is_root:
        title = " ".join(node["title_lines"]).strip()
        level = 0
    else:
        title = node["title"]
        level = node["level"]

    return {
        "id": "root" if is_root else _next_id(),
        "level": level,
        "title": title,
        "body": "\n\n".join(node["paras"]),   # 문단 사이는 빈 줄로 구분해 하나의 텍스트로 합침
        "split": True if is_root else (level <= DEFAULT_SPLIT_LEVEL),
        "children": [tree_to_editable(c) for c in node["children"]],
    }


@app.post("/parse")
async def parse_docx(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="현재는 .docx 파일만 지원합니다.")

    docx_bytes = await file.read()
    try:
        document = docx.Document(io.BytesIO(docx_bytes))
        tree = build_heading_tree(document)
        global _id_counter
        _id_counter = 0
        editable = tree_to_editable(tree, is_root=True)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"원고를 읽는 중 오류가 발생했습니다: {e}")

    extracted_title = editable["title"]
    return JSONResponse({
        "extracted_title": extracted_title,
        "root": editable,
    })


# ---------------------------------------------------------------------------
# 2단계: 수정된 구조(JSON) -> 실제 EPUB
# ---------------------------------------------------------------------------

class EditableNode(BaseModel):
    id: str
    level: int
    title: str
    body: str = ""
    split: bool
    children: List["EditableNode"] = []


EditableNode.model_rebuild()


class BuildRequest(BaseModel):
    root: EditableNode
    book_title: str = ""
    author: str = ""
    publisher: str = ""
    isbn: str = ""
    template_key: str = "essay"
    epub_version: str = "3"


def render_body_html(body_text: str) -> str:
    """편집 화면에서 받은 본문(빈 줄로 문단 구분)을 <p> 태그들로 변환."""
    if not body_text.strip():
        return ""
    paragraphs = [p for p in body_text.split("\n\n") if p.strip()]
    return "\n".join(f"<p>{plain_to_html(p.strip())}</p>" for p in paragraphs)


def build_epub_from_structure(req: BuildRequest) -> bytes:
    template = TEMPLATES.get(req.template_key, TEMPLATES["essay"])
    book_title = (req.book_title or req.root.title or "제목 없음").strip()
    book_author = req.author.strip() if req.author.strip() else "저자 미상"

    book = epub.EpubBook()
    book.set_identifier(req.isbn if req.isbn else str(uuid.uuid4()))
    book.set_title(book_title)
    book.set_language("ko")
    book.add_author(book_author)
    if req.publisher.strip():
        book.add_metadata("DC", "publisher", req.publisher.strip())

    style_item = epub.EpubItem(
        uid="style_default",
        file_name="style/default.css",
        media_type="text/css",
        content=template["css"],
    )
    book.add_item(style_item)

    all_chapters_flat = []
    counter = 0

    def make_chapter(title: str, body_text: str) -> "epub.EpubHtml":
        nonlocal counter
        counter += 1
        file_name = f"chap_{counter:03d}.xhtml"
        body_html = f"<h1>{escape_html(title)}</h1>\n" if title else ""
        body_html += render_body_html(body_text)
        c = epub.EpubHtml(title=title or f"챕터 {counter}", file_name=file_name, lang="ko")
        c.content = f"<html><head></head><body>{body_html}</body></html>"
        c.add_item(style_item)
        book.add_item(c)
        all_chapters_flat.append(c)
        return c

    def append_anchor(chapter, node: EditableNode) -> str:
        anchor_id = f"sec_{uuid.uuid4().hex[:8]}"
        heading_tag = f"h{min(max(node.level, 2) + 1, 6)}"
        html = f'<{heading_tag} id="{anchor_id}">{escape_html(node.title)}</{heading_tag}>\n'
        html += render_body_html(node.body)
        chapter.content = chapter.content.replace("</body>", html + "\n</body>")
        return anchor_id

    def walk(node: EditableNode, current_chapter):
        """current_chapter=None이면 반드시 새 챕터를 만들어야 하는 상황(=root 또는 split=True)."""
        toc_entry = None
        if current_chapter is None or node.split:
            chap = make_chapter(node.title, node.body)
            child_entries = []
            for child in node.children:
                entry = walk(child, chap)
                if entry is not None:
                    child_entries.append(entry)
            toc_entry = (epub.Section(node.title, href=chap.file_name), tuple(child_entries)) if child_entries else chap
        else:
            anchor_id = append_anchor(current_chapter, node)
            link = epub.Link(f"{current_chapter.file_name}#{anchor_id}", node.title, anchor_id)
            child_entries = []
            for child in node.children:
                entry = walk(child, current_chapter)
                if entry is not None:
                    child_entries.append(entry)
            toc_entry = (link, tuple(child_entries)) if child_entries else link
        return toc_entry

    toc_entries = [walk(req.root, None)]
    book.toc = tuple(toc_entries)

    book.add_item(epub.EpubNcx())
    if req.epub_version == "3":
        book.add_item(epub.EpubNav())
        book.spine = [all_chapters_flat[0], "nav"] + all_chapters_flat[1:]
    else:
        book.spine = list(all_chapters_flat)

    out = io.BytesIO()
    epub.write_epub(out, book)
    out.seek(0)
    return out.read()


@app.post("/build")
async def build(req: BuildRequest):
    try:
        epub_bytes = build_epub_from_structure(req)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"EPUB 생성 중 오류가 발생했습니다: {e}")

    from urllib.parse import quote
    base_name = req.book_title.strip() or "output"
    safe_title = re.sub(r"[^\w\-가-힣]", "_", base_name) or "output"
    encoded_name = quote(f"{safe_title}.epub")
    return StreamingResponse(
        io.BytesIO(epub_bytes),
        media_type="application/epub+zip",
        headers={
            "Content-Disposition": (
                f'attachment; filename="book.epub"; '
                f"filename*=UTF-8''{encoded_name}"
            )
        },
    )


# ---------------------------------------------------------------------------
# 화면 (업로드 -> 미리보기/수정 -> 다운로드, 전부 한 페이지에서 JS로 처리)
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def index():
    options_html = "".join(
        f'<option value="{key}">{val["label"]}</option>'
        for key, val in TEMPLATES.items()
    )
    css_map_json = json.dumps({key: val["css"] for key, val in TEMPLATES.items()})
    return f"""
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>다온 EPUB 변환기 (웹)</title>
<style>
  :root {{
    --bg:#0E1220; --panel:#161B2C; --panel2:#1D2338; --border:#2A3350;
    --text:#F5F9FF; --sub:#8C97B8; --accent:#4C8DFF; --accent-soft:rgba(76,141,255,0.14);
    --lv1:#4C8DFF; --lv2:#5EE6D9; --lv3:#F2C46D; --lv4:#F28D8D; --lv5:#C89BF2; --lv6:#8C97B8;
  }}
  * {{ box-sizing:border-box; }}
  body {{
    font-family:'Pretendard', -apple-system, sans-serif; background:var(--bg); color:var(--text);
    max-width:1560px; margin:0 auto; padding:0 32px 80px;
  }}
  h1 {{ font-size:21px; margin:28px 0 4px; }}
  .notice {{
    background:var(--accent-soft); border:1px solid rgba(76,141,255,0.35); padding:11px 14px;
    border-radius:8px; font-size:12.5px; color:#C7D6FF; margin-bottom:22px; line-height:1.5;
  }}
  label {{ display:block; font-size:12px; color:var(--sub); margin-bottom:5px; }}
  input[type=text], input[type=file], select, textarea {{
    width:100%; padding:9px 11px; border-radius:6px; border:1px solid var(--border);
    background:var(--panel2); color:#fff; box-sizing:border-box; font-family:inherit; font-size:14px;
    transition:border-color .15s ease;
  }}
  input[type=text]:focus, textarea:focus, select:focus {{ outline:none; border-color:var(--accent); }}
  textarea {{ min-height:64px; resize:vertical; line-height:1.65; }}

  button {{
    padding:10px 16px; background:var(--accent); border:none; border-radius:7px; color:#fff;
    font-size:13.5px; font-weight:600; cursor:pointer; transition:background .15s ease;
  }}
  button:hover {{ background:#6C9BFF; }}
  button:disabled {{ opacity:0.45; cursor:not-allowed; }}
  button.secondary {{ background:var(--panel2); border:1px solid var(--border); color:var(--sub); }}
  button.secondary:hover {{ background:#242B45; color:#fff; }}
  button.ghost {{
    background:transparent; border:1px solid var(--border); color:var(--sub);
    padding:5px 10px; font-size:11.5px; font-weight:500;
  }}
  button.ghost:hover {{ border-color:var(--accent); color:var(--accent); background:transparent; }}

  #status {{ margin-top:10px; font-size:12.5px; color:var(--sub); min-height:16px; }}

  /* --- 업로드 화면 --- */
  #uploadSection {{
    margin-top:8px; background:var(--panel); border:1px solid var(--border); border-radius:12px;
    padding:22px;
  }}
  #uploadSection button {{ margin-top:14px; width:100%; padding:13px; font-size:14.5px; }}

  /* --- 편집 화면 --- */
  #editSection {{ display:none; }}

  .meta-card {{
    background:var(--panel); border:1px solid var(--border); border-radius:12px; padding:18px;
    margin-top:18px;
  }}
  .meta-grid {{ display:grid; grid-template-columns:repeat(auto-fit, minmax(180px, 1fr)); gap:14px 18px; }}
  .meta-grid .full {{ grid-column:1 / -1; }}

  .toolbar {{
    position:sticky; top:0; z-index:10; background:rgba(14,18,32,0.92); backdrop-filter:blur(8px);
    padding:14px 0 12px; margin-top:4px; border-bottom:1px solid var(--border);
    display:flex; align-items:center; justify-content:space-between; gap:10px; flex-wrap:wrap;
  }}
  .toolbar .title {{ font-size:14px; font-weight:600; }}
  .toolbar .count {{ font-size:12px; color:var(--sub); font-weight:400; }}
  .toolbar-actions {{ display:flex; gap:8px; }}

  .layout {{ display:grid; grid-template-columns:1fr 1fr; gap:18px; align-items:start; margin-top:4px; }}
  @media (max-width:920px) {{ .layout {{ grid-template-columns:1fr; }} }}
  .editor-col {{ min-width:0; }}
  .preview-col {{ min-width:0; position:sticky; top:64px; }}
  @media (max-width:920px) {{ .preview-col {{ position:static; }} }}

  .preview-head {{
    padding:10px 2px 12px; font-size:13px; font-weight:600; color:var(--text);
  }}
  .preview-head .preview-sub {{
    font-size:11.5px; font-weight:400; color:var(--sub); margin-top:3px; line-height:1.5;
  }}

  .outline-box {{
    background:var(--panel); border:1px solid var(--border); border-radius:10px;
    padding:6px; max-height:360px; overflow-y:auto;
  }}
  .outline-row {{
    display:flex; align-items:center; gap:8px; padding:8px 10px; border-radius:6px;
    cursor:pointer; font-size:13px; font-weight:600;
  }}
  .outline-row:hover {{ background:var(--panel2); }}
  .outline-row.selected {{ background:var(--accent-soft); }}
  .outline-row .file-tag {{
    flex:none; font-size:10px; font-weight:700; color:var(--accent);
    background:rgba(76,141,255,0.14); padding:2px 7px; border-radius:20px; white-space:nowrap;
  }}
  .outline-row .row-title {{ overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
  .outline-row.sub {{
    margin-left:22px; font-weight:400; font-size:12px; color:var(--sub);
  }}
  .outline-row.sub .arrow {{ flex:none; opacity:0.55; }}
  .outline-row.sub.selected {{ color:var(--text); }}

  .selected-preview {{
    margin-top:14px; background:var(--panel); border:1px solid var(--border); border-radius:10px; overflow:hidden;
  }}
  .selected-preview .sp-label {{
    display:flex; align-items:center; justify-content:space-between; gap:8px;
    font-size:12px; padding:9px 12px; border-bottom:1px solid var(--border); background:var(--panel2);
  }}
  .selected-preview .sp-label .file-badge {{
    display:inline-flex; align-items:center; gap:6px; font-weight:700; color:var(--accent); font-size:11.5px;
  }}
  .selected-preview .sp-label .file-badge .dot {{
    width:6px; height:6px; border-radius:50%; background:var(--accent);
  }}
  .selected-preview .sp-label .file-name {{
    font-family:monospace; font-size:10px; color:var(--sub); white-space:nowrap;
  }}
  .selected-preview iframe {{ width:100%; border:none; display:block; background:#fff; min-height:120px; }}
  .empty-preview {{ font-size:12.5px; color:var(--sub); padding:20px 8px; text-align:center; }}

  .tree {{ margin-top:18px; }}

  .node {{
    background:var(--panel); border:1px solid var(--border); border-radius:10px;
    margin-top:10px; overflow:hidden;
  }}
  .node[data-level="0"] {{ border-left:4px solid var(--lv1); }}
  .node[data-level="1"] {{ border-left:4px solid var(--lv1); }}
  .node[data-level="2"] {{ border-left:4px solid var(--lv2); }}
  .node[data-level="3"] {{ border-left:4px solid var(--lv3); }}
  .node[data-level="4"] {{ border-left:4px solid var(--lv4); }}
  .node[data-level="5"] {{ border-left:4px solid var(--lv5); }}
  .node[data-level="6"] {{ border-left:4px solid var(--lv6); }}

  .node-head {{
    display:flex; align-items:center; gap:10px; padding:10px 12px; cursor:pointer;
    user-select:none;
  }}
  .node-head:hover {{ background:rgba(255,255,255,0.02); }}

  .caret {{
    width:16px; height:16px; flex:none; color:var(--sub); transition:transform .15s ease;
    display:flex; align-items:center; justify-content:center; font-size:11px;
  }}
  .node.collapsed > .node-head .caret {{ transform:rotate(-90deg); }}

  .level-badge {{
    flex:none; width:11px; height:11px; border-radius:50%; margin-right:1px;
  }}
  .node[data-level="0"] .level-badge {{ background:var(--lv1); }}
  .node[data-level="1"] .level-badge {{ background:var(--lv1); }}
  .node[data-level="2"] .level-badge {{ background:var(--lv2); }}
  .node[data-level="3"] .level-badge {{ background:var(--lv3); }}
  .node[data-level="4"] .level-badge {{ background:var(--lv4); }}
  .node[data-level="5"] .level-badge {{ background:var(--lv5); }}
  .node[data-level="6"] .level-badge {{ background:var(--lv6); }}

  .node-head input[type=text] {{
    flex:1 1 auto; min-width:0; background:transparent; border:1px solid transparent; padding:5px 7px;
    font-size:14px; font-weight:600; border-radius:5px;
  }}
  .node[data-level="0"] > .node-head input[type=text] {{ font-size:15.5px; }}
  .node-head input[type=text]:hover {{ border-color:var(--border); }}
  .node-head input[type=text]:focus {{ border-color:var(--accent); background:var(--panel2); }}

  .split-toggle {{
    flex:none; display:flex; align-items:center; gap:5px; font-size:11px; color:var(--sub);
    white-space:nowrap; padding:4px 8px; border-radius:20px; border:1px solid var(--border);
  }}
  .split-toggle.on {{ color:#9CFFEB; border-color:rgba(94,230,217,0.4); background:rgba(94,230,217,0.08); }}
  .split-toggle input {{ width:auto; margin:0; accent-color:var(--lv2); }}

  .node-body-wrap {{ padding:0 14px 14px 14px; }}
  .node-body-wrap textarea {{ font-size:13.5px; }}
  .node.collapsed .node-body-wrap, .node.collapsed .node-children {{ display:none; }}

  .node-children {{ padding:0 10px 10px 10px; }}

  hr.sep {{ border:none; border-top:1px solid var(--border); margin:26px 0; }}

  .empty-hint {{ font-size:12px; color:var(--sub); padding:10px 4px; }}
</style>
</head>
<body>
  <h1>다온 EPUB 변환기 (웹)</h1>
  <div class="notice">현재는 워드(.docx) 파일만 지원합니다. 변환 전에 제목ㆍ장 제목ㆍ본문ㆍ페이지 분리 여부를 직접 확인하고 수정할 수 있습니다.</div>

  <div id="uploadSection">
    <label>원고 파일 (.docx)</label>
    <input type="file" id="fileInput" accept=".docx">
    <button id="parseBtn">원고 불러오기 (수정 화면으로 이동)</button>
    <div id="uploadStatus" style="margin-top:10px; font-size:12.5px; color:var(--sub);"></div>
  </div>

  <div id="editSection">
    <div class="meta-card">
      <div class="meta-grid">
        <div>
          <label>책 제목 <span style="color:#556">(메타데이터용, 본문과 별개)</span></label>
          <input type="text" id="bookTitle">
        </div>
        <div>
          <label>저자</label>
          <input type="text" id="bookAuthor">
        </div>
        <div>
          <label>출판사</label>
          <input type="text" id="bookPublisher">
        </div>
        <div>
          <label>ISBN (선택)</label>
          <input type="text" id="bookIsbn">
        </div>
        <div>
          <label>디자인 템플릿</label>
          <select id="templateKey">{options_html}</select>
        </div>
        <div>
          <label>EPUB 버전</label>
          <select id="epubVersion">
            <option value="3">EPUB 3.0 (권장)</option>
            <option value="2">EPUB 2.0</option>
          </select>
        </div>
      </div>
    </div>

    <div class="toolbar">
      <div class="title">본문 구조 <span class="count" id="nodeCount"></span></div>
      <div class="toolbar-actions">
        <button class="ghost" id="expandAllBtn" type="button">모두 펼치기</button>
        <button class="ghost" id="collapseAllBtn" type="button">모두 접기</button>
        <button id="buildBtn">최종 EPUB 만들어서 다운로드</button>
      </div>
    </div>

    <div class="layout">
      <div class="editor-col">
        <div class="tree" id="chapterTree"></div>
        <hr class="sep">
        <button class="secondary" id="backBtn" type="button">← 다른 파일 다시 불러오기</button>
        <div id="status"></div>
      </div>
      <div class="preview-col">
        <div class="preview-head">
          <div>파일 구성 한눈에 보기</div>
          <div class="preview-sub">굵은 글씨 = 새 파일, 흐리고 들여쓰기된 글씨 = 같은 파일 안 소제목. <span id="fileCountNote"></span><br>항목을 클릭하면 아래에 실제 페이지 모습이 보입니다.</div>
        </div>
        <div class="outline-box" id="outlineList">
          <div class="empty-preview">원고를 불러오면 여기에 파일 구성이 표시됩니다.</div>
        </div>
        <div class="selected-preview" id="selectedPreview" style="display:none;"></div>
      </div>
    </div>
  </div>

<script>
const TEMPLATE_CSS = {css_map_json};
let rootData = null;
let previewDebounce = null;

function nodeLabel(level) {{
  if (level === 0) return '표지';
  return '레벨 ' + level + ' 제목';
}}

function countNodes(node) {{
  let n = 1;
  (node.children || []).forEach(c => n += countNodes(c));
  return n;
}}

function renderNode(node, container, depth) {{
  const wrap = document.createElement('div');
  wrap.className = 'node';
  wrap.dataset.nodeId = node.id;
  wrap.dataset.level = node.level;
  if (depth > 1) wrap.style.marginLeft = '14px';

  const head = document.createElement('div');
  head.className = 'node-head';

  const caret = document.createElement('span');
  caret.className = 'caret';
  caret.textContent = '▼';
  head.appendChild(caret);

  const badge = document.createElement('span');
  badge.className = 'level-badge';
  badge.title = nodeLabel(node.level);
  head.appendChild(badge);

  const titleInput = document.createElement('input');
  titleInput.type = 'text';
  titleInput.value = node.title;
  titleInput.dataset.field = 'title';
  titleInput.placeholder = '제목 없음';
  head.appendChild(titleInput);

  if (node.level > 0) {{
    const toggleWrap = document.createElement('label');
    toggleWrap.className = 'split-toggle' + (node.split ? ' on' : '');
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = node.split;
    cb.dataset.field = 'split';
    cb.addEventListener('change', () => {{
      toggleWrap.classList.toggle('on', cb.checked);
    }});
    toggleWrap.appendChild(cb);
    toggleWrap.appendChild(document.createTextNode(node.split ? '새 페이지' : '같은 페이지 안 소제목'));
    head.appendChild(toggleWrap);
  }}

  head.addEventListener('click', (e) => {{
    if (e.target.tagName === 'INPUT') return;
    wrap.classList.toggle('collapsed');
  }});

  wrap.appendChild(head);

  const bodyWrap = document.createElement('div');
  bodyWrap.className = 'node-body-wrap';
  const body = document.createElement('textarea');
  body.value = node.body || '';
  body.dataset.field = 'body';
  body.placeholder = '본문 내용 (문단 사이는 빈 줄로 구분됩니다)';
  bodyWrap.appendChild(body);
  wrap.appendChild(bodyWrap);

  const childrenWrap = document.createElement('div');
  childrenWrap.className = 'node-children';
  wrap.appendChild(childrenWrap);
  (node.children || []).forEach(child => renderNode(child, childrenWrap, depth + 1));

  container.appendChild(wrap);
}}

function collectNode(el) {{
  const id = el.dataset.nodeId;
  const level = parseInt(el.dataset.level, 10);
  const titleInput = el.querySelector(':scope > .node-head input[data-field="title"]');
  const splitInput = el.querySelector(':scope > .node-head input[data-field="split"]');
  const bodyInput = el.querySelector(':scope > .node-body-wrap textarea[data-field="body"]');
  const childrenWrap = el.querySelector(':scope > .node-children');
  const children = [];
  if (childrenWrap) {{
    childrenWrap.querySelectorAll(':scope > .node').forEach(childEl => {{
      children.push(collectNode(childEl));
    }});
  }}
  return {{
    id: id,
    level: level,
    title: titleInput ? titleInput.value : '',
    body: bodyInput ? bodyInput.value : '',
    split: splitInput ? splitInput.checked : true,
    children: children,
  }};
}}

// --- 아래부터는 미리보기 전용 함수들. 백엔드의 build_epub_from_structure()/walk()/
//     render_body_html()과 동일한 규칙으로 "실제로 몇 개의 파일이 만들어지는지"를
//     화면에서 그대로 흉내 냅니다. 서버에 요청하지 않고 브라우저에서 바로 계산합니다.
//
//     예전 버전은 모든 파일을 동시에 실제 페이지로 렌더링했는데, 파일 수가 많아지면
//     (1) iframe이 한 번에 너무 많이 생성되어 높이 계산이 깨지고
//     (2) 화면이 빽빽해져서 오히려 구조를 파악하기 어려운 문제가 있었습니다.
//     그래서 지금은 "목차 리스트"만 항상 보여주고, 실제 페이지 렌더링은 클릭한
//     항목 하나에 대해서만 수행합니다.

function escapeHtmlJs(text) {{
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}}

function renderBodyHtmlJs(bodyText) {{
  if (!bodyText || !bodyText.trim()) return '';
  return bodyText.split('\\n\\n')
    .map(p => p.trim())
    .filter(p => p.length > 0)
    .map(p => '<p>' + escapeHtmlJs(p).replace(/\\n/g, '<br/>') + '</p>')
    .join('\\n');
}}

let selectedFileIndex = null;

/**
 * node 트리를 순회하며 "실제로 새 파일이 되는 노드"만 pages 배열에 담고,
 * 동시에 화면에 뿌릴 목차 항목(outline)도 함께 만듭니다.
 * split=false인 노드는 새 파일을 만들지 않고, 가장 가까운 상위 파일의
 * html 뒤에 소제목(h2~h6)+본문으로 이어붙입니다. (백엔드 walk()와 동일한 규칙)
 */
function buildPreviewData(node, pages, outline, parentPageIdx) {{
  let pageIdx = parentPageIdx;
  if (parentPageIdx === null || node.split) {{
    const page = {{ title: node.title, html: '' }};
    page.html += node.title ? ('<h1>' + escapeHtmlJs(node.title) + '</h1>\\n') : '';
    page.html += renderBodyHtmlJs(node.body);
    pages.push(page);
    pageIdx = pages.length - 1;
    outline.push({{ type: 'file', title: node.title || '(제목 없음)', pageIdx: pageIdx }});
  }} else {{
    const page = pages[pageIdx];
    const tag = 'h' + Math.min(Math.max(node.level, 2) + 1, 6);
    page.html += '<div style="margin:18px 0 4px; padding:12px 14px; background:rgba(76,141,255,0.06); ' +
      'border-left:3px solid rgba(76,141,255,0.35); border-radius:0 6px 6px 0;">';
    page.html += '<div style="font-size:10.5px; color:#6b7fb0; margin-bottom:6px;">같은 파일 안 소제목 (새 파일 아님)</div>';
    page.html += '<' + tag + ' style="margin:0 0 6px;">' + escapeHtmlJs(node.title) + '</' + tag + '>\\n';
    page.html += renderBodyHtmlJs(node.body);
    page.html += '</div>';
    outline.push({{ type: 'sub', title: node.title || '(제목 없음)', pageIdx: pageIdx }});
  }}
  (node.children || []).forEach(child => buildPreviewData(child, pages, outline, pageIdx));
  return {{ pages, outline }};
}}

function showSelectedPage(pages) {{
  const box = document.getElementById('selectedPreview');
  if (selectedFileIndex === null || !pages[selectedFileIndex]) {{
    box.style.display = 'none';
    return;
  }}
  box.style.display = 'block';
  box.innerHTML = '';

  const page = pages[selectedFileIndex];
  const fileNum = String(selectedFileIndex + 1).padStart(3, '0');

  const labelBar = document.createElement('div');
  labelBar.className = 'sp-label';
  labelBar.innerHTML =
    '<span class="file-badge"><span class="dot"></span>파일 ' + (selectedFileIndex + 1) + '</span>' +
    '<span class="file-name">chap_' + fileNum + '.xhtml</span>';
  box.appendChild(labelBar);

  const iframe = document.createElement('iframe');
  iframe.setAttribute('scrolling', 'no');
  // load 리스너를 srcdoc 할당 "전"에 먼저 등록해야, 브라우저가 콘텐츠를 빨리
  // 그려버려서 load 이벤트를 놓치는 경우(=높이 계산이 안 되는 버그)를 막을 수 있습니다.
  iframe.addEventListener('load', () => {{
    try {{
      const h = iframe.contentWindow.document.body.scrollHeight;
      iframe.style.height = Math.max(120, h + 28) + 'px';
    }} catch (e) {{ iframe.style.height = '200px'; }}
  }});
  box.appendChild(iframe);

  const css = TEMPLATE_CSS[document.getElementById('templateKey').value] || '';
  const doc = '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>html,body{{margin:0;}} ' + css + '</style></head>' +
    '<body>' + page.html + '</body></html>';
  iframe.srcdoc = doc;
}}

function renderPreview() {{
  const outlineBox = document.getElementById('outlineList');
  if (!rootData) return;

  const rootEl = document.querySelector('#chapterTree > .node');
  if (!rootEl) return;
  const currentTree = collectNode(rootEl);

  const {{ pages, outline }} = buildPreviewData(currentTree, [], [], null);

  const fileCountNote = document.getElementById('fileCountNote');
  if (fileCountNote) {{
    fileCountNote.textContent = '지금 상태로는 총 ' + pages.length + '개 파일이 만들어집니다.';
  }}

  outlineBox.innerHTML = '';
  if (outline.length === 0) {{
    outlineBox.innerHTML = '<div class="empty-preview">표시할 항목이 없습니다.</div>';
  }} else {{
    outline.forEach(entry => {{
      const row = document.createElement('div');
      row.className = 'outline-row' + (entry.type === 'sub' ? ' sub' : '');
      if (entry.pageIdx === selectedFileIndex) row.classList.add('selected');

      if (entry.type === 'file') {{
        row.innerHTML = '<span class="file-tag">파일 ' + (entry.pageIdx + 1) + '</span>' +
          '<span class="row-title">' + escapeHtmlJs(entry.title) + '</span>';
      }} else {{
        row.innerHTML = '<span class="arrow">↳</span>' +
          '<span class="row-title">' + escapeHtmlJs(entry.title) + '</span>';
      }}
      row.addEventListener('click', () => {{
        selectedFileIndex = entry.pageIdx;
        renderPreview();
      }});
      outlineBox.appendChild(row);
    }});
  }}

  if (selectedFileIndex !== null && selectedFileIndex >= pages.length) {{
    selectedFileIndex = null;
  }}
  showSelectedPage(pages);
}}

function schedulePreview() {{
  clearTimeout(previewDebounce);
  previewDebounce = setTimeout(renderPreview, 220);
}}

document.getElementById('parseBtn').addEventListener('click', async () => {{
  const fileInput = document.getElementById('fileInput');
  const uploadStatus = document.getElementById('uploadStatus');
  if (!fileInput.files.length) {{
    uploadStatus.textContent = '먼저 워드 파일을 선택해주세요.';
    return;
  }}
  uploadStatus.textContent = '원고를 읽는 중입니다...';

  const formData = new FormData();
  formData.append('file', fileInput.files[0]);

  try {{
    const res = await fetch('/parse', {{ method: 'POST', body: formData }});
    if (!res.ok) {{
      const err = await res.json().catch(() => ({{}}));
      throw new Error(err.detail || '원고를 읽는 중 오류가 발생했습니다.');
    }}
    const data = await res.json();
    rootData = data.root;

    document.getElementById('bookTitle').value = data.extracted_title || '';
    const treeEl = document.getElementById('chapterTree');
    treeEl.innerHTML = '';
    renderNode(rootData, treeEl, 1);

    const total = countNodes(rootData);
    document.getElementById('nodeCount').textContent = '(총 ' + total + '개 항목)';

    document.getElementById('uploadSection').style.display = 'none';
    document.getElementById('editSection').style.display = 'block';
    uploadStatus.textContent = '';

    // 제목/본문/체크박스가 바뀔 때마다(입력 중에도) 미리보기를 다시 계산
    treeEl.addEventListener('input', schedulePreview);
    treeEl.addEventListener('change', schedulePreview);
    renderPreview();
  }} catch (err) {{
    uploadStatus.textContent = '오류: ' + err.message;
  }}
}});

document.getElementById('backBtn').addEventListener('click', () => {{
  document.getElementById('editSection').style.display = 'none';
  document.getElementById('uploadSection').style.display = 'block';
  document.getElementById('fileInput').value = '';
  document.getElementById('status').textContent = '';
}});

document.getElementById('templateKey').addEventListener('change', renderPreview);

document.getElementById('expandAllBtn').addEventListener('click', () => {{
  document.querySelectorAll('#chapterTree .node').forEach(el => el.classList.remove('collapsed'));
}});
document.getElementById('collapseAllBtn').addEventListener('click', () => {{
  document.querySelectorAll('#chapterTree .node[data-level]').forEach(el => {{
    if (el.dataset.level !== '0') el.classList.add('collapsed');
  }});
}});

document.getElementById('buildBtn').addEventListener('click', async () => {{
  const status = document.getElementById('status');
  const buildBtn = document.getElementById('buildBtn');
  buildBtn.disabled = true;
  status.textContent = 'EPUB 생성 중입니다...';

  const rootEl = document.querySelector('#chapterTree > .node');
  const editedRoot = collectNode(rootEl);

  const payload = {{
    root: editedRoot,
    book_title: document.getElementById('bookTitle').value,
    author: document.getElementById('bookAuthor').value,
    publisher: document.getElementById('bookPublisher').value,
    isbn: document.getElementById('bookIsbn').value,
    template_key: document.getElementById('templateKey').value,
    epub_version: document.getElementById('epubVersion').value,
  }};

  try {{
    const res = await fetch('/build', {{
      method: 'POST',
      headers: {{ 'Content-Type': 'application/json' }},
      body: JSON.stringify(payload),
    }});
    if (!res.ok) {{
      const err = await res.json().catch(() => ({{}}));
      throw new Error(err.detail || 'EPUB 생성 중 오류가 발생했습니다.');
    }}
    const blob = await res.blob();
    const disposition = res.headers.get('Content-Disposition') || '';
    const match = disposition.match(/filename\\*=UTF-8''([^;]+)/);
    const suggestedName = match ? decodeURIComponent(match[1]) : 'book.epub';

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
    buildBtn.disabled = false;
  }}
}});
</script>
</body>
</html>
"""



@app.get("/health")
async def health():
    return {"status": "ok"}
