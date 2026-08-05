"""
다온 EPUB 변환기 (웹 버전) - 단순 자동 변환

편집 화면 없이, 워드 파일을 올리면 바로 EPUB으로 변환해서 다운로드합니다.
아래 규칙으로 "사람이 손대지 않아도" 대체로 잘 나뉘도록 자동 처리합니다.

- 워드의 "Title" 스타일 문단을 진짜 책 제목으로 자동 인식 (제목 입력은 선택사항)
- 저자를 안 써도 "저자 미상"으로 자동 채움
- 문단 앞뒤 공백/탭 문자 자동 정리 (저자명 앞 공백, 소제목 사이 탭 등)
- Heading 1(부)·Heading 2(장)까지만 새 파일로 분리, Heading 3 이상은
  같은 파일 안 소제목으로 자동 병합 (MAX_SPLIT_LEVEL로 조정 가능)
- 목차(nav)가 표지보다 먼저 나오지 않도록 순서 고정
"""

import io
import re
import uuid
from pathlib import Path

from fastapi import FastAPI, Form, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse

import docx
from ebooklib import epub

app = FastAPI(title="다온 EPUB 변환기")

# 이 레벨(포함)까지만 새 파일(페이지)로 분리합니다. 나머지는 같은 파일 안 소제목으로 병합.
#   1 = 부(Part) 단위만 분리
#   2 = 부 + 장(Chapter) 단위까지 분리 (기본값, 대부분의 책 구조에 맞음)
MAX_SPLIT_LEVEL = 2

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


def escape_html(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def get_paragraph_html(para) -> str:
    """문단 텍스트를 가져오되, 수동 줄바꿈(Shift+Enter)은 <br/>로 보존합니다.
    탭 문자는 공백 하나로 정리하고, 문단 앞뒤 공백은 제거합니다
    (예: 저자명 앞에 정렬용으로 넣은 여러 칸 공백)."""
    parts = []
    for run in para.runs:
        for child in run._element:
            tag = child.tag.split("}")[-1]
            if tag == "t":
                parts.append(escape_html(child.text or ""))
            elif tag == "br":
                parts.append("<br/>")
            elif tag == "tab":
                parts.append(" ")
    return "".join(parts).strip()


def build_heading_tree(document: "docx.Document") -> dict:
    """'Heading 1'~'Heading 6' 스타일을 자동 인식해 트리 구조로 만듭니다.
    'Title' 스타일 문단은 진짜 책 제목으로 별도 수집합니다."""
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
                stack[-1]["paras"].append(html)

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

    extracted_title = " ".join(tree["title_lines"]).strip()
    book_title = extracted_title or title_override.strip() or "제목 없음"
    book_author = author.strip() if author.strip() else "저자 미상"

    book = epub.EpubBook()
    book.set_identifier(isbn if isbn else str(uuid.uuid4()))
    book.set_title(book_title)
    book.set_language(language or "ko")
    book.add_author(book_author)
    if publisher.strip():
        book.add_metadata("DC", "publisher", publisher.strip())

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

    def append_anchor(chapter, node: dict) -> str:
        """MAX_SPLIT_LEVEL을 넘는 소제목은 새 파일 대신 현재 챕터 파일 끝에 이어붙입니다."""
        anchor_id = f"sec_{uuid.uuid4().hex[:8]}"
        heading_tag = f"h{min(max(node['level'], 2) + 1, 6)}"
        anchor_html = f'<{heading_tag} id="{anchor_id}">{escape_html(node["title"])}</{heading_tag}>\n'
        anchor_html += "\n".join(f"<p>{html}</p>" for html in node["paras"])
        chapter.content = chapter.content.replace("</body>", anchor_html + "\n</body>")
        return anchor_id

    def process_node(node, current_chapter):
        if node["level"] <= MAX_SPLIT_LEVEL:
            chap = make_chapter(node["title"], node["paras"])
            entries = [process_node(c, chap) for c in node["children"]]
            entries = [e for e in entries if e is not None]
            if entries:
                return (epub.Section(node["title"], href=chap.file_name), tuple(entries))
            return chap
        else:
            anchor_id = append_anchor(current_chapter, node)
            link = epub.Link(f"{current_chapter.file_name}#{anchor_id}", node["title"], anchor_id)
            entries = [process_node(c, current_chapter) for c in node["children"]]
            entries = [e for e in entries if e is not None]
            if entries:
                return (link, tuple(entries))
            return link

    toc_entries = []
    intro_chap = make_chapter(book_title, tree["paras"])
    toc_entries.append(intro_chap)

    for child in tree["children"]:
        toc_entries.append(process_node(child, intro_chap))

    book.toc = tuple(toc_entries)

    book.add_item(epub.EpubNcx())
    if epub_version == "3":
        book.add_item(epub.EpubNav())
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
      body {{ font-family: 'Pretendard', -apple-system, sans-serif; background:#101423; color:#F5F9FF; max-width:560px; margin:60px auto; padding:0 20px; }}
      h1 {{ font-size:22px; margin-bottom:8px; }}
      .notice {{ background:rgba(79,182,198,0.12); border:1px solid rgba(79,182,198,0.4); padding:12px 16px; border-radius:6px; font-size:13px; margin-bottom:26px; line-height:1.5; }}
      .notice-guide {{ background:#181D2E; border-color:#333; margin-top:-12px; }}
      .notice-guide strong {{ display:block; margin-bottom:8px; font-size:13.5px; }}
      .notice-guide ul {{ margin:0 0 8px; padding-left:18px; }}
      .notice-guide li {{ margin-bottom:4px; }}
      label {{ display:block; margin-top:16px; font-size:13px; color:#B9C3D9; }}
      input, select {{ width:100%; padding:10px; margin-top:6px; border-radius:4px; border:1px solid #333; background:#181D2E; color:#fff; box-sizing:border-box; }}
      button {{ margin-top:26px; width:100%; padding:14px; background:#4C8DFF; border:none; border-radius:4px; color:#fff; font-size:15px; font-weight:600; cursor:pointer; }}
      button:hover {{ background:#6C93EF; }}
      button:disabled {{ opacity:0.6; cursor:not-allowed; }}
      #status {{ margin-top:12px; font-size:13px; color:#B9C3D9; min-height:18px; }}
    </style>
    </head>
    <body>
    <div id="login-gate" style="position:fixed;inset:0;background:#ffffff;z-index:9999;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:16px;font-family:sans-serif;">
  <h2 style="margin:0;">북다온 베타 서비스</h2>
  <p style="color:#555;">베타테스트 참여자 확인을 위해 구글 로그인이 필요합니다.</p>
  <button id="google-login-btn" style="padding:12px 28px;font-size:16px;border:none;border-radius:8px;background:#4285F4;color:white;cursor:pointer;">
    구글 계정으로 로그인
  </button>
</div>

<script src="https://www.gstatic.com/firebasejs/10.7.1/firebase-app-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.7.1/firebase-auth-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.7.1/firebase-firestore-compat.js"></script>
<script>
  const firebaseConfig = {{
    apiKey: "AIzaSyCSv6GO8mQOkZUpXTqj7X8IHQYM49npFM4",
    authDomain: "daonbooks-be849.firebaseapp.com",
    projectId: "daonbooks-be849"
  }};
  firebase.initializeApp(firebaseConfig);
  const db = firebase.firestore();

  function logEvent(eventType, toolName) {{
    const user = firebase.auth().currentUser;
    db.collection('events').add({{
      site: window.location.hostname,
      event_type: eventType,
      tool_name: toolName || null,
      user_email: user ? user.email : null,
      timestamp: firebase.firestore.FieldValue.serverTimestamp()
    }}).catch((err) => console.error('기록 실패', err));
  }}

  const gate = document.getElementById('login-gate');
  const loginBtn = document.getElementById('google-login-btn');
  const provider = new firebase.auth.GoogleAuthProvider();

  loginBtn.addEventListener('click', () => {{
    firebase.auth().signInWithPopup(provider).catch((err) => {{
      alert('로그인 실패: ' + err.message);
    }});
  }});

  firebase.auth().onAuthStateChanged((user) => {{
    if (user) {{
      gate.style.display = 'none';
      logEvent('visit', '이펍변환기');
    }} else {{
      gate.style.display = 'flex';
    }}
  }});
</script>
      <h1>다온 EPUB 변환기 (웹)</h1>
      <div class="notice">
  현재는 워드(.docx) 파일만 지원합니다.<br>
  제목·저자·챕터 분리는 자동으로 처리됩니다. 아래 칸은 비워되도 괜찮습니다.<br><br>
  ⚠️ 웹 데모 버전으로, 문서 구조에 따라 제목·챕터 인식이 정확하지 않을 수 있습니다.<br>
  변환 결과를 꼭 확인해주시고, 오류 발견 시 하단 홈페이지 제안함으로 알려주시면 감사하겠습니다.
</div>
      <div class="notice notice-guide">
        <strong>변환 전 워드 파일을 이렇게 써주세요</strong>
        <ul>
          <li>책 제목: 워드 상단 홈 탭 → 스타일에서 <b>제목</b> 선택</li>
          <li>부/장 제목: <b>제목 1</b> (예: "제1부"), <b>제목 2</b> (예: "1장")</li>
          <li>소제목: <b>제목 3</b> 이상 (예: "1. 에피소드")</li>
          <li>본문: <b>기본(표준)</b> 스타일 그대로 사용</li>
        </ul>
        스타일을 지정하지 않고 글자 크기·굵기만 바꾼 텍스트는 제목으로 인식되지 않아,
        본문과 구분 없이 하나로 합쳐질 수 있습니다.
      </div>
      <form id="convertForm">
        <label>원고 파일 (.docx)</label>
        <input type="file" name="file" accept=".docx" required>

        <label>제목 (비워두면 원고의 "Title" 스타일 문단을 자동으로 사용)</label>
        <input type="text" name="title">

        <label>저자 (비워두면 "저자 미상"으로 채워짐)</label>
        <input type="text" name="author">

        <label>출판사</label>
        <input type="text" name="publisher">

        <label>ISBN (선택)</label>
        <input type="text" name="isbn">

        <label>디자인 템플릿</label>
        <select name="template_key">{options_html}</select>

        <label>EPUB 버전</label>
        <select name="epub_version">
          <option value="3">EPUB 3.0</option>
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
            if (res.ok && typeof logEvent === 'function') logEvent('download', '이펍변환기');
            if (!res.ok) {{
              const err = await res.json().catch(() => ({{}}));
              throw new Error(err.detail || '변환 중 오류가 발생했습니다.');
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
