from pathlib import Path
from docx import Document


def load_document(file_path):
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"找不到文件: {file_path}")
    return Document(str(path))


def read_docx_paragraphs(document):
    paragraphs = []
    for idx, paragraph in enumerate(document.paragraphs):
        text = paragraph.text.strip()
        if text:
            paragraphs.append({
                "index": idx,
                "text": text,
                "paragraph": paragraph
            })
    return paragraphs
