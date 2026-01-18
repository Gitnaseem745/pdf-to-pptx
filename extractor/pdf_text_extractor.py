from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer

def extract_text_by_page(pdf_path: str):
    """
    Returns:
        [
            "Text from page 1",
            "Text from page 2",
            ...
        ]
    """
    pages_text = []

    for page_layout in extract_pages(pdf_path):
        page_text = []
        for element in page_layout:
            if isinstance(element, LTTextContainer):
                page_text.append(element.get_text())

        cleaned_text = "\n".join(page_text).strip()
        pages_text.append(cleaned_text)

    return pages_text


if __name__ == "__main__":
    pdf_path = "../input/slides.pdf"
    pages = extract_text_by_page(pdf_path)

    for i, text in enumerate(pages):
        print(f"\n--- Page {i+1} ---\n{text}")
