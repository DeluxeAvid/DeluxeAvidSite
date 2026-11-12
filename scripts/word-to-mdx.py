from pathlib import Path
import argparse
import re
import shutil
from datetime import datetime
from pypinyin import lazy_pinyin

from docx import Document
from lxml import etree
from docx.oxml.ns import qn
from docx.text.run import Run
from docx.text.paragraph import Paragraph


def parse_args():
    parser = argparse.ArgumentParser(
        description="Convert a Word document into Markdown."
    )
    parser.add_argument("input", type=Path, help="Source .docx file")
    parser.add_argument("--date", required=True, help="Publication date (YYYY-MM-DD)")
    parser.add_argument("--description", required=True, help="Article description")
    return parser.parse_args()


def escape_markdown(text):
    return re.sub(r"([\\`*_{}\[\]])", r"\\\1", text)


def format_text(text, run):
    text = escape_markdown(text)

    # 新增：處理底線與刪除線
    if run.underline:
        text = f"<u>{text}</u>"
    if getattr(run.font, "strike", False):
        text = f"~~{text}~~"

    if run.bold and run.italic:
        return f"***{text}***"
    if run.bold:
        return f"**{text}**"
    if run.italic:
        return f"*{text}*"

    return text


def extract_image(blip, paragraph, image_dir, image_map):
    relationship_id = blip.get(qn("r:embed"))

    if not relationship_id:
        return ""

    if relationship_id not in image_map:
        image_part = paragraph.part.related_parts[relationship_id]
        extension = Path(str(image_part.partname)).suffix.lower() or ".bin"

        filename = f"image-{len(image_map) + 1:02d}{extension}"
        image_path = image_dir / filename
        image_path.write_bytes(image_part.blob)

        image_map[relationship_id] = filename
        # print(f"[圖片 {len(image_map)}] {filename}") # 建議註解掉，避免終端機洗版

    filename = image_map[relationship_id]
    return f"\n\n![{filename}](./images/{filename})\n\n"


def convert_inline(element, paragraph, image_dir, image_map):
    result = []

    if element.tag == qn("w:r"):
        run = Run(element, paragraph)

        for child in element:
            if child.tag == qn("w:t"):
                result.append(format_text(child.text or "", run))

            elif child.tag in (qn("w:tab"),):
                result.append("    ")

            elif child.tag in (qn("w:br"), qn("w:cr")):
                # 一般段落內的軟換行，建議轉為兩個空白+換行，或直接保留 \n
                result.append("  \n")

            elif child.tag == qn("w:drawing"):
                for blip in child.iter(qn("a:blip")):
                    result.append(extract_image(blip, paragraph, image_dir, image_map))

    elif element.tag == qn("w:hyperlink"):
        relationship_id = element.get(qn("r:id"))
        text = "".join(node.text or "" for node in element.iter(qn("w:t")))

        if relationship_id:
            url = paragraph.part.rels[relationship_id].target_ref
            result.append(f"[{escape_markdown(text)}]({url})")
        else:
            result.append(escape_markdown(text))

    return "".join(result)


def convert_table_paragraph(paragraph_element, paragraph, image_dir, image_map):
    parts = []

    for child in paragraph_element:
        if child.tag == qn("w:r"):
            run_parts = []
            run = Run(child, paragraph)

            for run_child in child:
                if run_child.tag in (qn("w:br"), qn("w:cr")):
                    run_parts.append("<br>")

                elif run_child.tag == qn("w:t"):
                    run_parts.append(format_text(run_child.text or "", run))

                elif run_child.tag == qn("w:tab"):
                    run_parts.append("    ")

                elif run_child.tag == qn("w:drawing"):
                    for blip in run_child.iter(qn("a:blip")):
                        run_parts.append(
                            extract_image(blip, paragraph, image_dir, image_map)
                        )

            parts.append("".join(run_parts))

        else:
            parts.append(convert_inline(child, paragraph, image_dir, image_map))

    return "".join(parts).strip()


def convert_table(table_element, document, image_dir, image_map):
    """轉換表格，並直接回傳字串（使用單個換行符連接），確保 Markdown 解析正確"""
    rows = []

    for row_element in table_element.findall(qn("w:tr")):
        cells = []

        for cell_element in row_element.findall(qn("w:tc")):
            cell_parts = []

            for paragraph_element in cell_element.findall(qn("w:p")):
                paragraph = Paragraph(paragraph_element, document)

                cell_parts.append(
                    convert_table_paragraph(
                        paragraph_element,
                        paragraph,
                        image_dir,
                        image_map,
                    )
                )

            cell_text = "<br>".join(cell_parts)
            cell_text = cell_text.replace("|", r"\|")
            cells.append(cell_text)

        rows.append(cells)

    if not rows:
        return ""

    column_count = max(len(row) for row in rows)

    for row in rows:
        row.extend([""] * (column_count - len(row)))

    output = []
    output.append("| " + " | ".join(rows[0]) + " |")
    output.append("| " + " | ".join(["---"] * column_count) + " |")

    for row in rows[1:]:
        output.append("| " + " | ".join(row) + " |")

    # 重點修改：直接回傳用 \n 連接的完整表格字串
    return "\n".join(output)


def get_heading_level(paragraph):
    style = paragraph.style

    while style is not None:
        match = re.match(r"Heading (\d+)", style.name)

        if match:
            return int(match.group(1))

        style = style.base_style

    return None


def is_quote_paragraph(paragraph):
    style = paragraph.style
    while style is not None:
        if style.name.lower() in ("quote", "intense quote"):
            return True
        style = style.base_style
    return False


def get_list_marker(paragraph, document):
    p_pr = paragraph._p.pPr
    num_pr = p_pr.numPr if p_pr is not None else None

    if num_pr is None or num_pr.find(qn("w:numId")) is None:
        style = paragraph.style
        while style is not None:
            style_p_pr = style._element.pPr
            if style_p_pr is not None and style_p_pr.numPr is not None:
                style_num_pr = style_p_pr.numPr
                if style_num_pr.find(qn("w:numId")) is not None:
                    num_pr = style_num_pr
                    break
            style = style.base_style

    if num_pr is None:
        return None

    num_id_element = num_pr.find(qn("w:numId"))
    ilvl_element = num_pr.find(qn("w:ilvl"))

    if num_id_element is None:
        return None

    num_id = num_id_element.get(qn("w:val"))
    ilvl = int(ilvl_element.get(qn("w:val"))) if ilvl_element is not None else 0

    numbering = document.part.numbering_part.element

    num = next(
        (
            item
            for item in numbering.findall(qn("w:num"))
            if item.get(qn("w:numId")) == num_id
        ),
        None,
    )

    if num is None:
        return None

    abstract_id = num.find(qn("w:abstractNumId"))
    if abstract_id is None:
        return None

    abstract_num = next(
        (
            item
            for item in numbering.findall(qn("w:abstractNum"))
            if item.get(qn("w:abstractNumId")) == abstract_id.get(qn("w:val"))
        ),
        None,
    )

    if abstract_num is None:
        return None

    level = next(
        (
            item
            for item in abstract_num.findall(qn("w:lvl"))
            if int(item.get(qn("w:ilvl"))) == ilvl
        ),
        None,
    )

    if level is None:
        return None

    num_fmt = level.find(qn("w:numFmt"))
    if num_fmt is None:
        return None

    return {
        "num_id": num_id,
        "ilvl": ilvl,
        "format": num_fmt.get(qn("w:val")),
        "indent": "    " * ilvl,
    }


def yaml_quote(value):
    return "'" + value.replace("'", "''") + "'"


def main():
    args = parse_args()

    if not args.description.strip():
        raise ValueError("文章描述不可為空白，請提供有效的 --description。")

    input_path = args.input

    if not input_path.exists():
        raise FileNotFoundError(f"找不到檔案：{input_path}")

    if input_path.suffix.lower() != ".docx":
        raise ValueError("目前只支援 .docx 格式。")

    try:
        datetime.strptime(args.date, "%Y-%m-%d")
    except ValueError:
        raise ValueError("日期無效，請使用正確的 YYYY-MM-DD 格式。")

    document = Document(input_path)

    title = input_path.stem.replace("_", " ")

    filename = input_path.stem.lower()
    filename = re.sub(r"[\s_]+", "-", filename)

    pinyin_name = "-".join(lazy_pinyin(filename))
    slug = re.sub(r"[^a-z0-9]+", "-", pinyin_name)
    slug = re.sub(r"-+", "-", slug).strip("-")

    if not slug:
        slug = f"article-{args.date}"

    article_dir = Path("src/content/blog") / slug
    image_dir = article_dir / "images"

    if article_dir.exists():
        raise FileExistsError(
            f"文章資料夾已存在，為避免覆寫或混入既有檔案而中止轉換：{article_dir}"
        )

    try:
        article_dir.mkdir(parents=True, exist_ok=True)
        image_dir.mkdir(parents=True, exist_ok=True)

        print(f"已成功讀取：{input_path.name}")
        print(f"文章標題：{title}\n========== 開始轉換 ==========\n")

        image_map = {}
        output_blocks = []  # 改為儲存區塊 (Block)
        numbering_counters = {}
        heading_num_ids = set()

        # 用於處理連續列表的變數
        current_list_block = []

        def flush_list():
            if current_list_block:
                output_blocks.append("\n".join(current_list_block))
                current_list_block.clear()

        for paragraph in document.paragraphs:
            if get_heading_level(paragraph):
                list_info = get_list_marker(paragraph, document)
                if list_info and list_info["format"] != "bullet":
                    heading_num_ids.add(list_info["num_id"])

        for element in document.element.body:
            if element.tag == qn("w:tbl"):
                flush_list()  # 遇到表格前，先把累積的列表輸出
                table_output = convert_table(element, document, image_dir, image_map)
                if table_output:
                    output_blocks.append(table_output)
                continue

            if element.tag != qn("w:p"):
                continue

            paragraph = Paragraph(element, document)
            parts = []

            for child in element:
                parts.append(convert_inline(child, paragraph, image_dir, image_map))

            text = "".join(parts).strip()
            plain_text = paragraph.text.strip()

            if not text:
                continue

            heading_level = get_heading_level(paragraph)
            list_info = get_list_marker(paragraph, document)
            list_number = None

            if list_info and list_info["format"] != "bullet":
                counter_key = (list_info["num_id"], list_info["ilvl"])
                numbering_counters[counter_key] = (
                    numbering_counters.get(counter_key, 0) + 1
                )
                list_number = numbering_counters[counter_key]

            is_heading_sequence = bool(
                list_info
                and list_info["num_id"] in heading_num_ids
                and list_info["format"] != "bullet"
            )

            heading_prefix = ""
            if is_heading_sequence:
                heading_prefix = f"{list_number}. "

            # 根據類型加入輸出
            if is_heading_sequence:
                flush_list()
                # 修正：保留原本的標題層級，不再強制變成 ##
                level = heading_level if heading_level else 2
                output_blocks.append(f"{'#' * level} {heading_prefix}{plain_text}")
            elif heading_level:
                flush_list()
                output_blocks.append(f"{'#' * heading_level} {plain_text}")
            elif is_quote_paragraph(paragraph):
                flush_list()
                output_blocks.append(f"> {text}")
            elif list_info:
                # 列表不直接加入 output_blocks，而是先暫存到 current_list_block
                if list_info["format"] == "bullet":
                    current_list_block.append(f"{list_info['indent']}- {text}")
                else:
                    current_list_block.append(
                        f"{list_info['indent']}{list_number}. {text}"
                    )
            else:
                flush_list()
                output_blocks.append(text)

        flush_list()  # 文件結束前，把最後的列表輸出

        hero_image = ""
        if image_map:
            first_image = next(iter(image_map.values()))
            hero_image = f"heroImage: ./images/{first_image}\n"

        frontmatter = (
            "---\n"
            f"title: {yaml_quote(title)}\n"
            f"description: {yaml_quote(args.description)}\n"
            f"pubDate: {args.date}\n"
            f"{hero_image}"
            "---\n\n"
        )

        content = frontmatter + "\n\n".join(output_blocks) + "\n"

        # Detect whether the article requires MDX.
        mdx_components = (
            "Figure",
            "ImageGallery",
            "ImageCompare",
            "ImageLightbox",
            "HeaderLink",
        )

        requires_mdx = any(
            re.search(
                rf"<{re.escape(component)}(?=[\s/>])",
                content,
            )
            for component in mdx_components
        )

        extension = ".mdx" if requires_mdx else ".md"
        output_path = article_dir / f"index{extension}"

        output_path.write_text(
            content,
            encoding="utf-8",
        )

        print("\n========== 轉換完成 ==========")
        print(f"文章：{output_path.resolve()}")
        print(f"圖片：{len(image_map)} 張")

    except Exception:
        shutil.rmtree(article_dir, ignore_errors=True)
        raise


if __name__ == "__main__":
    try:
        main()
    except FileNotFoundError as error:
        print("\n========== 轉換失敗 ==========")
        print(error)
        print("請確認檔案名稱及所在位置。")
    except FileExistsError as error:
        print("\n========== 轉換失敗 ==========")
        print(error)
    except ValueError as error:
        print("\n========== 轉換失敗 ==========")
        print(error)
    except Exception as error:
        print("\n========== 發生未預期錯誤 ==========")
        print(f"{type(error).__name__}: {error}")
        raise
