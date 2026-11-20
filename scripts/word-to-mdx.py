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
    parser.add_argument(
        "--hero-image",
        help="Optional hero image filename located beside the source .docx file",
    )
    return parser.parse_args()


def escape_markdown(text):
    return re.sub(r"([\\`_{}\[\]])", r"\\\1", text)


def merge_adjacent_bold(text):
    """把 **A****B** 合并成 **AB**。"""
    pattern = re.compile(r"\*\*([^*]+)\*\*\*\*([^*]+)\*\*")
    previous = None
    while previous != text:
        previous = text
        text = pattern.sub(r"**\1\2**", text)
    return text


def format_text(text, run):
    if not text:
        return text

    # 1. 先把前后空白拆出来（含半角/全角空格、tab）
    stripped = text.strip()
    if not stripped:
        return text

    leading = text[: len(text) - len(text.lstrip())]
    trailing = text[len(text.rstrip()) :]

    # 2. 只对正文做转义和加格式标记
    body = escape_markdown(stripped)

    # 3. 先处理字体颜色
    color = None

    if run.font.color and run.font.color.rgb:
        rgb = str(run.font.color.rgb).upper()

        if rgb != "000000":
            color = f"#{rgb}"

    if color:
        body = f'<span style={{{{ color: "{color}" }}}}>{body}</span>'

    # 4. 再处理下划线、删除线、粗体和斜体
    if run.underline:
        body = f"<u>{body}</u>"

    if getattr(run.font, "strike", False):
        body = f"<del>{body}</del>"

    if run.bold and run.italic:
        body = f"***{body}***"
    elif run.bold:
        body = f"**{body}**"
    elif run.italic:
        body = f"*{body}*"

    # 5. 空白原样拼回标记外面
    return leading + body + trailing


def register_image(blip, paragraph, image_dir, image_map):
    """提取圖片並回傳可供後續排版使用的中繼資料。"""
    relationship_id = blip.get(qn("r:embed"))
    if not relationship_id:
        return None

    if relationship_id not in image_map:
        image_part = paragraph.part.related_parts[relationship_id]
        extension = Path(str(image_part.partname)).suffix.lower() or ".bin"
        filename = f"image-{len(image_map) + 1:02d}{extension}"
        (image_dir / filename).write_bytes(image_part.blob)
        image_map[relationship_id] = filename

    filename = image_map[relationship_id]
    return filename


def get_anchor_metadata(drawing, filename):
    """讀取 Word 浮動圖片的尺寸及定位資訊。"""
    anchor = drawing.find(".//" + qn("wp:anchor"))
    if anchor is None:
        return None

    extent = anchor.find(qn("wp:extent"))
    width = int(extent.get("cx", "0")) if extent is not None else 0
    height = int(extent.get("cy", "0")) if extent is not None else 0

    def read_position(tag):
        node = anchor.find(qn(tag))
        if node is None:
            return None, None
        if len(node):
            child = node[0]
            if child.tag == qn("wp:posOffset"):
                return node.get("relativeFrom"), int(child.text or "0")
            if child.tag == qn("wp:align"):
                return node.get("relativeFrom"), child.text
        return node.get("relativeFrom"), None

    h_relative, h_value = read_position("wp:positionH")
    v_relative, v_value = read_position("wp:positionV")

    return {
        "filename": filename,
        "width": width,
        "height": height,
        "h_relative": h_relative,
        "h_value": h_value,
        "v_relative": v_relative,
        "v_value": v_value,
    }


def image_markdown(filename):
    return f"\n\n![{filename}](./images/{filename})\n\n"


def convert_inline(element, paragraph, image_dir, image_map, floating_images=None):
    result = []

    if element.tag == qn("w:r"):
        run = Run(element, paragraph)

        for child in element:
            if child.tag == qn("w:t"):
                result.append(format_text(child.text or "", run))

            elif child.tag == qn("w:tab"):
                result.append("    ")

            elif child.tag in (qn("w:br"), qn("w:cr")):
                result.append("  \n")

            elif child.tag == qn("w:drawing"):
                anchor = child.find(".//" + qn("wp:anchor"))
                for blip in child.iter(qn("a:blip")):
                    filename = register_image(blip, paragraph, image_dir, image_map)
                    if not filename:
                        continue

                    if anchor is not None and floating_images is not None:
                        metadata = get_anchor_metadata(child, filename)
                        if metadata:
                            floating_images.append(metadata)
                    else:
                        result.append(image_markdown(filename))

    elif element.tag == qn("w:hyperlink"):
        relationship_id = element.get(qn("r:id"))
        text = "".join(node.text or "" for node in element.iter(qn("w:t")))

        if relationship_id:
            url = paragraph.part.rels[relationship_id].target_ref
            result.append(
                f'<a href="{url}" target="_blank" rel="noopener noreferrer">'
                f"{escape_markdown(text)}</a>"
            )
        else:
            result.append(escape_markdown(text))

    return "".join(result)


def render_gallery(images, columns):
    """將一組浮動圖片轉成 ImageGallery MDX。"""
    entries = []
    for item in sorted(images, key=lambda image: image["filename"]):
        filename = item["filename"]
        alias = Path(filename).stem.replace("-", "_")
        entries.append(f"    {{ src: {alias}, alt: {yaml_quote(filename)} }}")
    return (
        f"<ImageGallery columns={{{columns}}} images={{[\n"
        + ",\n".join(entries)
        + "\n]} />"
    )


def render_floating_image_groups(blocks, floating_blocks, page_width):
    """依照浮動圖片的相對位置及尺寸，辨識明確的並列圖片組。"""
    rendered = {}
    pending = []
    pending_markers = []
    block_positions = {block: index for index, block in enumerate(blocks)}

    def is_compact(item):
        return bool(item["width"] and item["width"] <= page_width * 0.55)

    def similar_size(items):
        if len(items) < 2:
            return True
        widths = [item["width"] for item in items]
        ratios = [item["width"] / item["height"] for item in items if item["height"]]
        return max(widths) <= min(widths) * 1.18 and (
            not ratios or max(ratios) <= min(ratios) * 1.18
        )

    def flush_pending_to_previous_marker():
        nonlocal pending, pending_markers
        if not pending or not pending_markers:
            pending = []
            pending_markers = []
            return
        target = pending_markers[-1]
        singles = "\n\n".join(
            image_markdown(item["filename"]).strip() for item in pending
        )
        previous = rendered.get(target, "")
        rendered[target] = "\n\n".join(part for part in (previous, singles) if part)
        pending = []
        pending_markers = []

    for marker in blocks:
        if marker not in floating_blocks:
            # 正文段落打斷了圖片組；不能把相隔一段文字的單圖錯誤合併。
            flush_pending_to_previous_marker()
            continue

        items = floating_blocks[marker]
        compact = [item for item in items if is_compact(item)]
        large = [item for item in items if not is_compact(item)]
        output_parts = []

        # 四張尺寸相近的圖片若同屬一個 Word 段落，通常是 2×2 圖片組。
        if len(compact) == 4 and similar_size(compact):
            if pending:
                for item in pending:
                    output_parts.append(image_markdown(item["filename"]).strip())
                pending = []
                pending_markers = []
            output_parts.append(render_gallery(compact, 2))
        elif (
            len(compact) == 3
            and not pending
            and similar_size(compact)
            and block_positions[marker] + 1 < len(blocks)
            and blocks[block_positions[marker] + 1] in floating_blocks
            and len(floating_blocks[blocks[block_positions[marker] + 1]]) == 2
            and all(
                is_compact(item)
                for item in floating_blocks[blocks[block_positions[marker] + 1]]
            )
            and similar_size(floating_blocks[blocks[block_positions[marker] + 1]])
        ):
            # 同段落三張小圖，且緊接段落另有兩張同尺寸圖片時，
            # 將前兩張組成雙欄，後三張組成三欄。
            output_parts.append(render_gallery(compact[:2], 2))
            pending = [compact[2]]
        else:
            # 同一水平位置的圖片視為一列；其餘圖片可與緊鄰段落合併。
            rows = []
            remaining = list(compact)
            while remaining:
                first = remaining.pop(0)
                row = [first]
                first_y = first.get("v_value")
                if isinstance(first_y, int):
                    matched = [
                        item
                        for item in remaining
                        if isinstance(item.get("v_value"), int)
                        and abs(item["v_value"] - first_y) <= 180000
                    ]
                    if matched:
                        row.extend(matched)
                        remaining = [item for item in remaining if item not in matched]
                rows.append(row)

            for row in rows:
                if len(row) >= 2 and similar_size(row):
                    combined = pending + row
                    if pending and len(combined) == 3 and similar_size(combined):
                        output_parts.append(render_gallery(combined, 3))
                        pending = []
                        pending_markers = []
                    else:
                        output_parts.append(render_gallery(row, min(len(row), 4)))
                else:
                    pending.extend(row)

            # 三張尺寸相近、跨相鄰段落的浮動圖片，視為一組三欄圖庫。
            if len(pending) == 3 and similar_size(pending):
                output_parts.append(render_gallery(pending, 3))
                pending = []
                pending_markers = []

            # 大圖代表前一組小圖已結束，先輸出尚未成組的圖片。
            if large and pending:
                output_parts.extend(
                    image_markdown(item["filename"]).strip() for item in pending
                )
                pending = []
                pending_markers = []

            output_parts.extend(
                image_markdown(item["filename"]).strip() for item in large
            )

        rendered[marker] = "\n\n".join(part for part in output_parts if part)
        if pending:
            pending_markers.append(marker)

    # 文件結尾仍未配成組的圖片，按普通單圖輸出。
    if pending:
        last_marker = pending_markers[-1] if pending_markers else None
        if last_marker:
            previous = rendered.get(last_marker, "")
            singles = "\n\n".join(
                image_markdown(item["filename"]).strip() for item in pending
            )
            rendered[last_marker] = "\n\n".join(
                part for part in (previous, singles) if part
            )

    result = []
    for block in blocks:
        if block in floating_blocks:
            value = rendered.get(block, "")
            if value:
                result.append(value)
        else:
            result.append(block)
    return result


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
                        filename = register_image(blip, paragraph, image_dir, image_map)
                        if filename:
                            run_parts.append(
                                f'<img src="./images/{filename}" alt="{filename}" />'
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

        custom_hero_image = None

        if args.hero_image:
            hero_source = input_path.parent / args.hero_image

            if not hero_source.is_file():
                raise FileNotFoundError(f"找不到指定的封面圖片：{hero_source}")

            custom_hero_image = hero_source.name
            shutil.copy2(hero_source, image_dir / custom_hero_image)

        print(f"已成功讀取：{input_path.name}")
        print(f"文章標題：{title}\n========== 開始轉換 ==========\n")

        image_map = {}
        output_blocks = []
        floating_blocks = {}
        floating_block_counter = 0
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
            paragraph_floating_images = []

            for child in element:
                parts.append(
                    convert_inline(
                        child,
                        paragraph,
                        image_dir,
                        image_map,
                        paragraph_floating_images,
                    )
                )

            text = "".join(parts).strip()
            text = merge_adjacent_bold(text)
            plain_text = paragraph.text.strip()

            if not text and not paragraph_floating_images:
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
                text = "\n".join(line.lstrip() for line in text.split("\n"))
                if text:
                    output_blocks.append(text)

            if paragraph_floating_images:
                flush_list()
                floating_block_counter += 1
                marker = f"__FLOATING_IMAGES_BLOCK_{floating_block_counter}__"
                floating_blocks[marker] = paragraph_floating_images
                output_blocks.append(marker)

        flush_list()  # 文件結束前，把最後的列表輸出

        page_width = (
            int(document.sections[0].page_width) if document.sections else 7772400
        )
        output_blocks = render_floating_image_groups(
            output_blocks, floating_blocks, page_width
        )

        hero_image = ""

        if custom_hero_image:
            hero_image = f"heroImage: ./images/{custom_hero_image}\n"
        elif image_map:
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

        gallery_images = []
        for marker_images in floating_blocks.values():
            for item in marker_images:
                # 僅為實際生成的 ImageGallery 圖片建立靜態 import。
                if (
                    item["filename"] in content
                    and f"src: {Path(item['filename']).stem.replace('-', '_')}"
                    in content
                ):
                    gallery_images.append(item["filename"])
        gallery_images = list(dict.fromkeys(gallery_images))
        if gallery_images:
            imports = [
                "import ImageGallery from '../../../components/ImageGallery.astro';"
            ]
            for filename in gallery_images:
                alias = Path(filename).stem.replace("-", "_")
                imports.append(f"import {alias} from './images/{filename}';")
            content = (
                frontmatter + "\n".join(imports) + "\n\n" + content[len(frontmatter) :]
            )

        # Detect whether the article requires MDX.
        mdx_components = (
            "Figure",
            "ImageGallery",
            "ImageCompare",
            "ImageLightbox",
            "HeaderLink",
            "span",
            "u",
            "del",
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
