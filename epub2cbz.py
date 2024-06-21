import os
import re
from bs4 import BeautifulSoup
from zipfile import ZipFile
from rich import print as rprint
from PIL import Image
from shutil import rmtree
from datetime import datetime


"""extract and rename images in spine order"""
btn_extract_images = 1
"""delete temp folders"""
btn_delete_temp = 1

"""create comicinfo.xml file"""
btn_comicinfo = 1
"""add title to comicinfo"""
btn_title = 1
"""add volume number to comicinfo"""
btn_volume_no = 1
"""add series to comicinfo"""
btn_series = 1
"""add author to comicinfo"""
btn_author = 1
"""add publisher to comicinfo"""
btn_publisher = 1
"""add language info to comicinfo"""
btn_language = 1
"""add release date to comicinfo"""
btn_date = 1
"""add reading direction to comicinfo"""
btn_reading_dir = 1
"""add description to comicinfo"""
btn_description = 1
"""add chapter info to comicinfo"""
btn_chapters = 1


def get_toc_file(epub, new_chapters, book_full, number):
    new_toc = []
    toc = []
    pattern_href = r'<a href="(.*?)">'

    i = 0
    while i < 4:
        try:
            alt_toc_file = [
                path
                for path in epub.namelist()
                if remove_starting_dots(book_full[int(number) + i]["page"]) in path
            ]
            with epub.open(alt_toc_file[0]) as toc_content:
                for line in toc_content:
                    line = line.decode("utf-8").strip()
                    if '<a href="' in line:
                        toc.append(line)
            for entry in toc:
                match = re.search(pattern_href, entry)
                if match:
                    for book in book_full:
                        if os.path.basename(book["page"]) == os.path.basename(match[1]):
                            new_chapters.append(
                                {
                                    "title": f"Page {book['number'] + 1}",
                                    "page": remove_starting_dots(match[1]),
                                    "image": "",
                                }
                            )
                            break
        except Exception as e:
            print(e)
            break
        i += 1

    seen = set()
    for new_chapter in new_chapters:
        if os.path.basename(new_chapter["page"].rsplit("#", 1)[0]) not in seen:
            seen.add(os.path.basename(new_chapter["page"].rsplit("#", 1)[0]))
            new_toc.append(new_chapter.copy())
    return new_toc


def parse_alternative_toc(epub_file, opf_path, chapters, book_full):
    is_between_tags = False
    guide = []
    new_toc = []
    pattern_toc = r'<reference type="toc"(.*?)href="(.*?)"/>'
    alt_toc_file = ""

    with ZipFile(epub_file, "r") as epub:
        with epub.open(opf_path) as opf_content:
            for line in opf_content:
                line = line.decode("utf-8").strip()
                if "<guide" in line:
                    is_between_tags = True
                elif "</guide" in line:
                    if is_between_tags:
                        guide.append(line)
                        break
                    is_between_tags = False
                if is_between_tags:
                    guide.append(line)
        for line in guide:
            matches = re.findall(pattern_toc, line)
            if matches:
                alt_toc_file = matches[0][1]
                break
        if alt_toc_file:
            for book in book_full:
                if os.path.basename(book["page"]) == os.path.basename(alt_toc_file):
                    new_toc = get_toc_file(
                        epub, chapters, book_full, int(book["number"])
                    )
                    break
        else:
            new_toc = chapters
    return new_toc


def parse_alternative_cover(epub_file, opf_path, book_full):
    pattern_cover = r'<meta name="cover"(.+?)content="(.+?)"(.+?)?/>'

    with ZipFile(epub_file, "r") as epub:
        opf_content = epub.read(opf_path).decode("utf-8")
        match_cover = re.findall(pattern_cover, opf_content)
        if match_cover:
            if (
                match_cover[0][1].lower().endswith(".jpg")
                or match_cover[0][1].lower().endswith(".jpeg")
                or match_cover[0][1].lower().endswith(".png")
            ):
                filename = match_cover[0][1]
                filename = remove_starting_dots(filename)
                filename = [path for path in epub.namelist() if filename in path]
                if filename and filename[0] in epub.namelist():
                    if filename[0] != book_full[0]["image"]:
                        book_full.insert(
                            0, {"page": "Cover", "number": "0", "image": filename[0]}
                        )
                        for i, book in enumerate(book_full):
                            if i > 0:
                                book_full[i]["number"] = i + 1
    return book_full


def parse_metadata(epub_file, opf_path):
    author = ""
    title = ""
    language = ""
    publisher = ""
    date = ""
    description = ""
    pattern_title = r"<dc:title(.+?)?>(.+?)</dc:title>"
    pattern_author = r"<dc:creator(.+?)?>(.+?)</dc:creator>"
    pattern_language = r"<dc:language(.+?)?>(.+?)</dc:language>"
    pattern_publisher = r"<dc:publisher(.+?)?>(.+?)</dc:publisher>"
    pattern_date = r"<dc:date(.+?)?>(.+?)</dc:date>"
    pattern_description = r"<dc:description(.+?)?>(.+?)</dc:description>"

    with ZipFile(epub_file, "r") as epub:
        opf_content = epub.read(opf_path).decode("utf-8")
        print(opf_content)
        match_title = re.findall(pattern_title, opf_content)
        match_author = re.findall(pattern_author, opf_content)
        match_language = re.findall(pattern_language, opf_content)
        match_publisher = re.findall(pattern_publisher, opf_content)
        match_date = re.findall(pattern_date, opf_content)
        match_description = re.findall(
            pattern_description, opf_content, flags=re.DOTALL
        )

        if match_title:
            title = match_title[0][1]
        if match_author:
            author = match_author[0][1]
        if match_language:
            language = match_language[0][1]
        if match_publisher:
            publisher = match_publisher[0][1]
        if match_date:
            date = match_date[0][1]
            if len(date) > 10:
                date = date[:10]
        if match_description:
            pattern_space = r"\s+"
            description = re.sub(pattern_space, " ", match_description[0][1])
    return author, title, language, publisher, date, description


def create_blank_image(dimension_x, dimension_y):
    image = Image.new("RGB", (dimension_x, dimension_y), color=(255, 255, 255))
    return image


def extract_images(epub_file, epub_filename, book_full):
    dimension_x = 0
    dimension_y = 0
    extension = ""

    with ZipFile(epub_file, "r") as epub:
        os.makedirs(epub_filename, exist_ok=True)
        extension = book_full[0]["image"].rsplit(".")[1]
        for i, book in enumerate(book_full):
            if book["image"] and i > 0:
                with epub.open(book["image"], "r") as zipimage:
                    image_dimension = Image.open(zipimage)
                    dimension_x, dimension_y = image_dimension.size
                break
        try:
            for i, book in enumerate(book_full):
                if book["image"]:
                    epub.extract(book["image"], epub_filename)
                    os.rename(
                        os.path.join(epub_filename, book["image"]),
                        os.path.join(
                            epub_filename,
                            str(i).zfill(len(str(len(book_full))))
                            + "."
                            + book["image"].rsplit(".")[1],
                        ),
                    )
                else:
                    image = create_blank_image(dimension_x, dimension_y)
                    image.save(
                        os.path.join(
                            epub_filename,
                            str(i).zfill(len(str(len(book_full)))) + "." + extension,
                        )
                    )
        except Exception:
            rprint(
                f"[red]Warning: Folder for book '{epub_filename}' not empty. Delete or empty and try again.[/]"
            )
    if btn_delete_temp:
        for root, dirs, _ in os.walk(epub_filename):
            for dir in dirs:
                try:
                    rmtree(os.path.join(root, dir))
                    print(f"Info: Cleaned up temp folder '{os.path.join(root, dir)}'")
                except Exception as e:
                    rprint(f"Exception deleting: [red]{e}[/]")


def parse_reading_direction(epub_file, opf_path):
    reading_direction = ""

    with ZipFile(epub_file, "r") as epub:
        opf_content = epub.read(opf_path).decode("utf-8")
        pattern = r'page-progression-direction="(.+?)"'
        match = re.findall(pattern, opf_content)
        if match and match[0] == "rtl":
            reading_direction = "YesAndRightToLeft"
        else:
            reading_direction = "No"
        return reading_direction


def parse_opf_pages(epub_file, opf_path, page_ids):
    pages = []
    images = []
    book_full = []
    match_ids = []
    cover_found = False
    css_path = get_css_file(epub_file, opf_path)

    with ZipFile(epub_file, "r") as epub:
        print(page_ids)
        with epub.open(opf_path) as opf_content:
            soup = BeautifulSoup(opf_content, features="xml")
            items = soup.find("manifest").find_all("item")
            for page_id in page_ids:
                for item in items:
                    match_id = item.attrs["id"]
                    print("matchid", item)
                    if match_id and match_id == page_id:
                        print("MATCH")
                        match_ids.append(match_id)
                        match_href = item.attrs["href"]
                        if match_href:
                            pages.append(match_href)
                            break
            print(pages)
            for i, page in enumerate(pages):
                image_path = find_image_path_in_file(epub, page)
                if image_path:
                    image_path = [
                        path
                        for path in epub.namelist()
                        if remove_starting_dots(image_path) in path
                    ]
                    images.append(image_path[0])
                    book = {"page": page, "number": i, "image": image_path[0]}
                    cover_found = True
                elif match_ids[i] and not cover_found:
                    css_image = find_image_path_in_css(epub, css_path, match_ids[i][0])
                    css_image = [
                        path
                        for path in epub.namelist()
                        if remove_starting_dots(css_image) in path
                    ]
                    book = {"page": page, "number": i, "image": css_image[0]}
                else:
                    book = {"page": page, "number": i, "image": ""}
                book_full.append(book)
    return book_full


def parse_epub_opf(epub_file, opf_path):
    pages = []
    with ZipFile(epub_file, "r") as epub:
        with epub.open(opf_path) as opf_content:
            soup = BeautifulSoup(opf_content, features="xml")
            itemrefs = soup.find("spine").find_all("itemref")

            for itemref in itemrefs:
                pages.append(itemref.attrs["idref"])

    book_full = parse_opf_pages(epub_file, opf_path, pages)
    return book_full


def parse_epub_toc(epub_file, opf_path):
    chapters = []

    with ZipFile(epub_file, "r") as epub:
        ncx_path = get_ncx_file(epub_file, opf_path)

        if ncx_path.endswith(".ncx"):
            toc_content = epub.read(ncx_path).decode("utf-8")
            soup = BeautifulSoup(toc_content, features="xml")
            nav_points = soup.find_all("navPoint")
            for nav_point in nav_points:
                title = nav_point.navLabel.text.strip()
                page = nav_point.content.attrs["src"].rsplit("#", 1)[0].strip()
                chapter = {"title": title, "page": page}
                image_path = find_image_path_in_file(epub, page)
                if image_path:
                    image_path = [
                        path
                        for path in epub.namelist()
                        if remove_starting_dots(image_path) in path
                    ]
                    chapter["image"] = image_path[0]
                chapters.append(chapter)
        elif ncx_path.endswith(".xhtml"):
            nav = []
            is_between_tags = False
            pattern = r'<a href="(.*?)">(.*?)</a>'
            with epub.open(ncx_path) as toc_content:
                for line in toc_content:
                    line = line.decode("utf-8").strip()
                    if 'epub:type="toc"' in line:
                        is_between_tags = True
                    elif "</nav" in line:
                        if is_between_tags:
                            nav.append(line)
                            break
                        is_between_tags = False
                    if is_between_tags:
                        nav.append(line)
                for line in nav:
                    matches = re.findall(pattern, line)
                    for match in matches:
                        chapter = {
                            "title": match[1].strip(),
                            "page": match[0].rsplit("#", 1)[0].strip(),
                        }
                        image_path = find_image_path_in_file(
                            epub, match[0].rsplit("#", 1)[0]
                        )
                        if image_path:
                            image_path = [
                                path
                                for path in epub.namelist()
                                if remove_starting_dots(image_path) in path
                            ]
                            chapter["image"] = image_path[0]
                        chapters.append(chapter)
    for i, chapter in enumerate(chapters):
        if i < (len(chapters) - 1) and chapter["page"] == chapters[i + 1]["page"]:
            chapter["title"] = chapter["title"] + " - " + chapters[i + 1]["title"]
            del chapters[i + 1]
    return chapters


def remove_starting_dots(path):
    if path.startswith("../") or path.startswith("./"):
        return path[2:]
    else:
        return path


def find_image_path_in_css(epub, filename, page_id):
    image_path = None
    filename = remove_starting_dots(filename)
    filename = [path for path in epub.namelist() if filename in path]

    if filename and filename[0] in epub.namelist():
        file_content = epub.read(filename[0]).decode("utf-8")
        image_path_patterns_css = [
            rf"#{page_id}(.*?)background-image:(.*?)url\(\"(.*?\.jpg|.*?\.jpeg|.*?\.png)\"\)"
        ]
        pattern_inner = r"url\(\"(.*?\.jpg|.*?\.jpeg|.*?\.png)\"\)"
        for pattern in image_path_patterns_css:
            image_match = re.search(pattern, file_content)
            if image_match and 'url("' in image_match[0]:
                match = re.search(pattern_inner, image_match[0])
                if match:
                    image_path = match.group(1)
                    break
    else:
        rprint(f"[blue]file not found {filename}[/]")
    return image_path


def find_image_path_in_file(epub, filename):
    image_path = None
    filename = remove_starting_dots(filename)
    filename = [path for path in epub.namelist() if filename in path]

    if (
        filename
        and filename[0] in epub.namelist()
        and (filename[0].endswith(".xhtml") or filename[0].endswith(".html"))
    ):
        file_content = epub.read(filename[0]).decode("utf-8")
        image_path_patterns = [
            r'src="(.*?\.jpg|.*?\.jpeg|.*?\.png)"',
            r'xlink:href="(.*?\.jpg|.*?\.jpeg|.*?\.png)"',
        ]
        for pattern in image_path_patterns:
            image_match = re.search(pattern, file_content)
            if image_match:
                image_path = image_match.group(1)
                break
    elif (
        filename
        and filename[0] in epub.namelist()
        and (
            filename[0].endswith(".jpg")
            or filename[0].endswith(".jpeg")
            or filename[0].endswith(".png")
        )
    ):
        image_path = filename[0]
    else:
        rprint(f"[blue]file not found {filename}[/]")
    return image_path


def extract_version(folder_name):
    v_index = folder_name.rfind("v")
    if v_index != -1 and v_index + 1 < len(folder_name):
        try:
            volume_number = int(folder_name[v_index + 1 :])
            return folder_name[:v_index].rstrip(), volume_number
        except ValueError:
            pass
    return folder_name.rstrip(), None


def convert_to_date(date_str):
    """
    Converts a string representing a date to day/month/year format.

    Args:
        date_str: The string to convert. Can be in YYYY format or full date format.

    Returns:
        A string in day/month/year format (e.g., "15/08/2023") or None if the format is invalid.
    """
    try:
        if len(date_str) == 4 and date_str.isdigit():
            return datetime.strptime(date_str, "%Y")
        else:
            return datetime.strptime(date_str, "%Y-%m-%d")
    except ValueError:
        return None


def write_chapters_to_txt(
    chapters, epub_filename, root_dir, reading_direction, book_full, metadata
):
    folder_name, volume_number = extract_version(os.path.basename(epub_filename))
    text_path = os.path.join(root_dir, os.path.basename(epub_filename), "ComicInfo.xml")
    os.makedirs(epub_filename, exist_ok=True)
    bookmark = ""
    author, title, language, publisher, date, description = metadata[0]

    with open(text_path, "w", encoding="utf-8") as text_file:
        text_file.write("<?xml version='1.0' encoding='utf-8'?>\n")
        text_file.write(
            '<ComicInfo xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">\n'
        )
        if btn_series:
            text_file.write(
                "  <Series>" + folder_name.replace("_", ":") + "</Series>\n"
            )
        if title and btn_title:
            text_file.write("  <Title>" + title.strip() + "</Title>\n")
        if volume_number and btn_volume_no:
            text_file.write("  <Volume>" + str(volume_number) + "</Volume>\n")
        if publisher and btn_publisher:
            text_file.write("  <Publisher>" + publisher.strip() + "</Publisher>\n")
        if language and btn_language:
            text_file.write("  <LanguageISO>" + language.strip() + "</LanguageISO>\n")
        if author and btn_author:
            text_file.write("  <Writer>" + author.strip() + "</Writer>\n")

        text_file.write("  <Pages>\n")

        for i, book in enumerate(book_full):
            text_file.write("    <Page ")
            if i == 0:
                bookmark = "Cover"
            else:
                bookmark = ""
                for chapter in chapters:
                    if os.path.basename(
                        chapter["page"].rsplit("#", 1)[0]
                    ) == os.path.basename(book["page"]):
                        bookmark = chapter["title"]

            if bookmark and btn_chapters:
                text_file.write(f'Bookmark="{bookmark.strip()}" ')
            text_file.write(f'Image="{i}" />\n')

        text_file.write("  </Pages>\n")
        if date and btn_date:
            date_full = convert_to_date(date)
            text_file.write("  <Day>" + str(date_full.day) + "</Day>\n")
            text_file.write("  <Month>" + str(date_full.month) + "</Month>\n")
            text_file.write("  <Year>" + str(date_full.year) + "</Year>\n")
        if btn_reading_dir:
            text_file.write(f"  <Manga>{reading_direction}</Manga>\n")
        if description and btn_description:
            text_file.write(f"  <Summary>{description.strip()}</Summary>\n")
        text_file.write("</ComicInfo>")


def get_opf_file(epub_file):
    container = "META-INF/container.xml"
    with ZipFile(epub_file, "r") as epub:
        container_content = epub.read(container).decode("utf-8")
        pattern = r'<rootfile full-path="(.+?)"'
        match = re.findall(pattern, container_content)
        return match[0]


def get_css_file(epub_file, opf_path):
    with ZipFile(epub_file, "r") as epub:
        opf_content = epub.read(opf_path).decode("utf-8")
        if "/" in opf_path:
            opf_path = opf_path.rsplit("/", 1)[0] + "/"
        else:
            opf_path = ""

        pattern = r'<item (.*?)media-type="text/css"(.*?)/>'

        try:
            matches = re.findall(pattern, opf_content)
            pattern_inner = r'href="(.*?)"'
            for item in matches:
                if 'href="' in item[0]:
                    match = re.search(pattern_inner, item[0])
                    if match:
                        link = match.group(1)
                        opf_path = opf_path + link
                elif 'href="' in item[1]:
                    match = re.search(pattern_inner, item[1])
                    if match:
                        link = match.group(1)
                        opf_path = opf_path + link
            return opf_path
        except IndexError:
            rprint(f"[red]couldnt find .css file in .opf of {epub_file}[/]")
            return
        rprint(f"[red]returned outside of try with {matches}[/]")
        return


def get_ncx_file(epub_file, opf_path):
    with ZipFile(epub_file, "r") as epub:
        opf_content = epub.read(opf_path).decode("utf-8")
        if "/" in opf_path:
            opf_path = opf_path.rsplit("/", 1)[0] + "/"
        else:
            opf_path = ""

        pattern = r'<item (.*?)media-type="application/x-dtbncx\+xml"(.*?)/>'
        pattern_nav = r'<item (.*?)?properties="nav"(.*?)?/>'

        try:
            matches = re.findall(pattern, opf_content)
            matches_nav = re.findall(pattern_nav, opf_content)
            pattern_inner = r'href="(.*?)"'

            if matches:
                for item in matches:
                    if 'href="' in item[0]:
                        match = re.search(pattern_inner, item[0])
                        if match:
                            link = match.group(1)
                            opf_path = opf_path + link
                            break
                    elif 'href="' in item[1]:
                        match = re.search(pattern_inner, item[1])
                        if match:
                            link = match.group(1)
                            opf_path = opf_path + link
                            break
                return opf_path
            elif matches_nav:
                for item in matches_nav:
                    if 'href="' in item[0]:
                        match = re.search(pattern_inner, item[0])
                        if match:
                            link = match.group(1)
                            opf_path = opf_path + link
                            break
                    elif 'href="' in item[1]:
                        match = re.search(pattern_inner, item[1])
                        if match:
                            link = match.group(1)
                            opf_path = opf_path + link
                            break
                return opf_path
        except IndexError:
            rprint(f"[red]couldnt find .ncx file in .opf of {epub_file}[/]")
            return
        rprint(f"[red]returned outside of try with {matches}[/]")
        return


def process_epub(epub_file, root_dir, opf_path):
    chapters = parse_epub_toc(epub_file, opf_path)
    epub_filename = epub_file.split(os.path.sep)[-1].rsplit(".")[0]
    book_full = parse_epub_opf(epub_file, opf_path)
    #
    metadata = [parse_metadata(epub_file, opf_path)]
    #
    book_full = parse_alternative_cover(epub_file, opf_path, book_full)
    #
    chapters = parse_alternative_toc(epub_file, opf_path, chapters, book_full)
    #
    if (
        os.path.basename(chapters[0]["page"].rsplit("#", 1)[0])
        == os.path.basename(book_full[1]["page"])
        and chapters[0]["title"] == "Cover"
    ):
        del book_full[1]
        rprint(
            f"[yellow]Info: Removed duplicate cover for '{os.path.basename(epub_filename)}'[/]"
        )
    #
    if btn_extract_images:
        extract_images(epub_file, epub_filename, book_full)
    #
    reading_direction = parse_reading_direction(epub_file, opf_path)
    #
    if btn_comicinfo:
        write_chapters_to_txt(
            chapters, epub_filename, root_dir, reading_direction, book_full, metadata
        )
    #
    rprint(f"[green]Processed '{os.path.basename(epub_filename)}'[/]")


def main():
    root_dir = os.getcwd()
    epub_paths = []

    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.endswith(".epub"):
                epub_path = os.path.join(dirpath, filename)
                opf_path = get_opf_file(epub_path)
                epub_paths.append((epub_path, root_dir, opf_path))

    # with Pool() as pool:
    # pool.starmap(process_epub, epub_paths)
    for pathi in epub_paths:
        print(pathi)
        process_epub(pathi[0], pathi[1], pathi[2])


if __name__ == "__main__":
    main()
