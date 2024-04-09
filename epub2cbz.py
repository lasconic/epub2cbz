import os
import re
from zipfile import ZipFile
from rich import print as rprint
from PIL import Image
import send2trash
from datetime import datetime

"""extract and rename images in spine order"""
btn_extract_images = 1
"""delete temp folders to recycle bin"""
btn_delete_temp = 1

"""create comicinfo.xml file"""
btn_comicinfo = 1
"""add title to comicinfo"""
btn_title = 1
"""add volume number to comicinfo"""
btn_volume_no = 1
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
"""add chapter info to comicinfo"""
btn_chapters = 1


def parse_metadata(epub_file, opf_path):
    author = ""
    title = ""
    language = ""
    publisher = ""
    date = ""
    pattern_title = r'<dc:title(.+?)?>(.+?)</dc:title>'
    pattern_author = r'<dc:creator(.+?)?>(.+?)</dc:creator>'
    pattern_language = r'<dc:language(.+?)?>(.+?)</dc:language>'
    pattern_publisher = r'<dc:publisher(.+?)?>(.+?)</dc:publisher>'
    pattern_date = r'<dc:date(.+?)?>(.+?)</dc:date>'
    
    with ZipFile(epub_file, 'r') as epub:
        opf_content = epub.read(opf_path).decode('utf-8')
        match_title = re.findall(pattern_title, opf_content)
        match_author = re.findall(pattern_author, opf_content)
        match_language = re.findall(pattern_language, opf_content)
        match_publisher = re.findall(pattern_publisher, opf_content)
        match_date = re.findall(pattern_date, opf_content)
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
    return author, title, language, publisher, date

def create_blank_image(dimension_x, dimension_y):
    image = Image.new("RGB", (dimension_x, dimension_y), color=(255, 255, 255))
    return image

def extract_images(epub_file, epub_filename, book_full):
    dimension_x = 0
    dimension_y = 0
    extension = ""
    with ZipFile(epub_file, 'r') as epub:
        os.makedirs(epub_filename, exist_ok=True)
        extension = book_full[0]['image'].rsplit(".")[1]
        for i, book in enumerate(book_full):
            if book['image'] and i > 0:
                with epub.open(book['image'], 'r') as zipimage:
                    image_dimension = Image.open(zipimage)
                    dimension_x, dimension_y = image_dimension.size
                break
        try:
            for book in book_full:
                if book['image']:
                    epub.extract(book['image'], epub_filename)
                    os.rename(os.path.join(epub_filename, book['image']), os.path.join(epub_filename, book['number'] + "." + book['image'].rsplit(".")[1]))
                else:
                    image = create_blank_image(dimension_x, dimension_y)
                    image.save(os.path.join(epub_filename, book['number'] + "." + extension))
        except Exception as e:
            rprint(f"[red]Warning: Folder for book {epub_filename} not empty. Delete or empty and try again.[/]")
    if btn_delete_temp:
        for root, dirs, _ in os.walk(epub_filename):
            for dir in dirs:
                try:
                    send2trash.send2trash(os.path.join(root, dir))
                    print(f"Info: Cleaned up temp folder \"{os.path.join(root, dir)}\" to recycle bin")
                except Exception as e:
                    rprint(f"Exception deleting: [red]{e}[/]")

def natural_keys(text):
    def convert(text):
        return float(text) if text.replace('.', '', 1).isdigit() else text.lower()
    return [ convert(c) for c in re.split('([0-9]+(?:\.[0-9]*)?)', text) ]

def parse_reading_direction(epub_file, opf_path):
    reading_direction = ""
    with ZipFile(epub_file, 'r') as epub:
        opf_content = epub.read(opf_path).decode('utf-8')
        pattern = r'page-progression-direction="(.+?)"'
        match = re.findall(pattern, opf_content)
        if match and match[0] == "rtl":
            reading_direction = "YesAndRightToLeft"
        else:
            reading_direction = "No"
        return reading_direction

def parse_opf_pages(epub_file, opf_path, page_ids):
    manifest = []
    pages = []
    images = []
    missing_images = []
    book_full = []
    is_between_tags = False
    pattern_id = r'id="(.+?)"'
    pattern_href = r'href="(.+?)"'

    with ZipFile(epub_file, 'r') as epub:
        with epub.open(opf_path) as opf_content:
            for line in opf_content:
                line = line.decode('utf-8').strip()
                if "<manifest" in line:
                    is_between_tags = True
                elif "</manifest" in line:
                    if is_between_tags:
                        manifest.append(line)
                    is_between_tags = False
                if is_between_tags:
                    manifest.append(line)

            for page_id in page_ids:
                for item in manifest:
                    match_id = re.findall(pattern_id, item)
                    if match_id and match_id == page_id:
                        match_href = re.search(pattern_href, item)
                        if match_href:
                            pages.append(match_href.group(1))
                            break
            for i, page in enumerate(pages):        
                image_path = find_image_path_in_file(epub, page)
                if image_path:
                    image_path = [path for path in epub.namelist() if remove_starting_dots(image_path) in path]
                    images.append(image_path[0])
                    book = {'page': page, 'number': str(i).zfill(len(str(len(pages)))), 'image': image_path[0]}
                else:
                    rprint(f"[yellow]Info: Image path for page '{page}' not found (page# {i})[/]")
                    missing_images.append(i)
                    book = {'page': page, 'number': str(i).zfill(len(str(len(pages)))), 'image': ""}
                book_full.append(book)
            if missing_images:
                rprint(f"[yellow]Info: missing pages: {len(missing_images)}[/]")
    return book_full

def parse_epub_opf(epub_file, opf_path):
    pages = []
    spine = []
    is_between_tags = False
    pattern = r'idref="(.*?)"'
    
    with ZipFile(epub_file, 'r') as epub:
        with epub.open(opf_path) as opf_content:
            for line in opf_content:
                line = line.decode('utf-8').strip()
                if "<spine" in line:
                    is_between_tags = True
                elif "</spine" in line:
                    if is_between_tags:
                        spine.append(line)
                    is_between_tags = False
                if is_between_tags:
                    spine.append(line)
        for line in spine:
            matches = re.findall(pattern, line)
            if matches:
                pages.append(matches)
    book_full = parse_opf_pages(epub_file, opf_path, pages)
    return book_full

def parse_epub_toc(epub_file, opf_path):
    chapters = []
    filenames = []
    image_count = 0

    with ZipFile(epub_file, 'r') as epub:
        toc_file = epub.namelist()
        for file in toc_file:
            extension = file.split('.')[-1].lower()
            if extension in ['jpg', 'jpeg', 'png']:
                filenames.append(file)
                image_count += 1

        ncx_path = get_ncx_file(epub_file, opf_path)
        toc_content = epub.read(ncx_path).decode('utf-8')
        
        pattern = r'<(^ncx:|(?!\/).*?)navPoint(.|\n)*?<(^ncx:|(?!\/).*?)navLabel>(.|\n)*?<(^ncx:|(?!\/).*?)text>(.*?)</(^ncx:|.*?)text>(.|\n)*?</(^ncx:|.*?)navLabel>(.|\n)*?<(^ncx:|(?!\/).*?)content src="(.+?)"/>'
        matches = re.findall(pattern, toc_content)
        for match in matches:
            chapter = {'title': match[5].strip(), 'page': match[11].strip()}          
            image_path = find_image_path_in_file(epub, match[11].rsplit("#", 1)[0])
            if image_path:
                image_path = [path for path in epub.namelist() if remove_starting_dots(image_path) in path]
                chapter['image'] = image_path[0]
            else:
                rprint(f"[yellow]Info: Image path for chapter '{match[5]}' not found[/]")
            chapters.append(chapter)

    filenames.sort(key=natural_keys)
    return chapters, filenames, image_count

def remove_starting_dots(path):
    if path.startswith("../") or path.startswith("./"):
        return path[2:]
    else:
        return path

def find_image_path_in_file(epub, filename):
    image_path = None
    filename = remove_starting_dots(filename)
    filename = [path for path in epub.namelist() if filename in path]
    if filename and filename[0] in epub.namelist():
        file_content = epub.read(filename[0]).decode('utf-8')
        image_path_patterns = [r'src="(.*?\.jpg|.*?\.jpeg|.*?\.png)"',r'xlink:href="(.*?\.jpg|.*?\.jpeg|.*?\.png)"']
        for pattern in image_path_patterns:
            image_match = re.search(pattern, file_content)
            if image_match:
                image_path = image_match.group(1)
                break

    else:
        rprint(f"[blue]file not found {filename}[/]")
    return image_path

def extract_version(folder_name):
  v_index = folder_name.rfind('v')
  if v_index != -1 and v_index + 1 < len(folder_name):
    try:
      volume_number = int(folder_name[v_index + 1:])
      return folder_name[:v_index].rstrip(), volume_number
    except ValueError:
      pass
  return folder_name.rstrip(), None

def write_chapters_to_txt(chapters, epub_filename, root_dir, reading_direction, book_full, metadata):
    folder_name, volume_number = extract_version(os.path.basename(epub_filename))
    text_path = os.path.join(root_dir, os.path.basename(epub_filename), "ComicInfo.xml")
    os.makedirs(epub_filename, exist_ok=True)
    bookmark = ""
    author, title, language, publisher, date = metadata[0]
    with open(text_path, "w", encoding="utf-8") as text_file:
        text_file.write('<?xml version=\'1.0\' encoding=\'utf-8\'?>\n')
        text_file.write('<ComicInfo xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">\n')
        text_file.write('  <Series>' + folder_name.replace("_", ":") + '</Series>\n')
        
        if title and btn_title:
            text_file.write('  <Title>' + title + '</Title>\n')
        if volume_number and btn_volume_no:
            text_file.write('  <Volume>' + str(volume_number) + '</Volume>\n')
        if publisher and btn_publisher:
            text_file.write('  <Publisher>' + publisher + '</Publisher>\n')
        if language and btn_language:
            text_file.write('  <LanguageISO>' + language + '</LanguageISO>\n')
        if author and btn_author:
            text_file.write('  <Writer>' + author + '</Writer>\n')
            
        text_file.write('  <Pages>\n')
        
        for i, book in enumerate(book_full):
            text_file.write('    <Page ')
            if i == 0:
                bookmark = "Cover"
            else:
                bookmark = ""
                for chapter in chapters:
                    if os.path.basename(chapter['page'].rsplit("#", 1)[0]) == os.path.basename(book['page']):
                        bookmark = chapter['title']

            if bookmark and btn_chapters:
                text_file.write(f"Bookmark=\"{bookmark}\" ")
            text_file.write(f"Image=\"{i}\" />\n")
        
        text_file.write('  </Pages>\n')
        if date and btn_date:
            date_full = datetime.strptime(date, "%Y-%m-%d")
            text_file.write('  <Day>' + str(date_full.day) + '</Day>\n')
            text_file.write('  <Month>' + str(date_full.month) + '</Month>\n')
            text_file.write('  <Year>' + str(date_full.year) + '</Year>\n')
        if btn_reading_dir:
            text_file.write(f"  <Manga>{reading_direction}</Manga>\n")
        text_file.write('</ComicInfo>')

def get_opf_file(epub_file):
    container = "META-INF/container.xml"
    with ZipFile(epub_file, 'r') as epub:
        container_content = epub.read(container).decode('utf-8')
        pattern = r'<rootfile full-path="(.+?)"'
        match = re.findall(pattern, container_content)
        return match[0]

def get_ncx_file(epub_file, opf_path):
    with ZipFile(epub_file, 'r') as epub:
        opf_content = epub.read(opf_path).decode('utf-8')
        if "/" in opf_path:
            opf_path = opf_path.rsplit("/", 1)[0] + "/"
        else:
            opf_path = ""

        pattern = r'<item (.*?)media-type="application/x-dtbncx\+xml"(.*?)/>'

        try:
            matches = re.findall(pattern, opf_content)
            pattern_inner = r'href="(.*?)"'
            i = 0
            for item in matches:
                if item[i] != "" and i <= len(matches):
                    match = re.search(pattern_inner, item[i])
                    if match:
                        link = match.group(1)
                        opf_path = opf_path + link
                elif item[i+1] != "" and i <= len(matches):
                    match = re.search(pattern_inner, item[i+1])
                    if match:
                        link = match.group(1)
                        opf_path = opf_path + link     
            return opf_path
        except IndexError:

            rprint(f"[red]couldnt find .ncx file in .opf of {epub_file}[/]")
            return
        rprint(f"[red]returned outside of try with {matches}[/]")
        return

def process_epub(epub_file, root_dir, opf_path):
    chapters, filenames, image_count = parse_epub_toc(epub_file, opf_path)
    epub_filename = epub_file.split(os.path.sep)[-1].rsplit(".")[0]
    book_full = parse_epub_opf(epub_file, opf_path)
    #
    if btn_extract_images:
        extract_images(epub_file, epub_filename, book_full)
    #
    reading_direction = parse_reading_direction(epub_file, opf_path)
    #
    metadata = [parse_metadata(epub_file, opf_path)]
    #
    if btn_comicinfo:
        write_chapters_to_txt(chapters, epub_filename, root_dir, reading_direction, book_full, metadata)
    #
    rprint(f"[green]Processed '{os.path.basename(epub_filename)}'[/]")

def read_mangalist(filename):
    with open(filename, "r", encoding="utf-8") as file:
        mangalist = file.read().splitlines()
    for i, item in enumerate(mangalist):
        if item.endswith(".cbz"):
            mangalist[i] = item[:-4] + ".epub"
    return mangalist

def main():
    root_dir = os.getcwd()
    mangalist = read_mangalist("mangalist.txt")
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.endswith(".epub") and filename in mangalist:
                epub_path = os.path.join(dirpath, filename)
                opf_path = get_opf_file(epub_path)
                process_epub(epub_path, root_dir, opf_path)

if __name__ == "__main__":
    main()
