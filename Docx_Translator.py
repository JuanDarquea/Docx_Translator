# docx_Translator
import os
import re

from googletrans import Translator # to translate text
import asyncio
from pathlib import Path
from tkinter import Tk
from tkinter import filedialog as fd
from dotenv import load_dotenv # to load environment variables from .env file
from zipfile import BadZipFile # to handle invalid .docx files
from docx import Document   # to read and write .docx files
from docx.table import Table
from docx.text.paragraph import Paragraph

# unused imports
import deepl # to translate text with contextual accuracy
import time

# load the .env file from the same directory as this script
try:
    env_path = Path(__file__).parent / "Project_env.env"
except NameError:
    env_path = Path.cwd() / "Project_env.env"

load_dotenv(env_path) # load environment variables from .env file

# define rate limit parameters
rate_limit_minute = 450  # translator plan rate limit: 450 requests per minute
delay_between_requests = 60/rate_limit_minute  # calculate delay between requests in seconds
max_retries = 5 # maximum number of retries for failed requests

# create google transalator obeject
translator = Translator()

NON_TRANSLATABLE_PATTERNS = [
    re.compile(r"https?://[^\s]+|www\.[^\s]+", re.IGNORECASE),  # URLs
    re.compile(r"\b[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}\b"),  # emails
    re.compile(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b"),  # dates
    re.compile(r"[$€£¥]\s?\d{1,3}(?:,\d{3})*(?:\.\d+)?|\b\d{1,3}(?:,\d{3})*(?:\.\d+)?\s?[$€£¥]\b"),  # currency
    re.compile(r"\b\d+(?:[.,]\d+)?%"),  # percentages
    re.compile(r"\b\d{1,3}(?:,\d{3})*(?:\.\d+)?\b"),  # numbers
]

# define a variable to select the file to be translated
def select_docx_file():
    """Open a file dialog and return the selected filepath to translate"""
    # create hidden root window
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    # create child window(file dialog)
    file_path = fd.askopenfilename( # assign a variable to open a dialog and select the file
        title="Choose a file to translate",
        filetypes=[
            ("Word Documents", "*.docx"), # shows only .docx files
            ("All FIles", "*.*") # show every type of file
        ],
        # use the environment variable to set the initial directory and a spare default value
        initialdir=os.getenv("lin_test_docs_dir", os.getenv("app_tools_dir"))
    )

    # destroy the root dialog window
    root.destroy()

    # Return None instead of empty string for better logic
    return file_path if file_path else None

def file_validation(file_path):
    """Validate if a file path was selected"""
    if file_path is None: # when no file is selected
        return
    elif not file_path.lower().endswith(".docx"): # validate file extension
        print("\nError!! The file selected must be a '.docx' file.")
        return
    else:
        try: # validate file existence
            # When file is selected
            print(f"\nFile selected to translate: {file_path}",
                    f"\nFile path: {os.path.dirname(file_path)}",
                    f"\nFile name: {os.path.basename(file_path)}",
                    f"\nFile size: {os.path.getsize(file_path)} KB", sep="")
            return True
        except FileExistsError: # file does not exist
            print(f"\nError!! The file {file_path} selected does not exist.")
            return
        except BadZipFile as e: # file is not a valid .docx file
            print(f"\nError!! The file selected is not a valid .docx file: {e}")
            return
        except PermissionError as e: # file access permission error
            print(f"\nError!! Permission denied to access the file: {e}")
            return
        except Exception as e: # other errors
            print(f"\nError validating the file: {e}")
        return

def read_document(file_path):
    """Read the .docx file and return it as an object"""
    selected_document = Document(file_path)
    doc = [] # create empty list to store paragraphs

    # extract all text paragraphs from the document
    try:
        print("\nReading document content...")
        for paragraph in selected_document.paragraphs:
#            if paragraph.text.strip() != "": # skip empty paragraphs
            doc.append(paragraph.text)
        # print document content read success message
        print("\nDocument content read successfully.")
    except Exception as e: # handle errors while reading document
        print(f"Error reading the document: {e}")
        return

    # print paragraph count
    total = len(selected_document.paragraphs)
    print(f"Paragraph count: {total}\n")

    # print all paragraphs with a paragraph index as test
#    for index, paragraph in enumerate(doc):
#        if paragraph.strip() != "": # skip empty paragraphs
#            print(index + 1,
#                  paragraph, f"Selected style: {selected_document.paragraphs[index].style.name}",
#                  f"Alignment: {selected_document.paragraphs[index].alignment}",
#                  f"Font: {selected_document.paragraphs[index].runs[0].font.name if selected_document.paragraphs[index].runs else "Default Font"}",
#                 f"{len(selected_document.paragraphs[index].text)} characters",
#                  sep = " - ")
#            print(f"P{index + 1}: {paragraph}") # alternative print format
#        else:
#            print(index + 1,"<Empty paragraph>", sep=" - ")
#            index - 1 # do not count empty paragraphs
    return selected_document if selected_document else None

def iter_document_blocks(document):
    """Yield body items in document order as ('paragraph'|'table', object)."""
    for child in document.element.body.iterchildren():
        if child.tag.endswith('}p'):
            yield "paragraph", Paragraph(child, document)
        elif child.tag.endswith('}tbl'):
            yield "table", Table(child, document)

def iter_container_blocks(container):
    """Yield paragraph/table blocks in order for document, header, or footer containers."""
    if hasattr(container, "element") and hasattr(container.element, "body"):
        parent_elm = container.element.body
    else:
        parent_elm = container._element

    for child in parent_elm.iterchildren():
        if child.tag.endswith('}p'):
            yield "paragraph", Paragraph(child, container)
        elif child.tag.endswith('}tbl'):
            yield "table", Table(child, container)

async def translate_text_preserving_whitespace(text, target_lang):
    """Translate text while preserving whitespace-only chunks unchanged."""
    if text is None or text == "":
        return text
    if text.strip() == "":
        return text

    leading_match = re.match(r"^\s*", text)
    trailing_match = re.search(r"\s*$", text)
    leading_ws = leading_match.group(0) if leading_match else ""
    trailing_ws = trailing_match.group(0) if trailing_match else ""
    core_end = len(text) - len(trailing_ws)
    core_text = text[len(leading_ws):core_end]

    if core_text == "":
        return text

    protected_text, replacements = protect_non_translatables(core_text)
    result = await translator.translate(protected_text, dest=target_lang)
    await asyncio.sleep(delay_between_requests)
    translated_core = restore_non_translatables(result.text, replacements)
    return f"{leading_ws}{translated_core}{trailing_ws}"

def protect_non_translatables(text):
    """
    Replace non-translatable fragments with placeholders and return:
    (protected_text, replacements_dict)
    """
    protected = text
    replacements = {}
    token_index = 0

    for pattern in NON_TRANSLATABLE_PATTERNS:
        def _replace(match):
            nonlocal token_index
            token = f"__NTX_{token_index}__"
            token_index += 1
            replacements[token] = match.group(0)
            return token

        protected = pattern.sub(_replace, protected)

    return protected, replacements

def restore_non_translatables(text, replacements):
    """Restore protected placeholders to their original values."""
    restored = text
    for token, original in replacements.items():
        restored = restored.replace(token, original)
    return restored

def fix_run_boundary_punctuation_from_texts(source_texts, target_runs):
    """
    Remove punctuation introduced at run boundaries when it did not exist in source.
    This avoids artifacts like: "ella. objetivos" when source boundary had no dot.
    """
    limit = min(len(source_texts), len(target_runs))

    for i in range(1, limit):
        src_prev = source_texts[i - 1].rstrip()
        src_next = source_texts[i].lstrip()
        tgt_prev = target_runs[i - 1]

        if not src_prev or not src_next:
            continue

        src_prev_ended_with_dot = src_prev.endswith(".")
        src_next_starts_word = src_next[0].isalpha() or src_next[0].isdigit()

        tgt_prev_text = tgt_prev.text
        tgt_prev_stripped = tgt_prev_text.rstrip()
        tgt_prev_has_dot = tgt_prev_stripped.endswith(".")

        if (not src_prev_ended_with_dot) and src_next_starts_word and tgt_prev_has_dot:
            # Remove only one trailing period and keep any trailing whitespace.
            m = re.search(r"\.(\s*)$", tgt_prev_text)
            if m:
                tgt_prev.text = tgt_prev_text[:m.start()] + m.group(1)

def is_page_number_run(run):
    """
    Detect runs that belong to page-number fields (PAGE/NUMPAGES/SECTIONPAGES).
    These runs must not be translated.
    """
    page_field_markers = ("PAGE", "NUMPAGES", "SECTIONPAGES")
    word_ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    # 1) <w:fldSimple w:instr="...PAGE...">
    parent = run._r.getparent()
    while parent is not None:
        if parent.tag.endswith('}fldSimple'):
            instr = (parent.get(f"{word_ns}instr") or "").upper()
            if any(marker in instr for marker in page_field_markers):
                return True
        parent = parent.getparent()

    # 2) Complex field instructions inside run descendants (<w:instrText>)
    for node in run._r.iter():
        if node.tag.endswith('}instrText'):
            instr_text = (node.text or "").upper()
            if any(marker in instr_text for marker in page_field_markers):
                return True

    return False

def has_field_code_nodes(run):
    """True when run contains field code XML nodes that must not be modified."""
    for node in run._r.iter():
        if node.tag.endswith('}fldChar') or node.tag.endswith('}instrText'):
            return True
    return False

async def translate_paragraph_in_place(paragraph, target_lang):
    """Translate one paragraph in place, preserving list/paragraph/run formatting."""
    if paragraph.runs:
        source_run_texts = [run.text for run in paragraph.runs]
        in_field = False

        for idx, run in enumerate(paragraph.runs):
            # Never rewrite field-code runs; it can break PAGE/NUMPAGES rendering.
            if has_field_code_nodes(run):
                for node in run._r.iter():
                    if node.tag.endswith('}fldChar'):
                        fld_char_type = node.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType")
                        if fld_char_type == "begin":
                            in_field = True
                        elif fld_char_type == "end":
                            in_field = False
                continue

            # Keep all field-result runs untouched while inside a field.
            if in_field or is_page_number_run(run):
                continue
            source_text = source_run_texts[idx]
            run.text = await translate_text_preserving_whitespace(source_text, target_lang)

        fix_run_boundary_punctuation_from_texts(source_run_texts, paragraph.runs)
        return

    if paragraph.text and paragraph.text.strip():
        paragraph.text = await translate_text_preserving_whitespace(paragraph.text, target_lang)

async def translate_table_in_place(table, target_lang):
    """Translate text in each table cell paragraph in place."""
    print(f"Translating table: {len(table.rows)} row(s), {len(table.columns)} col(s)")
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                await translate_paragraph_in_place(paragraph, target_lang)

async def translate_header_footer_in_place(document, target_lang):
    """Translate all unique headers and footers in place, preserving page-number fields."""
    visited = set()

    for section in document.sections:
        containers = (
            ("header_default", section.header),
            ("header_first", section.first_page_header),
            ("header_even", section.even_page_header),
            ("footer_default", section.footer),
            ("footer_first", section.first_page_footer),
            ("footer_even", section.even_page_footer),
        )

        for container_type, container in containers:
            key = (container_type, id(container._element))
            if key in visited:
                continue
            visited.add(key)

            for block_type, block in iter_container_blocks(container):
                if block_type == "paragraph":
                    await translate_paragraph_in_place(block, target_lang)
                elif block_type == "table":
                    await translate_table_in_place(block, target_lang)

async def translated_doc_creation(file_path, selected_document, target_lang="ES"):
    """
    Function to create an output file and write translated text into the new file.
    Returns:
    - output file path if successful or None if failed
    - prints error messages if any error occurs
    - creates a new .docx file with translated text and preserved formatting
    """
    print("\nCreating output file...")

    # Translate in place on source structure to preserve numbering/list definitions.
    trans_file = selected_document

    # Select original file path
    try:
        base_name = os.path.basename(file_path)
        name_no_ext, ext = os.path.splitext(base_name)
        print()
        print(f"Chosen base name: {base_name}")
    except Exception as e:
        print(f"Error reading the file: {e}")
        return

    try:
        output_dir = (os.getenv("lin_translated_docs_dir")
        or os.getenv("translated_docs_dir")
        or os.path.dirname(file_path))

        # check if output directory is defined
        if not output_dir:
            print("Error!! No output directory defined in environment variables.")
            return

        output_name = f"{name_no_ext}_{target_lang.upper()}{ext}"

        output_path = os.path.join(output_dir,
                                output_name)
        print(f"Chosen file dir: {output_dir}")
        print(f"New file name: {output_name}")
        print(f"New file path: {output_path}\n")
    except Exception as e:
        print(f"Error getting the output directory path: {e}")
        return

    try:
        for block_type, block in iter_document_blocks(trans_file):
            if block_type == "paragraph":
                await translate_paragraph_in_place(block, target_lang)
            elif block_type == "table":
                await translate_table_in_place(block, target_lang)

        await translate_header_footer_in_place(trans_file, target_lang)

        trans_file.save(output_path)
        print(f"Translated document saved successfully at: {output_path}")
        return output_path

    except Exception as e:
            print(f"Error reading translated document: {e}")
            return

def inspect_tables(selected_document):
    """"""
    if not selected_document.tables:
        print(f'\nNo tables found in the document.')
        return

    print(f'\nFound {len(selected_document.tables)} tables.\n')

 #   for t_idx, table in enumerate(selected_document.tables):
 #       t_idx += 1
 #       print(f'Table {t_idx}:',
 #             f'    Rows: {len(table.rows)}')

 #       for r_idx, row in enumerate(table.rows):
 #           print(f'        Row {r_idx + 1}: Cells = {len(row.cells)}')

 #           for c_idx, cell in enumerate(row.cells):
 #               print(f'            Cell {c_idx + 1}: Paragraphs = {len(cell.paragraphs)}')

  #              for p_idx, paragraph in enumerate(cell.paragraphs):
  #                  text = paragraph.text.strip()
  #                  print(f'                Paragraph {p_idx + 1}: "{text}"')

  #      print('-'*20, f'Table {t_idx} end', '-'*20) # spacing between tables

# the following instances are for debugging and learning purposes only
def debug_print_runs(selected_document):
    """
    Debug function to print all runs in a paragraph.
    Returns:
    - prints all runs in each paragraph with their formatting details
    """
    print()
    print("-"*20, "Run Debbug Start", "-"*20)

    for p_idx, paragraph in enumerate(selected_document.paragraphs, start=1):
        print(f"Paragraph {p_idx}:")
        print(f"Full text: {repr(paragraph.text)}")
        print(f"Run count: {len(paragraph.runs)}")

        for r_idx, run in enumerate(paragraph.runs, start=1):
            print(f"    Run {r_idx}: {repr(run.text)} | "
                  f"bold={run.bold}, italic={run.italic}, "
                  f"underline={run.underline}, "
                  f"font={run.font.name}, size={run.font.size}, "
                  )

        print("-"*50)
    print("-"*20, "Run Debbug End", "-"*20)

# the following instance is for debugging and learning purposes only
def debug_paragraph_runs(run_info):
    """
    Debug function to print runs of all paragraph
    Returns:
    - prints all runs in each paragraph with their formatting details
    """
    print()
    print("-"*20, "Paragraphs Run Info Start", "-"*20)

    if not run_info:
        print("No paragraphs found.")
        return

    for p_idx, paragraph in enumerate(run_info, start=1):
        print(f"Paragraph {p_idx} -")
        print(f"Alignment: {paragraph['alignment']}")
        print(f"Style: {paragraph['style']}")
        print(f"Run count: {len(paragraph['runs'])}")

        for r_idx, run in enumerate(paragraph["runs"], start=1):
            print(f"Run {r_idx}:")
            for key, value in run.items():
                print(f"    {key}: {value}")
            print()
#        break  # only the first paragraph
    print("-"*20, "All Paragraphs Run Info End", "-"*20)

async def main():
    print("Select a .docx file to translate...")

    # get file selected from user
    chosen_file = select_docx_file()

    # close if cancelled and no file selected
    if not chosen_file:
        print("\nUser closed the window before selecting a file.",
              "\nGoodbye!")
        return

    # validate file
    if file_validation(chosen_file) is None:
        return

    # read document
    selected_document = read_document(chosen_file)

    # inspect for tables
    inspect_tables(selected_document)

    # translate document and create output preserving structure and formatting
    output_path = await translated_doc_creation(chosen_file,
                                                selected_document,
                                                target_lang="ES")
    if output_path:
        os.system(f"open {output_path}")

if __name__=="__main__":
    asyncio.run(main())
