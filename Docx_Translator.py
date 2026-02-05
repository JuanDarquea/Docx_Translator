# docx_Translator
import os

from googletrans import Translator # to translate text
import asyncio
from pathlib import Path
from tkinter import Tk
from tkinter import filedialog as fd
from dotenv import load_dotenv # to load environment variables from .env file
from zipfile import BadZipFile # to handle invalid .docx files
from docx import Document   # to read and write .docx files
from datetime import datetime

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
    root.destroy

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
    for index, paragraph in enumerate(doc):
        if paragraph.strip() != "": # skip empty paragraphs
            print(index + 1,
                  paragraph, f"Selected style: {selected_document.paragraphs[index].style.name}",
                  f"Alignment: {selected_document.paragraphs[index].alignment}",
                  f"Font: {selected_document.paragraphs[index].runs[0].font.name if selected_document.paragraphs[index].runs else "Default Font"}",
                  f"{len(selected_document.paragraphs[index].text)} characters",
                  sep = " - ")
#            print(f"P{index + 1}: {paragraph}") # alternative print format
        else:
            print(index + 1,"<Empty paragraph>", sep=" - ")
#            index - 1 # do not count empty paragraphs
    return selected_document if selected_document else None

async def translate_text_googletrans(file_path, selected_document, target_lang="ES"):
    """
    Translate text using googletrans module.
    Returns:

    """
    file_text = selected_document
    if file_text is None:
        print("The file selected does not exist or could not be read.")
        return

    print(f"\nTranslating document to {target_lang} using googletrans...")
    translated_file = []
    partial_file = []
    async with Translator() as translator:

        try:
            for idx, paragraph in enumerate(file_text.paragraphs, start=1):
                style_name = paragraph.style.name if paragraph.style else None

                # Simulate an error for testing purposes
#                if idx == 7:
#                    raise Exception(f"Simulated error for testing purposes.")

                if paragraph.text.strip() == "": # skip empty paragraphs
                    translated_file.append("") # keep empty paragraphs
                    #print("<Empty paragraph> --> <Empty paragraph>")
                else:
                        result = await translator.translate(
                            paragraph.text,
                            dest=target_lang
                            )
                        #print(paragraph.text, result.text, style_name, sep="--> ")
                        translated_file.append(result.text)
                        await asyncio.sleep(delay_between_requests)  # to avoid hitting rate limits
        except Exception as e:
                print(f"\nError! Could not translate the paragraph {idx}, error type: {e}",
                      f"\nOriginal paragraph: {paragraph.text}")
                partial_file = translated_file.copy()

                # Create error log .txt file
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                output_dir = os.getenv("lin_error_logs_dir") or os.getenv("translated_docs_dir") or os.path.dirname(file_path)
                if not output_dir:
                    print("\nError!! No output directory defined in environment variables.")
                    return
                base_name = os.path.basename(file_path)
                name_no_ext, ext = os.path.splitext(base_name)
                error_log_file = f"{name_no_ext}_PARTIAL_ERROR_LOG.txt"
                error_log = os.path.join(output_dir,
                                    error_log_file)

                with open(error_log, "a+", encoding="utf-8") as log:
                    try:
                        line = f"\n[{timestamp}] - Paragraph #{idx} = {paragraph.text}- ERROR = {e}"
                        log.write(line)
                        print(f"\nError log saved successfully at: {error_log}")
                        return partial_file
                    except Exception as e:
                        print(f"\nError writing error log file: {e}")
                        return

                return partial_file

    print()
    print("\nThe file output is the following list:",
          f"\n{translated_file if translated_file else partial_file}")
    print()
    return translated_file if translated_file else partial_file

def build_run_spans(paragraph):
    """
    Build run spans for paragraph.
    Returns:
    - spans (list of dicts, one per run)
    - each dict contains text and formatting info
    """
    spans = []
    total_words = 0

    for run in paragraph.runs:
        words = run.text.split()
        if not words:
            continue

        count = len(words)
        total_words += count

        spans.append({
            "word_count": len(words),
            "bold": run.bold,
            "italic": run.italic,
            "underline": run.underline,
            "strikethrough": run.font.strike,
            "font_name": run.font.name,
            "font_size": run.font.size,
            "font_color": run.font.color.rgb,
        })
    return spans, total_words

def split_text_into_spans(text, n):
    """
    Split text into n spans.
    Returns:
    - list of n spans
    - each span is a substring of the original text
    """

    if n <= 1:
        return [text]

    words = text.split()
    avg = max(1,
              len(words) // n)
    spans = []
    idx = 0

    for i in range(n):
        if i == n - 1:
            spans.append(' '.join(words[idx:]))
        else:
            spans.append(' '.join(words[idx:idx + avg]))
        idx += avg

    return spans

def rebuild_run_from_spans(paragraph, translated_text, spans, original_total_words):
    """
    Rebuild runs in a paragraph based on translated text and original spans.
    Returns:
    - paragraph with rebuilt runs
    - each run has formatting from the corresponding original span
    """
    for run in paragraph.runs:
        run.text = "" # clear existing runs safelly

    translated_words = translated_text.split()
    total_words = len(translated_words)
    #print('Document total quantity of words =', total_words)

    if not translated_words or not  spans:
        paragraph.add_run(translated_text)
        return

    # number of spans that have content
    active_spans = len(spans)
    #print('Number of spans that have content =', active_spans)

    # compute ideal word count per span
    ideal_counts = [span["word_count"] / original_total_words * total_words
                    for span in spans]
    #print('Ideal span count =', ideal_counts)

    # allocate words safely
    allocations = []
    remaining_words = total_words
    remaining_spans = active_spans

    for i, ideal in enumerate(ideal_counts):
        remaining_spans -= 1

        if remaining_spans == 0:
            take = remaining_words
        else:
            take = max(1, int(round(ideal)))
            take = min(take, remaining_words - remaining_spans)

        allocations.append(take)
        remaining_words -= take

    # build runs
    index = 0
    for span, take in zip(spans, allocations):
        chunk = translated_words[index:index + take]
        index += take

        if not chunk:
            continue

        run = paragraph.add_run(" ".join(chunk) + " ")

        run.bold = span["bold"]
        run.italic = span["italic"]
        run.underline = span["underline"]
        run.font.strike = span["strikethrough"]
        run.font.name = span["font_name"]
        run.font.size = span["font_size"]
        if span["font_color"]:
            run.font.color.rgb = span["font_color"]

    # leftover words (translation longer than original)
    #if word_index < len(words):
    #    paragraph.add_run(" ".join(words[word_index:]))

def translated_doc_creation(file_path, translated_file, selected_document):
    """
    Function to create an output file and write translated text into the new file.
    Returns:
    - output file path if successful or None if failed
    - prints error messages if any error occurs
    - creates a new .docx file with translated text and preserved formatting
    """
    print("\nCreating output file...")

    # Create output file
    trans_file = Document()

    # Select original file path
    try:
        base_name = os.path.basename(file_path)
        name_no_ext, ext = os.path.splitext(base_name)
        print()
        print(f"Chosen base name: {base_name}")
    except Exception as e:
        print(f"Error reading the file: {e}")
        return

    # Assign output file path
    is_partial = len(translated_file) != len(selected_document.paragraphs)
    try:
        output_dir = (os.getenv("lin_translated_docs_dir")
        or os.getenv("translated_docs_dir")
        or os.path.dirname(file_path))

        # check if output directory is defined
        if not output_dir:
            print("Error!! No output directory defined in environment variables.")
            return

        # create output file name based on translation completeness
        if is_partial:
            timestamp = datetime.now().strftime("%y%m%d_%H%M")
            output_name = f"{name_no_ext}_PARTIAL_ES_{timestamp}{ext}"
        else:
            output_name = f"{name_no_ext}_ES{ext}"

        output_path = os.path.join(output_dir,
                                output_name)
        print(f"Chosen file dir: {output_dir}")
        print(f"New file name: {output_name}")
        print(f"New file path: {output_path}\n")
    except Exception as e:
        print(f"Error getting the output directory path: {e}")
        return

    # Write translated text into the new document
    try:
        for idx, text in enumerate(translated_file): # iterate over translated paragraphs
            orig_p = selected_document.paragraphs[idx] # get original paragraph

            p = trans_file.add_paragraph() # add new paragraph to the translated document
            # paragraph level style
            p.style = orig_p.style
            p.alignment = orig_p.alignment

            spans, total_words = build_run_spans(orig_p) # get original run formatting
            rebuild_run_from_spans(p, text, spans, total_words) # rebuild run with original formatting

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

    for t_idx, table in enumerate(selected_document.tables):
        t_idx += 1
        print(f'Table {t_idx}:',
              f'    Rows: {len(table.rows)}')

        for r_idx, row in enumerate(table.rows):
            print(f'        Row {r_idx + 1}: Cells = {len(row.cells)}')

            for c_idx, cell in enumerate(row.cells):
                print(f'            Cell {c_idx + 1}: Paragraphs = {len(cell.paragraphs)}')

                for p_idx, paragraph in enumerate(cell.paragraphs):
                    text = paragraph.text.strip()
                    print(f'                Paragraph {p_idx + 1}: "{text}"')

        print('-'*20, f'Table {t_idx} end', '-'*20) # spacing between tables

async def translate_table(table, target_lang="ES"):
    async with Translator() as translator:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    original_text = paragraph.text.strip()
                    if not original_text:
                        continue

                    translated_text = await translator.translate(original_text, target_lang)
                    spans, total_words = build_run_spans(paragraph)

                    # clear paragraph
                    for run in paragraph.runs:
                        run.text = ""

                    rebuild_run_from_spans(
                        paragraph,
                        translated_text.text,
                        spans,
                        total_words
                        )

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

def main():
    """
    Main function.
    Orchestrates the document translation process.
    1. Prompts user to select a .docx file
    2. Validates the selected file
    3. Reads the document content
    4. Gets paragraph styles information
    5. Translates the document text
    6. Creates a new document with the translated text
    7. Saves the translated document to the specified directory
    8. Handles errors and logs them if translation fails
    9. Prints debug information if needed
    10. Exits gracefully if user cancels file selection
    11. Uses asynchronous translation to improve performance
    12. Preserves original formatting in the translated document
    13. Supports environment variables for configuration
    14. Provides informative console output throughout the process
    15. Implements rate limitting to avoid hitting API limits
    """
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

    # translate tables if there is any
    if selected_document.tables:
        print('\nTranslating tables...')
        for table in selected_document.tables:
            asyncio.run(translate_table(table))

    # temporary table file inspection
    temp_path = os.path.join(
        os.getenv("lin_translated_docs_dir"),
        'TEMP_TABLE_INSPECTION.docx'
    )
    selected_document.save(temp_path)
    print(f'\nTemporary table inspecition file saved at: {temp_path}')

    # debug print runs information
#    debug_print_runs(selected_document)

    # debug print paragraph runs information
#    run_info = paragraphs_style_info(selected_document)
#    debug_paragraph_runs(run_info)

    # translate document and save to translated files directory
    translated_file = asyncio.run(
        translate_text_googletrans(chosen_file,
                                   selected_document,
                                   target_lang="ES")
    )

    # call function to create a new document with the translated text
    translated_doc_creation(chosen_file,
                            translated_file,
                            selected_document
                            )

if __name__=="__main__":
    main()
