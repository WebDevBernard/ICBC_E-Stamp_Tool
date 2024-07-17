import os
import re
import time
import timeit
import warnings
from collections import namedtuple, defaultdict
from datetime import datetime
from pathlib import Path
import sys
import fitz
import pandas as pd

warnings.simplefilter("ignore")

# font size and font options

font_size = 9
font = "SpaceMono"
fonts = {
    "FiraMono": ["fimo", "fimbo"],
    "SpaceMono": ["spacemo", "spacembo"],
    "NotoSans": ["notos", "notosbo"],
    "Ubuntu": ["ubuntu", "ubuntubo"],
    "Cascadia": ["cascadia", "cascadiab"],
}

# <=========================================Coordinates and Locations=========================================>

# Names of all the keywords to search for
search_keywords = [
    "is_icbc",
    "transaction_timestamp",
    "agency_number",
    "licence_plate",
    "insured_name",
    "owner_name",
    "applicant_name",
    "customer_copy_pages",
    "validation_stamp",
    "time_of_validation",
    "top",
]

# Search parameters using either indexing, or coordinates, or regex
search_params = [
    "target_keyword",
    "first_index",
    "second_index",
    "target_coordinates",
    "append_duplicates",
    "join_list",
]

# This sets a default value of None if unused
KeywordSearch = namedtuple(
    "SearchKeywords", search_keywords, defaults=(None,) * len(search_keywords)
)
SearchParams = namedtuple(
    "SearchParams", search_params, defaults=(None,) * len(search_params)
)

# Keyword dictionary and their search params
keyword_dict = {
    "ICBC": KeywordSearch(
        # no keyword + coordinate search params
        is_icbc=SearchParams(
            target_coordinates=(
                409.97900390625,
                63.84881591796875,
                576.0,
                83.7454833984375,
            ),
        ),
        customer_copy_pages=SearchParams(
            target_coordinates=(
                498.43798828125,
                751.9528198242188,
                578.1806640625,
                769.977294921875,
            ),
        ),
        top=SearchParams(
            target_coordinates=(230.3990020751953, 36.0, 573.6226196289062, 48.2890625),
        ),
        # keyword + index search params
        transaction_timestamp=SearchParams(
            target_keyword="Transaction Timestamp", first_index=0, second_index=0
        ),
        agency_number=SearchParams(
            target_keyword="Agency Number", first_index=0, second_index=0
        ),
        insured_name=SearchParams(
            target_keyword="Name of Insured (surname followed by given name(s))",
            first_index=0,
            second_index=1,
        ),
        owner_name=SearchParams(target_keyword="Owner ", first_index=0, second_index=1),
        applicant_name=SearchParams(
            target_keyword="Applicant", first_index=0, second_index=1
        ),
        # regex keyword + index search params
        licence_plate=SearchParams(
            target_keyword=re.compile(r"(?<!Previous )\bLicence Plate Number\b"),
            first_index=0,
            second_index=0,
        ),
        # keyword + coordinates search params
        validation_stamp=SearchParams(
            target_keyword="NOT VALID UNLESS STAMPED BY",
            target_coordinates=(
                -4.247998046875011,
                13.768768310546875,
                1.5807250976562273,
                58.947509765625,
            ),
        ),
        time_of_validation=SearchParams(
            target_keyword="TIME OF VALIDATION",
            target_coordinates=(0.0, 10.354278564453125, 0.0, 40),
        ),
    )._asdict(),
}


# <=========================================Helper Functions=========================================>


# append words list used in search_for_keyword function
def append_word_to_dict(word_list, field_dict, append_duplicates):
    for words in word_list:
        word = words[4].strip().split("\n")
        if append_duplicates:
            field_dict.append(word)
        if word and word not in field_dict:
            field_dict.append(word)


# Clean and transforms list items
def dd(dict_items, field_dict, key, regex=None, strip_char=None):
    for items in dict_items[key]:
        for item in items:
            if regex:
                item = re.sub(re.compile(regex), "", item)
            if strip_char:
                item = item.rstrip(strip_char)
                field_dict["insured_name"] = item
            else:
                field_dict[key] = item


# used in Excel sheet Function
def ee(df, row, default):
    value = df.at[row, 1]
    return default if isinstance(value, float) else value


# makes all strings into a list
def ll(dictionary):
    for key, value in dictionary.items():
        if value and isinstance(value[0], str):
            dictionary[key] = [value]
    return dictionary


def format_transaction_timestamp(timestamp_str):
    year = int(timestamp_str[0:4])
    month = int(timestamp_str[4:6])
    day = int(timestamp_str[6:8])
    hour = int(timestamp_str[8:10])
    minute = int(timestamp_str[10:12])
    second = int(timestamp_str[12:14])
    datetime_obj = datetime(year, month, day, hour, minute, second)
    return datetime_obj


# used in check_if_matching_transaction_timestamp to find matching transaction timestamps
def find_matching_paths(target_filename, paths):
    matching_paths = [path for path in paths if path.stem.split()[0] == target_filename]
    return matching_paths


# adds a # next to identical file names
def unique_file_name(path):
    filename, extension = os.path.splitext(path)
    counter = 1
    while Path(path).is_file():
        path = filename + " (" + str(counter) + ")" + extension
        counter += 1
    return path


# https://stackoverflow.com/questions/3160699/python-progress-bar
def progressbar(it, prefix="", size=60, out=sys.stdout): # Python3.6+
    count = len(it)
    start = time.time() # time estimate start
    def show(j):
        x = int(size*j/count)
        # time estimate calculation and string
        remaining = ((time.time() - start) / j) * (count - j)
        mins, sec = divmod(remaining, 60) # limited to minutes
        time_str = f"{int(mins):02}:{sec:03.1f}"
        print(f"{prefix}[{u'█'*x}{('.'*(size-x))}] {j}/{count} Est wait {time_str}", end='\r', file=out, flush=True)
    show(0.1) # avoid div/0
    for i, item in enumerate(it):
        yield item
        show(i+1)
    print(flush=True, file=out)
# <=========================================Main Functions=========================================>


# This Excel sheet reads the user inputs
def get_excel_data():
    root_dir = Path(__file__).parent.parent
    excel_path = root_dir / "BM3KXR.xlsx"
    defaults = {
        "number_of_pdfs": 5,
        "agency_name": "",
        "toggle_timestamp": "Timestamp",
        "toggle_customer_copy": "No"
    }

    if not excel_path.exists():
        return defaults

    try:
        df_excel = pd.read_excel(excel_path, sheet_name=0, header=None)
        data = {
            "number_of_pdfs": defaults["number_of_pdfs"] if isinstance(ee(df_excel, 2, defaults["number_of_pdfs"]), str) else ee(df_excel, 2, defaults["number_of_pdfs"]),
            "agency_name": ee(df_excel, 4, defaults["agency_name"]),
            "toggle_timestamp": ee(df_excel, 6, defaults["toggle_timestamp"]),
            "toggle_customer_copy": ee(df_excel, 8, defaults["toggle_customer_copy"]),
        }
    except KeyError:
        return None

    return data

(
    number_of_pdfs,
    agency_name,
    toggle_timestamp,
    toggle_customer_copy,
) = get_excel_data().values()


# Get input directory pdf's sorted by last modified date
def get_sorted_input_dir():
    downloads_dir = Path.home() / "Downloads"
    list_of_pdf_files = list(downloads_dir.glob("*.pdf"))
    sorted_pdf_files_by_last_modified_date = sorted(
        list_of_pdf_files, key=lambda file: Path(file).lstat().st_mtime, reverse=True
    )
    return sorted_pdf_files_by_last_modified_date


sorted_input_dir = get_sorted_input_dir()

# output directory
icbc_e_stamp_copies_dir = (
        Path.home() / "Desktop" / "ICBC E-Stamp Copies (this folder can be deleted)"
)
icbc_e_stamp_copies_dir.mkdir(exist_ok=True)
unsorted_e_stamp_copies_dir = (
        Path.home()
        / "Desktop"
        / "ICBC E-Stamp Copies (this folder can be deleted)"
        / "Unsorted E-Stamp Copies"
)
if toggle_customer_copy == "No":
    unsorted_e_stamp_copies_dir.mkdir(exist_ok=True)


def root_folder_filename(df):
    if df["licence_plate"].at[0] != "NONLIC":
        return f"{df['licence_plate'].at[0]} (Customer Copy).pdf"
    else:
        return f"{df['insured_name'].at[0].title()} (Customer Copy).pdf"


def sub_folder_filename(df):
    if df["licence_plate"].at[0] != "NONLIC":
        return f"{df['licence_plate'].at[0]}.pdf"
    else:
        return f"{df['insured_name'].at[0].title()}.pdf"


# open the first page of the pdf and scan a set coordinate to determine if it is an ICBC doc
def identify_icbc_pdf(doc):
    page_one = doc[0].get_text(
        "text", clip=keyword_dict["ICBC"]["is_icbc"].target_coordinates
    )
    if "Transaction Timestamp ".casefold() in page_one.casefold():
        return "ICBC"


# Returns all text, their coordinates, and which page number they are found on
def get_all_text(doc):
    field_dict = {}
    for page_num in range(len(doc)):
        page = doc[page_num - 1]
        wlist = page.get_text("blocks")
        text_boxes = [
            list(filter(None, inner_list[4].split("\n"))) for inner_list in wlist
        ]
        text_coords = [inner_list[:4] for inner_list in wlist]
        field_dict[page_num] = [
            [elem1, elem2] for elem1, elem2 in zip(text_boxes, text_coords)
        ]
    return field_dict


# Search for keyword matches, this one is hard to understand, will refactor later
def locate_keywords(all_text, type_of_pdf):
    field_dict = defaultdict(lambda: defaultdict(list))
    for pg_num, pg in all_text.items():
        for i, word_list in enumerate(pg):
            for j, target in keyword_dict[type_of_pdf].items():
                if target is not None:
                    try:
                        # This if statement list with [page number, (coordinates)] for the validation stamp position
                        if (
                                target.target_keyword
                                and target.target_coordinates
                                and any(target.target_keyword in s for s in word_list[0])
                        ):
                            coordinates = tuple(
                                x + y
                                for x, y in zip(
                                    all_text[pg_num][i][1],
                                    target.target_coordinates,
                                )
                            )
                            page_and_coordinates = [pg_num - 1, coordinates]
                            field_dict[type_of_pdf][j].append(page_and_coordinates)

                        # This if statement is used to find keywords other than license plate number
                        elif isinstance(target.target_keyword, str) and any(
                                target.target_keyword in s for s in word_list[0]
                        ):
                            word = all_text[pg_num][i + target.first_index][0][
                                target.second_index
                            ]
                            if target.append_duplicates:
                                field_dict[type_of_pdf][j].append(word)
                            elif target.join_list:
                                field_dict[type_of_pdf][j].append(
                                    " ".join(word).split(", ")
                                )
                            elif (
                                    word and word not in field_dict[type_of_pdf][j]
                            ):
                                field_dict[type_of_pdf][j].append(word)

                        #  This if statement is used to find the license plate number
                        elif isinstance(target.target_keyword, re.Pattern):
                            word = all_text[pg_num][i + target.first_index][0][
                                target.second_index
                            ]
                            if (
                                    re.search(target.target_keyword, word)
                                    and word not in field_dict[type_of_pdf][j]
                            ):
                                field_dict[type_of_pdf][j].append(word)
                    except IndexError:
                        continue
    return field_dict


# Removes whitespace and non-relevant words
def format_keywords(matching_keywords):
    field_dict = {}
    if not matching_keywords["licence_plate"]:
        field_dict["licence_plate"] = "NONLIC"
    else:
        dd(matching_keywords, field_dict, "licence_plate", r"Licence Plate Number ")
    dd(matching_keywords, field_dict, "transaction_timestamp", r"Transaction Timestamp ")
    dd(matching_keywords, field_dict, "agency_number", r"Agency Number ")
    dd(matching_keywords, field_dict, "owner_name", strip_char=".")
    dd(matching_keywords, field_dict, "applicant_name", strip_char=".")
    dd(matching_keywords, field_dict, "insured_name", strip_char=".")
    return field_dict


# find matching transaction timestamps in the input and output folder
def check_if_matching_transaction_timestamp(
        processed_timestamps,
        icbc_file_name,
        timestamp,
):
    # check for duplicates in input folder
    if timestamp in processed_timestamps:
        pass
    processed_timestamps.add(timestamp)
    # checks for duplicates in output folder
    output_dir_paths = list(Path(icbc_e_stamp_copies_dir).rglob("*.pdf"))
    output_dir_file_names = [path.stem.split()[0] for path in output_dir_paths]
    target_filename = Path(icbc_file_name).stem.split()[0]
    matching_transaction_ids = []
    if target_filename in output_dir_file_names:
        matching_paths = find_matching_paths(target_filename, output_dir_paths)
        for path_name in matching_paths:
            with fitz.open(path_name) as doc:
                target_transaction_id = doc[0].get_text(
                    "text",
                    clip=keyword_dict["ICBC"]["is_icbc"].target_coordinates,
                )
                if target_transaction_id:
                    match = int(
                        re.match(re.compile(r".*?(\d+)"), target_transaction_id).group(
                            1
                        )
                    )
                    matching_transaction_ids.append(match)
    return matching_transaction_ids


# Checks if stamp will fit if agency name entered is too long
stamp_does_not_fit = False


# Stamp the location where the string "NOT VALID UNLESS STAMPED BY" are found
def find_stamp_location(stamp_location, timestamp_date, page, agency_number):
    global stamp_does_not_fit
    fontname = fonts[font][0]
    fontname_bold = fonts[font][1]
    agency_name_coordinates = (3, 7, -3, 0)
    agency_name_factor = 60
    agency_number_coordinates = (0, 10, 0, 0)
    time_stamp_coordinates = (0, 13, 0, 0)
    formatted_date = (
        timestamp_date.strftime("%b %d, %Y")
        if toggle_timestamp == "Timestamp"
        else datetime.today().strftime("%I:%M")
    )
    agency_name_location = tuple(
        x + y for x, y in zip(stamp_location.coordinates, agency_name_coordinates)
    )

    agency_number_location = tuple(
        x + y for x, y in zip(stamp_location.coordinates, agency_number_coordinates)
    )

    date_location = tuple(
        x + y for x, y in zip(agency_number_location, time_stamp_coordinates)
    )
    agency_name_str = str(agency_name)

    rect = fitz.Rect(stamp_location.coordinates)
    fs = font_size * (min(rect.width, rect.height) / agency_name_factor)

    if len(agency_name_str) > 0:
        stamp_with_agency_name = page.insert_textbox(
            agency_name_location,
            f"{agency_name_str}\n{agency_number}\n{formatted_date}",
            align=fitz.TEXT_ALIGN_CENTER,
            fontname=fontname,
            fontsize=fs,
        )
        if stamp_with_agency_name < 0:
            stamp_does_not_fit = True

    else:
        page.insert_textbox(
            agency_number_location,
            agency_number,
            align=fitz.TEXT_ALIGN_CENTER,
            fontname=fontname_bold,
            fontsize=font_size,
        )
        page.insert_textbox(
            date_location,
            formatted_date,
            align=fitz.TEXT_ALIGN_CENTER,
            fontname=fontname,
            fontsize=font_size,
        )


# Stamp the location where the string "TIME OF VALIDATION" are found
def find_time_of_validation_location(time_location, timestamp_date, page):
    time_of_validation_am = (0, 0.5, 0, 0)
    time_of_validation_pm = (0, 21.9, 0, 0)
    formatted_date = (
        timestamp_date.strftime("%I:%M")
        if toggle_timestamp == "Timestamp"
        else datetime.today().strftime("%I:%M")
    )
    current_time = (
        timestamp_date.hour
        if toggle_timestamp == "Timestamp"
        else datetime.today().hour
    )
    time_location = tuple(
        x + y
        for x, y in zip(
            time_location.coordinates,
            time_of_validation_am if current_time < 12 else time_of_validation_pm,
        )
    )
    page.insert_textbox(
        time_location,
        formatted_date,
        align=fitz.TEXT_ALIGN_RIGHT,
        fontname=fonts["SpaceMono"][0],
        fontsize=6,
    )


# finds the pages that need to be stamped
def stamp_policy(
        timestamp,
        matching_keywords,
        doc,
        agency_number,
):
    timestamp_date = format_transaction_timestamp(timestamp)
    ValidationStamp = namedtuple("ValidationStamp", ["page_num", "coordinates"])
    validation_stamp = [
        ValidationStamp(page_num=t[0], coordinates=t[1])
        for t in matching_keywords["ICBC"]["validation_stamp"]
    ]
    for stamp_location in validation_stamp:
        page = doc[stamp_location.page_num]
        find_stamp_location(stamp_location, timestamp_date, page, agency_number)

    TimeOfValidation = namedtuple("TimeOfValidation", ["page_num", "coordinates"])
    time_of_validation = [
        TimeOfValidation(page_num=t[0], coordinates=t[1])
        for t in matching_keywords["ICBC"]["time_of_validation"]
    ]
    for time_location in time_of_validation:
        page = doc[time_location.page_num]
        find_time_of_validation_location(time_location, timestamp_date, page)


# Return page numbers of all non customer copy pages
def not_customer_copy_page_numbers(pdf):
    pages = []
    with fitz.open(pdf) as doc:
        top = False
        top_block = doc[0].get_text(
            "text", clip=keyword_dict["ICBC"]["top"].target_coordinates
        )
        if (
                "Temporary Operation Permit and Owner’s Certificate of Insurance".casefold()
                in top_block.casefold()
        ):
            top = True
        for page_num in range(len(doc)):
            page = doc[page_num]
            text_block = page.get_text(
                "text",
                clip=keyword_dict["ICBC"]["customer_copy_pages"].target_coordinates,
            )
            if not "Customer Copy".casefold() in text_block.casefold():
                pages.append(page_num)
        if top:
            del pages[-1]
    return list(reversed(pages))


def copy_policy(df, doc, pdf):
    root_folder_output_path = icbc_e_stamp_copies_dir / root_folder_filename(df)
    sub_folder_output_path = unsorted_e_stamp_copies_dir / sub_folder_filename(df)
    if toggle_customer_copy == "No":
        doc.save(
            unique_file_name(sub_folder_output_path),
            garbage=4,
            deflate=True,
        )
    doc.delete_pages(not_customer_copy_page_numbers(pdf))
    if len(doc) > 0:
        doc.save(
            unique_file_name(root_folder_output_path),
            garbage=4,
            deflate=True,
        )


# <=========================================Begin code execution=========================================>
timer = 0
def main():
    global timer
    loop_counter = 0
    copy_counter = 0
    found_icbc = 0
    # Stores the timestamps in input folder to avoid duplicate copies of the same pdf
    processed_timestamps = set()
    # Step 1: open the specified number of pdfs in the input directory (downloads folder)
    for pdf in progressbar(sorted_input_dir[:number_of_pdfs], prefix="Progress: ", size=40):
        loop_counter += 1
        with fitz.open(pdf) as doc:
            # Step 2 open each pdf and identify if it is an ICBC policy doc
            is_icbc_pdf = identify_icbc_pdf(doc)
            if is_icbc_pdf:
                found_icbc += 1
                # Step 3 if is ICBC doc, extract all text, their coordinates and page number where they are found
                all_text = get_all_text(doc)
                # Step 4 Search for keywords using the keyword dictionary
                matching_keywords = locate_keywords(all_text, is_icbc_pdf)
                # Step 5 Remove white space and irrelevant words in keywords
                formatted_keywords = format_keywords(ll(matching_keywords[is_icbc_pdf]))
                # Step 6 Save into a Pandas Data frame, data analysis to find matching transaction timestamp
                df = pd.DataFrame([formatted_keywords])
                # Step 7 Check if matching transaction timestamp in input and output folder
                timestamp = df["transaction_timestamp"].at[0]
                matching_timestamp = check_if_matching_transaction_timestamp(
                    processed_timestamps,
                    root_folder_filename(df),
                    int(timestamp)
                )
                # Step 8 Stamp ICBC
                if int(timestamp) not in matching_timestamp:
                    stamp_policy(
                        timestamp,
                        matching_keywords,
                        doc,
                        df["agency_number"].at[0],
                    )
                    if not stamp_does_not_fit:
                        # Step 9 Copy files to folder
                        copy_policy(df, doc, pdf)
                        copy_counter += 1
    if stamp_does_not_fit:
        print("Agency name is over the 17 character limit")
        os.system('pause')
        return
    elif found_icbc > 0:
        print(f"Scanned: {found_icbc} out of {loop_counter} documents")
        print(f"Copied: {copy_counter} out of {found_icbc} documents")
        timer = 3
    elif found_icbc == 0:
        print(f"There are no policy documents in the Downloads folder")
        timer = 3


if __name__ == "__main__":
    time_taken = timeit.timeit(lambda: main(), number=1)
    try:
        if not stamp_does_not_fit:
            print(f"Time taken: {time_taken} seconds")
            if timer > 0:
                time.sleep(timer)
    except Exception as e:
        print(str(e))
        time.sleep(3)
