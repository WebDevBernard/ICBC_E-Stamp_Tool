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
]

# This sets a default value of None if unused
KeywordSearch = namedtuple(
    "SearchKeywords", search_keywords, defaults=(None,) * len(search_keywords)
)
SearchParams = namedtuple(
    "SearchParams", search_params, defaults=(None,) * len(search_params)
)

# Keyword dictionary and their search params
keyword_dict = (
    KeywordSearch(
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
        top=SearchParams(target_keyword="Temporary Operation Permit and Owner’s Certificate of Insurance",
                         first_index=0, second_index=0),
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
    )._asdict())


# <=========================================Helper Functions=========================================>


# makes all strings into a list
def make_string_to_list(dictionary):
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


# adds a (+=1) next to identical file names
def unique_file_name(path):
    filename, extension = os.path.splitext(path)
    counter = 1
    while Path(path).is_file():
        path = filename + " (" + str(counter) + ")" + extension
        counter += 1
    return path


# https://stackoverflow.com/questions/3160699/python-progress-bar
def progressbar(it, prefix="", size=60, out=sys.stdout):  # Python3.6+
    count = len(it)
    start = time.time()  # time estimate start

    def show(j):
        x = int(size * j / count)
        # time estimate calculation and string
        remaining = ((time.time() - start) / j) * (count - j)
        mins, sec = divmod(remaining, 60)  # limited to minutes
        time_str = f"{int(mins):02}:{sec:03.1f}"
        print(f"{prefix}[{u'█' * x}{('.' * (size - x))}] {j}/{count} Est wait {time_str}", end='\r', file=out,
              flush=True)

    if len(it) > 0:
        show(0.1)  # avoid div/0
        for i, item in enumerate(it):
            yield item
            show(i + 1)
        print(flush=True, file=out)


# <=========================================Main Functions=========================================>


# This Excel sheet reads the user inputs
def get_excel_data():
    # used in Excel sheet Function
    def find_excel_values(df, row, default):
        value = df.at[row, 1]
        return default if isinstance(value, float) else value

    defaults = {
        "number_of_pdfs": 5,
        "agency_name": "",
        "broker_number": "",
        "toggle_timestamp": "Timestamp",
        "toggle_customer_copy": "No"
    }
    root_dir = Path(__file__).parent.parent
    excel_path = root_dir / "BM3KXR.xlsx"

    if not excel_path.exists():
        return defaults

    try:
        df_excel = pd.read_excel(excel_path, sheet_name=0, header=None)
        data = {
            # special if statement in case user enters a string
            "number_of_pdfs": defaults["number_of_pdfs"] if isinstance(
                find_excel_values(df_excel, 2, defaults["number_of_pdfs"]), str) else find_excel_values(df_excel, 2,
                                                                                                        defaults[
                                                                                                            "number_of_pdfs"]),
            "agency_name": find_excel_values(df_excel, 4, defaults["agency_name"]),
            "agency_number": find_excel_values(df_excel, 6, defaults["broker_number"]),
            "toggle_timestamp": find_excel_values(df_excel, 8, defaults["toggle_timestamp"]),
            "toggle_customer_copy": find_excel_values(df_excel, 10, defaults["toggle_customer_copy"]),
        }
    except KeyError:
        return defaults

    return data


(
    number_of_pdfs,
    agency_name,
    broker_number,
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

output_dir_paths = list(Path(icbc_e_stamp_copies_dir).rglob("*.pdf"))
output_dir_file_names = [path.stem.split()[0] for path in output_dir_paths]


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
        "text", clip=keyword_dict["is_icbc"].target_coordinates
    )
    if "Transaction Timestamp " in page_one:
        return True


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


# Search for keyword matches
def locate_keywords(all_text):
    field_dict = defaultdict(list)
    for page_num, pages in all_text.items():
        # text_index is the index of all the words extracted from the pdf
        for text_index, text_single_page in enumerate(pages):
            # params_key is the keyword_dict key names (transaction_timestamp, agency_number, etc.)
            for params_key, params in keyword_dict.items():
                try:
                    # This if statement list with [page number, (coordinates)] for the validation stamp position
                    if (
                            params.target_keyword
                            and params.target_coordinates
                            and any(params.target_keyword in s for s in text_single_page[0])
                    ):
                        coordinates = tuple(
                            x + y
                            for x, y in zip(
                                all_text[page_num][text_index][1],
                                params.target_coordinates,
                            )
                        )
                        page_and_coordinates = [page_num - 1, coordinates]
                        field_dict[params_key].append(page_and_coordinates)

                    # This if statement is used to find keywords other than license plate number
                    elif isinstance(params.target_keyword, str) and any(
                            params.target_keyword in s for s in text_single_page[0]
                    ):
                        keyword = all_text[page_num][text_index + params.first_index][0][
                            params.second_index
                        ]
                        if keyword and keyword not in field_dict[params_key]:
                            field_dict[params_key].append(keyword)

                    #  This if statement is used to find the license plate number
                    elif isinstance(params.target_keyword, re.Pattern):
                        keyword = all_text[page_num][text_index + params.first_index][0][
                            params.second_index
                        ]
                        if re.search(params.target_keyword, keyword) and keyword not in field_dict[params_key]:
                            field_dict[params_key].append(keyword)
                except IndexError:
                    continue
    return field_dict


# Removes whitespace and non-relevant words
def format_keywords(matching_keywords):
    # Returns the first match to a keyword
    def filter_keywords(key, regex=None, strip_char=None):
        for items in matching_keywords[key]:
            if regex:
                items[0] = re.sub(re.compile(regex), "", items[0])
            if strip_char:
                field_dict["insured_name"] = items[0].rstrip(strip_char)
            else:
                field_dict[key] = items[0]

    field_dict = {}
    if not matching_keywords["licence_plate"]:
        field_dict["licence_plate"] = "NONLIC"
    else:
        filter_keywords("licence_plate", r"Licence Plate Number ")

    if len(matching_keywords["top"]) > 0 and "Temporary Operation Permit and Owner’s Certificate of Insurance" in \
            matching_keywords["top"][0]:
        field_dict["top"] = True
    else:
        field_dict["top"] = False
    filter_keywords("transaction_timestamp", r"Transaction Timestamp ")
    filter_keywords("agency_number", r"Agency Number ")
    filter_keywords("owner_name", strip_char=".")
    filter_keywords("applicant_name", strip_char=".")
    filter_keywords("insured_name", strip_char=".")
    return field_dict


# find matching transaction timestamps in the input and output folder
def check_if_matching_transaction_timestamp(
        processed_timestamps,
        icbc_file_name,
):
    # used to find matching transaction timestamps in output folder
    def find_matching_paths(paths):
        return [path for path in paths if path.stem.split()[0] == target_filename]

    # checks for duplicates in output folder
    target_filename = Path(icbc_file_name).stem.split()[0]
    if target_filename in output_dir_file_names:
        matching_paths = find_matching_paths(output_dir_paths)
        for path_name in matching_paths:
            with fitz.open(path_name) as doc:
                target_transaction_id = doc[0].get_text(
                    "text",
                    clip=keyword_dict["is_icbc"].target_coordinates,
                )
                match = int(
                    re.match(re.compile(r".*?(\d+)"), target_transaction_id).group(
                        1
                    )
                )
                processed_timestamps.add(match)
    return processed_timestamps


# Checks if stamp will fit if agency name entered is too long
stamp_does_not_fit = False


# Stamp the location where the string "NOT VALID UNLESS STAMPED BY" are found
def find_stamp_location(stamp_location, timestamp_date, page, agency_number):
    global stamp_does_not_fit
    font_size = 9
    font = "SpaceMono"
    fonts = {
        "SpaceMono": ["spacemo", "spacembo"]
    }
    fontname = fonts[font][0]
    fontname_bold = fonts[font][1]
    agency_name_coordinates = (3, 7, -3, 0)
    agency_name_factor = 60
    agency_number_coordinates = (0, 10, 0, 0)
    time_stamp_coordinates = (0, 13, 0, 0)
    formatted_date = (
        timestamp_date.strftime("%b %d, %Y")
        if toggle_timestamp == "Timestamp"
        else datetime.today().strftime("%b %d, %Y")
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
            str(agency_number),
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
    time_of_validation_am = (0, 0.7, 0, 0)
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
        fontname="helv",
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
        for t in matching_keywords["validation_stamp"]
    ]
    for stamp_location in validation_stamp:
        page = doc[stamp_location.page_num]
        find_stamp_location(stamp_location, timestamp_date, page, agency_number)

    TimeOfValidation = namedtuple("TimeOfValidation", ["page_num", "coordinates"])
    time_of_validation = [
        TimeOfValidation(page_num=t[0], coordinates=t[1])
        for t in matching_keywords["time_of_validation"]
    ]
    for time_location in time_of_validation:
        page = doc[time_location.page_num]
        find_time_of_validation_location(time_location, timestamp_date, page)


def copy_policy(df, doc):
    # Return page numbers of all non customer copy pages
    def not_customer_copy_page_numbers():
        pages = []
        # need top because it does not print customer copy on that page
        top = df["top"].at[0]
        for page_num in range(len(doc)):
            page = doc[page_num]
            text_block = page.get_text(
                "text",
                clip=keyword_dict["customer_copy_pages"].target_coordinates,
            )
            if "Customer Copy" not in text_block:
                pages.append(page_num)
        if len(pages) > 0 and top:
            del pages[-1]
        return list(reversed(pages))

    non_customer_copy = (not_customer_copy_page_numbers())
    root_folder_output_path = icbc_e_stamp_copies_dir / root_folder_filename(df)
    sub_folder_output_path = unsorted_e_stamp_copies_dir / sub_folder_filename(df)

    if toggle_customer_copy == "No":
        doc.save(
            unique_file_name(sub_folder_output_path),
            garbage=4,
            deflate=True,
        )
    if len(non_customer_copy) > 0:
        doc.delete_pages(non_customer_copy)
    if len(doc) > 0:
        doc.save(
            unique_file_name(root_folder_output_path),
            garbage=4,
            deflate=True,
        )


# <=========================================Start of loop=========================================>
timer = 0


def main():
    global timer
    loop_counter = 0
    scan_counter = 0
    copy_counter = 0
    # Stores the timestamps in input and output folder to avoid duplicate copies of the same pdf
    processed_timestamps = set()
    # Step 1: open the specified number of pdfs in the input directory (downloads folder)
    for pdf in progressbar(sorted_input_dir[:number_of_pdfs], prefix="Progress: ", size=40):
        loop_counter += 1
        with fitz.open(pdf) as doc:
            # Step 2 open each pdf and identify if it is an ICBC policy doc
            is_icbc_pdf = identify_icbc_pdf(doc)
            if not is_icbc_pdf:
                continue
            scan_counter += 1
            # Step 3 if is ICBC doc, extract all text, their coordinates and page number where they are found
            all_text = get_all_text(doc)
            # Step 4 Search for keywords using the keyword dictionary
            matching_keywords = locate_keywords(all_text)
            # Step 5 Remove white space and irrelevant words in keywords
            formatted_keywords = format_keywords(make_string_to_list(matching_keywords))
            # Step 6 Save into a Pandas DF, data analysis to find matching transaction timestamp
            df = pd.DataFrame([formatted_keywords])
            # Step 7 Check if matching transaction timestamp in input and output folder
            timestamp_str = df["transaction_timestamp"].at[0]
            timestamp_int = int(df["transaction_timestamp"].at[0])
            check_if_matching_transaction_timestamp(
                processed_timestamps,
                root_folder_filename(df)
            )
            if timestamp_int in processed_timestamps:
                continue
            # Used when user uses the broker_number in the Excel sheet
            agency_number_or_broker_number = df["agency_number"].at[0] if len(
                str(broker_number)) == 0 else broker_number
            # Step 8 Stamp ICBC
            stamp_policy(
                timestamp_str,
                matching_keywords,
                doc,
                agency_number_or_broker_number,
            )
            if stamp_does_not_fit:
                continue
            # Step 9 Copy files to folder
            # After stamping, add it to the set so any duplicates in input (downloads folder) will be ignored
            processed_timestamps.add(timestamp_int)
            copy_policy(df, doc)
            copy_counter += 1
    if stamp_does_not_fit:
        print("Agency name is over the 17 character limit")
        os.system('pause')
        return
    elif scan_counter > 0:
        print(f"Scanned: {scan_counter} out of {loop_counter} documents")
        print(f"Copied: {copy_counter} out of {scan_counter} documents")
        timer = 3
    elif scan_counter == 0:
        print(f"There are no policy documents in the Downloads folder")
        timer = 3


if __name__ == "__main__":
    time_taken = timeit.timeit(lambda: main(), number=1)
    print(f"Time taken: {time_taken} seconds")
    time.sleep(timer)
