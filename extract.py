import os
import re
import sys
from datetime import datetime

import pandas as pd
import pdfplumber


def read_pdf(pdf_path, page_no=1):
    """
    Extracts text from a specified page of a PDF file.
    """
    try:
        text = ""
        with pdfplumber.open(pdf_path) as pdf:
            text += pdf.pages[page_no - 1].extract_text()
        return text
    except:
        print("Error: Not a pdf file!\n")
        return None


def slice_text(text, begining="", ending=""):
    """
    Slices and returns a string, given a begining & ending part.
    """
    return text.split(begining)[1].strip().split(ending)[0].strip()


def extract_fields(text, file_path):
    """
    Extracts the fields from given text and returns a pandas dataframe.
    """
    # Extract all repeating fields:
    # ----------------------------
    try:
        order_no = text.split("Order No:")[1].strip().split(" ")[0].strip()
    except:
        print(f"Couldn't extract Order No. from {file_path}")
        order_no = "N/A"

    try:
        product_description = (
            text.split("Product Description:")[1].strip().split("\n")[0].strip()
        )
    except:
        print(f"Couldn't extract Product Description from {file_path}")
        product_description = "N/A"

    try:
        season = text.split("Season:")[1].strip().split("\n")[0].strip()
    except:
        print(f"Couldn't extract Season from {file_path}")
        season = "N/A"

    try:
        type_of_construction = (
            text.split("Type of Construction:")[1].strip().split("\n")[0].strip()
        )
    except:
        print(f"Couldn't extract Type of Construction from {file_path}")
        type_of_construction = "N/A"

    try:
        no_of_pieces = int(
            text.split("No of Pieces:")[1].strip().split("\n")[0].strip()
        )
    except:
        print(f"Couldn't extract No of Pieces from {file_path}")
        no_of_pieces = "N/A"

    try:
        sales_mode = text.split("Sales Mode:")[1].strip().split("\n")[0].strip()
    except:
        print(f"Couldn't extract Sales Mode from {file_path}")
        sales_mode = "N/A"

    # Extract all country codes with their Invoice Average Price:
    # -------------------------
    try:
        # 1. Find the text slice
        text_slice = slice_text(
            text,
            "Invoice Average Price Country",
            "By accepting and performing under this Order, the Supplier acknowledges:",
        )

        # 2. Format into a python list:
        country_codes_with_price = []
        for line in text_slice.split("\n"):
            try:
                price = float(line.strip().split("USD")[0].strip())
            except:
                price = "N/A"

            try:
                codes = line.strip().split("USD")[1].strip()
            except:
                codes = "N/A"

            country_codes_with_price.extend(
                [{"code": code.strip(), "price": price} for code in codes.split(", ")]
            )
    except:
        print("!!!!!!!!!!!!!!!!!!")
        print(
            f"Couldn't extract Country Codes with their Invoice Average Prices from {file_path}"
        )
        print("!!!!!!!!!!!!!!!!!!")
        country_codes_with_price = []

    # Extract all country codes with their Time of Delivery
    try:
        # 1. Find the text slice
        text_slice = slice_text(
            text, "Time of Delivery Planning Markets Quantity % Total Qty", "Total:"
        )

        # 2. Format into a python list:
        country_codes_with_time = []
        for item in text_slice.split("\n"):
            # Extract and format the date
            date_match = re.search(r"\d{2} [A-Za-z]{3}, \d{4}", item)
            if date_match:
                original_date_str = date_match.group(0)
                try:
                    dt_object = datetime.strptime(original_date_str, "%d %b, %Y")
                    formatted_date = dt_object.strftime("%Y-%m-%d")
                except ValueError:
                    formatted_date = "N/A"
            else:
                formatted_date = "N/A"

            # Extract country codes before all parenthesized patterns
            all_parentheses_matches = re.findall(
                r"([A-Z]+)\s*\(", item
            )  # captures the country code just before the parentheses

            temp_codes = []
            # Process the extracted patterns
            for outside_code in all_parentheses_matches:
                if outside_code:
                    temp_codes.append(outside_code)

            # If no country codes were found, handle the 'N/A' case
            if not temp_codes:
                country_codes_with_time.append((formatted_date, "N/A"))
            else:
                # Combine the formatted date with each country code found
                for code in temp_codes:
                    country_codes_with_time.append((formatted_date, code))
    except:
        print("!!!!!!!!!!!!!!!!!!")
        print(
            f"Couldn't extract Country Codes with their Time of Delivery from {file_path}"
        )
        print("!!!!!!!!!!!!!!!!!!")
        country_codes_with_time = []

    data = {
        "Order No": [],
        "Country": [],
        "Product Description": [],
        "Season": [],
        "Type of Construction": [],
        "No. of Pieces": [],
        "Sales Mode": [],
        "Time of Delivery": [],
        "Invoice Average Price": [],
    }

    for country_code_with_price in country_codes_with_price:
        # Dynamic fields
        data["Country"].append(country_code_with_price["code"])
        data["Invoice Average Price"].append(country_code_with_price["price"])

        try:
            time_of_delivery = list(
                filter(
                    lambda x: x[1] == country_code_with_price["code"],
                    country_codes_with_time,
                )
            )[0][0]
        except:
            time_of_delivery = "N/A"
        data["Time of Delivery"].append(time_of_delivery)

        # Static fields
        data["Order No"].append(order_no)
        data["Product Description"].append(product_description)
        data["Season"].append(season)
        data["Type of Construction"].append(type_of_construction)
        data["No. of Pieces"].append(no_of_pieces)
        data["Sales Mode"].append(sales_mode)

    df = pd.DataFrame(data)
    return df


def save_excel(df, file_name, sheet_name, font_size=12):
    """
    Creates a new sheet in the given excel file and saves it after applying auto-formatting.
    """
    try:
        df.to_excel(
            writer, sheet_name=sheet_name, index=False, header=False, startrow=1
        )

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        general_format = workbook.add_format(
            {"font_size": font_size, "font_name": "Calibri"}
        )

        header_format = workbook.add_format(
            {
                "bold": False,
                "font_size": font_size,
                "font_name": "Calibri",
            }
        )

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Auto-size all columns based on content and apply the universal format.
        for i, col in enumerate(df.columns):
            # Calculate the maximum length of the column name and all data points.
            max_len = max(df[col].astype(str).map(len).max(), len(col))

            # Set the width of the column to fit the content plus a small buffer.
            worksheet.set_column(i, i, max_len + 1, general_format)

        print(f"Successfully saved and formatted {file_name}.\n")

    except Exception as e:
        print(f"An error occurred: {e}\n")


def process_pdf_file(file_path, output_path):
    """
    Processes a pdf file to extract the required fields and saves it to a excel file.
    """
    file_name = file_path.split("/")[-1].strip()
    print(f"Found {file_name}!")
    print("---------------------------")

    print("Trying to extract text from page 1...")
    pdf_text = read_pdf(file_path)

    if pdf_text:
        print("Extracting required fields...")
        df = extract_fields(pdf_text, file_path)

        print("Saving to output.xlsx...")
        save_excel(df, output_path, file_name[:31])


# Check if a file path argument was provided
if len(sys.argv) < 2:
    print("Error: No file path provided.")
    print('Usage: python extract.py "/path/to/your/pdf/file"')
    sys.exit(1)

# The first command-line argument is the script name itself,
# so the second argument (index 1) is the file path.
file_path = sys.argv[1]
try:
    output_path = sys.argv[2]
except:
    output_path = "output.xlsx"


if os.path.exists(file_path):
    print("Valid: The path exists.")

    # If it's a single file
    if os.path.isfile(file_path):
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            process_pdf_file(file_path, output_path)

    # If it's a directory with multiple files
    elif os.path.isdir(file_path):
        files_inside = os.listdir(file_path)
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            for path in files_inside:
                process_pdf_file(os.path.join(file_path, path), output_path)
else:
    print("Invalid: The path does not exist.")
