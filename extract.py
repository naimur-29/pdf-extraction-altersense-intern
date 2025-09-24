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


def extract_field(field_indetifier, text, file_path):
    """
    Extracts the static field from given text.
    """
    try:
        value = text.split(field_indetifier)[1].strip().split("\n")[0].strip()

        # Handle exceptional fields
        if field_indetifier == "Order No:":
            value = text.split(field_indetifier)[1].strip().split(" ")[0].strip()
        elif field_indetifier == "No of Pieces:":
            value = int(value)
    except:
        print(f"Couldn't extract {field_indetifier} from {file_path}")
        value = "N/A"
    return value


def extract_static_fields(text, file_path):
    """
    Extracts the static fields from given text.
    """
    fields = [
        ("order_no", "Order No:"),
        ("product_description", "Product Description:"),
        ("season", "Season:"),
        ("type_of_construction", "Type of Construction:"),
        ("no_of_pieces", "No of Pieces:"),
        ("sales_mode", "Sales Mode:"),
        ("order_no", "Order No:"),
        ("order_no", "Order No:"),
        ("order_no", "Order No:"),
        ("order_no", "Order No:"),
    ]
    static_fields = {}
    for field_name, field_indetifier in fields:
        static_fields[field_name] = extract_field(field_indetifier, text, file_path)

    return static_fields


def extract_country_codes_with_prices(text):
    """
    Extract all country codes with their Invoice Average Prices.
    """
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
        print(
            f"Couldn't extract Country Codes with their Invoice Average Prices from {file_path}"
        )
        country_codes_with_price = []

    return country_codes_with_price


def extract_country_codes_with_delivery_times(text):
    """
    Extract all country codes with their Times of Delivery.
    """
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
        print(
            f"Couldn't extract Country Codes with their Time of Delivery from {file_path}"
        )
        country_codes_with_time = []

    return country_codes_with_time


def extract(text, file_path):
    """
    Extracts the fields from given text, creates & returns a pandas dataframe.
    """
    static_fields = extract_static_fields(text, file_path)
    country_codes_with_price = extract_country_codes_with_prices(text)
    country_codes_with_delivery_times = extract_country_codes_with_delivery_times(text)

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
                    country_codes_with_delivery_times,
                )
            )[0][0]
        except:
            time_of_delivery = "N/A"
        data["Time of Delivery"].append(time_of_delivery)

        # Static fields
        data["Order No"].append(static_fields["order_no"])
        data["Product Description"].append(static_fields["product_description"])
        data["Season"].append(static_fields["season"])
        data["Type of Construction"].append(static_fields["type_of_construction"])
        data["No. of Pieces"].append(static_fields["no_of_pieces"])
        data["Sales Mode"].append(static_fields["sales_mode"])

    return pd.DataFrame(data)


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
        print(f"An error occurred during excel file saving: {e}\n")


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
        df = extract(pdf_text, file_path)

        print("Saving to output.xlsx...")
        save_excel(df, output_path, file_name[:31])


if __name__ == "__main__":
    """
    Main function to handle command-line arguments and process files.
    """
    if len(sys.argv) < 2:
        print("Error: No file path provided.")
        print(
            'Usage: python extract.py "/path/to/your/pdf/file_or_directory" [output.xlsx]'
        )
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = "output.xlsx"
    if len(sys.argv) > 2:
        output_path = sys.argv[2]

    if os.path.exists(input_path):
        print("Valid: The path exists.")

        # If it's a single file
        if os.path.isfile(input_path):
            with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                process_pdf_file(input_path, output_path)

        # If it's a directory with multiple files
        elif os.path.isdir(input_path):
            file_paths = os.listdir(input_path)
            with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                for file_path in file_paths:
                    process_pdf_file(os.path.join(input_path, file_path), output_path)
    else:
        print("Invalid: The path does not exist.")
