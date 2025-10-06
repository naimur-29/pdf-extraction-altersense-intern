# PDF Data Extractor

This script extracts specific data from PDF files and saves it into a formatted Excel spreadsheet. It can process a single PDF file or all PDF files within a directory.

-----

## 🚀 Getting Started

### Prerequisites

Make sure you have **Python 3.x** installed on your system.

### Installation

1.  **Clone the Repository**
    Start by cloning the project's repository to your local machine.

    ```bash
    git clone [repository_url]
    cd [repository_name]
    ```

2.  **Create a Virtual Environment**
    It's a good practice to use a virtual environment to manage project dependencies.

    ```bash
    python -m venv venv
    ```

      - **Activate the environment:**
          - **Windows:** `venv\Scripts\activate`
          - **macOS/Linux:** `source venv/bin/activate`

3.  **Install Dependencies**
    Install the required libraries from the `requirements.txt` file.

    ```bash
    pip install -r requirements.txt
    ```

-----

## ✍️ Usage

Run the script from your terminal using the following command.

```bash
python extract.py [input_path] [output_path]
```

### Arguments

| Argument | Description | Example |
| :--- | :--- | :--- |
| **`[input_path]`** | **Required**. The path to a single PDF file or a directory containing multiple PDF files. | `C:\Users\docs\my_pdf.pdf` or `C:\Users\docs` |
| **`[output_path]`** | **Optional**. The path and name for the output Excel file. If not provided, the output will be saved as `output.xlsx` in the same directory. | `C:\Users\reports\extracted_data.xlsx` |

### Examples

  - **Process a single PDF:**

    ```bash
    python extract.py "C:/Users/documents/invoice.pdf"
    ```

    This will create an `output.xlsx` file with data from `invoice.pdf`.

  - **Process multiple PDFs in a directory:**

    ```bash
    python extract.py "C:/Users/documents/invoices_folder" "all_invoices_data.xlsx"
    ```

    This command will process all PDFs in `invoices_folder` and save the extracted data into a single Excel file named `all_invoices_data.xlsx`. Each PDF's data will be stored on a separate sheet within the Excel file.

-----

## ❗ Important Notes

  - The script is designed to extract data from a specific PDF layout. It may not work as expected on other PDF structures.
  - The output Excel file is automatically formatted with column headers and adjusted column widths for better readability.
  - Errors encountered during extraction will be printed to the console.
