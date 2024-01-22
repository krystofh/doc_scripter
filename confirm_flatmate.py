import sys
import json
import os
from docx import Document


def display_help():
    print(
        "Usage: python confirm_flatmate.py <word_document_filename> <config_json_file>"
    )
    sys.exit(1)


def load_config(config_file):
    try:
        with open(config_file, "r") as file:
            config_data = json.load(file)
        return config_data
    except FileNotFoundError:
        print(f"Error: Config file '{config_file}' not found.")
        sys.exit(1)
    except json.JSONDecodeError:
        print(f"Error: Invalid JSON format in config file '{config_file}'.")
        sys.exit(1)


def print_rows(doc, table_number: int):
    """Print all rows of a table specified by its index. Useful for playing around in order to find what element you are looking for.

    Parameters
    ----------
    doc : docx.Document
        Word document to be parsed
    table_number : int
        index of the table
    """
    for i, row in enumerate(doc.tables[table_number].rows):
        print(f"Row {i}: ", end="")
        print(row.cells[0].text)


def check_keyword(config_data: dict, keyword: str):
    pass


def replace_keywords(doc, config_data: dict, mode="table"):
    if mode == "paragraph":
        pass
    elif mode == "table":
        # Used if the document has a table structure
        # Look up each data keyword to be changed in the document (see config.json)
        # "obj": {"property1": {"keyword": KEY, "value": VALUE}, "property2": ...}
        # Example object: "main_tenant": {"firstname": {"keyword": "FIRSTNAME_MAIN_TENANT","value": "Max"}
        keywords = []
        values = []

        for obj, properties in config_data.items():
            for obj_property in properties:
                keyword = config_data[obj][obj_property]["keyword"]
                value = config_data[obj][obj_property]["value"]
                print(f"Looking for {keyword}: {value}")

        # Look in each table
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    # Each cell can have multiple paragraphs with different formatting
                    # In order not to overwrite formatting, edit the paragraphs, not entire cell
                    for cell_paragraph in cell.paragraphs:
                        if key in cell_paragraph.text:
                            cell_paragraph.text = cell_paragraph.text.replace(
                                key, str(value)
                            )
    else:
        raise AttributeError("Unsopported operation mode for keyword search")


def main():
    # Check for the correct number of arguments
    if len(sys.argv) != 3:
        display_help()

    # Get file names from command line arguments
    docx_filename = sys.argv[1]
    json_filename = sys.argv[2]

    # Check if files exist
    if not os.path.exists(docx_filename):
        print(f"Error: Word document file '{docx_filename}' not found.")
        sys.exit(1)

    config_data = load_config(json_filename)

    # Load the Word document
    document = Document(docx_filename)
    # print_rows(document, 0)

    # Replace keywords in the document
    replace_keywords(document, config_data)

    # Save the modified document
    output_filename = f"modified_{docx_filename}"
    document.save(output_filename)
    print(f"Success: Document saved as '{output_filename}'.")


if __name__ == "__main__":
    main()
