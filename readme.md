# Doc scripter

Scripted editing of Word documents with some examples providing a re-usable template for document automation.
The motivation is to be able to create a template document in Word with defined formatting and then be able to programatically 
edit its content to produce new documents. 

The provided example `confirm_flatmate.py` serves the purpose of generating "*Wohnungsgeberbestätigung*" - a form needed in Germany to confirm sub-lease a room in a shared flat.

## Installation

- Prerequisities:
  - python 3.10 installed
- install with: 

```bash
python3 -m pipenv sync
```

- alternativelly, install the package `python-docx` directly with: `python3 -m pip install python-docx`

## Usage

- edit the configuration in `config.json` for your purpose

  ```bash
  python3 -m pipenv run confirm_flatmate.py <word_document_filename> <config_json_file>
  ```

- using the example:

  ```bash
  python3 -m pipenv run confirm_flatmate.py wohnungsgeber.docx config.json
  ```

## How it works

- in the JSON configuration file, the necessary information of the tenant, main tenant and flat owner is mapped using various keywords
- the Word document is loaded and parsed using `python-docx` module
- the script looks for the keywords in the document to replace them with the values as defined in the config
- if the document uses tables, then the content inside can't be retrieved directly using `Document.paragraphs` but `Document.tables` and then the paragraphs. Note that every time formatting is changed (eveny inside one user-defined paragraph or one uninterrupted sentence), a new XML paragraph is created. 
- new, modified, document is saved with the preposition `modified_`

## Resources

- Documentation of `python-docx` module: https://python-docx.readthedocs.io/en/latest/ 