# Word Document Image Appender

This is a simple Streamlit application designed to help users easily append images to Word documents. It was created for a friend to streamline the process of adding images to documents.

## Features

- Create a new Word document.
- Upload an existing Word document.
- Append images to the document with options for auto-numbering or custom naming.
- Download the updated document.

## Installation

To run this application, you need to have Python installed. Follow the steps below to set up the project:

1. Clone the repository:
   ```bash
   git clone https://github.com/srps/streamlit-doc-enhancer.git
   cd streamlit-doc-enhancer
   ```

2. Install the required dependencies:
   ```bash
   uv sync
   ```

## Usage

To start the application, run the following command:
```bash
uv run streamlit run main.py
```

### Application Interface

1. **Create New Document**: Click the "Create New Document" button to start with a blank document.
2. **Upload Existing Document**: Use the file uploader to select an existing Word document.
3. **Append Images**: Upload images and choose between auto-numbering or custom naming for the images. Click "Append Images" to add them to the document.
4. **Download Updated Document**: Once the images are appended, click the "Download Updated Document" button to save the changes.

## Dependencies

- [Streamlit](https://streamlit.io/)
- [python-docx](https://python-docx.readthedocs.io/en/latest/)
- [Pillow](https://python-pillow.org/)

## License

This project is licensed under the MIT License.