# Electrochemistry Data Extractor

A Streamlit application to extract specific step/cycle data from `.xlsb` (Excel Binary Workbook) files for CV and GCD analysis.

## Features
- **CV Mode**: Filters to the second-to-last active cycle and extracts the `Voltage_V` and `_Current_A` columns.
- **GCD Mode**: Filters Cycle 2 (Step 2) and Cycle 3 (Step 1), concatenates them smoothly with continuous elapsed time (`StepTime_s`), and extracts `Voltage_V`.
- **Bulk Processing**: Upload multiple `.xlsb` files at once.
- **Excel Export**: Download all extracted data consolidated into a single `.xlsx` file.

## Running Locally

1. Install Python 3.9+ (if not already installed).
2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the Streamlit app:
   ```bash
   streamlit run app.py
   ```
4. A browser window will automatically open (default: `http://localhost:8501`).

## Deploying to Streamlit Community Cloud (GitHub)

1. Go to your GitHub account and create a new repository (e.g., `electrochemistry-extractor`).
2. Upload the files in this folder (`app.py`, `requirements.txt`, and `README.md`) to the repository.
3. Sign up/Log in to [Streamlit Community Cloud](https://share.streamlit.io/).
4. Click on **New app**.
5. Give Streamlit permission to access your GitHub repositories.
6. Select your repository (`electrochemistry-extractor`) and set the **Main file path** to `app.py`.
7. Click **Deploy!**
8. Now you can share the link to your app with anyone, and they can upload their `.xlsb` files directly to the web app.
