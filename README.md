# Newsletter Agent

This repository contains the automated generation script for the **Pravartiya Newsletter**, focusing on listed companies and regulatory compliance updates (e.g., from SEBI).

## Overview

The main script (`agent.py`) reads data from Excel files (`.xlsx`) provided in the `data/` directory, extracts the necessary regulations, circulars, and updates, and formats them into a PowerPoint presentation (`.pptx`) based on the provided template.

## Folder Structure

- `agent.py`: The core script that performs data extraction and PowerPoint generation.
- `Template/`: Contains the PowerPoint template file used for generation (`Pravartiya - Template (1).pptx`).
- `data/`: Place your input Excel files here. The script expects the files to contain `Summary`, `Month`, `Year`, and other metadata. **Note**: SEBI-related Excel files are processed first.
- `output/`: The generated presentation (`output.pptx`) is saved here after successfully running the script.
- `documents/`: For any reference documents or PDFs.
- `scripts/`: Assorted debugging, inspection (`inspect_*.py`), and evaluation scripts.
- `tests/`: Various test files (`test_*.py`) and older test output presentations.

## Requirements

Ensure you have the necessary Python libraries installed. Primary dependencies include:

- `pandas`
- `python-pptx`

Install dependencies using pip:

```bash
pip install pandas python-pptx
```

## How to Run

1. Place your data files (Excel) inside the `data/` directory.
2. Ensure the template is located in `Template/`.
3. Run the agent script from the project root:
   ```bash
   python agent.py
   ```
4. The generated newsletter will be saved to `output/output.pptx`.
