Detailed Description

In this project, I developed a Python script to automate the task of matching customer service case numbers with corresponding transcripts from Otter.ai. The automation was designed to handle various data processing tasks, including file extraction, validation, and integration with Excel.

**Data Cleaning and Preparation:**

- Drop NaN Values: Removed rows with missing values to ensure data integrity.
- Conditional Row Removal: Deleted rows based on specific conditions to refine the dataset.
- Column Type Conversion: Used methods like to_numeric, to_datetime, and astype to convert columns to appropriate data types for accurate processing.

**Matching Case Numbers with Transcripts:**

- Reading Excel and Text Files: The script read the input Excel file containing customer service case numbers and listed all text files from the transcripts folder.
- Finding Matches: A custom function was used to find matching text files for each case number based on file names.
- Integrating Transcripts: The content of matching text files was read and integrated into the corresponding rows in the Excel sheet.

**Output:**

**Updated Excel File:** The final output was an updated Excel file where each case number was matched with its corresponding transcript, significantly reducing the manual effort required to perform this task weekly.

The automation process not only improved efficiency by reducing the time taken to match case numbers with transcripts but also ensured greater accuracy and consistency in the reports generated. This project demonstrates the power of Python in automating repetitive tasks and enhancing data processing workflows.
