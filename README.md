Automate a PowerPoint Presentation using AutoPptReport. This does need some requirements for the PPTX template namely having "[DATE-RANGE] and [TASK COMPLETED]" 
as well as having photos (upto only 6) per date range and the folder to be named the same date range of the pdf report.

Example:

input_folder
 March 3-8/
 Surname - March 3-8 - Weekly Accomplishment Report.pdf

You also need to install the following Python libraries by typing this in the CLI:

pip install python-pptx pdfplumber dotenv

You need to then install Gemini API for Python with this command:

pip install -q -U google-generativeai
