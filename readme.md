# Auto Bundler.py
This program takes source pdfs, cover pages and indices for the "bundles" and merges the source pdfs into one output document called `<Witness_Name> - statement and exhibits`

# Pre requisites
1. Create a folder with the witness's name as the name of the folder in the src folder
2. Create saved searches for the list of witnesses
   1. The saved search should have these fields in order
   2. Control number
   3. Witness ID
   4. Document date
   5. Content description
   6. Undated/Estimated Date
   7. Witness document type
   8. Date Range or Partial Date
3. Download the pdfs as produced images and _NOT OCR_ versions of thee pdfs and save to the witness's named folder
4. Export to file 
   1. Change the file extention to CSV
   2. Save it as export to the witness's named folder
5. Repeat for each witness

# Setup
1. Create a file named order.csv in the source folder
2. Column names:
   1. Witness name
   2. Date
   3. should_skip
   4. is_draft
3. Populate these fields with the information about the witnesses 
4. Date - the date that will be displayed on the cover page
5. should_skip - a boolean value represented by 0 or 1
6. is_draft - a boolean value represented by 0 or 1
7. Witness name - must match the name of the folder you create in the src folder

# Running
```console
conda activate pdf
cd path
python bundler.py
```
When running the cd command use the path to the folder containing `bundler.py`.
For example: `<path-to-repo>/py-pdf/bundler.py`

When the program runs the name of each witness will show up before the bundler runs on it.

# Output
The program outputs a folder that is named by the witness in the output folder.
the folder contains:
1. A cover page word doc and pdf  
2. An index page word doc and pdf if there are exhibits
3. An exhibits pdf that contains the cover page, index page, and exhibits
4. A stand alone statements pdf that has all statements from the witness
5. A complete bundle named `<Witness_Name> - Statement and Exhibits.pdf` 
   
If you need to change any of the information that is produced _DO NOT_ change it in the output folder - change the information in the src folder and re-run.

# FAQs
<<<<<<< HEAD
Permission denied:  
   This generally means that you have a file open in an external program which is holding a write-lock (e.g. Adobe Acrobat, Microsoft Word etc.). If the issue persists restart your machine because Microsoft Word sucks.

Bundle Index out of range:  
Check your `order.csv` files to make sure it is well formed e.g. errors like missing commas.

Errors with opening output file:
This is most likely because an OCR version of an input pdf was included.
=======
## Permission denied:  
This generally means that you have a file open somehwere that it is trying to write to close it. if the issue persists restart your machine because Mircrosoft Word sucks

## Bundle Index out of range:  
The way to fix this is to double check your order.csv files make sure you don't have any errors like missing commas 

## Output file errors on open:  
This is most likely because you didn't download the none ocr versions of the pdfs and will not work.
>>>>>>> temp



