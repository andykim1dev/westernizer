## Westernizer

Westernizer (OS) is a Python-based desktop tool designed to automate the preparation of Western blot documentation for research presentations and experimental recordkeeping. It combines data entry, image organization, and slide generation into a single, interactive pipeline. The tool outputs a structured PowerPoint (.pptx) file and a companion Excel (.xlsx) file for users to complete experimental metadata.

Westernizer (OS) is ideal for researchers who frequently prepare Western blot figures and associated metadata for meetings, lab reports, or publications. By automating routine formatting and layout tasks, the tool saves time, ensures consistent presentation, and reduces human error in documentation.

---

## Features

- **Automated Image Processing**: Scans a folder of Western blot images and extracts image metadata, like filenames, to infer antibody names.
- **Excel Template Generation**: Automatically creates a formatted Excel spreadsheet with prefilled antibody names and placeholder columns for user input (e.g., volumes, gel percentages, antibody dilutions, and experimental conditions).
- **PowerPoint Generation**: Builds a clean, slide-based summary of the experiment. Each slide includes the Western blot image, optional marker image, and condition labels.
- **Marker Matching**: Matches Western blot images with marker images based on user input from the Excel sheet.
- **Condition Annotations**: Places experimental condition labels as aligned text boxes and visual connectors on the second slide of the presentation.
- **Interactive Prompts**: Uses `tkinter` to display clear prompts and instructions, making it easy to use without requiring command-line interaction.
- **Formatted Output**: Ensures that both the Excel and PowerPoint outputs are professional and presentation-ready, with aligned text, borders, highlighting, and readable font settings.

---

## How It Works

1. The script begins by creating a directory structure with folders for inserting Western blot images (`Insert WB`), marker images (`Insert MARKER`), and for saving outputs (`Output Excel` and `Output PowerPoint`).
2. A GUI message prompts the user to place all Western blot images into the `Insert WB` folder.
3. The script scans the folder, removes any image with `'lad'` in the filename (usually ladder markers), and extracts the filenames (without extensions) to use as antibody names.
4. These names are used to auto-generate an Excel sheet (`WB Template.xlsx`) where users can fill in volumes, gel percentages, marker names, antibody catalog numbers, dilutions, and experimental conditions.
5. The Excel file is opened automatically for user input. After the user finishes and clicks "OK" in a pop-up, the script rereads the file.
6. The conditions column is extracted and used to annotate slides with text labels and lines.
7. A PowerPoint is generated:
   - The title slide contains the antibody names and the current date.
   - Each subsequent slide contains a Western blot image centered and scaled to fit.
   - If a matching marker image is found (based on Excel input), it is added to the same slide, usually aligned to the right or bottom.
   - On the second slide, condition labels are added in a structured and evenly spaced fashion for visual clarity.
   - Each slide also receives annotations like filenames and size markers (e.g., "kDa") near the image.

---

## Usage

To run the program:

```bash
python3 "Westernizer (OS).py"

1.	A pop-up will appear prompting you to place your Western blot images into the Insert WB directory. Supported formats include .tif, .jpg, .png, and .jpeg.
2.	After clicking OK, the script will:
    - Scan the Insert WB directory.
    - Ignore any image files containing 'lad' in the filename.
    - Use the remaining filenames (excluding extensions) as antibody names.
3.	An Excel file named WB Template.xlsx will be generated automatically in the Output Excel folder. Open it and fill in the following columns:
    - Volume (uL): Enter the sample volume used per lane.
    - Gel (%): Specify the gel concentration (e.g., 10, 12.5).
    - Marker: Provide the marker image filename (without file extension) corresponding to each blot.
    - 1′ Antibody Catalog #: Enter the catalog number of the primary antibody.
    - 1′ Antibody Dilution (1:x): Input the dilution ratio (e.g., 1:1000).
    - (Optional placeholder column): A spacer column is included but unused.
    - Conditions: Define experimental conditions for labeling (e.g., Control, KO, Drug A, etc.).
4.	After completing the Excel sheet, save it and click OK on the second pop-up window.
5.	The script will now:
    - Parse the filled Excel sheet.
    - Generate a dictionary linking antibody names to their marker image identifiers.
    - Extract and sanitize the list of experimental conditions.
    - Create a PowerPoint presentation that includes:
    - A title slide summarizing all antibody names and the current date.
    - One slide per Western blot image, centered and scaled.
    - A marker image placed in the lower-right corner (if matched).
    - A “kDa” size label and line annotation near the top left.
    - On the second slide, a row of condition labels with connector lines, evenly spaced and center-aligned.
    - A small header on each image slide showing the original filename.
6.	The generated .pptx file will be saved to the Output PowerPoint directory. A message will confirm completion, and the file will open automatically.
7.	The Excel and PowerPoint files are timestamped with the current date and time for traceability.


