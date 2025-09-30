# PDF Book Tagger

**PDF Book Tagger** is a PowerShell tool that automatically embeds bibliographic metadata into PDF eBooks. It inserts the book's title, author(s), and Library of Congress Classification (LCC) ([reference](https://www.loc.gov/aba/publications/FreeLCC/freelcc.html#ABX)) into the PDF file’s *Title*, *Author*, and *Subject* properties.

## Features
The tool supports:
- Processing of single PDF files
- Batch processing of multiple PDFs in a directory
- Direct metadata lookup by ISBN

## Prerequisites
- Windows with PowerShell 5.0 or later  
- Administrator privileges (required on first run to install [IText7Module](https://www.powershellgallery.com/packages/IText7Module/1.0.34))  
- ~750 MB of disk space on first run to download and extract LCC data (reduced to ~85 MB after schedules are generated)

**Tip:** To make custom file properties (such as *Title*, *Author*, and *Comments*) visible in Windows Explorer (File Explorer), you can install [PDF Property Extension](https://coolsoft.altervista.org/en/pdfpropertyextension).

## Usage
Run the script from PowerShell 5.1 (Windows PowerShell) or PowerShell 7+ (pwsh):
```powershell
.\pdf-book-tagger.ps1 <input>
```

Where `<input>` is one of:
- `<pdf-path>`: Path to a single PDF file (absolute or relative).
- `<directory-path>`: Path to a directory containing PDF files.
- `<ISBN>`: ISBN-10 or ISBN-13 (hyphens and spaces allowed).
- Active internet connection for metadata retrieval

To check the current version:
```powershell
.\pdf-book-tagger.ps1 -version
```

## Examples
Directory processing:<br>
<img src="docs/Demo_Directory.gif" alt="Usage of PDF Book Tagger" width="800">

Single PDF processing:<br>
<img src="docs/Demo_PDF.gif" alt="Usage of PDF Book Tagger" width="800">

Metadata lookup by ISBN:<br>
<img src="docs/Demo_ISBN.gif" alt="Usage of PDF Book Tagger" width="800">


## Output
If processing is successful, a new PDF with embedded metadata is created in an auto-generated *Success* directory on your desktop. The new PDF file is renamed to match the eBook's title, while the original remains unchanged.

If processing fails:
- For a single ISBN or PDF path: an error message is displayed.
- For a directory path: an error message is displayed, and all unprocessed PDFs are moved to an auto-generated *Failure* directory on the desktop.


## Metadata sources
To maximize accuracy, it queries multiple metadata providers (APIs and websites):

- [Library of Congress](https://lx2.loc.gov/)
- [CERN Library Catalogue](https://catalogue.library.cern/)  
- [Open Library](https://openlibrary.org/)  
- [Google Books API](https://www.googleapis.com/)  
- [CiNii](https://cir.nii.ac.jp/)  
- [Yale Library](https://library.yale.edu/)  
- [Prospector (Colorado Alliance of Research Libraries)](https://prospector.coalliance.org/)  

Library of Congress results are preferred for LCC classification.


## How It Works
On first run, the tool automatically downloads the latest [Library of Congress Classification (LCC)](https://www.loc.gov/cds/products/MDSConnect-classification.html) file into the *Resources* directory.
It then parses this data using *"lib\SchedulesExtractor.jar"* to generate a `.txt` schedule for each [main class](https://www.loc.gov/catdir/cpso/lcco/) in the same directory.
Once the schedules are created, the original classification file is deleted to conserve disk space.
The generated schedule files are as follows:
| Main Class                                                        | Schedule File |
|-------------------------------------------------------------------|---------------|
| A - General Works                                                 | A.txt         |
| B - Philosophy. Psychology. Religion                              | B.txt         |
| C - Auxiliary Sciences of History                                 | C.txt         |
| ...                                                               | ...           |
| Z - Bibliography. Library Science. Information Resources (General)| Z.txt         |

Each schedule file contains the full subject hierarchy as defined in the official [LOC text files](https://www.loc.gov/aba/publications/FreeLCC/freelcc.html#ABX), starting from its main class.

For example, *QA241* ([LCC_Q2025TEXT.pdf](https://www.loc.gov/aba/publications/FreeLCC/freelcc.html#ABX)) corresponds to:

<img src="docs/Example LCC.png" alt="Example subject hierarchy from LOC PDF" width="500">

This generates the following entry in the `Q.txt` schedule file:

*QA241 - Science/Mathematics/Algebra/Number theory/General works, treatises, and textbooks*

Here, *Science* represents the *Q* main class.


Once the schedules have been generated, each PDF is processed through the following five steps:

1. **Extract ISBN**  
   Scans the first 10 pages of the PDF to detect valid ISBNs (ISBN-10 or ISBN-13).

2. **Fetch Metadata**  
   Queries multiple sources to retrieve candidate titles, authors, and LCC numbers.

3. **Resolve Conflicts**  
   Selects the most reliable metadata by
   - Prioritizing the longest repeated title across sources
   - Choosing the most frequently occurring LCC (preferring Library of Congress when available)
   - Filtering similar author name variants.
   Prompts the user when results are ambiguous or tied.

4. **Lookup Subject Hierarchy**  
   Retrieves the full subject path from the corresponding schedule file using the selected LCC number.

5. **Write Metadata**  
   Embeds the finalized metadata directly into the PDF file.


## Estimated Processing Time per PDF
Processing a PDF typically takes around 13 seconds, depending on connection speed, response time from metadata sources, and computer performance.


## Installation

PDF Book Tagger is a fully functional PowerShell script. Download the source code and ensure the following files are present, with the structure maintained as follows:
<pre>
pdf-book-tagger
    ├── pdf-book-tagger.ps1
    ├── lib/
    │    └── SchedulesExtractor.jar
    └── data/
         └── CanonicalIdentifiers.json
</pre>
These are the only necessary files for the script to work (no compilation required).

However, the repository also contains a SchedulesExtractor directory with the Java source code. This is provided for users who want to inspect or modify the Java code used for processing the Library of Congress Classification (LCC) schedules.

For normal usage, you do not need to compile or run the Java code; the included prebuilt Java binaries are automatically used by the PowerShell script.

## Troubleshooting

- **No valid ISBN found in PDF**<br>
    The PDF may not contain an ISBN in the first 10 pages. Try running the script with the ISBN manually: `.\pdf-book-tagger.ps1 <ISBN>`


- **Failed to open PDF. The file may be corrupted**<br>
  Try repairing the PDF using [iLovePDF Repair](https://www.ilovepdf.com/repair-pdf)

- **Book not found from all sources**<br>
  The ISBN may be incorrect or the book may not be in any of the databases. Insert the data manually.
