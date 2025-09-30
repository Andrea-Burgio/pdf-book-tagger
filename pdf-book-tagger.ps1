param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$inputArg,

    [Parameter(Mandatory=$false)]
    [switch]$Version
)

$SCRIPT_VERSION = "0.9.0"

$divisionLineLong = "`n<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>`n"
$divisionLineShort = "`n________________________________________`n"


$desktopPath = [Environment]::GetFolderPath("Desktop")
$outputDirectorySuccess = Join-Path $desktopPath "Success"
$outputDirectoryFailure = Join-Path $desktopPath "Failure"
$dataDirectory = Join-Path $PSScriptRoot "data"
$canonicalIdentifiers = $null






# ========== Resources directory setup: LOC classification data acquisition ==========

$jarUrl = Join-Path $PSScriptRoot "lib/SchedulesExtractor.jar"

# Directory containing all schedule files (e.g., Q.txt, T.txt) and used for downloading classification files.
$resourcesDirectory = Join-Path $PSScriptRoot "Resources"

# Gzipped archive containing Classification.xml, downloaded from LOC.
$classificationGz = Join-Path $resourcesDirectory "Classification.xml.gz"

# Extracted XML file containing all LOC subjects and LCC codes, used to generate schedule files.
$classificationXml = Join-Path $resourcesDirectory "Classification.xml"


<#
  Verifies that all the extracted LOC subject schedules are present in $resourcesDirectory and match the entries
  in Classification.xml.
#>
function AllSubjectsPresent {

    Write-Host "Checking subjects currently present in $resourcesDirectory directory..."

    $LOCSubjects = @('A','B','C','D','E','F','G','H','J','K','L','M','N','P','Q','R','S','T','U','V','Z')
    $existanceToScheduleTxtFileMap = @{
        Missing  = $LOCSubjects | Where-Object { -not (Test-Path (Join-Path $resourcesDirectory "$_.txt")) }
        Present  = $LOCSubjects | Where-Object { Test-Path (Join-Path $resourcesDirectory "$_.txt") }
    }

    if ($existanceToScheduleTxtFileMap['Missing'].Count -gt 0){
        if ($existanceToScheduleTxtFileMap['Missing'].Count -ne $LOCSubjects.Count) {
            Write-Host "Some schedules are missing: $($existanceToScheduleTxtFileMap['Missing'] -join ', ')" -ForegroundColor Yellow
        }
        else {
            Write-Host "All schedules are missing." -ForegroundColor Yellow
        }

        Write-Host "Expected: $($LOCSubjects.Count) Found: $($existanceToScheduleTxtFileMap['Present'].Count)" -ForegroundColor Yellow
        return $false
    }

    return $true

}

# Extracts the Classification.xml to $resourcesDirectory from a .gz archive and deletes the original compressed file.
function ExtractAndRemoveGZip{

    Write-Host "Extracting $classificationGz..."
    $inStream = [System.IO.File]::OpenRead($classificationGz)
    $gzipStream = New-Object System.IO.Compression.GzipStream($inStream, [System.IO.Compression.CompressionMode]::Decompress)
    $outStream = [System.IO.File]::Create($classificationXml)
    $gzipStream.CopyTo($outStream)

    $gzipStream.Dispose()
    $inStream.Dispose()
    $outStream.Dispose()

    # Remove Classification.gz
    Write-Host "Deleting $classificationGz file..."
    Remove-Item $classificationGz -Force
    Write-Host "File deleted." -ForegroundColor Green

    Write-Host "Extraction completed. $classificationXml extracted." -ForegroundColor Green

}

# Parses LOC classification HTML to find the latest year's Classification XML file download link.
function GetClassificationDownloadUrl($htmlContent) {

    # Extract all years from <h2> tags
    $years = [regex]::Matches($htmlContent, '<h2>(\d{4})</h2>') | ForEach-Object { $_.Groups[1].Value }
    if (-not $years) {
        throw "Could not find any year headings on $url"
    }
    Write-Host "Years found: $($years -join ', ')"

    # Pick the newest year
    $latestYear = ($years | Sort-Object {[int]$_} -Descending | Select-Object -First 1)
    Write-Host "Latest year available: $latestYear"

    # Find the <ul> following the latest year
    $patternUL = "<h2>$latestYear</h2>\s*<ul.*?>(.*?)</ul>"
    $ulMatch = [regex]::Match($htmlContent, $patternUL, [System.Text.RegularExpressions.RegexOptions]::Singleline)

    if (-not $ulMatch.Success) {
        throw "No UL found for year $latestYear."
    }

    $ulContent = $ulMatch.Groups[1].Value

    # Extract the first <a> tag with inner text 'XML'
    $patternXML = '<a\s+href="([^"]+)">XML</a>'
    $xmlMatch = [regex]::Match($ulContent, $patternXML)
    if (-not $xmlMatch.Success) {
        throw "No XML link found in $latestYear UL."
    }

    return "https://www.loc.gov" + $xmlMatch.Groups[1].Value # Download URL

}

<#
  Removes duplicate codes from .txt files in $ResourcesDirectory, keeping the last (most complete) line for each code.
  For example:
  QA612 - Science/Mathematics/Geometry/Topology/Algebraic topology. Combinatorial topology
  QA612 - Science/Mathematics/Geometry/Topology/Algebraic topology. Combinatorial topology/General works
  Only the last occurrence is needed.
#>
function RemoveDuplicateCodes {

    Write-Host "Removing duplicates.."
    Get-ChildItem -Path $ResourcesDirectory -Filter *.txt | ForEach-Object {
        $lines = Get-Content $_.FullName
        $codeMap = [ordered]@{}

        foreach ($line in $lines) {
            $code = ($line -split '\s+')[0]

            # Store/overwrite with the current line
            $codeMap[$code] = $line
        }

        # Write all unique lines (keeping last occurrence of each code)
        $codeMap.Values | Set-Content $_.FullName -Encoding UTF8
        Write-Host "Completed processing: $($_.Name)"
    }

    Write-Host "Duplicates removed" -ForegroundColor Green

}

# Downloads the latest Library of Congress Classification XML file in $resourcesDirectory.
function DownloadXMLClassificationFile {

    # URL of the LOC classification page
    $url = "https://www.loc.gov/cds/products/MDSConnect-classification.html"
    Write-Host "Searching for latest classification file..."

    $downloadUrl = GetClassificationDownloadUrl (Invoke-WebRequest -Uri $url).Content

    # Download
    $client = New-Object System.Net.WebClient
    $client.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")

    try {
        Write-Host "Downloading from $downloadUrl..."
        $client.DownloadFile($downloadUrl, $classificationGz)
    }
    catch {
        throw "Couldn't download file from $downloadUrl.`n$($_.Exception.Message)"
    }

    Write-Host "Latest Classification.xml.gz file downloaded successfully: $classificationGz" -ForegroundColor Green

}

<#
  Generates subject-based schedule .txt files from Classification.xml, ensuring completeness and removing duplicates.
  Steps:
  1. Checks for Classification.xml; if missing, downloads the .gz archive and extracts it.
  2. Runs Java tool ($jarUrl) to generate per-subject schedule files (e.g., Q.txt).
  3. Verifies that all expected subjects are present and removes duplicate codes from each schedule file.
     Throws an error if any subjects are missing.
  4. Deletes the extracted Classification.xml file after successful processing.
#>
function GenerateScheduleTxtFiles{

    $xmlExists = Test-Path ($classificationXml)
    $gzExists = Test-Path ($classificationGz)

    # Get Classification.xml file
    if (-not $xmlExists){
        Write-Host "$classificationXml file not found." -ForegroundColor Yellow

        if (-not $gzExists){
            Write-Host "$classificationGz not found. Downloading latest..." -ForegroundColor Yellow

            DownloadXMLClassificationFile

            ExtractAndRemoveGZip($classificationGz)
        }
        elseif ($gzExists){
            Write-Host "Found: $classificationGz "
            ExtractAndRemoveGZip($classificationGz)
        }
    }

    <# Extracts classification data into separate output files based on LOC subject (A.xml, B.xml, etc.).
       Example:
       In Classification.xml file:
       <record>
         <leader>00256cw  a2200121n  4500</leader>
         <controlfield tag="001">CF 01446090</controlfield>
         <controlfield tag="003">DLC</controlfield>
         <controlfield tag="005">20010404152408.0</controlfield>
         <controlfield tag="008">010404aaaaaaaa</controlfield>
         <datafield tag="010" ind1=" " ind2=" ">
           <subfield code="a">CF 01446090</subfield>
         </datafield>
         <datafield tag="040" ind1=" " ind2=" ">
           <subfield code="a">DLC</subfield>
           <subfield code="c">DLC</subfield>
         </datafield>
         <datafield tag="084" ind1="0" ind2=" ">
           <subfield code="a">lcc</subfield>
         </datafield>
         <datafield tag="153" ind1=" " ind2=" ">
           <subfield code="a">QC23.2</subfield>
           <subfield code="h">Physics</subfield>
           <subfield code="h">Elementary textbooks</subfield>
           <subfield code="j">2001-</subfield>
         </datafield>
       </record>

       Output in Q.txt (since QC23.2):
       QC23.2 - Science/Physics/Elementary textbooks/2001-
     #>
    & java -jar $jarUrl $classificationXml

    if ($LASTEXITCODE -ne 0) {
        throw "Schedules extraction failed. Exit code $LASTEXITCODE."
    }

    if (-not (AllSubjectsPresent)) {
           throw "Schedules extraction failed. Some subjects were not extracted."
    }

    RemoveDuplicateCodes

    Write-Host "Deleting $classificationXml file..."
    Remove-Item $classificationXml -Force
    Write-Host "File deleted." -ForegroundColor Green

}

<#
  Prepares the environment for metadata extraction and PDF processing:
  - Ensures the IText7Module PowerShell module is installed. Installs it if missing.
  - Ensures the $resourcesDirectory exists. Creates the directory if missing.
  - Ensures that all LOC schedule files are present. Clears the directory and regenerates them
    from the latest Classification.xml (downloading and extracting it if necessary) if missing.
#>
function SetUp {
    # Import or, if missing, install IText7Module module
    if (-not (Get-Module -ListAvailable -Name 'IText7Module')) {
        Write-Host 'IText7Module not installed. Installing...' -ForegroundColor Yellow
        Install-Module -Name IText7Module -Force
        Write-Host "IText7Module installed"
    }

    # Ensure Resources directory exists
    if (-not (Test-Path $resourcesDirectory)) {
        Write-Host "$resourcesDirectory directory not found. Creating directory..." -ForegroundColor Yellow
        New-Item -ItemType Directory -Path "Resources" | Out-Null
        Write-Host "Directory created." -ForegroundColor Green
    }

    # Generate, if missing, LOC schedule files from latest classification data
    if (-not (AllSubjectsPresent)) {
        Write-Host "Preparing $resourcesDirectory..."
        Get-ChildItem -Path $resourcesDirectory -Filter "*.txt" | ForEach-Object { Remove-Item $_.FullName -Force }
        Write-Host "$resourcesDirectory is ready." -ForegroundColor Green
        try{
            GenerateScheduleTxtFiles
        }
        catch {
            Write-Host "Schedule generation failed. Removing Resources directory..." -ForegroundColor Red
            Remove-Item -Path $resourcesDirectory -Recurse -Force
            exit
        }
    }

    Write-Host "`nSetup complete.`n" -ForegroundColor Green
}







# ========== Utils ==========

function NormalizeIsbn([string]$isbn){
    return ($isbn.Trim() -replace '[^\dXx]', '').ToUpperInvariant()
}

function IsLcc([string]$targetString) {

    if ([string]::IsNullOrWhiteSpace($targetString) -or $targetString.Length -lt 1) {
        return $false
    }

    $normalizedTargetString = ($targetString -replace '\s+', ' ').Trim()

    $pattern = '^[A-Z]{1,3}\s*\d{1,4}\s*(\.\d+)?(\s*\.[A-Z]\d+[A-Z0-9]*)*(\s+\d{4})?$'

    $isCodeValidCode = [regex]::IsMatch($normalizedTargetString.ToUpperInvariant(), $pattern)

    $line = $null
    if ($isCodeValidCode){
        try {
            $line = GetLineFromSchedule($targetString)
        }
        catch {
            Write-Host "Failed to get line from schedule for '$targetString': $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }


    return (-not [string]::IsNullOrWhiteSpace($line) -and $isCodeValidCode)
}

<# Validates whether a string is a correct ISBN-10 or ISBN-13.
   Only two ISBN formats exist: ISBN-10 and ISBN-13
   ISBN-13:
   - Must contain 13 digits
   - Begins with 978 or 979
   - To find the last digit: multiply first 12 digits by a weight that alternates 1, 3, 1, 3, …. Add the 12 products.
     Apply modulo 10.
   ISBN-10:
   - Must contain 10 characters
   - Last character can be an X (ten)
   - Multiply the first nine digits by descending weights 10 down to 2. Sum the products. Apply the formula
     (11 - ($sum % 11)) % 11. The result is the final character. If the result is 10, the final character is X
 #>
function IsIsbn([string]$isbn) {

    if (-not $isbn) { return $false }

    # Normalize: trim and remove everything except digits and X/x
    $isbn = NormalizeIsbn $isbn

    if ($isbn.Length -ne 10 -and $isbn.Length -ne 13) {
        return $false
    }

    # ISBN-10 validation
    if ($isbn.Length -eq 10) {
        $sum = 0
        # Multiply the first nine digits by descending weights 10 down to 2. Sum the products.
        for ($i = 0; $i -lt 9; $i++) {
            if (-not [char]::IsDigit($isbn[$i])) {
                return $false
            }
            $sum += [char]::GetNumericValue($isbn[$i]) * (10 - $i)
        }

        $calculatedCheck = (11 - ($sum % 11)) % 11
        $lastChar = $isbn[9]

        if ($calculatedCheck -eq 10) {
            return ($lastChar -eq 'X')
        } else {
            if (-not [char]::IsDigit($lastChar)) {
                return $false
            }
            return [char]::GetNumericValue($isbn[9]) -eq $calculatedCheck
        }
    }
    # ISBN-13 validation
    else {
        $sum = 0
        for ($i = 0; $i -lt 12; $i++) {
            if (-not [char]::IsDigit($isbn[$i])) {
                return $false
            }
            $weight = if ($i % 2 -eq 0) { 1 } else { 3 }
            $sum += [char]::GetNumericValue($isbn[$i]) * $weight
        }

        $calculatedCheck = (10 - ($sum % 10)) % 10
        $lastChar = $isbn[12]

        if (-not [char]::IsDigit($lastChar)) {
            return $false
        }
        $checkDigit = [char]::GetNumericValue($lastChar)

        return $checkDigit -eq $calculatedCheck
    }
}

# ========== Fetchers utils ==========

<#
  Normalizes a given LCC code for consistency:
    - Converts to uppercase.
    - Extracts the primary LCC (e.g., 'QA300' from 'QA300 .A18 2015').
    - Handles special cases (e.g., 'QA299.6-433', 'B-BJ6a.1', 'P-PZ20.161.N52').
  Returns the normalized LCC string.
#>
function NormalizeLcc([string]$lcc){

    # Capitalize LCC
    $lcc = $lcc.ToUpperInvariant()

    #Get primary LCC. E.g.: 'QA300' from 'QA300 .A18 2015'
    $lcc = ($lcc -split '\s+')[0].Trim()

    # Cern library catalogue may contain LCCs such as 'QA299.6-433'.
    # Schedules may contain LCCs such as B-BJ6a.1 and P-PZ20.161.N52
    # Extract first part of the LCC.
    if ($lcc -match "^B-" -or $lcc -match "^P-"){
        $lcc = ($lcc -split '-')[0].Trim()
    }
    else {
        $lcc = ($lcc -split ' - ')[0].Trim()
    }

    return $lcc

}

<#
  Filters a list of input LCC codes, retaining only those that exist in the schedule files.
  Returns an array of valid, non-empty LCC codes.
#>
function FilterValidLccs([string[]]$lccs){

    # Filter only LCCs that exist in the schedule
    $filteredLccs = $lccs | Where-Object { $_ -and $_.Trim() -ne "" } |
                            Where-Object {
                                $line = GetLineFromSchedule $_
                                $null -ne $line -and $line -ne ""
                            }

    return $filteredLccs

}

<#
  Compares two LCC codes to determine their relative order.
  - Splits each code into alternating letter and number groups (e.g., "QA76.73.J38" → ["QA", "76.73", "J", "38"]).
  - Compares each group pairwise: numeric groups are compared numerically, letter groups lexicographically.
  - Returns:
      - -1 if $lccA comes before $lccB in classification order
      -  1 if $lccA comes after $lccB
      -  0 if both are equal
#>
function CompareLcc([string]$lccA, [string]$lccB) {

    if ($lccA -eq $lccB) {
        return 0
    }

    # Split into letter and number groups
    $pattern = '[A-Z]+|\d+(?:\.\d+)?'
    $groupsA = [regex]::Matches($lccA, $pattern) | ForEach-Object { $_.Value }
    $groupsB = [regex]::Matches($lccB, $pattern) | ForEach-Object { $_.Value }

    # Compare only up to the smallest group count
    $minGroupSize = [Math]::Min($groupsA.Count, $groupsB.Count)

    for ($i = 0; $i -lt $minGroupSize; $i++) {
        $groupA = $groupsA[$i]
        $groupB = $groupsB[$i]

        $isGroupANum = $groupA -match '^\d'
        $isGroupBNum = $groupB -match '^\d'

        if ($isGroupANum -and $isGroupBNum) {
            # Compare numerically
            if ([double]$groupA -lt [double]$groupB) {
                return -1
            }
            if ([double]$groupA -gt [double]$groupB) {
                return 1
            }
        }
        else {
            # Compare lexicographically (letters only)
            if ($groupA -lt $groupB) {
                return -1
            }
            if ($groupA -gt $groupB) {
                return 1
            }
        }
    }

    # If the groups have a different count, the shorter one is smaller, otherwise they are equal
    if ($groupsA.Count -lt $groupsB.Count) {
        return -1
    }
    elseif ($groupsA.Count -gt $groupsB.Count) {
        return 1
    }
    else {
        return 0
    }

}


<#
  Determines whether rangeA is a proper subrange of rangeB.
  Both ranges are in the format "PrefixStart-End" (e.g. PL8453.8-8453.895) and comparison is based on LCC
#>
function isSubrange([string]$rangeA, [string]$rangeB) {

    $rangeAParts = $rangeA -split '-'
    $rangeAStart = $rangeAParts[0]
    $rangeAPrefix = ([regex]::Match($rangeAStart, '^[A-Z]+')).Value
    $rangeAEnd = $rangeAPrefix + $rangeAParts[1]

    $rangeBParts = $rangeB -split '-'
    $rangeBStart = $rangeBParts[0]
    $rangeBPrefix = ([regex]::Match($rangeBStart, '^[A-Z]+')).Value
    $rangeBEnd = $rangeBPrefix + $rangeBParts[1]

    return ((CompareLcc $rangeAStart $rangeBStart) -eq 1) -and ((CompareLcc $rangeAEnd $rangeBEnd) -eq -1)

}

<#
  Checks whether a given LCC code falls within a specified range.
  The range is defined by a start and end code, and comparison is based on LCC ordering.
#>
function IsLccInRange([string]$lcc, [string]$rangeStart, [string]$rangeEnd){

    $rangePrefix = ([regex]::Match($rangeStart, '^[A-Z]+')).Value
    $rangeEnd = $rangePrefix + $rangeEnd

    return ((CompareLcc $lcc $rangeStart) -eq 1) -and ((CompareLcc $lcc $rangeEnd) -eq -1)

}

<#
  Retrieves the most appropriate subject line from an LCC schedule file for a given LCC code.
  - Determines the LOC subject prefix from the first character of the LCC (e.g., "Q").
  - Loads the corresponding schedule file (e.g., "Q.txt") from the resources directory.
  - Searches for an exact match first.
  - If no exact match is found, checks for:
    - Collapsed A-Z ranges (e.g., "QA9.A5-Z") and matches based on prefix similarity.
    - Numeric collapsed ranges (e.g., "PG7157.P47-7157.P472") and selects the most specific match.
    If matched via a collapsed range, returns a constructed line combining the original LCC and the subject hierarchy.
  Throws an error if the schedule file is missing.
#>
function GetLineFromSchedule([string]$lcc) {

    # Get LOCSubject (e.g., Q)
    $prefix = [string]$lcc[0]

    # Validate prefix: must be A-Z
    if ($prefix -notmatch '^[A-Z]$') {
        return $null
    }

    # Get the resource file (e.g., "Q.txt") based on the first character of the LCC code
    $file = Get-ChildItem -Path $resourcesDirectory -File -Filter "$prefix.txt" | Select-Object -First 1

    if (-not $file) {
        throw "No '$prefix.txt' in $resourcesDirectory"
    }

    $lines = Get-Content -Path $file.FullName

    $resultLine = $null
    $mostSpecificRange = $null
    foreach ($line in $lines) {
        $parts = $line -split ' - ', 2  # Split line into [LCC code, subject hierarchy]
        $lineCode = $parts[0].Trim()

        # Case 1: exact match
        if ($lineCode -eq $lcc) {
            return $line
        }

        # Case 2: Collapsed A-Z range, e.g. QA9.A5-Z or QL638.875.A-Z
        if ($lineCode -match "^([A-Z]+\d+(?:\.[A-Z]?\d+)*)(?:\.A\d+-Z)$") {
            if ($lcc -like "$($Matches[1]).*") {
                $subjectHierarchy = $parts[1].Trim()
                $lcc = NormalizeLcc $lcc
                $resultLine = "$lcc - $subjectHierarchy"
            }
        }

        # Case 3: Collapsed range, e.g. PG7157.P47-7157.P472 or PL8455.7-8455.795
        if ($lineCode -match "^([A-Z]+\d+(?:\.[A-Z]?\d+)*?)-(\d+(?:\.[A-Z]?\d+)*)$") {
            if (IsLccInRange $lcc $Matches[1] $Matches[2]){
                if (-not $mostSpecificRange -or (isSubrange $Matches[0] $mostSpecificRange)){
                    $mostSpecificRange = $Matches[0]
                    $subjectHierarchy = $parts[1].Trim()
                    $resultLine = "$lcc - $subjectHierarchy"
                }
            }
        }
    }

    return $resultLine
}

<#
  Reconstruct full name from comma-separated parts:
  - Handles formats like "Last, First", "First, Jr.", and "Last, First, Jr."
  - Applies suffix logic for Jr., I, II, III, IV
  Examples:
    "Smith, John"             -> "John Smith"
    "Callahan, James J., Jr." -> "James J. Callahan Jr."
    "Doe, Jane, III"          -> "Jane Doe III"
#>
function BuildNameFromParts([string]$author){

    $fullName = $author

    $lastName = $null
    $firstName = $null

    $parts = $fullName -split ',', 3

    # Handle two-part names: "Last, First" or "First, Jr."
    if ($parts.Count -eq 2){
        # Hande standard case: "Last, First"
        if ($parts[1].Trim() -notmatch '(?i)Jr') {
            $firstName = $parts[1].Trim()
            $lastName = $parts[0].Trim()
            $fullName = "$firstName $lastName"
        }
        # Handle special case (suffix without last name): "First, Jr."
        elseif ($parts[1].Trim() -match '^(Jr\.?|I|II|III|IV)$') {
            $firstName = $parts[0].Trim()
            $suffix = $parts[1].Trim()
            $fullName = "$firstName $suffix"
        }
    }
    # Handle three-part names: "Last, First, Jr."
    elseif ($parts.Count -eq 3 -and $parts[2].Trim() -match '^(Jr\.?|I|II|III|IV)$') {
        $firstName = $parts[1].Trim()
        $lastName = $parts[0].Trim()
        $fullName = "$firstName $lastName $($parts[2].Trim())"
    }
    # three-Part names: "Bohn, John L., 1965-"
    elseif ($parts.Count -eq 3){
        $firstName = $parts[1].Trim()
        $lastName = $parts[0].Trim()
        $fullName = "$firstName $lastName"
    }

    return $fullName

}

<# Wrong names example:
 Callahan, James J., Jr.
 James (James J.) Callahan
 E. J. (Edward J.) Norminton,
 P. O. J. (Philipp O. J.) Scherer
 John M. H. (John Meigs Hubbell) Olmsted
 in collaboration with Herbert Kreyszig
 IN COLLABORATION WITH Herbert Kreyszig
 George S.  Boolos
 by George Boolos
 WILLIAM. STALLINGS
 Joseph O.'rourke
 Volker Kuhn:
 David S. Richeson!
 Abraham Silberschatz, Professor
 Philipp O.J. Scherer
#>
function NormalizeAuthor([string]$author) {

    if (-not $author) {
        return $null
    }

    # Remove trailing comma
    $fullName = $author -replace ',\s*$', ''

    $fullName = BuildNameFromParts $fullName

    # Remove birth/death year ranges (e.g., "1945-2000", "1980-")
    $fullName = $fullName -replace '\s+\d{4}-?(\d{4})?', ''

    # Fix casing (WILLIAM. STALLINGS -> William. Stallings)
    $fullName = (Get-Culture).TextInfo.ToTitleCase($fullName.ToLower())

    <#
      Replace name before parentheses with name in parentheses
      Example: James (James J.) Callahan -> James J. Callahan
    #>
    if ($fullName -match "^(.*?)\((.*?)\)(.*)$") {
        # $matches[2] = content inside parentheses
        # $matches[3] = string after parentheses
        $fullName = "$($matches[2])$($matches[3])"
    }
    <#
      Correct misformatted surnames where an initial and apostrophe are separated by a period
      Example: Joseph O.'Rourke -> Joseph O'Rourke
    #>
    if ($fullName -match "([A-Za-z\s\.]{2,})?([A-Za-z])\.?('[A-Za-z]{2,})\b"){
        <#
          $matches[1] = optional leading name content (e.g., "Joseph ")
          $matches[2] = single-letter initial before the apostrophe (e.g., "O")
          $matches[3] = apostrophe and surname (e.g., "'Rourke")
        #>
        $capitalizedMatch3 = (Get-Culture).TextInfo.ToTitleCase($($matches[3].ToLower()))
        $fullName = "$($matches[1])$($matches[2])$($capitalizedMatch3)"
    }

    <#
      Remove punctuation after 3+ letter words only if followed by a space or end of string
      Example: William. Stallings -> William Stallings
    #>
    $fullName = $fullName -replace '\b([A-Za-z]{3,})[,;:\.!](?=\s|$)', '$1'

    <#
      Remove non-authorial phrases and academic titles commonly found in metadata
      Example: in collaboration with Herbert Kreyszig -> Herbert Kreyszig
      Case insensitive by default
    #>
    $fullName = $fullName -replace 'in collaboration with', '' `
                          -replace '\bby\s', '' `
                          -replace '\bprofessor', '' `
                          -replace '\s\.+\s\[Et Al\.*\]\.*', '' `
                          -replace 'Dr\.', '' `
                          -replace "(?i)\[And\s[a-z]{3,}\sOthers\]"

    # Replace various Unicode dash characters (– — ‒ − ‐) with a standard ASCII hyphen (-)
    $fullName = $fullName -replace '[\u2013\u2014\u2012\u2212\u2010]', '-'

    <#
      Add a period after single letters
      Example: Paul R Halmos -> Paul R. Halmos
    #>
    $fullName = $fullName -replace '\b([A-Za-z])(?!\.)(?=\s|$)', '$1.'

    <#
      Insert a space after a single-letter initial followed by a period
      Example: Philipp O.J. Scherer -> O. J. Scherer
    #>
    $fullName = $fullName -replace '(\b[A-Za-z]\.)(?=[A-Za-z]|$)', '$1 '

    # Remove common HTML entities that may appear in metadata
    $fullName = $fullName -replace '&amp', '' `
                          -replace '&quot', ''
    <#
      Final whitespace and punctuation cleanup:
      - Collapse multiple spaces into one
      - Remove stray periods surrounded by spaces
      - Trim leading whitespace
      - Remove trailing period
    #>
    $fullName = $fullName -replace '\s{2,}', ' ' `
                          -replace '\s+\.\s*', '' `
                          -replace '^\s', '' `
                          -replace '’', ''' `
                          -replace "\.$", ''

    # Ensure suffix "Jr" ends with a period
    $fullName = $fullName -replace 'Jr$', 'Jr.'

    $fullName = $fullName.Trim()

    return $fullName

}

# Removes ordinal-based edition markers like "2nd Edition" or "(Twenty-Third Edición)" from a title string.
function RemoveEditionFromTitle([string]$title){

    $cleanedTitle = $title

    # List of ordinal representationsRemove edition information (ordinal numbers + "Edition")
    $ordinals = @(
        'First','1st','1.',             'Second','2nd','2.',            'Third','3rd','3.',
        'Fourth','4th','4.',            'Fifth','5th','5.',             'Sixth','6th','6.',
        'Seventh','7th','7.',           'Eighth','8th','8.',            'Ninth','9th','9.',
        'Tenth','10th','10.',           'Eleventh','11th','11.',        'Twelfth','12th','12.',
        'Thirteenth','13th','13.',      'Fourteenth','14th','14.',      'Fifteenth','15th','15.',
        'Sixteenth','16th','16.',       'Seventeenth','17th','17.',     'Eighteenth','18th','18.',
        'Nineteenth','19th','19.',      'Twentieth','20th','20.',       'Twenty-First','21st','21.',
        'Twenty-Second','22nd','22.',   'Twenty-Third','23rd','23.',    'Twenty-Fourth','24th','24.',
        'Twenty-Fifth','25th','25.',    'Twenty-Sixth','26th','26.',    'Twenty-Seventh','27th','27.',
        'Twenty-Eighth','28th','28.',   'Twenty-Ninth','29th','29.',    'Thirtieth','30th','30.'
    )

    foreach ($ordinal in $ordinals) {
        $cleanedTitle = $cleanedTitle -replace "(?i)\s*\(?$ordinal\s+(?:Edition|Edición)\)?", ''
    }

    return $cleanedTitle

}

# Returns a list of unique acronyms found in a title.
function GetAcronymsInTitle([string]$title){

    <#
      Extract unique acronyms from the title, including dotted forms and optional trailing 's' (e.g., "U.S.A.", "PDFs")
      May include false positives (e.g., "AND", "ART").
    #>
    $acronymPattern = '(?:\b[A-Z]{2,5}(?:\.[A-Z]{2,5})*s?\b)'
    $acronyms = [regex]::Matches($title, $acronymPattern) | ForEach-Object { $_.Value } | Select-Object -Unique

    $notAcronyms = @(
        "AND", "THE", "FOR", "BUT", "NOT", "YOU", "ALL", "CAN", "WHO", "HOW",
        "WAS", "HAD", "HAS", "HER", "HIM", "HIS", "OUT", "NOW", "ONE", "TWO",
        "NEW", "END", "SET", "TOP", "GET", "RUN", "USE", "TIP", "LAW", "ART"
    )

    # Filter out common false positives (e.g., short all-caps words like "AND", "ART") from the acronym list
    $acronyms = $acronyms | Where-Object { $notAcronyms -notcontains $_ }

    return $acronyms

}

<#
  Returns the canonical form of a word if it matches known technical identifiers or Roman numerals.
  The identifiers are loaded from JSON on first use and cached for later calls.
#>
function GetCanonicalIdentifier([string]$word) {

    # Lazy-load the canonical identifiers from JSON if they haven't been loaded yet
    if (-not $script:canonicalIdentifiers){
        $jsonPath = Join-Path $dataDirectory "CanonicalIdentifiers.json"
        $script:canonicalIdentifiers = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json
    }

    $standardIdentifiers = $script:canonicalIdentifiers.technicalIdentifiers +
                           $script:canonicalIdentifiers.romanNumerals

    foreach ($standardId in $standardIdentifiers) {
        <#
          Generate common variants of the identifier for matching:
          - Uppercase form (e.g., "API")
          - Plural form (e.g., "APIS")
          - Parenthesized plural (e.g., "(APIS)")
        #>
        $standardIdVariants = @(
            $standardId.ToUpper(),
            $standardId.ToUpper() + "S",
            "($($standardId.ToUpper())S)"
        )
        # Check if $word matches any of the known variants. If so, return it
        if ($standardIdVariants -contains $word.ToUpper()) {
            return $standardId
        }
    }
    return $null

}

# Identifies the role of a title word (e.g., Acronym, Version, Hyphenated, etc.)
function GetWordType([int]$index, [string[]]$words, [string[]]$acronyms, [string[]]$lowercaseWords){

    $word = $words[$index]

    $word = $word -replace ':', ''

    if ($acronyms -contains $word) {
        return 'Acronym'
    }

    if (GetCanonicalIdentifier $word) {
        return 'StandardIdentifier'
    }

    # Version number (e.g., "2.0", "1.2.3")
    if ($word -match '^\d+(\.\d+)*$') {
        return 'Version'
    }

    # If it's a hyphenated compound word (e.g., "real-time", "end-to-end")
    if ($word -match '^[A-Za-z]+-[A-Za-z]+$') {
        return 'Hyphenated'
    }

    $prevEndsWithPunctuation = $index -gt 0 -and $words[$index - 1] -match '[:\-]$'
    if ($word.ToLower() -in $lowercaseWords -and -not $prevEndsWithPunctuation) {
        return 'Lowercase'
    }

    return 'Standard'

}

<#
   Formats a hyphenated word for title casing.
   Capitalizes subwords based on position and exception list:
    - First/last word in the title -> always capitalized
    - First/last subword in compound -> capitalized
    - Subwords not in $lowercaseWords -> capitalized
    - Otherwise -> lowercase
   Examples:
   "state-of-the-art" -> "State-of-the-art" (if first word)
   "mother-in-law" -> "Mother-in-law" (middle word)
   "jack-of-all-trades" -> "Jack-of-all-Trades" (if last word)
#>
function FormatHyphenatedWord([string]$word, [int]$index, [int]$totalWords, [string[]]$lowercaseWords){

    $subWords = $word -split '-'
    $formattedSubWords = @()

    for ($j = 0; $j -lt $subWords.Length; $j++) {
        $subword = $subWords[$j]
        $subLower = $subword.ToLower()

        $isFirstOrLastInTitle = ($index -eq 0 -or $index -eq ($totalWords - 1))
        $isFirstOrLastInCompound = ($j -eq 0 -or $j -eq ($subWords.Length - 1))

        if ($isFirstOrLastInTitle -or $isFirstOrLastInCompound -or ($subLower -notin $lowercaseWords)) {
            $formattedSubWords += $subword.Substring(0,1).ToUpper() + $subword.Substring(1).ToLower()
        } else {
            $formattedSubWords += $subLower
        }
    }

    return ($formattedSubWords -join '-')

}

<#
  Formats a single word for title casing.
  Capitalizes based on position, punctuation context, and exception list:
  - First/last word in the title -> always capitalized
  - Word following punctuation (e.g., colon, dash) -> capitalized
  - Word not in $lowercaseWords -> capitalized
  - Otherwise -> lowercase
  Examples:
  - "the quick brown fox" -> "The quick brown Fox"
  - "life: the journey begins" -> "Life: The journey Begins"
#>
function FormatStandardWord([int]$index, [string[]]$words, [string[]]$lowercaseWords) {
    $word = $words[$index]
    if ([string]::IsNullOrEmpty($word)){
        return $null
    }

    $lowerWord = $word.ToLower()
    $isFirstOrLastInTitle = ($index -eq 0 -or $index -eq ($words.Length - 1))
    $prevEndsWithPunctuation = $index -gt 0 -and $words[$index - 1] -match '[:\-]$'

    $formattedWord = $null
    if ($isFirstOrLastInTitle -or ($lowerWord -notin $lowercaseWords) -or $prevEndsWithPunctuation) {
        if ($word.Length -gt 1) {
            $formattedWord = $word.Substring(0,1).ToUpper() + $word.Substring(1).ToLower()
        }
        elseif ($word.Length -eq 1) {
            $formattedWord = $word.ToUpper()
        }
    }

    return $formattedWord

}

<#
  Cleans and formats a raw title:
  - Removes noise (e.g. edition markers, escape chars, metadata)
  - Applies title casing preserving acronyms and identifiers
  - Normalizes punctuation and spacing
#>
function NormalizeTitle([string]$title) {

    if ([string]::IsNullOrWhiteSpace($title)){
       return $null
    }

    <#
      Normalize title formatting:
      - Replace tabs, newlines, and carriage returns with spaces
      - Collapse multiple spaces into one
      - Remove literal "\n" escape sequences (e.g., "Intro\\nto Math" -> "Intro to Math")
      - Remove "[electronic resource]" tag (e.g., "Calculus [electronic resource]" -> "Calculus")
      - Replace musical sharp symbol (♯) with ASCII "#"
      - Remove leading hyphen-only prefixes (e.g., "- Title" -> "Title")
      - Remove sequences of two or more periods (e.g., "Title...." -> "Title")
      - Strip trailing slash and anything after it (e.g., "Title / Author Name" -> "Title")
      - Normalize long dash sequences to " - " (e.g., "Title--Subtitle" -> "Title - Subtitle")
      - Replace various Unicode dash characters (– — ‒ − ‐) with ASCII hyphen (-)
      - Collapse multiple hyphens into a single hyphen (e.g., "Title---Subtitle" -> "Title-Subtitle")
    #>
    $title = $title -replace '[\r\n\t]', ' '`
                    -replace '\s+', ' '`
                    -replace '\\n', '' `
                    -replace '\[electronic resource\]', ''`
                    -replace '\[global Edition\]', '' `
                    -replace '♯', '#'`
                    -replace '^\s*-\s*', '' `
                    -replace '(?:\.){2,}', '' `
                    -replace ' /.*', '' `
                    -replace '-{2,}', ' - ' `
                    -replace '[\u2013\u2014\u2012\u2212\u2010]', '-' `
                    -replace '-+', '-'

    # Trim leading and trailing spaces, hyphens, and colons from the title (e.g., " -Title: " -> "Title")
    $title = $title.Trim(' -:')

    $title = RemoveEditionFromTitle $title

    # Preload acronym and lowercase words for casing decisions during per-word title normalization
    $acronyms = GetAcronymsInTitle $title
    $lowercaseWords = @( # Words that should remain lowercase in title case (articles, prepositions, conjunctions)
        'a', 'an', 'the', 'and', 'but', 'or', 'for', 'nor', 'on', 'at', 'to', 'from',
        'by', 'of', 'in', 'into', 'with', 'without', 'through', 'over', 'under',
        'above', 'below', 'up', 'down', 'out', 'off', 'between', 'among', 'during',
        'before', 'after', 'since', 'until', 'while', 'as', 'if', 'than', 'that',
        'which', 'who', 'whom', 'whose', 'where', 'when', 'why', 'how', 'what'
    )

    # Convert to title case while preserving acronyms and identifiers
    $finalTitleWords = @()  # Holds processed words for final title reconstruction
    $words = [regex]::Split($title, '\s+')

    foreach ($i in 0..($words.Length - 1)) {

        $word = [string]$words[$i]
        $type = GetWordType $i $words $acronyms $lowercaseWords

        switch ($type) {
            'Acronym' {
                $finalTitleWords += $word
            }
            'StandardIdentifier' {
                $finalTitleWords += GetCanonicalIdentifier $word
            }
            'Version' {
                $finalTitleWords += $word
            }
            'Hyphenated' {
                $finalTitleWords += FormatHyphenatedWord $word $i $words.Length $lowercaseWords
            }
            'Standard' {
                $finalTitleWords += FormatStandardWord $i $words $lowercaseWords
            }
            'Lowercase' {
                $finalTitleWords += $word.ToLower()
            }
            Default {
                throw "Couldn't identify word type for: $word"
            }
        }
    }

    # Join words and clean up punctuation spacing
    $reconstructedTitle = ($finalTitleWords -join ' ').Trim()

    $finalTitle = $reconstructedTitle -replace '\s*:\s*', ': ' `
                                      -replace '\s+-\s+', ' - ' `
                                      -replace '\s*,\s*', ', ' `
                                      -replace '\s*-:\s*', ': ' `
                                      -replace '\s+', ' ' `
                                      -replace ':\s*$', '' `
                                      -replace '\.\s*$', '' `
                                      -replace ',\s*$', '' `
                                      -replace '\s*-\s*$', '' `
                                      -replace '\s\[by\].*', '' `
                                      -replace '^\s+', ''

    # Capitalize the first letter. It may be part of $lowercaseWords. E.g., 'a Hilbert Space Problem Book'
    $finalTitle = $finalTitle.Substring(0,1).ToUpper() + $finalTitle.Substring(1)

    return $finalTitle

}

<#
  Determines whether a given author string is valid by checking against common exclusion patterns
  (e.g., company names, universities, titles, URLs, ...).
  Returns $true if the author appears valid, $false if it matches any exclusion pattern.
#>
function IsValidAuthor($author) {
    $exclusionPatterns = @(
        "(?i)\bfirm\b",
        "(?i)\bUniversity\b",
        "(?i)\bbased( in part)? on\b",
        "(?i)\btranslated by\b",
        "(?i)\bWith Contributions From\b",
        "(?i)\bInstitute Of\b",
        "[\[\]\(\)]"
        "\.(com|org|net|edu|gov|info|biz|co|io|ai|us|uk|de|fr|it|jp|cn)\b"
        "(?i)Founder"
        "(?i)ceo"
        "(?i)\bLaboratory\b"
        "^\s*\d{1,}-?(\d{1,})?$"
    )

    foreach ($pattern in $exclusionPatterns) {
        if ($author -match $pattern) {
            return $false
        }
    }

    return $true
}







# ========== Fetchers ==========

# Represents book metadata to be inserted into the PDF file.
class BookData {
    [string]$Title
    [string]$Subject
    [string[]]$Authors

    BookData() {
        $this.Title = $null
        $this.Subject = $null
        $this.Authors = @()
    }

    BookData([string]$title, [string]$subject, [string[]]$authors) {
        $this.Title = $title
        $this.Subject = $subject
        $this.Authors = $authors
    }
}

<#
# ----- Search Works -----
function GetData_SearchWorks([string]$targetBook) {

    function GetBookUrlFromResults([string]$resultsContent){

        # Find the URL of the first result
        # Group 1 = uri (no base). 2nd group = Book title
        $linksPattern = '(?i)<a[^>]*data-context-href="[^"]*"[^>]*itemprop="name"[^>]*href="([^"]*)"[^>]*>([^<]*)</a>'
        $anchorElements = [regex]::Matches($resultsContent, $linksPattern)

        if ($anchorElements.Count -eq 0) {
            throw "Book not found"
        }

        $hrefAttribute = $anchorElements[0].Groups[1].Value  # Extract the href attribute of the first matching link

        return "https://searchworks.stanford.edu" + $hrefAttribute

    }

    function GetLibrarianViewContent([string]$bookPageContent) {
        $librarianViewLinkPattern = '<a[^>]*id="librarianLink"[^>]*href="([^"]*)"[^>]*>'
        $librarianViewAnchorElement = [regex]::Match($bookPageContent, $librarianViewLinkPattern)
        $librarianUrl = "https://searchworks.stanford.edu$($librarianViewAnchorElement.Groups[1].Value)"
        Write-Host "Opening librarian view: $librarianUrl"

        $librarianViewContent = $null
        try {
            $librarianViewContent = (Invoke-WebRequest $librarianUrl -ErrorAction Stop).Content
        }
        catch {
            throw "Error opening librarian view at '$librarianUrl'. $_"
        }

        return $librarianViewContent
    }

    function GetFolioJsonContent([string]$librarianViewContent) {

        $folioJsonPattern = "<summary[^>]*>\s*FOLIO JSON\s*</summary>\s*<pre[^>]*>(.*?)</pre>"

        $match = [regex]::Match($librarianViewContent, $folioJsonPattern, 'Singleline')

        if (-not $match.Success) {
            throw "Couldn't find FOLIO JSON"
        }

        $jsonText = $match.Groups[1].Value.Trim()
        $decodedJsonText = [System.Net.WebUtility]::HtmlDecode($jsonText)

        return $decodedJsonText | ConvertFrom-Json -ErrorAction Stop

    }

    <# FOLIO JSON content example
     {
        "instance": {
            "title": "Antenna theory and design / Warren L. Stutzman, Gary A. Thiele.",
            "identifiers": [
                {"value": "0471025909 (cloth : alk. paper)", ...},
                {"value": "9780471025900 (cloth : alk. paper)", ...},
                ...
            ],
            "contributors": [
                {"name": "Stutzman, Warren L", "primary": true, ...},
                {"name": "Thiele, Gary A", "primary": false, ...}
            ],
            "classifications": [
                {
                    "classificationNumber": "TK7874.6 .S79 1998",
                    "classificationTypeId": "ce176ace-a53e-4b4d-aa89-725ed7b2edac"
                },
                ...
            ],
            ...
        }
     }
     Notice: The title may be uncorrect
    # >
    function BuildBookData_SearchWorks($folioInstance){

        # Handle title
        if (-not [string]::IsNullOrWhiteSpace($folioInstance.title)) {
            $title = NormalizeTitle $($folioInstance.title)
        }

        # Handle authors
        $authors = ( $folioInstance.contributors |
                     Where-Object {
                        <# "2b94c631-fca9-4892-a730-03ee529ffe2a" person name (not corporate entity
                           e.g, Safari, an O'Reilly Media Company)# >
                        ($_.contributorNameTypeId -eq "2b94c631-fca9-4892-a730-03ee529ffe2a") -and
                        (IsValidAuthor $_.name)
                     } |
                     ForEach-Object { NormalizeAuthor $_.name} )

        # Handle LCC
        $lcc = $null
        $subject = $null
        foreach ($classification in $folioInstance.classifications) {
            $classificationTypeId = "ce176ace-a53e-4b4d-aa89-725ed7b2edac"
            if ($classification.classificationTypeId -eq $classificationTypeId) {
                $candidateLcc = NormalizeLcc $classification.classificationNumber
                $lcc = FilterValidLccs @($candidateLcc) | Select-Object -First 1

                if ($lcc){
                    $subject = GetLineFromSchedule $lcc
                }

                break
            }
        }

        return [BookData]::new($title, $subject, $authors)

    }

    # Search book
    $url = "https://searchworks.stanford.edu/?search_field=search&q=$targetBook"

    Write-Host "Fetching $url" -ForegroundColor Blue

    $resultsContent = $null
    try {
        $resultsContent = (Invoke-WebRequest $url -ErrorAction Stop).Content
    }
    catch {
        throw "Error fetching content from $url. $_"
    }


    [string]$url = GetBookUrlFromResults $resultsContent

    # Open book
    Write-Host "Opening book: $url" -ForegroundColor Blue
    $bookPageContent = $null
    try {
        $bookPageContent = (Invoke-WebRequest $url -ErrorAction Stop).Content
    }
    catch {
        throw "Error opening book page at '$url'. $_"
    }

    # Open librarian view
    $librarianViewContent = GetLibrarianViewContent $bookPageContent

    # Get FOLIO JSON
    $folioContent = GetFolioJsonContent $librarianViewContent

    return BuildBookData_SearchWorks $folioContent.instance
}
#>
# ----- Library of Congress -----
function GetData_LibraryOfCongress([string]$targetBook) {

    <# Response example:
       <zs:searchRetrieveResponse
           xmlns:zs="http://www.loc.gov/zing/srw/">
           ...
           <zs:numberOfRecords>1</zs:numberOfRecords>
           <zs:records>
               <zs:record>
                   <zs:recordData>
                       <mods ...>
                           <titleInfo>
                               <title>Antenna theory and design</title>
                           </titleInfo>
                           <name type="personal" usage="primary">
                               <namePart>Stutzman, Warren L.</namePart>
                           </name>
                           <name type="personal">
                               <namePart>Thiele, Gary A.</namePart>
                           </name>
                           <note type="statement of responsibility">Warren L. Stutzman, Gary A. Thiele.</note>
                           <classification authority="lcc">TK7874.6 .S79 2013</classification>
                           <identifier type="isbn">9780470576649</identifier>
                           <physicalDescription>
                                <form authority="marcform">print</form>
                                ...
                           </physicalDescription>
                           ...
                       </mods>
                   </zs:recordData>
                   <zs:recordPosition>1</zs:recordPosition>
               </zs:record>
           </zs:records>
           <zs:echoedSearchRetrieveRequest>...</zs:echoedSearchRetrieveRequest>
       </zs:searchRetrieveResponse>

       - <titleInfo> may contain a <nonSort> element, representing the prefix (e.g., "A", "The", ...)
         Example:
         <titleInfo>
            <nonSort xml:space="preserve">A </nonSort>
            <title>combinatorial introduction to topology</title>
        </titleInfo>

       - It may also contain a <partName> element, representing the subtitle. Example:
        <titleInfo>
            <title>Java</title>
            <partName>The complete reference</partName>
        </titleInfo>

       - It may contain an alternative title:
        <titleInfo>
            <title>Getting started with sensors</title>
        </titleInfo>
        <titleInfo type="alternative" displayLabel="Cover title :">
            <title>Make --getting started with sensors</title>
        </titleInfo>

       - Names can contain multiple namePart elements:
        <name type="personal">
            <namePart>McGuire, Paul</namePart>
            <namePart type="termsOfAddress">(Computer programmer),</namePart>
            <role>
                <roleTerm type="text">author</roleTerm>
            </role>
        </name>

        - Multiple print formats may be available (book 9780134073545). If so, prompt the user. Example:
        <zs:records>
            <zs:record>
                <zs:recordData>
                    <mods ...>
                        ...
                        <originInfo>
                            <edition>Twelth edition.</edition>
                        </originInfo>
                        <originInfo eventType="publication">
                            <dateIssued>[2017]</dateIssued>
                        </originInfo>
                    </mods>
                </zs:recordData>
            </zs:record>
            <zs:record>
                <zs:recordData>
                    <mods ...>
                        ...
                        <originInfo>
                            <edition>Thirteenth edition.</edition>
                        </originInfo>
                        <originInfo eventType="publication">
                            <dateIssued>[2019]</dateIssued>
                        </originInfo>
                    </mods>
                </zs:recordData>
            </zs:record>
        </zs:records>
     #>
    function BuildBookData_LibraryOfCongress($mod){

        # Handle Lcc
        $lccNode = $mod.classification | Where-Object { $_.authority -eq 'lcc' }
        $candidateLcc = NormalizeLcc ($lccNode.'#text')
        $lcc = FilterValidLccs @($candidateLcc) | Select-Object -First 1
        $subject = $null
        if ($lcc){
            $subject = GetLineFromSchedule $lcc
        }

        # Handle authors
        $personalNames = $mod.name |  Where-Object { $_.type -eq 'personal' }

        $authors = @()
        # For each <name type="personal"> element
        foreach ($personalName in $personalNames){
            $author = $null
            # Find the default namePart (no type) and assigns it as the base name.
            foreach ($namePart in $personalName.namePart){
                if ($null -eq $namePart.type){
                    $author = $namePart
                }
                # Append any termsOfAddress (e.g., "Jr.") if present.
                if ($namePart.type -eq 'termsOfAddress'){
                    $author = "$author $($namePart.'#text')"
                }
            }
            # Validate the combined name and add the normalized result to $authors.
            if (IsValidAuthor $author){
                $authors += NormalizeAuthor $author
            }
        }

        # Handle title
        # Extract the main title, not the alternative one
        $mainTitleInfo = $mod.titleInfo | Where-Object { -not $_.type }
        if (-not [string]::IsNullOrWhiteSpace($mainTitleInfo.title)){
            $title = NormalizeTitle $mainTitleInfo.title
        }


        if ($mainTitleInfo.nonSort) {
            $title = ($mainTitleInfo.nonSort.'#text' + $title).Trim()
        }

        if ($mainTitleInfo.partName){
            $title = $title + ": " + $mainTitleInfo.partName
        }

        return [BookData]::new($title, $subject, $authors)

    }

    # Call API
    $url = "http://lx2.loc.gov:210/lcdb?version=1.1&operation=searchRetrieve&maximumRecords=7&recordSchema=mods&query=$targetBook"

    Write-Host "Fetching $url" -ForegroundColor Blue

    $bookDataResponse = $null
    try {
        $bookDataResponse = Invoke-RestMethod -Uri $url -ErrorAction Stop
    }
    catch {
        throw "Error fetching content from $url. $_"
    }

    $recordCount = [int]$bookDataResponse.searchRetrieveResponse.numberOfRecords
    if ($recordCount -eq 0) {
        throw "Book not found"
    }

    # If both the electronic and print formats are available, use the print format
    $responseMods = @()
    foreach ($record in $bookDataResponse.searchRetrieveResponse.records.record) {
        $mod = $record.recordData.mods
        foreach ($form in $mod.physicalDescription.form) {
            $authority = $null
            if ($form.authority) {
                $authority = $form.authority
            } else {
                $authority = $form.GetAttribute('authority')
            }

            $text = $null
            if ($form.'#text') {
                $text = $form.'#text'
            }
            else {
                $text = $form.InnerText
            }

            if ($authority -eq "marcform" -and $text -eq "print") {
                $responseMods += $mod
                break
            }
        }
    }

    $responseMod = $null    # responseMod to process in BuildBookData_LibraryOfCongress
    if ($responseMods.Count -gt 1) {
        Write-Host "Multiple editions found:" -ForegroundColor Yellow

        for ($i = 0; $i -lt $responseMods.Count; $i++) {
            # Select the originInfo without eventType from the current record
            $mainOriginInfo = $responseMods[$i].originInfo | Where-Object { -not $_.HasAttribute("eventType") }

            $label = $null
            $value = $null

            if ($mainOriginInfo.edition) {
                $label = "Edition"
                $value = $mainOriginInfo.edition
            }
            else {
                # Fallback: use dateIssued
                $pubInfo = $responseMods[$i].originInfo | Where-Object { $_.eventType -eq 'publication' }

                if ($pubInfo.dateIssued) {
                    $label = "Issued on"
                    $value = $pubInfo.dateIssued
                }
                elseif ($responseMods[$i].originInfo.dateIssued) {
                    $label = "Issued on"
                    $value = $responseMods[$i].originInfo.dateIssued.'#text'
                    if ([string]::IsNullOrWhiteSpace($value)){
                        $value = $responseMods[$i].originInfo.dateIssued
                    }
                }
                else {
                    $label = "Edition"
                    $value = "Unknown (check the response)"
                }
            }

            Write-Host "$($i + 1) - $label`: $value"
        }

        do {
            $choice = Read-Host "Pick the correct one (1-$($responseMods.Count))"
        } while ((-not $choice) -or ([int]$choice -lt 1) -or ([int]$choice -gt $responseMods.Count))

        $responseMod = $responseMods[[int]$choice - 1]
    }
    else {
        $responseMod = $responseMods[0]
    }

    return BuildBookData_LibraryOfCongress $responseMod

}

# ----- CERN Library Catalogue -----
function GetData_CernLibraryCatalogue([string]$targetBook) {

    # Not used for authors

    <# Response examples:
       {
         "hits": {
           "hits": [
             {
               "metadata": {
                 "alternative_titles": [{"type": "SUBTITLE", "value": "a modern approach"}],
                 "authors": [{"full_name": "Kurama, Vamsi", "roles": ["AUTHOR"],"type": "PERSON"}, ...],
                 ...,
                 "subjects": [{"scheme": "LOC", "value": "QA76.73.P98"}, ...],
                 "title": "Python programming"
               },
               ...
             },
             ...
           ],
           ...
         },
         ...
       }

       {
         "hits": {
           "hits": [
             {
               "metadata": {
                 "relations": {
                   "edition": [
                     {
                       "record_metadata": {
                         "edition": "10th",
                         "publication_year": "2016",
                         "title": "Fundamentals",
                         ...
                       },
                       "relation_type": "edition"
                     },
                     {
                       "record_metadata": {
                         "edition": "7th",
                         "publication_year": "2005",
                         "title": "Core Java 2",
                         ...
                       },
                       "relation_type": "edition",
                       ...
                     },
                   "multipart_monograph": [
                     {
                       "record_metadata": {
                         "edition": "11th",
                         "publication_year": "2019",
                         "title": "Core Java",
                         ...
                       },
                       "relation_type": "multipart_monograph",
                       "volume": "1"
                     }
                   ]
                 },
                 "title": "Fundamentals"
               }
             }
           ],
         }
       }

    - Authors may contain inaccurate values. Example:
      {"full_name": "Safari, an O'Reilly Media Company", "roles": ["AUTHOR"], "type": "PERSON"}
    #>
    function BuildBookData_CernLibraryCatalogue($metadata){

        # Handle title
        [string]$title = $metadata.title

        # Check if multiple candidate titles exist in relations
        if ($metadata.alternative_titles.Count -gt 0) {
            $title += ": " + $metadata.alternative_titles[0].value
        }
        if (-not [string]::IsNullOrWhiteSpace($title)){
            $title = NormalizeTitle $title
        }

        # Always add the default title first
        $candidateTitles += @($title)

        # Add multipart_monograph variants (series title + volume + part title)
        if ($metadata.relations.multipart_monograph.Count -gt 0) {
            foreach ($relation in $metadata.relations.multipart_monograph) {
                # Extract title from record_metadata
                $seriesTitle = $relation.record_metadata.title
                if ($seriesTitle -isnot [string] -and $seriesTitle.title) {
                    $seriesTitle = $seriesTitle.title
                }

                $volume = $relation.volume
                $composedTitle = "$seriesTitle. Volume $volume - $($metadata.title)"
                if ($candidateTitles -notcontains $composedTitle -and -not [string]::IsNullOrWhiteSpace($composedTitle)) {
                    $candidateTitles += NormalizeTitle $composedTitle
                }
            }
        }

        # Add edition variants
        if ($metadata.relations.edition.Count -gt 1) {
            foreach ($relation in $metadata.relations.edition) {
                $recordTitle = $relation.record_metadata.title
                if ($candidateTitles -notcontains $recordTitle -and -not [string]::IsNullOrWhiteSpace($recordTitle)) {
                    $candidateTitles += NormalizeTitle $recordTitle
                }
            }
        }

        # If there is more than one candidate, prompt the user
        if ($candidateTitles.Count -gt 1) {
            Write-Host "Multiple possible titles found:" -ForegroundColor Yellow
            for ($i = 0; $i -lt $candidateTitles.Count; $i++) {
                Write-Host "$($i + 1). $($candidateTitles[$i])"
            }

            do {
                $choice = Read-Host "Pick the correct one (1-$($candidateTitles.Count))"
            }
            while ((-not $choice) -or ([int]$choice -lt 1) -or ([int]$choice -gt $candidateTitles.Count))

            $title = $candidateTitles[[int]$choice - 1]
        }
        else {
            $title = $candidateTitles[0]
        }

        if (-not [string]::IsNullOrWhiteSpace($title)){
            $title = NormalizeTitle $title
        }

        # Handle LCC
        # Extract first part of each LCC (e.g. extract "TK7874.6" from "TK7874.6 .S79 2012") and get the most common one
        $normalizedLccs = @($metadata.subjects |
                Where-Object scheme -eq 'LOC' |
                ForEach-Object { NormalizeLcc $_.value })

        $lcc = FilterValidLccs ($normalizedLccs) | Select-Object -First 1

        $subject = $null
        if ($lcc){
            $subject = GetLineFromSchedule $lcc
        }

        return [BookData]::new($title, $subject, $null)     # Authors may contain inaccurate values.

    }


    # Call API
    $url = "https://catalogue.library.cern/api/literature/?q=$targetBook"

    Write-Host "Fetching $url" -ForegroundColor Blue

    $bookDataResponse = $null
    try {
        $bookDataResponse = Invoke-RestMethod -Uri $url -ErrorAction Stop
    }
    catch {
        throw "Error fetching content from $url. $_"
    }

    if ($bookDataResponse.hits.hits.Count -eq 0) {
        throw "Book not found"
    }

    $responseMetadata = $bookDataResponse.hits.hits[0].metadata

    return BuildBookData_CernLibraryCatalogue $responseMetadata

}

# ----- Open Library -----
function GetData_OpenLibrary([string]$targetBook){

    <# Example responses:
       {
           "ISBN:9780470576649": {
               "title": "Antenna theory and design",
               "authors": [{ ..., "name": "Warren L. Stutzman" }],
               "by_statement": "Warren L. Stutzman, Gary A. Thiele",
               "classifications": {"lc_classifications": ["TK7874.6 .S79 2012", "TK7871.6", ...], ...},
               "subjects": [{"name": "Antennas (Electronics)", "url": "https://..."},],
               ...
           }
       }
       {
            "ISBN:0393969452": {
                    "title": "C programming",
                    "subtitle": "a modern approach",
                    "authors": [{ ..., "name": "K. N. King"}],
                    "by_statement": "K.N. King.",
                    "classifications": {"lc_classifications": ["QA76.73.C15 K49 1996", ...], ...},
                    ...
                }
       }
       {
            "ISBN:1565924649": {
                "title": "Learning Python",
                "authors": [{ ..., "name": "Mark Lutz"}, { ..., "name": "David Ascher"}],
                "by_statement": "Mark Lutz and David Ascher.",
                "classifications": {"lc_classifications": ["QA76.73.P98 L877 1999", ...], ...},
                ...
            }
       }
       {
          "ISBN:9780470458365": {
            "title": "Advanced engineering mathematics",
            "authors": [{ "name": "Erwin Kreyszig", ...}],
            "by_statement": "Erwin Kreyszig ; in collaboration with Herbert Kreyszig, Edward J. Norminton",
            "classifications": {"lc_classifications": ["QA401 .K7 2011","TA330"]},
            ...
          }
        }
        {
          "ISBN:9781439049327": {
            "title": "A small-scale Approach to organic laboratory techniques",
            "authors": [{"name": "Donald L. Pavia", ...}],
            "by_statement": "Donald L. Pavia ... [et al.].",
            "classifications": {"lc_classifications": ["", "QD261 .I543 2011"]},
          }
        }

       Notice:
       - Some authors may be missing from 'authors' (first example).
       - 'by_statement' may contain additional words (e.g., "and", etc) (third/fourth example) or may be missing
       - 'subjects' field may contain LCC codes in the 'name' subfield.
       - lcc may contain empty strings (last example)
    #>
    function BuildBookData_OpenLibrary($bookDataResponse){

        # Handle title
        [string]$title = $bookDataResponse.title
        if ($bookDataResponse.subtitle) {
            $title += ": " + $bookDataResponse.subtitle
        }

        if (-not [string]::IsNullOrWhiteSpace($title)) {
            $title = NormalizeTitle $title
        }

        # Handle authors
        $authorsStructured = @()
        $authorsByStatement = @()

        if ($bookDataResponse.authors) {
            $authorsStructured = @($bookDataResponse.authors | ForEach-Object { $_.name.Trim() })
        }

        if ($bookDataResponse.by_statement) {
            $splitPattern = ',| and | & |;'
            $authorsByStatement = @(
            $bookDataResponse.by_statement -split $splitPattern |
                    ForEach-Object { NormalizeAuthor($_) } |
                    Where-Object { $_ -ne "" }
            )
        }

        $authors = ($authorsStructured + $authorsByStatement) |
                Where-Object { IsValidAuthor $_ } |
                ForEach-Object { NormalizeAuthor $_ } |
                Sort-Object -Unique

        # Handle LCCs
        $lcc = $null
        $subject = $null

        if ($bookDataResponse.classifications.lc_classifications){

            $normalizedLccs = $bookDataResponse.classifications.lc_classifications | ForEach-Object { NormalizeLcc $_ }

            $lcc = FilterValidLccs $normalizedLccs | Select-Object -First 1

            $subject = $null
            if ($lcc){
                $subject = GetLineFromSchedule $lcc
            }
        }

        if (-not $lcc) {
            $lcc = $bookDataResponse.subjects |
                    Where-Object { IsLcc $_.name } |
                    Select-Object -ExpandProperty name -First 1
        }

        return [BookData]::new($title, $subject, $authors)

    }

    # Call API
    $url = "https://openlibrary.org/api/books?bibkeys=ISBN:$targetBook&format=json&jscmd=data"

    Write-Host "`nFetching $url" -ForegroundColor Blue

    $response = $null
    try {
        $response = Invoke-RestMethod -Uri $url -ErrorAction Stop
    }
    catch {
        throw "Error fetching content from $url. $_"
    }

    $bookDataResponse = $($response."ISBN:$targetBook");

    if (-not $bookDataResponse) {
        throw "Book not found"
    }

    return BuildBookData_OpenLibrary $bookDataResponse

}

# ----- Google API -----
function GetData_GoogleApis([string]$targetBook){

    # Google Books API does not provide LCC data.

    <#
    Response example https://www.googleapis.com/books/v1/volumes?q=isbn:9781801828642:
    {
     "kind": "books#volumes",
     "totalItems": 1000000,
     "items": [
       {
         "volumeInfo": {
           "title": "Coding",
           "subtitle": "3 Books in 1: \"Python Coding and Programming + Linux for Beginners + Learn Python Programming\"",
           "authors": [ "Michael Clark", "Michael Learn" ],
           "industryIdentifiers": [
             { "type": "ISBN_10", "identifier": "1801828644" },
             { "type": "ISBN_13", "identifier": "9781801828642" }
           ],
           ...
         },
       },
       ...
     ]
    }
    #>
    function BuildBookData_GoogleApis($responseVolumeInfo){

        # Handle title
        [string]$title = $responseVolumeInfo.title
        if ($responseVolumeInfo.subtitle){
            $title += ": " + $responseVolumeInfo.subtitle
        }

        if (-not [string]::IsNullOrWhiteSpace($title)) {
            $title = NormalizeTitle $title
        }

        # Handle authors
        [string[]]$authors = $responseVolumeInfo.authors |
                Where-Object { IsValidAuthor $_ } |
                ForEach-Object { NormalizeAuthor $_ }

        return [BookData]::new($title, $null, $authors)

    }


    # Call API
    $url = "https://www.googleapis.com/books/v1/volumes?q=isbn:$targetBook"

    Write-Host "Fetching $url" -ForegroundColor Blue

    try {
        $bookDataResponse = Invoke-RestMethod -Uri $url -ErrorAction Stop
    }
    catch {
        throw "Error fetching content from $url"
    }

    if ($bookDataResponse.totalItems -eq "0"){
        throw "Book not found"
    }

    $responseVolumeInfo = $bookDataResponse.items[0].volumeInfo

    # Return book metadata
    return BuildBookData_GoogleApis $responseVolumeInfo

}

# ----- CiNii -----
function GetData_CiNii([string]$targetBook){

    function BuildBookData_CiNii($bookPageContent) {

        # Handle Title
        $titleMatches = [regex]::Matches($bookPageContent, '(?i)<dt class="key">Title</dt>\s*<dd class="value">"([^"]*)"</dd>')

        if ($titleMatches.Count -eq 0){
            throw "Book not found"
        }

        $title = [System.Net.WebUtility]::HtmlDecode($titleMatches[0].Groups[1].Value)
        if (-not [string]::IsNullOrWhiteSpace($title)) {
            $title = NormalizeTitle $title
        }

        # Handle LCC
        $lccMatches = [regex]::Matches($bookPageContent, '(?i)<li>LCC:[^>]*>([^>]*)</a>')

        $lcc = $null
        if ($lccMatches.Count -gt 0){
            $lcc = NormalizeLcc $lccMatches[0].Groups[1].Value
            $lcc = FilterValidLccs @($lcc) | Select-Object -First 1
        }

        $subject = $null
        if ($lcc){
            $subject = GetLineFromSchedule $lcc
        }

        # Handle Authors
        $authorsMatches = [regex]::Matches($bookPageContent, '(?i)<dt class="key">Statement of Responsibility</dt>\s*<dd class="value">([^<]*)</dd>')

        $authors = $null
        if ($authorsMatches.Count -gt 0){
            $authors = $authorsMatches[0].Groups[1].Value
        }


        if ($authors) {
            $authors = $authors -replace '(based( in part)? on lectures|translated|edited|foreword|introduction|commentary)\s+by\s+[^;]*(;|$)', ''

            $splitPattern = ' with |,| and | & |;'
            $authors = @(
            $authors -split $splitPattern |
                    Where-Object { IsValidAuthor $_ } |
                    ForEach-Object { NormalizeAuthor $_ } |
                    Where-Object { $_ -ne "" }
            )
        }

        return [BookData]::new($title, $subject, $authors)

    }

    function GetBookUrlFromResults([string]$resultsContent){

        # Find the URL of the first result
        # Group 1 = uri (no base). 2nd group = Book title
        $linksPattern = '(?i)<a[^>]*class="taggedlink"[^>]*href="/crid/([^"]*)"[^>]*>([^<]*)</a>'
        $anchorElements = [regex]::Matches($resultsContent, $linksPattern)

        if ($anchorElements.Count -eq 0) {
            throw "Book not found"
        }

        $cridNumber = $anchorElements[0].Groups[1].Value  # Extract the href attribute of the first matching link

        return "https://cir.nii.ac.jp/crid/$cridNumber`?lang=en"
    }

    # Search book
    $url = "https://cir.nii.ac.jp/all?q=$targetBook"
    Write-Host "Fetching $url" -ForegroundColor Blue

    $resultsContent = $null
    try {
        $resultsContent = (Invoke-WebRequest $url -ErrorAction Stop).Content
    }
    catch {
        throw "Error fetching content from $url. $_"
    }

    [string]$url = GetBookUrlFromResults $resultsContent

    # Open book
    Write-Host "Opening book: $url" -ForegroundColor Blue


    $bookPageContent = $null
    try {
        $bookPageContent = (Invoke-WebRequest $url -ErrorAction Stop).Content
    }
    catch {
        throw "Error opening book page at '$url'. $_"
    }

    return BuildBookData_CiNii $bookPageContent

}

# ----- Yale Library -----
function GetData_Yale([string]$targetBook){

    function GetBookUrlFromResults([string]$resultsContent){

        # Find the URL of the first result
        # Group 1 = uri (no base). 2nd group = Book title
        $linksPattern = '<div class=''result_title''><a\s*data-context-href="/catalog/\d{7,}/[^;]*;document_id=\d{7,}[^"]*"\s*href="/catalog/(\d{7,})[^>]*>[^<]*</a></div>'

        $anchorElements = [regex]::Matches($resultsContent, $linksPattern)

        if ($anchorElements.Count -eq 0) {
            throw "Book not found"
        }

        $bookId = $anchorElements[0].Groups[1].Value  # Extract the href attribute of the first matching link

        return "https://search.library.yale.edu/catalog/$bookId"
    }

    function BuildBookData_Yale([xml]$marcXml) {
        # Handle title
        $tag245 = $marcXml.record.datafield | Where-Object { $_.tag -eq '245' }
        $primaryTitle = ($tag245.subfield | Where-Object { $_.code -eq 'a' }).'#text'
        $secondaryTitle = ($tag245.subfield | Where-Object { $_.code -eq 'b' }).'#text'
        $title = NormalizeTitle ($primaryTitle + $secondaryTitle)

        # Handle LCC
        $tag50 = $marcXml.record.datafield | Where-Object { $_.tag -eq '050' }
        $lcc = NormalizeLcc ($tag50.subfield | Where-Object { $_.code -eq 'a' }).'#text'

        $subject = $null
        if ($lcc){
            $subject = GetLineFromSchedule $lcc
        }

        # Handle authors
        $tag100 = $marcXml.record.datafield | Where-Object { $_.tag -eq '100' }
        $primaryAuthor = ($tag100.subfield | Where-Object { $_.code -eq 'a' }).'#text'

        $tags700 = $marcXml.record.datafield | Where-Object { $_.tag -eq '700' }
        $secondaryAuthors = $tags700 | ForEach-Object {($_.subfield | Where-Object { $_.code -eq 'a' }).'#text'}

        $authors = @($secondaryAuthors) + @($primaryAuthor)

        $authors = $authors | Where-Object { IsValidAuthor $_ } |
                ForEach-Object { NormalizeAuthor $_ } |
                Where-Object { $_ -ne "" }

        return [BookData]::new($title, $subject, $authors)
    }

    # Search book
    $url = "https://search.library.yale.edu/quicksearch?commit=Search&q=$targetBook"
    Write-Host "Fetching $url" -ForegroundColor Blue

    $resultsContent = $null
    try {
        $resultsContent = (Invoke-WebRequest $url -ErrorAction Stop).Content
    }
    catch {
        throw "Error fetching content from $url. $_"
    }

    [string]$url = (GetBookUrlFromResults $resultsContent) + ".marcxml"

    # Open book
    Write-Host "Opening book: $url" -ForegroundColor Blue


    $marcXml = $null
    try {
        [xml]$marcXml = (Invoke-WebRequest $url -ErrorAction Stop).Content
    }
    catch {
        throw "Error opening book page at '$url'. $_"
    }

    return BuildBookData_Yale $marcXml
}

# ----- Prospector -----
function GetData_Prospector([string]$targetBook) {

    # Not used for authors

    function GetBookUrlFromResults([string]$resultsContent){
        $resultPattern = '(?i)<h2 class="briefcitTitle">\s*<a\s*href="([^"]*)">'
        $resultsMatches = [regex]::Matches($resultsContent, $resultPattern)

        # Return first result URL
        return "https://prospector.coalliance.org/$($resultsMatches[0].Groups[1].Value)"

    }

    function BuildBookData_Prospector($bookPageContent) {

        # Handle title
        $titlePattern = '(?i)<td[^>]*>\s*Title\s*</td>\s*<td[^>]*>\s*<strong>([^>]*)</strong>\s*</td>'
        $titleMatches = [regex]::Matches($bookPageContent, $titlePattern)
        $title = $null
        if ($titleMatches.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($titleMatches[0].Groups[1].Value)){
            $title = NormalizeTitle $titleMatches[0].Groups[1].Value
        }

        # Handle LCC
        $lccPattern = '(?i)<td[^>]*>\s*LC\s*#\s*</td>\s*<td[^>]*>\s*([^>]*)\s*</td>'
        $lccMatches = [regex]::Matches($bookPageContent, $lccPattern)

        $lcc = $null
        if ($lccMatches.Count -gt 0){
            $lcc = NormalizeLcc $lccMatches[0].Groups[1].Value
            $lcc = FilterValidLccs @($lcc) | Select-Object -First 1
        }

        $subject = $null
        if ($lcc){
            $subject = GetLineFromSchedule $lcc
        }

        return [BookData]::new($title, $subject, $null)

    }

    # Search book
    $url = "https://prospector.coalliance.org/search/?searchtype=X&SORT=D&searcharg=$targetBook"
    Write-Host "Fetching $url" -ForegroundColor Blue

    $resultsContent = $null
    try {
        $resultsContent = (Invoke-WebRequest $url -ErrorAction Stop).Content
    }
    catch {
        throw "Error fetching content from $url. $_"
    }

    if ($resultsContent -match "1 result found"){
        $bookPageContent = $resultsContent
    }
    elseif ($resultsContent -match "NO ENTRIES FOUND") {
        throw "Book not found"
    }
    else {
        [string]$url = GetBookUrlFromResults $resultsContent

        # Open book
        Write-Host "Opening book: $url" -ForegroundColor Blue

        $bookPageContent = $null
        try {
            $bookPageContent = (Invoke-WebRequest $url -ErrorAction Stop).Content
        }
        catch {
            throw "Error opening book page at '$url'. $_"
        }
    }

    return BuildBookData_Prospector $bookPageContent

}







# ========== Main operations Utils ==========

<#
  Displays a numbered menu of options and lets the user select one.
  Adds a "Custom input..." option to allow free text entry.
  Returns the selected option or the user-provided custom input.
#>
function ChooseSingleOption([string[]]$options) {

    $options += "Custom input..."

    # Show menu
    for ($i = 0; $i -lt $options.Count; $i++) {
        Write-Host ("$(($i + 1)). $($options[$i])")
    }

    Write-Host ""

    # Make choice
    $choice = ""
    $enterPressed = $false
    do {
        $key = [System.Console]::ReadKey($true)

        switch ($key.Key) {
            'Escape' {
                Write-Host "`nAborting..."
                return $null
            }
            'Enter' {
                if ($choice -eq ""){
                    continue
                }
                $enterPressed = $true
            }
            'Backspace' {
                if ($choice.Length -gt 0) {
                    $choice = $choice.Substring(0, $choice.Length - 1)
                    [System.Console]::Write("`b `b")
                }
            }
            default {
                if ($key.KeyChar -match '\d') {
                    $testChoice = $choice + $key.KeyChar

                    if ([int]$testChoice -ge 1 -and [int]$testChoice -le $options.Count) {
                        $choice = $testChoice
                        [System.Console]::Write($key.KeyChar)
                    }
                    # If invalid number, ignore the keystroke
                }
            }
        }
    } while (-not $enterPressed)

    Write-Host ""

    # Handle choice
    if ($options[$choice - 1] -eq "Custom input...") {
        return (Read-Host "Enter custom input")
    }
    else {
        return $options[$choice - 1]
    }

}

<#
  Displays the book's LCC, Title, and Authors returned by a fetcher.
  Prints "Not Found" when a field is missing.

  Special cases:
   - LCC: Only marked missing if the fetcher is not GoogleAPI (GoogleAPI rarely provides LCC data).
   - Authors: Only marked missing if the fetcher is not CernLibraryCatalogue or Prospector, since these sources
     often omit or return unclean author data.
#>
function LogBookDataResponse([BookData]$bookData, [string]$fetcherName, [string]$color) {

    # LCC
    if ($bookData.Subject) {
        Write-Host "LCC: " -NoNewline
        Write-Host "$($bookData.Subject)" -ForegroundColor $color
    }
    elseif ($fetcherName -ne "GoogleAPI") {
        Write-Host "LCC: " -NoNewline
        Write-Host "Not Found" -ForegroundColor Yellow
    }

    # Title
    if ($bookData.Title) {
        Write-Host "Title: " -NoNewline
        Write-Host "$($bookData.Title)" -ForegroundColor $color
    }
    else {
        Write-Host "Title: " -NoNewline
        Write-Host "Not Found" -ForegroundColor Yellow
    }

    # Authors
    if ($bookData.Authors -and ($bookData.Authors.Count -gt 0)) {
        Write-Host "Authors: " -NoNewline
        Write-Host "$($bookData.Authors -join '; ')" -ForegroundColor $color
    }
    elseif (($fetcherName -ne "CernLibraryCatalogue") -and ($fetcherName -ne "Prospector")) {
        Write-Host "Authors: " -NoNewline
        Write-Host "Not Found" -ForegroundColor Yellow
    }

}

<#
  Prompts the user to manually enter an LCC. Allows typing 'g' to open a Google search for help.
  Repeats until a valid LCC or an empty input is provided.
#>
function InsertLCCManually([string]$targetBook) {

    $result = $null
    do {
        $result = Read-Host "LCC (enter 'g' to Google)"
        if ($result -eq "g") {
            # Query example: "LC" "1565924649"
            $query = "`"LC`" `"$targetBook`""
            $encodedQuery = [uri]::EscapeDataString($query)
            Start-Process msedge "https://www.google.com/search?q=$encodedQuery"
        }
        $result = $result.ToUpper()
        if ($result -ne "" -and $result -ne "G" -and -not (IsLcc $result)) {
            Write-Host "Enter a valid LCC" -ForegroundColor Yellow
        }
    } while ($result -ne "" -and $result -eq "G" -and -not (IsLcc $result))

    if ($result -ne ""){
        $result = GetLineFromSchedule $result
    }

    return $result

}

<#
  Resolves the final LCC or title from a list of candidates.
  The 'fieldName' parameter specifies which field is being resolved ("LCC" or "Title").
  If no candidates exist, prompts the user to enter a value manually.
  If one candidate occurs most frequently, selects it automatically.
  If there is a tie, asks the user to choose one.
#>
function SelectMostFrequentOrPrompt([string[]]$candidates, [string]$fieldName, [string]$isbn) {

    # Candidates may include empty strings because $null values are cast to [string] and become '' when passed in
    $candidates = $candidates | Where-Object { -not [string]::IsNullOrEmpty($_) }

    # Group identical candidates and sort them by frequency (most common first)
    $groupedCandidatesByFreq = $candidates | Group-Object | Sort-Object Count -Descending

    # No LCCs or titles
    $hasCandidates = $candidates.Count -gt 0

    $hasWinner = $null
    # Either a single candidate exists, or the top candidate has a higher frequency than the runner-up
    if ($groupedCandidatesByFreq.Count -eq 1) {
        $hasWinner = $true
    }
    elseif ($hasCandidates -and $groupedCandidatesByFreq.Count -ge 2) {
        $hasWinner = $groupedCandidatesByFreq[0].Count -ne $groupedCandidatesByFreq[1].Count
    }

    $result = $null
    if (-not $hasCandidates) {
        switch ($FieldName) {
            "LCC" {
                Write-Host "Couldn't find any $fieldName. Insert it manually (press Enter to skip)."
                $result = InsertLCCManually $isbn
            }

            "Title" {
                Write-Host "Couldn't find any title. Insert it manually (press Enter to skip)."
                $result = Read-Host "Title"
            }
        }
    }
    elseif ($hasWinner) {
        $result = $groupedCandidatesByFreq[0].Name
    }
    else {
        Write-Host "Tie for $fieldName candidates. Select one:"
        $result = ChooseSingleOption ($candidates | Select-Object -Unique)
        Write-Host $divisionLineLong
    }

    return $result

}

# Removes all characters from the input string that are invalid in Windows file names.
function RemoveInvalidFileNameChars([string]$Name) {

    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars() -join ''
    $regex = "[$([Regex]::Escape($invalidChars))]"

    return ($Name -replace $regex, '')

}

# Extracts and returns text from the specified page range of a PDF file as a single string.
function GetPdfContentInPageRange([string]$filePath, [int[]]$range) {

    if (-not (Test-Path $filePath)) {
        throw "File not found: $filePath"
    }

    $pdfReader = [iText.Kernel.Pdf.PdfReader]::new($filePath)
    $pdfDocument = [iText.Kernel.Pdf.PdfDocument]::new($pdfReader)

    $stringBuilder = New-Object System.Text.StringBuilder

    foreach ($pageNumber in $range) {
        $page = $pdfDocument.GetPage($pageNumber)
        $strategy = [iText.Kernel.Pdf.Canvas.Parser.Listener.SimpleTextExtractionStrategy]::new()
        $text = [iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor]::GetTextFromPage($page, $strategy)
        [void]$stringBuilder.AppendLine($text)  # [void] prevents the output from being printed to the console
    }

    $pdfDocument.Close()

    return $stringBuilder.ToString()

}

<#
  Extracts all valid ISBN-10 and ISBN-13 values from a text string.
  Supports both prefixed (ISBN, ISBN-10, ISBN-13) and standalone formats.
  Returns an array of valid ISBNs or $null if none found.
#>
function GetIsbnsFromString([string]$text) {

    if (-not $text) {
        return $null
    }

    # Group 1 = ISBN (for standalone cases Groups[0] (full match) = Groups[1] (first and only capture))
    $isbnPatterns = @(
        "(?:ISBN[-:]?[ \t]*(?:13[-:]?[ \t]*)?)?(97[89](?:[- \t]?\d){10})",             # ISBN-13 with prefix
        "(?:ISBN[-:]?[ \t]*(?:10[-:]?[ \t]*)?)?(\d(?:[- \t]?\d){8}(?:[- \t]?[\dX]))"  # ISBN-10 with prefix
    )

    $isbns = @()
    foreach ($pattern in $isbnPatterns) {
        # $isbnPatterns expect uppercase "ISBN" but input may vary
        $isbnMatches = [regex]::Matches($text, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        if ($isbnMatches.Count -gt 0) {
            $isbnMatches = $isbnMatches | ForEach-Object { NormalizeIsbn $_.Groups[1].Value } | Select-Object -Unique
            foreach ($isbnMatch in $isbnMatches) {
                $isbnCandidate = $isbnMatch

                if (IsIsbn $isbnCandidate) {
                    Write-Host "Found ISBN: $isbnCandidate" -ForegroundColor Magenta
                    $isbns += $isbnCandidate
                }
            }
        }
    }

    if ($isbns.Count -eq 0) {
        Write-Host "No ISBNs found." -ForegroundColor Yellow
    }

    return $isbns

}

<#
  Attempts to extract valid ISBNs from a PDF file by scanning its pages in stages. Early pages
  (cover, copyright, introduction, ...) are more likely to contain the correct ISBN, so they are
  prioritized to reduce false positives and improve performance.
#>
function GetIsbnsFromPdf([string]$pdf) {

    Write-Host $divisionLineLong
    Write-Host "Finding ISBN in $pdf..."

    $extractionAttempts = @(
        @{ Pages = 1..3 },
        @{ Pages = 4..6 },
        @{ Pages = 7..10 }
    )

    $isbns = @()

    foreach ($attempt in $extractionAttempts) {
        Write-Host "Reading PDF pages $($attempt.Pages -join ', ')..."

        $content = GetPdfContentInPageRange $pdf $attempt.Pages

        if ($content.Trim().Length -gt 0) {
            $attemptIsbns = GetIsbnsFromString $content
            if ($attemptIsbns.Count -gt 0) {
                $isbns += $attemptIsbns
            }
        }
        else {
            Write-Host "Failed reading pages $($attempt.Pages -join ', ')" -ForegroundColor yellow
        }
    }

    if ($isbns.Count -eq 0) {
        throw "No valid ISBN found in '$pdf' within pages 1-10. Unable to process PDF."
    }

    return $isbns

}

<#
  Checks whether the output PDF size is reasonable compared to the original. Prompts the user if the file is
  more than 30% smaller to confirm it is safe to save. Throws an error if the output is empty.
#>
function ConfirmPdfSizeIntegrity([string]$inputPdf, [byte[]]$bytes){
    $estimatedSize = $bytes.Length
    $inputFileSize = (Get-Item $inputPdf).Length
    $lossTolerance = 0.3
    $minAcceptableSize = $inputFileSize * (1 - $lossTolerance)

    if ($estimatedSize -eq 0) {
        throw "No data could be written to the PDF output."
    }

    if ($estimatedSize -lt $minAcceptableSize) {
        Write-Host "Output is more than $($lossTolerance*100)% smaller:"  -ForegroundColor Yellow
        Write-Host "- Original: $([math]::Round($inputFileSize/1MB, 2))MB" -ForegroundColor Yellow
        Write-Host "- Estimated: $([math]::Round($estimatedSize/1MB, 2))MB" -ForegroundColor Yellow
        Write-Host "Proceed with saving (y/n)? " -NoNewline -ForegroundColor Yellow

        $response = Read-Host
        if ($response -eq 'n') {
            throw "Operation cancelled."
        }
    }
}

<#
  Updates a PDF with the given title, authors, and subject, then saves it using the title as the output file
  name (after removing invalid characters). Verifies input file existence, writes changes in memory, checks
  for suspicious size reduction (prompts user if needed), and handles corrupted or empty output.
#>
function SetBookDataAndSave([string]$inputPdf, [string]$title, [string[]]$authors, [string]$subject) {

    if (-not (Test-Path $inputPdf)) {
        throw "File not found: $inputPdf"
    }

    # Prepare a memory stream
    $memoryStream = New-Object System.IO.MemoryStream
    $pdfReader = [iText.Kernel.Pdf.PdfReader]::new($inputPdf)
    [void]$pdfReader.SetUnethicalReading($true)
    $pdfWriter = [iText.Kernel.Pdf.PdfWriter]::new($memoryStream)

    try {
        $pdfDocument = [iText.Kernel.Pdf.PdfDocument]::new($pdfReader, $pdfWriter)
    } catch {
        $pdfReader.Close()
        throw "Failed to open PDF. The file may be corrupted or malformed. Try repairing it with https://www.ilovepdf.com/repair-pdf and try again"
    }

    $info = $pdfDocument.GetDocumentInfo()

    # Clear existing metadata
    $emptyXmp = [iText.Kernel.XMP.XMPMetaFactory]::Create()
    $pdfDocument.SetXmpMetadata($emptyXmp)

    [void]$info.SetTitle("")
    [void]$info.SetAuthor("")
    [void]$info.SetSubject("")
    [void]$info.SetKeywords("")
    [void]$info.SetCreator("")
    [void]$info.SetProducer("")
    [void]$info.SetMoreInfo("CreationDate", $null)
    [void]$info.SetMoreInfo("ModDate", $null) # Necessary to clear author field

    # Set new metadata
    if ($title) {
        [void]$info.SetTitle($title)
    }

    if ($authors -and $authors.Count -ge 1) {
        [void]$info.SetAuthor($authors -join "; ")
    }

    if ($subject) {
        [void]$info.SetSubject($subject)
    }

    try {
        # Finalize all changes, flush data to the memory stream, and release iText resources.
        # The MemoryStream is not disposed.
        $pdfDocument.Close()
    } catch {
        throw "Failed to finalize PDF. The file may be corrupted or malformed. Try repairing it with https://www.ilovepdf.com/repair-pdf and try again"
    }
    $bytes = $memoryStream.ToArray()
    $memoryStream.Dispose()
    $pdfReader.Close()


    ConfirmPdfSizeIntegrity $inputPdf $bytes

    # Build output file name
    $outputFileName = RemoveInvalidFileNameChars $title

    if ($outputFileName -eq "") {
        Write-Host  "Missing book title. Enter a name for the output PDF before proceeding: " -NoNewline -ForegroundColor Yellow
        $outputFileName = Read-Host
    }
    $outputPdf = "$outputDirectorySuccess\$outputFileName.pdf"

    if (-not (Test-Path -Path $outputDirectorySuccess)) {
        Write-Host "Directory $outputDirectorySuccess doesn't exist. Creating directory..."
        New-Item -ItemType Directory -Path $outputDirectorySuccess | Out-Null
        Write-Host "Directory created." -ForegroundColor Green
    }

    # Generate unique filename if one already exists
    $baseName = $outputFileName
    $counter = 1
    $outputPdf = "$outputDirectorySuccess\$outputFileName.pdf"
    while (Test-Path $outputPdf) {
        $outputFileName = "$baseName ($counter)"
        $outputPdf = "$outputDirectorySuccess\$outputFileName.pdf"
        $counter++
    }

    # Save the modified PDF bytes to disk
    [System.IO.File]::WriteAllBytes($outputPdf, $bytes)

    Write-Host "PDF saved at: $outputPdf" -ForegroundColor Green

}

<#
  Groups similar author names into variant groups.

  Authors are considered variants of the same person if:
  1) Their names match when normalized (diacritics removed)
  2) Their last names are identical after normalization
  3) First name comparison:
     - If either first name ends with a period (e.g., "J."), their prefixes (excluding the period) must match
     - If neither name ends with a period, the full first tokens must be equal

  Returns:
    A list of arrays, each containing names considered variants of the same author.
#>

function GroupBySameAuthor([string[]]$authors) {

    function SameAuthor([String]$authorA, [String]$authorB){
        # Strip diacritics
        $normalizedA = $authorA.Normalize([Text.NormalizationForm]::FormD) -replace '\p{M}', ''
        $normalizedB = $authorB.Normalize([Text.NormalizationForm]::FormD) -replace '\p{M}', ''

        if ($normalizedA -eq $normalizedB){
            return $true
        }

        $authorATokens = $normalizedA -split "\s+"
        $authorBTokens = $normalizedB -split "\s+"

        $lastAToken = $authorATokens[-1]
        $lastBToken = $authorBTokens[-1]

        if ($lastAToken -ne $lastBToken) {
            return $false
        }

        $firstAToken = $authorATokens[0]
        $firstBToken = $authorBTokens[0]

        <#
          If either first token is an initial (ends with a period), compare their prefixes (excluding the period).
          If the prefixes differ, assume the authors are not the same.
        #>
        $min = [Math]::Min($firstAToken.Length, $firstBToken.Length)
        $prefixA = $firstAToken.Substring(0, ($min - 1))
        $prefixB = $firstBToken.Substring(0, ($min - 1))
        if (($firstAToken -match "\.$" -or $firstBToken -match "\.$") -and ($prefixA -ne $prefixB)) {
            return $false
        }

        if ($firstAToken -notmatch "\.$" -and $firstBToken -notmatch "\.$" -and
                ($firstAToken -ne $firstBToken)){
            return $false
        }

        return $true
    }

    $variantGroups = New-Object System.Collections.ArrayList
    $authors = $authors | Sort-Object

    foreach ($author in $authors) {
        $foundGroup = $false

        # If this author belongs to an existing group, add it to that group
        for ($i = 0; $i -lt $variantGroups.Count; $i++) {
            if (SameAuthor $variantGroups[$i][0] $author) {
                $variantGroups[$i].Add($author) | Out-Null
                $foundGroup = $true
                break
            }
        }

        # If no matching group found, create a new group
        if (-not $foundGroup) {
            $newGroup = New-Object System.Collections.ArrayList
            $newGroup.Add($author) | Out-Null
            $variantGroups.Add($newGroup) | Out-Null
        }
    }

    return ,$variantGroups
}

# Selects the "best" author from a list by choosing the longest name.
function SelectBestAuthor([string[]]$authors) {

    if (-not $authors -or $authors.Count -eq 0) {
        return $null
    }

    $result = $authors | Sort-Object { $_.Length } -Descending | Select-Object -First 1
    return $result

}

# Prompts the user to manually enter authors when none were found. Returns the list of entered names.
function InsertAuthorsManually {

    Write-Host "Couldn't find any authors. Insert them manually (press Enter to skip)."
    $manualAuthors = [System.Collections.Generic.List[string]]::new()
    $i = 1

    while ($true) {
        $author = Read-Host "Author $i"
        if ([string]::IsNullOrWhiteSpace($author)) {
            break
        }
        $manualAuthors.Add($author)
        $i++
    }

    return $manualAuthors

}

<#
  Used to select the final list of authors when multiple candidates are present.
  If:
   1) More than half of the author-capable fetchers returned author data,
   2) An author was returned by only one fetcher, and
   3) That author belongs to a singleton variant group,
  it prompts the user to confirm whether to include the author.
  Otherwise, all candidate authors are included automatically since no clear
  consensus can be determined.
#>
function SelectFinalAuthors([hashtable]$fetcherToBookDataMap, $variantGroups, [string[]]$candidateAuthors){

    $authors = [System.Collections.Generic.List[string]]::new() # Authors to be returned
    $excludedAuthors = @()  # Track authors permanently rejected. Only used for logging.
    <# CernLibraryCatalogue, and Prospector do not provide author data.
       Total fetchers = 8 = $fetcherToBookDataMap.Keys.Count
       Fetchers capable of returning authors = 6 = $fetcherToBookDataMap.Keys.Count - $nonAuthorFetchers.Count
     #>
    $nonAuthorFetchers = @("CernLibraryCatalogue", "Prospector")
    $threshold = [math]::floor(($fetcherToBookDataMap.Keys.Count - $nonAuthorFetchers.Count)/2)

    # Get fetchers that returned at least one author
    $fetchersWithAuthors = $fetcherToBookDataMap.GetEnumerator() | Where-Object { $_.Value.Authors }

    if ($fetchersWithAuthors.Count -ge $threshold){
        foreach ($candidateAuthor in $candidateAuthors){
            # Count the fetchers that returned $author
            $authorOccurrences = ($fetcherToBookDataMap.GetEnumerator() | Where-Object { $_.Value.Authors -contains $candidateAuthor }).Count

            # Check if $author belongs to a singleton variant group
            $isAuthorInSingletonGroup = ($variantGroups | Where-Object { $_ -contains $candidateAuthor }).Count -eq 1


            if ($authorOccurrences -eq 1 -and $isAuthorInSingletonGroup){
                $choice = $null
                do {
                    Write-Host "Author '$candidateAuthor' may be incorrect (found in 1/$($fetchersWithAuthors.Count) sources). Include it? (y/n)"
                    # Show remaining authors excluding the current one and any authors previously rejected
                    Write-Host "Other authors:"
                    $authorsToShow = $candidateAuthors | Where-Object { $_ -ne $candidateAuthor -and
                                                                        $excludedAuthors -notcontains $_}
                    $authorsToShow | ForEach-Object { Write-Host "- $_" }
                    $choice = Read-Host
                }
                while ($choice.ToLower() -ne "y" -and $choice.ToLower() -ne "n")

                if ($choice -eq "y") {
                    $authors.Add($candidateAuthor)
                }
                else {
                    $excludedAuthors += $candidateAuthor
                }
            }
            else{
                # Automatically include the author if no confirmation is needed
                $authors.Add($candidateAuthor)
            }
        }
    }
    else {
        $candidateAuthors | ForEach-Object { $authors.Add($_) }
    }

    return $authors

}

<#
  Returns the longest title that appears more than once in the input array.
  If no title is repeated, returns $null. Only considers non-empty strings.
#>
function SelectLongestRepeatedTitle([string[]]$candidates) {

    # Filter out null or empty strings
    $candidates = $candidates | Where-Object { -not [string]::IsNullOrEmpty($_) }

    if (-not $candidates) {
        return $null
    }

    # Find maximum title length
    $maxLength = ($candidates | Measure-Object -Property Length -Maximum).Maximum

    # Select titles having that maximum length
    $longestTitles = $candidates | Where-Object { $_.Length -eq $maxLength }

    # Group by identical title
    $titleGroups = $longestTitles | Group-Object | Where-Object { $_.Count -ge 2 }

    # If no repeated title exists, return null
    if (-not $titleGroups) {
        return $null
    }

    # Return the longest repeated title (first if multiple)
    return $titleGroups[0].Name

}

<#
  Attempts to determine the most accurate and consistent metadata returned by all fetchers.
  - For title:
    Prefers the longest title that appears across multiple sources, aiming to capture complete and consistent metadata.
    If no such title is shared by multiple fetchers falls back to frequency-based selection or prompts the user if no
    clear majority exists.

  - For LCC:
    Prioritizes the official Library of Congress subject if available.
    Otherwise, filters out range-based LCCs (e.g., "QA1-999") and selects the most frequent specific code.
    If no clear winner among specific codes, falls back to full list and prompts the user if needed.

  - For authors:
    Collects all returned author names, groups similar variants (e.g., "Doe, John" vs. "John Doe"),
    selects the best representation from each group, and determines the final list of authors. If an author appears
    in fewer than half of the fetchers capable of providing authors, the user is prompted to confirm whether it
    should be included.
#>
function BuildFinalBookData([hashtable]$fetcherToBookDataMap, [string]$isbn) {

    # Handle authors
    # Get 'unique' authors (some 'unique' authors, may represent the same author)
    $allAuthors = ($fetcherToBookDataMap.Values.ForEach({ $_.Authors }) | Where-Object { $_ }) | Select-Object -Unique
    $authors = [System.Collections.Generic.List[string]]::new() # Authors to be returned

    switch ($allAuthors.Count) {

        0 { $authors = InsertAuthorsManually }

        1 { $authors = $allAuthors }

        <#
          If $allAuthors.Count is greater than 1, it means multiple author names were returned.
          Some of these may represent the same author with different spellings or formats.
          To ensure accuracy:
            1. Group similar author names together (e.g., "Doe, John" and "John Doe").
            2. Select the best variant from each group using SelectBestAuthor.
            3. Pass the resulting list to SelectFinalAuthors, which determines the definitive set of authors.
        #>
        default {
            $variantGroups = GroupBySameAuthor $allAuthors
            foreach ($group in $variantGroups) {
                $bestAuthor = SelectBestAuthor $group
                $authors.Add($bestAuthor)
            }

            if ($authors.Count -ge 2){
                $authors = SelectFinalAuthors $fetcherToBookDataMap $variantGroups $authors
            }
        }
    }

    # Handle titles


    # Try to get the longest repeated title first
    $title = SelectLongestRepeatedTitle $fetcherToBookDataMap.Values.Title

    # If no repeated longest title found, fall back to frequency-based selection
    if ([string]::IsNullOrEmpty($title)) {
        $title = SelectMostFrequentOrPrompt $fetcherToBookDataMap.Values.Title "Title" $isbn
    }

    # Handle LCCs
    # Use official source if possible
    $subject = $null
    if ($fetcherToBookDataMap["LibraryOfCongress"].Subject){
        $subject = $fetcherToBookDataMap["LibraryOfCongress"].Subject
    }
    # Use SearchWorks or CernLibraryCatalogue as secondary sources
    else {
        $subjects = @(
            $fetcherToBookDataMap["SearchWorks"].Subject
            $fetcherToBookDataMap["CernLibraryCatalogue"].Subject
            $fetcherToBookDataMap["CiNii"].Subject
            $fetcherToBookDataMap["Prospector"].Subject
            $fetcherToBookDataMap["Yale"].Subject
            $fetcherToBookDataMap["OpenLibrary"].Subject
        )

        $noRangesSubject = $subjects | Where-Object {
            $code = ($_ -split ' - ')[0]
            # Collapsed A-Z range, e.g. QA9.A5-Z or QL638.875.A-Z
            $azRangePattern = "^([A-Z]+\d+(?:\.[A-Z]?\d+)*)(?:\.A\d+-Z)$"
            # Collapsed range, e.g. PG7157.P47-7157.P472 or PL8455.7-8455.795
            $collapsedRangePattern = "^([A-Z]+\d+(?:\.[A-Z]?\d+)*?)-(\d+(?:\.[A-Z]?\d+)*)$"

            $code -notmatch $azRangePattern -and $code -notmatch $collapsedRangePattern
        }

        $noRangesSubject = $noRangesSubject | Where-Object { -not [string]::IsNullOrEmpty($_) }

        $subjectCandidates = if ($noRangesSubject.Count -gt 0) {
            $noRangesSubject
        } else {
            $subjects
        }

        $subject = SelectMostFrequentOrPrompt $subjectCandidates "LCC" $isbn

    }

    return [BookData]::new($title, $subject, $authors)

}







# ========== Main operations  ==========

function FindBookDataFromSources([string]$isbn) {

    # Normalize ISBN
    $isbn = ($isbn.Trim() -replace '[^\dXx]', '').ToUpperInvariant()
    Write-Host "Fetching metadata for $isbn..."

    $fetchers = [Ordered]@{
        GoogleAPI            = "GetData_GoogleApis"
        LibraryOfCongress    = "GetData_LibraryOfCongress"
        CernLibraryCatalogue = "GetData_CernLibraryCatalogue"
        OpenLibrary          = "GetData_OpenLibrary"
        CiNii                = "GetData_CiNii"
        Yale                 = "GetData_Yale"
        Prospector           = "GetData_Prospector"
    }

    $fetcherToBookDataMap = @{}

    foreach ($fetcher in $fetchers.GetEnumerator()) {
        try {
            [BookData]$bookDataResponse = & $fetcher.Value $isbn    # Get data from source
        }
        catch{
            # Used to ensure the map retains a consistent number of entries
            $fetcherToBookDataMap[$fetcher.Key] = $null

            $exceptionMessage = $($_.Exception.Message)

            if ($exceptionMessage -match "Internal server error \(500\)"){
                $exceptionMessage = "Internal server error (500). Something went wrong on our end."
            }

            Write-Warning "$($fetcher.Key): $exceptionMessage"
            # Uncomment to debug:
            # Write-Host "StackTrace:`n$($_.ScriptStackTrace)" -ForegroundColor Yellow

            Write-Host $divisionLineShort
            continue
        }
        LogBookDataResponse $bookDataResponse $fetcher.Key "Magenta"
        $fetcherToBookDataMap[$fetcher.Key] = $bookDataResponse

        Write-Host $divisionLineShort
    }

    [BookData]$bookData = BuildFinalBookData $fetcherToBookDataMap $isbn

    Write-Host "`nMetadata"
    LogBookDataResponse $bookData $null "Green"

    return $bookData

}

# Logs missing metadata fields (title, subject, author)
function LogMissingMetadata([BookData]$bookData){

    if (-not $bookData.Title){
        Write-Host "Couldn't find any title" -ForegroundColor Yellow
    }

    if (-not $bookData.Subject){
        Write-Host "Couldn't find any subject" -ForegroundColor Yellow
    }

    if (-not $bookData.Authors -or $bookData.Authors.Count -eq 0){
        Write-Host "Couldn't find any author" -ForegroundColor Yellow
    }

}

# Moves a PDF to the failure directory and logs the reason for failure
function MoveToFailureDirectory($pdf, $reason){

    if (-not (Test-Path -Path $outputDirectoryFailure)) {
        Write-Host "Directory $outputDirectoryFailure doesn't exist. Creating directory..."
        New-Item -ItemType Directory -Path $outputDirectoryFailure | Out-Null
        Write-Host "Directory created"
    }

    Write-Host $reason -ForegroundColor Red
    Move-Item -Path $pdf -Destination $outputDirectoryFailure
    Write-Host "$pdf moved to $outputDirectoryFailure"
    Write-Host $divisionLineLong

}

# Main pipeline for processing a single PDF: extracts ISBN, fetches metadata, and updates the file
function ProcessPDF([string]$pdf) {

    try {
        $isbns = GetIsbnsFromPdf $pdf
    }
    catch {
        if (Test-Path $inputArg -PathType Container){
            MoveToFailureDirectory $pdf ($_.Exception.Message)
        }
        throw ""
    }

    $isIsbn = $true;

    Write-Host $divisionLineLong -ForegroundColor DarkCyan
    Write-Host "Fetching data..."

    $pdfProcessed = $false
    foreach ($isbn in $isbns){
        try {
            $bookData = FindBookDataFromSources $isbn
        }
        catch {
            throw ("Error fetching data from sources: $($_.Exception.Message)")
        }

        if ($bookData.Title -or $bookData.Subject -or ($bookData.Authors -or $bookData.Authors.Count -gt 0)){
            Write-Host $divisionLineLong
            Write-Host "Updating metadata..."
            LogMissingMetadata $bookData

            # Update metadata in PDF
            try {
                SetBookDataAndSave -inputPdf $pdf `
                           -title $bookData.Title `
                           -author $bookData.Authors `
                           -subject $bookData.Subject
            }
            catch {
                if (Test-Path $inputArg -PathType Container) {
                    MoveToFailureDirectory $pdf $_.Exception.Message
                }
            }
            $pdfProcessed = $true
            break
        }

        Write-Host "`nMetadata unavailable (no title, subject, or author) for ISBN: $isbn. Trying next ISBN.`n" -ForegroundColor Yellow

    }

    if (-not $pdfProcessed -and (Test-Path $inputArg -PathType Container)) {
        MoveToFailureDirectory $pdf "`nMetadata unavailable (no title, subject, or author) found. Skipping further processing."
    }

    $isIsbn = $false; # Reset flag

}

# Processes all PDF files in a directory
function ProcessDirectory([string]$directoryPath) {

    $files = Get-ChildItem -Path $directoryPath -File | Where-Object { $_.Extension -ieq '.pdf' }

    if ($files.Count -eq 0) {
        Write-Host "No PDF files found in: $directoryPath" -ForegroundColor Yellow
        return
    }

    foreach ($file in $files) {
        $fileName = $file.FullName
        try {
            ProcessPDF $fileName   # force full path string
            Write-Host "`n`n`n`n`n`n- - - - - - -`n`n`n`n`n`n"
        }
        catch {
            Write-Host "Error processing: $fileName" -ForegroundColor Red
            Write-Host ($_.Exception.Message) -ForegroundColor Red
            Write-Host "`n`n`n`n`n`n- - - - - - -`n`n`n`n`n`n"
        }
    }

    Write-Host "Processing complete. Output files have been saved to: $desktopPath`n`n" -ForegroundColor Green

}







# ========== Main Control Flow  ==========

# Main entry point. Validates input and routes to appropriate processing logic.
function Run {

    $isPdf = [System.IO.Path]::GetExtension($inputArg) -ieq '.pdf'

    # Process PDF
    if ((Test-Path $inputArg -PathType Leaf) -and $isPdf) {
        ProcessPDF $inputArg
    }
    # Process directory
    elseif (Test-Path $inputArg -PathType Container) {
        ProcessDirectory $inputArg
    }
    # Search metadata
    elseif ($inputArg -and (IsIsbn $inputArg)) {
        $isIsbn = $true
        FindBookDataFromSources $inputArg
    }
    # PDF file does not exist
    elseif (-not (Test-Path $inputArg -PathType Leaf) -and $isPdf){
        Write-Host "File: '$inputArg' does not exist." -ForegroundColor Red
    }

}

# If the $resourcesDirectory doesn't exist, initiate setup to create it and prepare required files.
if (-not (Test-Path $resourcesDirectory)) {

    Write-Host "$resourcesDirectory not found. Starting setup.."
    Setup

}

if ($Version) {
    Write-Host "PDF Book Tagger v$SCRIPT_VERSION" -ForegroundColor Cyan
    exit 0
}

if ([string]::IsNullOrWhiteSpace($inputArg)) {
    Write-Host "Error: Input argument is required." -ForegroundColor Red
    Write-Host "Usage:" -ForegroundColor Yellow
    Write-Host ".\script.ps1 <File Path> [Book ISBN]" -ForegroundColor Yellow
    Write-Host ".\script.ps1 <Book ISBN | Directory Path>" -ForegroundColor Yellow
    exit 1
}

Import-Module IText7Module
Add-Type -AssemblyName System.Web
initialiseIText7 *>$null

Run