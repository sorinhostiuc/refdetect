# RefDetect

RefDetect examines the citation data stored inside Microsoft Word and OpenDocument files and reports which reference manager most likely created them. The document is read for analysis and is not modified.

## Detected reference managers

The detector recognizes citation structures associated with Zotero, Mendeley, EndNote, Paperpile, RefWorks, Citavi, and JabRef. It reports the evidence found in the document and assigns a confidence level to the result.

## Installation

Download `refdetect-1.0.0.xpi` from the latest release. In Zotero, open `Tools > Plugins`, choose `Install Add-on From File`, and select the XPI. The current release supports Zotero 7, 8, and 9.

## Checking a document

Open `Tools > Detect Reference Manager...` and select a `.docx` or `.odt` file. RefDetect parses the document package, counts the signatures left by supported citation tools, and displays the likely source together with the confidence score and supporting details.

A document whose citations have been flattened to plain text contains no field metadata; RefDetect cannot identify the original manager in that case.

## Building from source

On Windows, run `npm run build:windows`. On Unix-like systems, run `npm run build`. The XPI is written to the repository root.

## License

RefDetect is released under the MIT License. See [LICENSE](LICENSE).
