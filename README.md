# Scripts for Adobe Indesign

Some scripts I've written to do useful things in Adobe Indesign.

## Contents

1. [Copy Things](#copy-things)
1. [Fix Links](#fix-links)
1. [Numbered Markers](#numbered-markers)
1. [Combine Documents](#combine-documents)
1. [Style Highlighter](#style-highlighter)
1. [Generate Underlines](#generate-underlines)
1. [Synchronize All Documents Text Selection](#synchronize-all-documents-text-selection)

---

## Copy Things

[![Download Copy Things script](https://img.shields.io/badge/Download_Script-*_FREE!_*_-F50?style=flat-square)](https://raw.githubusercontent.com/mark1bean/scripts-for-adobe-indesign/main/Copy%20Things.js)   ![Language: ExtendScript](https://img.shields.io/badge/Language-ExtendScript-99B?style=flat-square)   ![Version: 2025-05-12](https://img.shields.io/badge/Version-2025--05--12-5A5?style=flat-square)   [![Donate](https://img.shields.io/badge/Donate-PayPal-blue?style=flat-square)](https://www.paypal.com/donate?hosted_button_id=SBQHVWHSSTA9Q)

A script for copying/updating styles from one document to one *or more* destination documents.

### Features

- Quickly access and select multiple things to copy in one go.
- Copy to as many documents as you have open in one go.
- Will update existing things, if they already exists.
- Takes account of the hierarchic structure in the source document, and replicates it in the destination (ie. keeps your folder structure intact).

![Copy Things script's UI](docs/images/copy-things-ui-1.png)

### Usage

1. Choose the thing type: *Character Styles, Paragraph Styles, Object Styles, Cell Styles, TableStyles, Swatches*.
1. Filter the list of things, if it helps.
1. Choose the things to copy.
1. Choose the destination documents (will show all open documents).
1. Press the Copy button to perform the copying.

### Limitations

- Not much testing done! If you find a bug, please [start a new Issue](https://github.com/mark1bean/scripts-for-adobe-indesign/issues) and always include a link to a demo document that shows the error.
- The scripting API doesn't provide control over the *order* of the copied things, so you may need to sort them afterwards.

---

## Fix Links

[![Download Fix Links script](https://img.shields.io/badge/Download_Script-*_FREE!_*_-F50?style=flat-square)](https://raw.githubusercontent.com/mark1bean/scripts-for-adobe-indesign/main/Fix%20Links.js)   ![Language: ExtendScript](https://img.shields.io/badge/Language-ExtendScript-99B?style=flat-square)   ![Version: 2025-10-03](https://img.shields.io/badge/Version-2025--10--03-5A5?style=flat-square)   [![Donate](https://img.shields.io/badge/Donate-PayPal-blue?style=flat-square)](https://www.paypal.com/donate?hosted_button_id=SBQHVWHSSTA9Q)

A super-helpful script for matching up missing links after moving platforms or servers.

For example Bob has the original files here:

`/Users/bob/Dropbox/...`

But for Alice the path to the same files is:

`/Users/alice/Library/CloudStorage/Dropbox/...`

This is a big problem, and when the linked files are in a complex file structure, it can be very time consuming to fix via Indesign's native tools.

#### How to use

1. Open your document or documents (script will process all open documents).
1. Run script.
1. Cilck a link from the Broken Links list.
1. Click "Find" button and manually relink to that correct file.
1. Script will use the new link to try to fix broken links.
1. Repeat step 2 and 3 until Broken Links list is empty.
1. Click "Fix [N] Links" button to perform the changes to the open document(s).

> Tip: You can store the settings in the document(s) so the next time you need it, you won't have to perform steps 3 and 4 again. This applies in reverse, so if Alice sends the file back to Bob, when Bob runs the script it will be ready to return the paths to match Bob's environment.

**Important Note**: this script won't help much if the links are randomly missing in various places. It is specially designed to help when relinking a parent folder to a consistent set of sub-folders, such as when moving a job folder across platforms or servers.

---

## Numbered Markers

[![Download Copy Things script](https://img.shields.io/badge/Download_Script-*_FREE!_*_-F50?style=flat-square)](https://raw.githubusercontent.com/mark1bean/scripts-for-adobe-indesign/main/Numbered%20Markers.js)   ![ExtendScript](https://img.shields.io/badge/Language-ExtendScript-99B?style=flat-square)   ![Version: 2025-05-15](https://img.shields.io/badge/Version-2025--05--15-5A5?style=flat-square)   [![Donate](https://img.shields.io/badge/Donate-PayPal-blue?style=flat-square)](https://www.paypal.com/donate?hosted_button_id=SBQHVWHSSTA9Q)

A script for managing numbered markers linked to a numbered list. For example: numbered place markers on a map, linked to the list of place names.

Why is it good? It saves you from manually re-numbering all your markers when the list changes order, or when items are removed.

Read the [quick tutorial](docs/numbered-markers-quick-tutorial.md). Or the ultra-quick tutorial.

 You will need:

1. A marker — a text frame **with an object style applied**, eg. a circle text frame with a number in it.
1. A numbered list — paragraphs **with a paragraph style applied**, that *must use Indesign's numbering system*, eg. a list of place names for a map legend.

Run this script, choosing the list paragraph style and the marker object style and make sure that "Create missing markers" is ON.

Result: the script will create as many markers as needed and they will be linked to your list. Now you can position them freely and they will update every time you subsequently run the script.

![Numbered Markers script's UI](docs/images/numbered-markers-17.png)

#### Limitations

- Not much testing done! If you find a bug, please [start a new Issue](https://github.com/mark1bean/scripts-for-adobe-indesign/issues) and always include a link to a demo document that shows the error.
- If you already have markers set up, see the last section of the [quick tutorial](doc/numbered-markers-quick-tutorial.md).

---

## Combine Documents

[![Download Combine Documents script](https://img.shields.io/badge/Download_Script-*_FREE!_*_-F50?style=flat-square)](https://raw.githubusercontent.com/mark1bean/scripts-for-adobe-indesign/main/Combine%20Documents.js)   ![Language: ExtendScript](https://img.shields.io/badge/Language-ExtendScript-99B?style=flat-square)   ![Version: 2025-05-17](https://img.shields.io/badge/Version-2025--05--17-5A5?style=flat-square)   [![Donate](https://img.shields.io/badge/Donate-PayPal-blue?style=flat-square)](https://www.paypal.com/donate?hosted_button_id=SBQHVWHSSTA9Q)

A script for combining multiple documents into one.

#### How to use

1. Add all documents to a **book**, if not already. This is how you will tell the script in which *order* you want to combine the documents. Once the documents are combined, you don't need to keep this book; it is just a stepping stone.
1. Run script and choose the book (if more than one) and set the options you want.
1. Click "Combine"

![Combine Documents script's UI](docs/images/combine-documents-ui-1.png)

#### Limitations

- Not much testing done! If you find a bug, please [start a new Issue](https://github.com/mark1bean/scripts-for-adobe-indesign/issues) and always include a link to a demo document that shows the error.

---

## Style Highlighter

[![Download Combine Documents script](https://img.shields.io/badge/Download_Script-*_FREE!_*_-F50?style=flat-square)](https://raw.githubusercontent.com/mark1bean/scripts-for-adobe-indesign/main/Style%20Highlighter.js)   ![Language: ExtendScript](https://img.shields.io/badge/Language-ExtendScript-99B?style=flat-square)   ![Version: 2025-06-17](https://img.shields.io/badge/Version-2025--06--17-5A5?style=flat-square)   [![Donate](https://img.shields.io/badge/Donate-PayPal-blue?style=flat-square)](https://www.paypal.com/donate?hosted_button_id=SBQHVWHSSTA9Q)

Highlights all text in a chosen paragraph or character style; a non-damaging visual tool to spot where styles have gone astray.

#### How to use

1. Run script to show the UI.
1. Select one (or more) paragraph or character styles.
1. Select a highlight color.
1. Click "Add highlighter".

![Style Highlighter script's UI](docs/images/style-highlighter-ui-1.png)

#### Notes

- The script *does not modify your styles* at all.
- The script makes use of **conditions** to perform the highlighting. It shouldn't interfere with existing conditions, because multiple conditions can be applied to the same text without a problem. The script uses a naming prefix to avoid collisions.
- It is okay for your styles to have conditions in them, the script won't bother them.
- The highlighting is only visible when in normal screen mode, and disappears in preview mode.
- You can leave the highlighting conditions active for as long as you want.
- To remove the highlighting you can either (a) manually remove the Highlighter Style condition from the Conditions Panel, or (b) run the script and choose "Remove all Highlighters" to remove them all.

#### Limitations

- The highlighting operation happens when you press the "Add Highlighter" button; it will not highlight new paragraphs even if they use the highlighted style.
- Not much testing done! If you find a bug, please [start a new Issue](https://github.com/mark1bean/scripts-for-adobe-indesign/issues) and always include a link to a demo document that shows the error.

---

## Generate Underlines

[![Download Copy Things script](https://img.shields.io/badge/Download_Script-*_FREE!_*_-F50?style=flat-square)](https://raw.githubusercontent.com/mark1bean/scripts-for-adobe-indesign/main/Generate%20Underlines.js)   ![Language: ExtendScript](https://img.shields.io/badge/Language-ExtendScript-99B?style=flat-square)   ![Version: 2025-07-19](https://img.shields.io/badge/Version-2025--07--19-5A5?style=flat-square)   [![Donate](https://img.shields.io/badge/Donate-PayPal-blue?style=flat-square)](https://www.paypal.com/donate?hosted_button_id=SBQHVWHSSTA9Q)

A script for drawing multiple, overlapping underlines.

### Features

- Assign underlining using simple Conditions.
- Styles underlines using Object Styles.
- To update underlines after text re-flow, re-run script.

### Usage

1. Start with the text you want to underline.

![Two text frames in Indesign](docs/images/generate-underlines-1.png)

1. Apply the "underline" Conditions to the text as needed.

![Two text frames in Indesign showing applied Conditions as highlighted colors](docs/images/generate-underlines-2.png)

1. Run the script, which will draw the underlines, grouped directly below the text frame(s).

![The text frames now have underlining](docs/images/generate-underlines-3.png)

Note: the underlines are normal drawn lines (path items).

![The underlines, without the text](docs/images/generate-underlines-4.png)

Note: every time you run the script, it will remove any previous underlining before generating the new result.

---

## Synchronize All Documents Text Selection

[![Download Synchronize All Documents Text Selection script](https://img.shields.io/badge/Download_Script-*_FREE!_*_-F50?style=flat-square)](https://raw.githubusercontent.com/mark1bean/scripts-for-adobe-indesign/main/Synchronize%20All%20Documents%20Text%20Selection.js)   ![Language: ExtendScript](https://img.shields.io/badge/Language-ExtendScript-99B?style=flat-square)   ![Version: 2025-03-15](https://img.shields.io/badge/Version-2025--03--15-5A5?style=flat-square)   [![Donate](https://img.shields.io/badge/Donate-PayPal-blue?style=flat-square)](https://www.paypal.com/donate?hosted_button_id=SBQHVWHSSTA9Q)

This script can be incredibly useful when working on **multiple documents** that all have matching headings or subheadings.

Let's say you have 20 documents open and you need to edit the technical specs section of each. In one of the documents, you select the subheading "Technical Specs" and run this script. All those documents will show the "Technical Specs" text with the same zoom level. This saves HUGE amounts of time navigating around each open document.

> TIP: Use the window cycling shortcut key:  command-tilde (MacOS) control-tab (Windows).

Important note: script will work more reliably if your text, eg. "Technical Specs" subheading in the example above, has a sensible paragraph style applied. The script will also match the paragraph style (and character style) as well as the text contents. This avoids cases where the script might match the string "Technical Specs" in the body text or somewhere else in the document.

---

## More scripts coming

I will be sharing more of my scripts when I get time. Click "Watch" to get notified of updates!

[![Star](https://img.shields.io/github/stars/mark1bean/scripts-for-adobe-indesign.svg?style=social&label=Star)](https://github.com/mark1bean/scripts-for-adobe-indesign)

---

## Installation

Step 1: Download the individual scripts (see buttons under script names), or
[![Download](https://img.shields.io/badge/Download_all_scripts_(.zip)-*_FREE!_*-F50?style=flat-square)](https://github.com/mark1bean/scripts-for-adobe-indesign/archive/refs/heads/main.zip)

(Note: If the script shows as raw text in the browser, save it to your computer with the extension ".js".)

Step 2: Place the Scripts in the Appropriate Folder

See [How To Install Scripts in Adobe Indesign](https://creativepro.com/how-to-install-scripts-in-indesign).

---

## Author

Created by Mark Bean (Adobe Community Expert "[m1b](https://community.adobe.com/t5/user/viewprofilepage/user-id/13791991)").

If any of these scripts will save you time, please consider supporting me!

[![Donate](https://img.shields.io/badge/Donate-PayPal-blue.svg?style=flat-square)](https://www.paypal.com/donate?hosted_button_id=SBQHVWHSSTA9Q)

![Profile picture](https://github.com/mark1bean.png)

---

## License

These scripts are open-source and available under the MIT License. See the [LICENSE](LICENSE) file for details.
