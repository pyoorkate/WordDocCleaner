# WordDocCleaner
A python script that allows you to clean up the styles in a word document

When producing epubs or other documents based on word files it's useful to force word files to clean up their act.

Often folks will unintentionally or unknowingly misuse styles, resulting in one style having multiple definitions through the document

You might find Normal has multiple fonts, or fonts in multiple sizes, with different line spacing, etc.
If a document has been worked on in any combination of LibreOffice, GoogleDocs and Word then it can have multiple style definitions
If folks have worked on it who have their language set to different locales, you can find that it has multiple different language settings in one file

This script will take a docx format file and:
- Identify the styles in it, then recursively let you change them to: Normal, Heading1 or Heading2 (or ignore them)
- Reset the default font back to the default
- Force the dictionary to the one you specify

At least that's the theory.

Call it with: py path/to/file/style_simplifier.py
