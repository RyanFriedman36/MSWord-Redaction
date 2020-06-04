# docx-Redaction

A desktop application written in python for the purpose of redacting sensitive information from a Microsoft Word document. The UI will prompt you to select a docx file and a .txt file. The .txt file should be simply a comm-separated list of full names of any sensitive proper nouns (do not include middle initials).

This program will create a new docx file so that there is no metadata that indicates what data was redacted. This is not an executable, the necessary python libraries must be installed on the local machine to run. 

Inspiration:

My mother works as a hearing officer and she has now been tasked with redacting all of the sensitive information from her decisions before she can publish them. In honor of Mother's day I decided to write her a program that could that for her!


Not implemented yet:
- support for headers and footers
- support for footnotes/references
- support for tables
- support for images

Current Bugs:
- document style hierarchy not completely retained
- ~~off-by-one issue on some redactions~~
