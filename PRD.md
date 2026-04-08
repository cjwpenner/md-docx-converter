#MD-DCOX

I want a tool that is able to generate a word document (.docx) from a .md document, and a .md  from a word document.
It should be python based and have a shortcut from my desktop to invoke it.
The UI should ask for the full file path of the document to convert and then create the alternative of it ({{name}}.md for a docx, and {{name.docx}} for a .md) in the same directory. If the file already exists it should ask if you want to replace it. If the file path is not a .md or a .docx it should give a warning that this is not a valid path.
The UI is simple enough to be provided as a CLI, so use that if appropriate.
Use the rules in the MardownSyntax.md document that I have included to help with this - it is taken from here: https://www.jetbrains.com/help/hub/markdown-syntax.html#backslash-escapes
The .md should be turned into a "normal" template word document. My MS Word Normal template is found here: "C:\Users\Chris\AppData\Roaming\Microsoft\Templates\Normal.dotm"
The hierarchy of Microsoft word headings is based on the style: Title > Heading 1 > Heading 2 > etc.
The heading Hierarchy in .md is # > ## > ### etc. 
But there is a rule that needs checking - if there is only a single "#" at the top of a .md file, then it is a Title. If there a re multiple "#" lines in the .md then they are "Heading 1" in Word. Likewise in Word, if there is no Title, the treat the Heading 1 as a "#", if there is a Title, then the Title becomes the "#" and the Heading 1 becomes "##".
Anything in the .md that is a bullet point list or numbered list should become a bullet point list or numbered list style in the Word document and vice versa.
Tables need special handling. The python should check the number of columns and the subsequent number of lines with the same number of columns, then recreate the table in a Word table.


