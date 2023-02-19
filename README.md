# **Docx to Discord**
## **Bold**, __Underline__, and *Italic* is Currently Supported

The 'converter.py' file will convert text inside a docx file into a simple markdown format that can be pasted into Discord. Just save your text file to the converter folder and run the script. The result will be printed.

Punctuation is a bit iffy due to python-docx's difficulties distinguishing runs from one another. Even the slightest change in format can separate runs, making punctuation appear awkwardly. String slicing and such could be used to fix this, but it will probably be easier to either build this script in VBA or use a docx to markdown converter, which I might automate with a Discord bot in the future. We'll see.

I mainly made this for fun, but feel free to use it. Again, there's probably better options.
