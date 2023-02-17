from docx import Document
from string import punctuation
import os


def detect_format(r):
    check = r.bold, r.italic, r.underline
    return {'b': r.bold,
            'i': r.italic,
            'u': r.underline}


def check_run(r):
    form = lambda x, b: x if b else ''
    check = \
        (form('*' , r.italic),
         form('**', r.bold),
         form('__', r.underline))
    return ''.join(check)


def create_text(x, c):
    string = '%s'*3 % (c, x, c)
    if x[0].isspace():
        return ' ' + string
    if x[-1].isspace():
        return string + ' '
    return string


def check_paragraphs(doc):
    base = []
    for para in doc.paragraphs:
        new = str()
        for r in para.runs:
            if r.text.isspace():
                new += r.text
                continue
            mod = check_run(r)
            new += create_text(r.text, mod)     
        base.append(new)
    return base


def main():
    doc = Document('testing\\test_doc.docx')
    for i in check_paragraphs(doc):
        print(i)


if __name__ == '__main__':
    main()
