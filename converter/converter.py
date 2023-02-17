from docx import Document
from string import punctuation


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
    string = '%s'*3 % \
        (c, x.strip(), c)
    return string + ' '


def check_paragraphs(doc):
    base = []
    for para in doc.paragraphs:
        new = str()
        for r in para.runs:
            if r.text.isspace():
                continue
            if r.text in punctuation:
                new = new[:-1] + r.text
                continue
            mod = check_run(r)
            new += create_text(r.text, mod)     
        base.append(new)
    return base


def main():
    doc = Document('converter\\example_doc.docx')
    print('-', ' -', '  -', sep='\n')
    for i in check_paragraphs(doc):
        print(i)
    print('  -', ' -', '-', sep='\n')


if __name__ == '__main__':
    main()