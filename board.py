#
# References:
#   https://symfoware.blog.fc2.com/blog-entry-768.html
#

import sys
import re
import csv
import datetime
from reportlab.pdfgen import canvas
from reportlab.platypus.doctemplate import PageTemplate, BaseDocTemplate
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import PageBreak
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus.frames import Frame
from reportlab.lib.units import cm


class Comment():
    def __init__(self, index, submitter, dt, body):
        self.index = index
        self.submitter = submitter
        self.dt = datetime.datetime.strptime(dt, '%Y/%m/%d %H:%M')
        self.body = body

    def __str__(self):
        return '[{}] {} ({})\n{}'.format(self.index, self.submitter, self.dt, self.body)

class Board():
    def __init__(
            self,id, title, body, creator, create_time, updator, update_time, raw_comments
        ):
        self.id = id
        self.title = title
        self.body = body
        self.creator = creator
        self.create_time = create_time
        self.updator = updator
        self.update_time = update_time
        self.comments = self.get_comments(raw_comments)

    def __str__(self):
        title = '{} [{}]'.format(self.title, self.id)
        creator = '{} ({})'.format(self.creator, self.create_time)
        updator = '{} ({})'.format(self.updator, self.update_time)

        return '{}\n{}/{}\n{}\n'.format(title, creator, updator, self.body)

    def get_comments(self, raw_comments):
        comments = []
        body = []
        marker = False
        next_comment = None

        for raw_comment in raw_comments.splitlines()[:-1]:
            if raw_comment == '--------------------------------------------------':
                marker = True
                continue

            m = re.match(
                r'(\d+): (.*?) (\d{4})/(\d{1,2})/(\d{1,2}) (.*) (\d{1,2}):(\d{1,2})',
                raw_comment
            )

            if marker and m:
                index = m.group(1)
                submitter = m.group(2)
                year = m.group(3)
                month = m.group(4)
                date = m.group(5)
                hour = m.group(7)
                minute = int(m.group(8))
                dt = '{}/{}/{} {}:{}'.format(year, month, date, hour, minute)

                if next_comment:
                    next_comment.body = '<br />\n'.join(body[1:-1])
                    comments.append(next_comment) 

                next_comment = Comment(index, submitter, dt, '')
                body = []

            else:
                if next_comment:
                    body.append(raw_comment)

            marker = False

        if next_comment:
            body.append(raw_comment)
            next_comment.body = '<br />\n'.join(body[1:-1])
            comments.append(next_comment) 

        return comments

def read_csv(path):
    boards = []
    with open(path, encoding='utf-8') as f:
        reader = csv.reader(f)
        skip = True
        for row in reader:
            if skip:
                skip = False
                continue

            board = Board(
                row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7])
            boards.append(board)

    return boards

class BoardDocTemplate(BaseDocTemplate):
    def __init__(self, filename, **kw):
        self.allowSplitting = 0
        BaseDocTemplate.__init__(self, filename, **kw)
        template = PageTemplate('normal', [Frame(2.5*cm, 2.5*cm, 15*cm, 25*cm, id='F1')])
        self.addPageTemplates(template)

def gen_board_pdf(boards):
    doc = BoardDocTemplate('output.pdf')
    pdfmetrics.registerFont(TTFont('IPA Gothic', './ipaexg.ttf'))

    story = []

    header_style = ParagraphStyle(
        fontName='IPA Gothic',
        fontSize=12,
        name='header',
        spaceAfter=10
    )

    comment_header_style = ParagraphStyle(
        fontName='IPA Gothic',
        fontSize=10,
        name='comment_header',
        spaceAfter=6
    )

    body_style = ParagraphStyle(
        fontName='IPA Gothic',
        fontSize=8,
        name='body',
        spaceAfter=4,
        justifyBreaks=1
    )

    for board in boards:
        title = '{} [{}]'.format(board.title, board.id)
        creator = '{} ({})'.format(board.creator, board.create_time)
        header = '{} / {}'.format(title, creator)

        story.append(Paragraph(header, header_style))
        story.append(Paragraph(board.body, body_style))

        for comment in board.comments:
            comment_header = '[{}] {} ({})'.format(comment.index, comment.submitter, comment.dt)
            story.append(Paragraph(comment_header, comment_header_style))
            story.append(Paragraph(comment.body, body_style))

        story.append(PageBreak())

    doc.multiBuild(story)


def main(path):
    boards = read_csv(path)
    gen_board_pdf(boards)

def usage():
    print('board.py [path to exported csv]')

if __name__ == '__main__':
    if len(sys.argv) != 2:
        usage()
        sys.exit(1)

    main(sys.argv[1])
    sys.exit(0)
