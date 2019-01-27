#
# References:
#   https://symfoware.blog.fc2.com/blog-entry-768.html
#   https://www.reportlab.com/snippets/8/
#

import sys
import re
import csv
import datetime
import argparse
from reportlab.pdfgen import canvas
from reportlab.platypus.doctemplate import PageTemplate, BaseDocTemplate
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import PageBreak
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus.frames import Frame
from reportlab.lib.units import cm
from reportlab.lib.fonts import addMapping
from reportlab.lib.colors import Color

HEADERS = {
    'Board': ['ID','タイトル','本文','作成者','作成日時','更新者','更新日時','コメント'],
    'Todo': ['ID','タイトル','本文','作成者','作成日時','更新者','更新日時','ステータス','優先度','担当者','期日','コメント'],
    'Member': ['姓','名','よみがな姓','よみがな名','メールアドレス'],
}

DEFAULT_FONT = 'IPA Gothic'
DEFAULT_FONT_FILE = 'ipaexg.ttf'

class DocTemplate(BaseDocTemplate):
    def __init__(self, filename, **kw):
        self.allowSplitting = 0
        BaseDocTemplate.__init__(self, filename, **kw)
        template = PageTemplate('normal', [Frame(2.5*cm, 2.5*cm, 15*cm, 25*cm, id='F1')])
        self.addPageTemplates(template)

    def afterFlowable(self, flowable):
        if flowable.__class__.__name__ == 'Paragraph':
            text = flowable.getPlainText()
            style = flowable.style.name
            if style == 'header':
                self.notify('TOCEntry', (0, text, self.page, flowable._bookmark))

class CommentGenerator():
    def __init__(self, comments):
        self.comments = comments

        self.header_style = ParagraphStyle(
            fontName=DEFAULT_FONT,
            fontSize=10,
            name='comment_header',
            spaceAfter=6
        )

        self.body_style = ParagraphStyle(
            fontName=DEFAULT_FONT,
            fontSize=8,
            name='body',
            spaceAfter=4,
            justifyBreaks=1
        )

    def convert(self):
        story = []
        for comment in self.comments:
            comment_header = '[{}] {} ({})'.format(comment.index, comment.submitter, comment.dt)
            story.append(Paragraph(comment_header, self.header_style))
            story.append(Paragraph(comment.body, self.body_style))
        return story

class BoardGenerator():
    def __init__(self, boards):
        self.boards = boards

        self.header_style = ParagraphStyle(
            fontName=DEFAULT_FONT,
            fontSize=12,
            name='header',
            spaceAfter=10,
            borderColor=Color(0, 0, 0, 1),
            borderPadding=3,
            borderWidth=1,
        )

        self.body_style = ParagraphStyle(
            fontName=DEFAULT_FONT,
            fontSize=8,
            name='body',
            spaceAfter=4,
            justifyBreaks=1,
            backColor=Color(0, 0, 0, 0.1),
        )

    def does_support_toc(self):
        return True

    def convert(self):
        import xxhash
        story = []
        for board in self.boards:
            title = '{} [{}]'.format(board.title, board.id)
            creator = '{} ({})'.format(board.creator, board.create_time)
            digest = xxhash.xxh32_hexdigest(title)
            header = '{} / {} <a name={} />'.format(title, creator, digest)

            hp = Paragraph(header, self.header_style)
            hp._bookmark = digest
            story.append(hp)
            story.append(Paragraph(board.body, self.body_style))
            story.extend(CommentGenerator(board.comments).convert())
            story.append(PageBreak())

        return story

class TodoGenerator():
    def __init__(self, todos):
        self.todos = todos 

        self.header_style = ParagraphStyle(
            fontName=DEFAULT_FONT,
            fontSize=12,
            name='header',
            spaceAfter=10,
            borderColor=Color(0, 0, 0, 1),
            borderPadding=3,
            borderWidth=1,
        )

        self.body_style = ParagraphStyle(
            fontName=DEFAULT_FONT,
            fontSize=8,
            name='body',
            spaceAfter=4,
            justifyBreaks=1,
            backColor=Color(0, 0, 0, 0.1),
        )

    def does_support_toc(self):
        return True

    def convert(self):
        import xxhash
        story = []
        for todo in self.todos:
            title = '{} [{}]'.format(todo.title, todo.id)
            creator = '{} ({})'.format(todo.creator, todo.create_time)
            status = '{}, {}, {}, {}'.format(todo.status, todo.priority, todo.pic, todo.due)
            digest = xxhash.xxh32_hexdigest(title)
            header = '{} / {} / {} <a name={} />'.format(title, creator, status, digest)

            hp = Paragraph(header, self.header_style)
            hp._bookmark = digest
            story.append(hp)
            story.append(Paragraph(todo.body, self.body_style))
            story.extend(CommentGenerator(todo.comments).convert())
            story.append(PageBreak())

        return story

class MemberGenerator():
    def __init__(self, members):
        self.members = members

        self.body_style = ParagraphStyle(
            fontName=DEFAULT_FONT,
            fontSize=8,
            name='body',
            spaceAfter=4,
            justifyBreaks=1
        )

    def does_support_toc(self):
        return False

    def convert(self):
        story = []
        for member in self.members:
            kanji = '{} {}'.format(member.sei, member.mei)
            kana = '{} {}'.format(member.sei_kana, member.mei_kana)
            person = '{} ({}) / {}'.format(kanji, kana, member.mail)

            story.append(Paragraph(person, self.body_style))

        return story

class Comment():
    def __init__(self, index, submitter, dt, body):
        self.index = index
        self.submitter = submitter
        self.dt = dt
        self.body = body

    def __str__(self):
        return '[{}] {} ({})\n{}'.format(self.index, self.submitter, self.dt, self.body)

    @staticmethod
    def parse(data):
        comments = []
        body = []
        marker = False
        next_comment = None

        for raw_comment in data.splitlines()[:-1]:
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
                year = int(m.group(3))
                month = int(m.group(4))
                date = int(m.group(5))
                hour = int(m.group(7))
                minute = int(m.group(8))
                dt = datetime.datetime(year, month, date, hour, minute)

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

class Board():
    def __init__(
            self, id, title, body, creator, create_time, updator, update_time, comments
        ):
        self.id = id
        self.title = title
        self.body = body.replace('\n', '<br />\n')
        self.creator = creator
        self.create_time = datetime.datetime.strptime(create_time, '%Y/%m/%d %H:%M')
        self.updator = updator
        self.update_time = datetime.datetime.strptime(update_time, '%Y/%m/%d %H:%M')
        self.comments = Comment.parse(comments)

    def __str__(self):
        title = '{} [{}]'.format(self.title, self.id)
        creator = '{} ({})'.format(self.creator, self.create_time)
        updator = '{} ({})'.format(self.updator, self.update_time)

        return '{}\n{}/{}\n{}\n'.format(title, creator, updator, self.body)

class Todo():
    def __init__(
            self, id, title, body, creator, create_time, updator, update_time, status, priority, pic, due, comments
        ):
        self.id = id
        self.title = title
        self.body = body
        self.creator = creator
        self.create_time = datetime.datetime.strptime(create_time, '%Y/%m/%d %H:%M')
        self.updator = updator
        self.update_time = datetime.datetime.strptime(update_time, '%Y/%m/%d %H:%M')
        self.status = status
        self.priority = priority
        self.pic = pic
        self.due = due
        self.comments = Comment.parse(comments)

    def __str__(self):
        title = '{} [{}]'.format(self.title, self.id)
        creator = '{} ({})'.format(self.creator, self.create_time)
        updator = '{} ({})'.format(self.updator, self.update_time)

        return '{}\n{}/{}\n{}\n'.format(title, creator, updator, self.body)

class Member():
    def __init__(self, sei, mei, sei_kana, mei_kana, mail):
        self.sei = sei
        self.mei = mei
        self.sei_kana = sei_kana
        self.mei_kana = mei_kana
        self.mail = mail

    def __str__(self):
        return '{} {} ({})\n'.format(self.sei, self.mei, self.mail)

def analyze(header):
    if(header == HEADERS['Board']):
        return 'Board'
    elif(header == HEADERS['Todo']):
        return 'Todo'
    elif(header == HEADERS['Member']):
        return 'Member'
    else:
        print('Unsupported CSV')
        sys.exit(1)

def read_csv(path, from_date=None, to_date=None):
    entities = []
    with open(path, encoding='utf-8') as f:
        reader = csv.reader(f)
        analyzed = False
        for row in reader:
            if not analyzed:
                class_name = analyze(row)
                analyzed = True
                continue

            entity = globals()[class_name](*row)
            try:
                if from_date and entity.create_time < from_date:
                    continue
                elif to_date and entity.create_time > to_date:
                    continue
            except AttributeError:
                pass

            entities.append(entity)

    entities.sort(key=lambda x: x.create_time, reverse=True)
    return (entities, class_name)

def gen_pdf(generator, output, toc=True):
    addMapping(DEFAULT_FONT, 1, 1, DEFAULT_FONT)

    story = []

    if toc:
        title_style = ParagraphStyle(
            fontName=DEFAULT_FONT,
            fontSize=15,
            name='TOC',
            spaceAfter=10
        )
        story.append(Paragraph('目次', title_style))

        toc = TableOfContents()
        toc.levelStyles = [
            ParagraphStyle(
                fontName=DEFAULT_FONT,
                fontSize=8,
                name='body',
                spaceAfter=4,
                justifyBreaks=1
            )
        ]

        story.append(toc)
        story.append(PageBreak())

    story.extend(generator.convert())

    doc = DocTemplate(output)
    pdfmetrics.registerFont(TTFont(DEFAULT_FONT, DEFAULT_FONT_FILE))
    doc.multiBuild(story)

def main(path, output, from_date=None, to_date=None):
    (entities, class_name) = read_csv(path, from_date, to_date)
    generator = globals()[class_name + 'Generator'](entities) #TODO: converter name mapping
    gen_pdf(generator, output, generator.does_support_toc())


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='')
    parser.add_argument('csv', help='input csv file path')
    parser.add_argument('pdf', help='output pdf file name')
    parser.add_argument('-f', '--from-date', action='store', metavar='yyyy/mm/dd', help='')
    parser.add_argument('-t', '--to-date', action='store', metavar='yyyy/mm/dd', help='')
    args = parser.parse_args()

    from_date = None
    to_date = None
    try:
        if args.from_date:
            from_date = datetime.datetime.strptime(args.from_date, '%Y/%m/%d')
        if args.to_date:
            to_date = datetime.datetime.strptime(args.to_date, '%Y/%m/%d')
    except ValueError:
        print(f'invalud date format')
        sys.exit(1)

    main(args.csv, args.pdf, from_date, to_date)
    sys.exit(0)
