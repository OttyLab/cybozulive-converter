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

HEADERS = {
    'Board': ['ID','タイトル','本文','作成者','作成日時','更新者','更新日時','コメント'],
    'Todo': ['ID','タイトル','本文','作成者','作成日時','更新者','更新日時','ステータス','優先度','担当者','期日','コメント'],
    'Member': ['姓','名','よみがな姓','よみがな名','メールアドレス'],
}

class DocTemplate(BaseDocTemplate):
    def __init__(self, filename, **kw):
        self.allowSplitting = 0
        BaseDocTemplate.__init__(self, filename, **kw)
        template = PageTemplate('normal', [Frame(2.5*cm, 2.5*cm, 15*cm, 25*cm, id='F1')])
        self.addPageTemplates(template)

class CommentGenerator():
    def __init__(self, comments):
        self.comments = comments

        self.header_style = ParagraphStyle(
            fontName='IPA Gothic',
            fontSize=10,
            name='comment_header',
            spaceAfter=6
        )

        self.body_style = ParagraphStyle(
            fontName='IPA Gothic',
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
            fontName='IPA Gothic',
            fontSize=12,
            name='header',
            spaceAfter=10
        )

        self.body_style = ParagraphStyle(
            fontName='IPA Gothic',
            fontSize=8,
            name='body',
            spaceAfter=4,
            justifyBreaks=1
        )

    def convert(self):
        story = []
        for board in self.boards:
            title = '{} [{}]'.format(board.title, board.id)
            creator = '{} ({})'.format(board.creator, board.create_time)
            header = '{} / {}'.format(title, creator)

            story.append(Paragraph(header, self.header_style))
            story.append(Paragraph(board.body, self.body_style))
            story.extend(CommentGenerator(board.comments).convert())
            story.append(PageBreak())

        return story

class TodoGenerator():
    def __init__(self, todos):
        self.todos = todos 

        self.header_style = ParagraphStyle(
            fontName='IPA Gothic',
            fontSize=12,
            name='header',
            spaceAfter=10
        )

        self.body_style = ParagraphStyle(
            fontName='IPA Gothic',
            fontSize=8,
            name='body',
            spaceAfter=4,
            justifyBreaks=1
        )

    def convert(self):
        story = []
        for todo in self.todos:
            title = '{} [{}]'.format(todo.title, todo.id)
            creator = '{} ({})'.format(todo.creator, todo.create_time)
            status = '{}, {}, {}, {}'.format(todo.status, todo.priority, todo.pic, todo.due)
            header = '{} / {} / {}'.format(title, creator, status)

            story.append(Paragraph(header, self.header_style))
            story.append(Paragraph(todo.body, self.body_style))
            story.extend(CommentGenerator(todo.comments).convert())
            story.append(PageBreak())

        return story

class MemberGenerator():
    def __init__(self, members):
        self.members = members

        self.body_style = ParagraphStyle(
            fontName='IPA Gothic',
            fontSize=8,
            name='body',
            spaceAfter=4,
            justifyBreaks=1
        )

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
        self.dt = datetime.datetime.strptime(dt, '%Y/%m/%d %H:%M')
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

class Board():
    def __init__(
            self, id, title, body, creator, create_time, updator, update_time, comments
        ):
        self.id = id
        self.title = title
        self.body = body
        self.creator = creator
        self.create_time = create_time
        self.updator = updator
        self.update_time = update_time
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
        self.create_time = create_time
        self.updator = updator
        self.update_time = update_time
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

def read_csv(path):
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
            entities.append(entity)
    return (entities, class_name)

def gen_pdf(generator, output):
    doc = DocTemplate(output)
    pdfmetrics.registerFont(TTFont('IPA Gothic', './ipaexg.ttf'))
    doc.multiBuild(generator.convert())

def main(path, output):
    (entities, class_name) = read_csv(path)
    generator = globals()[class_name + 'Generator'](entities) #TODO: converter name mapping
    gen_pdf(generator, output)

def usage():
    print('board.py [csv file name] [output file name]')

if __name__ == '__main__':
    if len(sys.argv) != 3:
        usage()
        sys.exit(1)

    main(sys.argv[1], sys.argv[2])
    sys.exit(0)
