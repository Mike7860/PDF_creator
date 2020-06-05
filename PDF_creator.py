import os
import openpyxl
import time
import calendar
import PyPDF2
import textwrap

from datetime import date
from openpyxl.workbook import Workbook
#from openpyxl.styles import Font
#from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
from reportlab.platypus import BaseDocTemplate, SimpleDocTemplate, Frame, Paragraph, NextPageTemplate, PageBreak, PageTemplate, Spacer, Table, TableStyle
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
#from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors

style = getSampleStyleSheet()
Datas = []
dir = os.getcwd()
print('You are in "' +dir+ '" Do you want change directory? y/n')
dirchoise = input()

if dirchoise =='y':
    print('Please provide new directory')
    os.chdir(input())
if dirchoise =='n':
    dir = dir

workbook = openpyxl.load_workbook('Excel.xlsx', 'rb', data_only=True)
firstsheet = workbook['Arkusz1']
secondsheet = workbook['Arkusz2']
thirdsheet = workbook['Arkusz3']

print('Please provide title of created document')
title = input()

allcells = firstsheet.max_row
cellswithdata = allcells - 1
print('All raws: ', cellswithdata)
r = 'C'+str(allcells)

class Raw:
    rawone = 0
    rawtwo = 0
    rawthird = 0

OBJ = []

pdffile = "pdf.pdf"
doc = BaseDocTemplate(pdffile, showBoundary=0)
date = date.today()

class NumofPages(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._page_curr_num = []

    def showPage(self):
        self._page_curr_num.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        page_count = len(self._page_curr_num)
        for page in self._page_curr_num:
            self.__dict__.update(page)
            self.draw_page_number(page_count)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_number(self, page_count):
        self.setFont('Helvetica', 8)
        self.drawString(9 * cm, 2 * cm, "Page %s of %s" % (self._pageNumber, page_count))

def firstpage(canvas,doc):
    canvas.line(70, 800, 530, 800)
    canvas.setFont('Helvetica', 22)
    canvas.drawCentredString(9.5 * cm, 24 * cm, title)
    canvas.line(70, 100, 530, 100)

def restpages(canvas,doc):
    canvas.line(70, 790, 530, 790)

pdfmetrics.getRegisteredFontNames()
datas2 = []

datas2.append(['First','Second','Third'])

stylesoftable = TableStyle([('ALIGN',(0,0),(2,-1),'CENTRE'),
                            ('FONT',(0,0),(2,0),'Helvetica',10),
                            ('FONT',(0,1),(-1,-1),'Helvetica',8),
                            ('INNERGRID',(0,0),(-1,-1),0.25,colors.black),
                            ('BOX',(0,0),(-1,-1),0.25,colors.black)])

stylesoftable.add('BACKGROUND', (0,0),(2,0),colors.lightcoral)

for i in range (cellswithdata):
    OBJ = Raw()

cells = firstsheet['A2':r]

for A2, B2, C2 in cells:
    OBJ.rawone=A2.value
    OBJ.rawtwo=B2.value
    OBJ.rawthird=C2.value

    datas2.append([OBJ.rawone,OBJ.rawtwo,OBJ.rawthird])
datatable = Table(datas2, [5 * cm, 5 * cm, 5 * cm], repeatRows=1)
datatable.hAlign = 'CENTER'
datatable.setStyle(stylesoftable)

frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal")
Datas.append(NextPageTemplate('Table'))
Datas.append(PageBreak())
Datas.append(Paragraph("", style['Normal']))
Datas.append(datatable)
Datas.append(Spacer(1,16))

doc.addPageTemplates([PageTemplate(id='', frames=frame, onPage=firstpage), PageTemplate(id='Table', frames=frame, onPage=restpages), Datas])
doc.build(Datas, canvasmaker=NumofPages)