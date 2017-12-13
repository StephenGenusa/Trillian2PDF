#!/usr/bin/env python
#coding:utf-8

"""
CreateTrillianPDFHistory.py

By Stephen Genusa
http://www.github.com/StephenGenusa

Copyright © 2017 by Stephen Genusa

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice,
this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation
and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
POSSIBILITY OF SUCH DAMAGE.
"""

#
# Works with Reportlab 3.2 but not 3.3. Throws exception in parser in 3.3
#
import os
import re

import sys
import datetime
import urllib
from cStringIO import StringIO
import traceback

from BeautifulSoup import BeautifulSoup
from HTMLParser import HTMLParser
import PIL
from reportlab.pdfgen import canvas
from reportlab.lib.styles import ParagraphStyle as PS
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak
from reportlab.platypus.doctemplate import PageTemplate
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.platypus.frames import Frame
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle, StyleSheet1
from reportlab.lib.units import inch, cm
from reportlab.lib.pagesizes import LETTER
import requests

_stylesheet = None
user_name = ""
buddy_dict = {}

def stylesheet():
    """ Create the style sheet """
    global _stylesheet
    if _stylesheet is not None:
        return _stylesheet
    stylesheet = StyleSheet1()
    stylesheet.add(ParagraphStyle(name = 'Normal',
                                  alignment=TA_LEFT,
                                  fontName = 'Courier',
                                  fontSize = 10,
                                  leading = 12,
                                  spaceAfter = 2))
    stylesheet.add(ParagraphStyle(name = 'Heading1',
                                  parent = stylesheet['Normal'],
                                  pageBreakBefore = 0,
                                  add_to_toc=True,
                                  keepWithNext = 1,
                                  fontName = 'Courier-Bold',
                                  fontSize = 16,
                                  leading = 22,
                                  spaceBefore = 0,
                                  spaceAfter = 0),
                                  alias = 'h1')
    stylesheet.add(ParagraphStyle(name = 'Heading2',
                                  parent = stylesheet['Normal'],
                                  frameBreakBefore = 0,
                                  keepWithNext = 1,
                                  fontName = 'Courier',
                                  fontSize = 12,
                                  leading = 18,
                                  spaceBefore = 8,
                                  spaceAfter = 0),
                                  alias = 'h2')
    _stylesheet = stylesheet
    return stylesheet

class FooterCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self.pages = []

    def showPage(self):
        self.pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        page_count = len(self.pages)
        for page in self.pages:
            self.__dict__.update(page)
            self.draw_canvas(page_count)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_canvas(self, page_count):
        page = "Page %s of %s" % (self._pageNumber, page_count)
        x = 128
        self.saveState()
        self.setStrokeColorRGB(0, 0, 0)
        self.setLineWidth(0.5)
        self.line(66, 78, LETTER[0] - 66, 78)
        self.setFont('Times-Roman', 10)
        self.drawString(LETTER[0]-x, 65, page)
        self.restoreState()    


class MyDocTemplate(SimpleDocTemplate):
     def __init__(self, filename, **kw):
         self.allowSplitting = 0
         apply(SimpleDocTemplate.__init__, (self, filename), kw)

     def afterFlowable(self, flowable):
         global user_name
         #"Registers TOC entries."
         if flowable.__class__.__name__ == 'Paragraph':
             text = flowable.getPlainText()
             style = flowable.style.name
             if style == 'Heading1':
                 self.notify('TOCEntry', (0, text, self.page))
                 self.canv.bookmarkPage(user_name)
                 #print "toc entry"
             if style == 'h2':
                 self.notify('TOCEntry', (1, text, self.page))
                 
        
def clean_filename(dirty_filename):

    if dirty_filename[0:4] == "Re: ":
        dirty_filename = dirty_filename[4:90]
    return re.sub('[^\w\-_\. ]', '', dirty_filename).strip()


class MLStripper(HTMLParser):
    def __init__(self):
        self.reset()
        self.fed = []
    def handle_data(self, d):
        self.fed.append(d)
    def get_data(self):
        return ''.join(self.fed)

def strip_tags(html):
    #print "html=",html
    s = MLStripper()
    s.feed(html)
    strText = s.get_data()
    strText = '\n'.join(x.strip() for x in strText.split('\n'))
    return strText

    
centered = PS(name = 'centered',
    fontSize = 30,
    leading = 16,
    alignment = 1,
    spaceAfter = 20)
 
styles=stylesheet()
curTopic = -1
Story = []
ptext = ""


toc = TableOfContents()
toc.levelStyles = [
    PS(fontName='Times-Bold', fontSize=20, name='TOCHeading1', leftIndent=20, firstLineIndent=-20, spaceBefore=10, leading=16),
    PS(fontSize=18, name='TOCHeading2', leftIndent=40, firstLineIndent=-20, spaceBefore=5, leading=12),
]
Story.append(toc)
Story.append(Paragraph('<b>Table of contents</b>', centered))
Story.append(PageBreak())


def parseAttrib(attr_name, attr_str):
    start_pos = attr_str.find(attr_name + '="')
    if start_pos > -1:
        start_pos += len(attr_name) + 2
        end_pos = attr_str.find('"', start_pos + 1)
        if end_pos > -1:
            return attr_str[start_pos:end_pos]
        else:
          #print "parse error 1", attr_name, attr_str
          return ""
    else:
        #print "parse error 2", attr_name, attr_str
        return ""

def buildBuddyDict(buddy_path):
    global buddy_dict
    buddy_files = [name for name in os.listdir(buddy_path) if name.lower().find("buddies") > -1]
    for buddy_file in buddy_files:
        with open(os.path.join(buddy_path, buddy_file)) as f:
            lines = f.read().splitlines()
            for line in lines:
                cur_line = urllib.unquote(line).strip().replace(chr(0), '')
                if cur_line[:10] == "<buddy uri":
                    buddy_uri = parseAttrib("buddy uri", cur_line)
                    if buddy_uri:
                        buddy_uri = buddy_uri[8:-2]
                        buddy_email = buddy_uri.split(":")[0]
                        buddy_name = buddy_uri.split(":")[1]
                        if buddy_email not in buddy_dict:
                            buddy_dict[buddy_email] = buddy_name
                elif cur_line[:17] == "<groupchat medium":
                    buddy_email = parseAttrib("name", cur_line)
                    searchObj = re.findall(r'renamed="\d">(.*)</groupchat', cur_line)
                    if searchObj:
                        buddy_name = searchObj[0]
                        if buddy_name:
                            if buddy_email not in buddy_dict:
                                buddy_dict[buddy_email] = buddy_name

def GetUserFromDict(bd, id):
    if id in bd:
        return bd[id]
    else:
        return id
    
def maxSize(image, maxSize):
    """ im = maxSize(im, (maxSizeX, maxSizeY), method = Image.BICUBIC)
    Adapted from : https://mail.python.org/pipermail/image-sig/2006-January/003724.html

    Resizes a PIL image to a maximum size specified while maintaining
    the aspect ratio of the image.  Similar to Image.thumbnail(), but allows
    usage of different resizing methods and does NOT modify the image in-place."""

    imAspect = float(image.size[0])/float(image.size[1])
    outAspect = float(maxSize[0])/float(maxSize[1])
    if imAspect >= outAspect:
        return image.resize((maxSize[0], int((float(maxSize[0])/imAspect) + 0.5)), PIL.Image.ANTIALIAS)
    else:
        return image.resize((int((float(maxSize[1])*imAspect) + 0.5), maxSize[1]), PIL.Image.ANTIALIAS)    


print "Stephen Genusa's Trillian Message to PDF Exporter"    
start_path = os.path.expanduser('~\\AppData\Roaming\\Trillian\\Users\\')
user_dirs = [name for name in os.listdir(start_path) if os.path.isdir(os.path.join(start_path, name)) and name.find("%40") > -1]
for user_dir in user_dirs:
    trillian_path = os.path.join(start_path, user_dir)
    #print "Trillian user path is", trillian_path
    
    buildBuddyDict(trillian_path)
    
    Story = []
    for person_group in ['Query', 'Channel']:
        astra_logs_path = os.path.join(trillian_path, "logs", "ASTRA", person_group)
        astra_log_files = [name for name in os.listdir(astra_logs_path) if os.path.splitext(name)[-1].lower() == ".xml" and name.find("assets") == -1] #   and name.find("ann") > -1
        for cur_astra_log in astra_log_files:
            current_trillian_user = urllib.unquote(user_dir)
            cur_remote_user = cur_astra_log.replace(".xml", "")
            bol_printed_user = False
            with open(os.path.join(astra_logs_path, cur_astra_log)) as f:
                lines = f.read().splitlines()
                bol_log_contains_msgs = False
                for line in lines:
                    if "<message " in line:
                        bol_log_contains_msgs = True
                        break
                if bol_log_contains_msgs:
                    print "  Processing:", cur_astra_log
                    for line in lines:
                        try:
                            while line[0] != "<":
                                line = line[1:]
                            cur_line = line.strip().replace(chr(0), '')
                            # "<groupmessage", "<filetransfer", "<history vers"

                            msg_time = int(parseAttrib("time", cur_line))
                            msg_type = parseAttrib("type", cur_line)
                            msg_text = urllib.unquote(parseAttrib("text", cur_line))
                            msg_from = urllib.unquote(parseAttrib("from", cur_line))

                            d = datetime.datetime.fromtimestamp(msg_time)

                            user_name = GetUserFromDict(buddy_dict, cur_remote_user)
                            
                            if not bol_printed_user:
                                htext = "Trillian messages with " + user_name
                                Story.append(Paragraph(htext, styles["h1"]))
                                bol_printed_user = True
                                    
                            if cur_line[:8] == "<session":
                                if msg_type == "start":
                                    sess_text = "Session start " + d.strftime('%m/%d/%Y %I:%M%p')
                                            
                            if cur_line[:8] == "<message":
                                if sess_text:
                                    Story.append(Paragraph(sess_text, styles["h2"]))
                                    sess_text = ""
                                
                                if msg_type in ["outgoing_privateMessage", "outgoing_groupMessage"]:
                                    user_name = GetUserFromDict(buddy_dict, current_trillian_user)
                                elif msg_type == "incoming_privateMessage":
                                    user_name = GetUserFromDict(buddy_dict, cur_remote_user)
                                elif msg_type == "incoming_groupMessage":
                                    user_name = GetUserFromDict(buddy_dict, msg_from)
                                    
                                ptext = d.strftime('%I:%M%p') + " " + user_name + ": " + msg_text.replace("\n", "<br />").replace("</a>", "</a><br />")
                                bs4obj = BeautifulSoup(ptext)
                                ptext = bs4obj.prettify()
                                if ptext.count("<a href") > ptext.count("</a>"):
                                    ptext.replace("</a>", "</a></a>")
                                Story.append(Paragraph(ptext, styles["Normal"]))
                                msg_temp = strip_tags(msg_text)
                                try:
                                    if "trillian.im" in msg_text:
                                        url_regex = "(http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F])))"
                                        html = BeautifulSoup(msg_text)
                                        urls = html('a')
                                        for url in urls:
                                            url = strip_tags(str(url))
                                            requests.adapters.DEFAULT_RETRIES = 5
                                            r = requests.get(url, stream=True)
                                            if r.status_code == 200:
                                                print "    Getting image..."
                                                with open("tmp.img", 'wb') as f:
                                                    for chunk in r.iter_content(1024):
                                                        f.write(chunk)
                                                img1 = StringIO()
                                                pil_image = PIL.Image.open("tmp.img")
                                                #print pil_image.size, "==>", url
                                                if pil_image.size[0] > 400 or pil_image.size[1] > 400:
                                                    pil_image = maxSize(pil_image, (450, 450))
                                                pil_image.save(img1, 'PNG')
                                                img1.seek(0)
                                                Story.append(Image(img1))
                                                try:
                                                    os.remove("tmp.img")
                                                except:
                                                    pass
                                except:
                                    #print "Exception:"
                                    #print '-'*60
                                    #traceback.print_exc(file=sys.stdout)
                                    #print '-'*60
                                    pass
                        except:
                            #print "Exception:"
                            #print '-'*60
                            #traceback.print_exc(file=sys.stdout)
                            #print '-'*60
                            pass
            if bol_log_contains_msgs:
                Story.append(PageBreak())
            sess_text = ""
    print "Building PDF: please wait..."
    strFileName = "Trillian-History.pdf"
    if os.path.isfile(strFileName):
        os.remove(strFileName)
    doc = MyDocTemplate(strFileName, pagesize=letter)
    doc.multiBuild(Story, canvasmaker=FooterCanvas)
    print "         PDF:", strFileName, "ready."
