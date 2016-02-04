import smtplib

from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.mime.text import MIMEText
from email import Encoders

import StringIO
import pdfminer.pdfinterp
import pdfminer.pdfpage
import pdfminer.converter

# Pdf conversion to HTML so the mails have a body
def pdf2html(data):
    outfp = StringIO.StringIO()
    infp = StringIO.StringIO(data)
    rsrcmgr = pdfminer.pdfinterp.PDFResourceManager(caching=True)
    device = pdfminer.converter.HTMLConverter(rsrcmgr, outfp, codec='utf-8', scale=1,
                           layoutmode='normal',
                           laparams=pdfminer.layout.LAParams(),
                           imagewriter=None)
    interpreter = pdfminer.pdfinterp.PDFPageInterpreter(rsrcmgr, device)
    for page in pdfminer.pdfpage.PDFPage.get_pages(infp, set(),
                                                   maxpages=0, password=None,
                                                   caching=True, check_extractable=True):
        interpreter.process_page(page)

    device.close()
    res = outfp.getvalue()
    outfp.close()
    return res

class Backend(object):

    def __init__(self, config):
        self.mailto = config["to"]
        self.mailfrom = config["from"]
        self.mailserver = config["server"]

    def new_collection(self):
        return DocumentCollection(self)


class DocumentCollection(object):

    def __init__(self, backend):
        self.backend = backend
        self.msg = MIMEMultipart()
        self.msg['To'] = backend.mailto

    def set_metadata(self, subject, sender, date):
        self.msg['Subject'] = subject
        self.msg['From'] = self.backend.mailfrom % sender

    def attach(self, filename, mimetype, binary):
        html = pdf2html(binary)
        self.msg.attach(MIMEText(html, 'html'))

        part = MIMEBase(*mimetype.split("/"))
        part.set_payload(binary)
        Encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="%s"' % filename)
        self.msg.attach(part)

    def execute(self):
        server = smtplib.SMTP(self.backend.mailserver)
        server.sendmail(self.msg['From'], self.backend.mailto, self.msg.as_string())
