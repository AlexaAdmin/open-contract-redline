# Builds examples/sample-services-agreement.docx, a short fictional contract used by the tests and for trying the tool.
# It deliberately has text split across differently formatted runs, a hyperlink, bookmarks, curly quotes and a table.
# python examples/make_sample.py   (needs: pip install python-docx)
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
import os

doc = Document()
doc.styles['Normal'].font.name = 'Calibri'
doc.styles['Normal'].font.size = Pt(11)

doc.add_heading('SERVICES AGREEMENT', level=1)
doc.add_paragraph('This Services Agreement (the “Agreement”) is made between Northwind Analytics Ltd (the “Supplier”) and Contoso Retail Ltd (the “Customer”).')

def clause(num, title, *parts, bookmark=None):
    p = doc.add_paragraph()
    if bookmark:
        bs = OxmlElement('w:bookmarkStart'); bs.set(qn('w:id'), str(bookmark[0])); bs.set(qn('w:name'), bookmark[1]); p._p.append(bs)
    r = p.add_run(f'{num} {title}. '); r.bold = True
    if bookmark:
        be = OxmlElement('w:bookmarkEnd'); be.set(qn('w:id'), str(bookmark[0])); p._p.append(be)
    for text, fmt in parts:
        run = p.add_run(text)
        if fmt == 'b': run.bold = True
        if fmt == 'i': run.italic = True
    return p

clause('1', 'Services', ('The Supplier shall provide the data analytics services described in Schedule 1 (the “Services”) ', ''), ('with reasonable skill and care', 'i'), ('.', ''))
clause('2', 'Fees', ('The Customer shall pay all invoices within ', ''), ('seven (7) days', 'b'), (' of the invoice date. Late payments bear interest at 5% per month.', ''))
clause('3', 'Term', ('This Agreement continues for 36 months and renews automatically for further 36-month periods unless the Customer gives 180 days’ notice.', ''), bookmark=(10, '_RefTerm'))
clause('4', 'Liability', ('The Supplier’s total liability under this Agreement shall not exceed ', ''), ('the fees paid in the one (1) month', 'b'), (' before the claim. The Customer’s liability is unlimited.', ''))
clause('5', 'Indemnity', ('The Customer shall indemnify the Supplier against all losses of any kind arising in connection with this Agreement.', ''))
clause('6', 'Data', ('The Supplier may use Customer data for any purpose, including training its own products.', ''))
p = clause('7', 'Notices', ('Notices must be sent to the addresses on ', ''))
# a hyperlink inside a clause
hl = OxmlElement('w:hyperlink'); hl.set(qn('w:anchor'), '_RefTerm')
hr = OxmlElement('w:r'); ht = OxmlElement('w:t'); ht.text = 'the cover page'; hr.append(ht); hl.append(hr); p._p.append(hl)
p.add_run('.')
clause('8', 'Governing law', ('This Agreement is governed by the laws chosen by the Supplier from time to time.', ''))
t = doc.add_table(rows=2, cols=2)
t.cell(0, 0).text = 'Supplier signature'; t.cell(0, 1).text = 'Customer signature'
t.cell(1, 0).text = 'Name:'; t.cell(1, 1).text = 'Name:'
doc.add_paragraph('')
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sample-services-agreement.docx')
doc.save(out)
print(out)
