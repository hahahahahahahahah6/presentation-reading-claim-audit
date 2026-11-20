import json
import subprocess
import sys
import tempfile
from pathlib import Path
import unittest
from presentation_reading_claim_audit.audit import audit, evidence
from presentation_reading_claim_audit.phrases import PRESENTATIONREADING, NOPRESENTATIONREADING, GENERIC_NAMES

YES = ('presentation reading', 'read presentation', 'reads presentation', 'reading presentation', 'read presentations', 'reads presentations', 'reading presentations', 'load presentation', 'loads presentation', 'loading presentation', 'load presentations', 'loads presentations', 'loading presentations', 'open presentation', 'opens presentation', 'opening presentation', 'open presentations', 'opens presentations', 'opening presentations', 'parse presentation', 'parses presentation', 'parsing presentation', 'parse presentations', 'parses presentations', 'parsing presentations', 'import presentation', 'imports presentation', 'importing presentation', 'import presentations', 'imports presentations', 'importing presentations', 'presentation reader', 'presentation parser', 'presentation importer', 'presentation loading', 'read pptx', 'reads pptx', 'reading pptx', 'read pptx file', 'reads pptx file', 'load pptx', 'loads pptx', 'loading pptx', 'open pptx', 'opens pptx', 'opening pptx', 'read ppt', 'reads ppt', 'reading ppt', 'load pptx file')
NO = ('no presentation reading', 'does not read presentation', 'does not read presentations', 'does not support presentation reading', 'without presentation reading', 'no presentation reader', 'does not load presentation', 'does not load presentations', 'does not open presentation', 'does not open presentations', 'does not parse presentation', 'does not parse presentations', 'does not import presentation', 'does not import presentations', 'no presentation parser', 'does not read pptx', 'does not read pptx file', 'does not load pptx', 'does not open pptx', 'does not read ppt', 'does not read ppt file', 'does not load pptx file', 'presentation reading not supported', 'no presentation reading support')
GENERIC = ('read_presentation', 'read_presentations', 'load_presentation', 'load_presentations', 'open_presentation', 'open_presentations', 'parse_presentation', 'parse_presentations', 'import_presentation', 'import_presentations', 'read_pptx', 'read_pptx_file', 'load_pptx', 'open_pptx', 'read_ppt', 'read_ppt_file', 'load_pptx_file', 'parse_pptx', 'parse_ppt', 'read_slides', 'load_slides', 'open_slides', 'presentation_reader', 'read_pptx_document')

class AuditTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
    def write(self, name, text):
        p = self.root / name
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(text, encoding='utf-8')
    def cli(self, *args):
        return subprocess.run([sys.executable, '-m', 'presentation_reading_claim_audit', *args], capture_output=True, text=True)
    def test_phrase_oracle(self):
        self.assertEqual(PRESENTATIONREADING, YES)
        self.assertEqual(NOPRESENTATIONREADING, NO)
        self.assertEqual(GENERIC_NAMES, GENERIC)
    def test_findings(self):
        self.write('README.md', 'read presentation')
        self.assertEqual(audit(self.root)['findings'][0]['rule'], 'documented-presentation-reading-claim-false')
        self.assertEqual(self.cli('audit', str(self.root)).returncode, 2)
        self.write('tests/test_a.py', 'import pptx\npptx.Presentation("x")')
        self.assertEqual(audit(self.root)['findings'][0]['severity'], 'MEDIUM')
        self.assertEqual(self.cli('audit', str(self.root)).returncode, 0)
        self.assertEqual(self.cli('audit', str(self.root), '--fail-on', 'medium').returncode, 2)
        self.write('app.py', 'import pptx\npptx.Presentation("x")')
        self.assertEqual(audit(self.root)['verdict'], 'PASS')
        self.write('README.md', 'does not read presentation')
        self.assertEqual(audit(self.root)['findings'][0]['rule'], 'documented-no-presentation-reading-claim-false')
        self.assertEqual(self.cli('audit', str(self.root)).returncode, 2)
        self.write('README.md', 'unrelated')
        self.assertEqual(audit(self.root)['verdict'], 'PASS')
    def test_docs(self):
        self.write('README.md', 'does not read presentation')
        self.assertEqual(audit(self.root)['verdict'], 'PASS')
        self.write('README.md', '```\nread pptx\n```')
        self.assertEqual(audit(self.root)['verdict'], 'PASS')
        self.write('README.md', '`read pptx`')
        self.assertEqual(audit(self.root)['verdict'], 'FAIL')
        self.write('README.md', 'nothing')
        self.write('docs/a.md', 'read pptx')
        self.assertEqual(audit(self.root, ['docs/'])['verdict'], 'PASS')
    def test_routes(self):
        cases = ['from pptx import Presentation as P\nP("x")', 'import pptx\npptx.Presentation("x")', 'import pptx as p\np.Presentation("x")', 'from pptx import Presentation\nPresentation(pptx="x")', 'from pkg import read_presentation as f\nf()']
        cases += ['x.' + n + '()' for n in GENERIC]
        for source in cases:
            with self.subTest(source=source): self.assertTrue(evidence(source))
    def test_non_routes(self):
        cases = ['import ' + n for n in ('pptx', 'spire', 'aspose', 'odf', 'pypandoc', 'textract', 'win32com')]
        cases += ['from pptx import Presentation\nPresentation()', 'from pptx import Presentation\np=Presentation()\np.add_slide()\np.save("x")', 'import spire.presentation\nspire.presentation.Presentation("x")', 'import aspose.slides\naspose.slides.Presentation("x")', 'from odf.opendocument import load\nload("x")', 'import pypandoc\npypandoc.convert_file("x")', 'import textract\ntextract.process("x")', 'import win32com.client\nwin32com.client.Dispatch("PowerPoint.Application")', 'import subprocess\nsubprocess.run(["soffice", "x"])', 'import subprocess\nsubprocess.run(["pandoc", "x"])', 'from docx import Document\nDocument("x")', 'from openpyxl import load_workbook\nload_workbook("x")', 'import pandas\npandas.read_excel("x")', 'import pypdf\npypdf.extract_text()', 'import pdfplumber\npdfplumber.extract_text()', 'open("x.txt").read()', 'from pathlib import Path\nPath("x").read_text()', 'from PIL import Image\nImage.open("x")', '"presentation"', 'mindee.parse(InvoiceV4)', 'azure.begin_analyze_document("prebuilt-read")', 'veryfi.Client()', 'documentai.process_document()', 'textract.analyze_document()', 'eden.universal_ai()', 'llamaparse.parse()', 'nanonets.ocr(model_id="opaque")', 'docsumo.upload()', 'import pptx\npptx = object()\npptx.Presentation("x")', 'from pptx import Presentation\ndef Presentation(x): pass\nPresentation("x")', 'def read_pptx(): pass\nx.read_pptx()', 'from pptx import Presentation\nwith ctx() as Presentation: pass\nPresentation("x")', 'from pptx import Presentation\nPresentation', 'from pptx import Presentation\ninvalid !']
        for source in cases:
            with self.subTest(source=source): self.assertFalse(evidence(source))
    def test_exclusions_json(self):
        self.write('README.md', 'read pptx')
        self.write('part.py', 'import pptx')
        self.write('other.py', 'pptx.Presentation("x")')
        self.write('examples/e.py', 'import pptx\npptx.Presentation("x")')
        self.write('fixtures/f.py', 'import pptx\npptx.Presentation("x")')
        (self.root / 'link.py').symlink_to(self.root / 'examples/e.py')
        result = json.loads(self.cli('audit', str(self.root), '--json').stdout)
        self.assertEqual(result['verdict'], 'FAIL')
        self.assertIn('message', result['findings'][0])
        self.assertEqual(self.cli('audit', str(self.root / 'missing')).returncode, 1)
        self.assertEqual(self.cli('bad').returncode, 1)
    def test_version(self):
        self.assertEqual(self.cli('--version').stdout.strip(), '0.1.0')

if __name__ == '__main__': unittest.main()
