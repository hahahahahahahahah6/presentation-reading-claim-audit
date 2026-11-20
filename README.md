# Documentation claim auditor

This offline static auditor compares documented claims about the checked capability with Python source patterns. It examines the root README and Markdown under docs, excluding fenced code; inline code counts. It reports unsupported affirmative claims and denials contradicted by production evidence. Tests are reported separately.

```sh
presentation-reading-claim-audit audit . --json --fail-on medium --ignore vendor/
python3 -m presentation_reading_claim_audit --version
```

Exit status 0 means PASS, 2 means a finding meets the threshold, and 1 indicates an input or usage error. The default threshold is high; use medium to include test-only evidence. Ignore prefixes may be repeated. Source is parsed, never executed; symlinks, examples, fixtures, and documentation are excluded from source evidence.

Findings use these identifiers:

```
documented-presentation-reading-claim-false
documented-no-presentation-reading-claim-false
```

## Differentiation

This tool checks a distinct file-access capability. Auditors for Word documents, spreadsheets, PDF extraction, file generation, and document extraction cover different operations. Other out-of-scope claims include invoice, receipt, identity, bank check, bank statement, tax form, pay stub, health insurance card, business card, credit card, contract, mortgage, resume, bill of lading, delivery note, carte grise, energy bill, marriage certificate, bank account details, nutrition facts label, material certificate, US mail, invoice splitting, multiple receipt detection, French health card, Indian passport, driver license, and vaccination card extraction; document cropping, electronic signatures, and payment processing are also outside the phrase lists.

## Limitations

Exact phrases can describe opening an application, importing already processed data, a job title, or a wrapper around a writer. The abbreviation for parts per thousand can also trigger a claim. A file object may already have been opened elsewhere. External converters are deliberately excluded, so a project relying on them can be flagged. Static patterns count dead code, may miss alias chains, and require each route in one module. Denial handling depends on the listed phrases.

The constructor below with an argument counts even if no slides, shapes, or text are subsequently accessed. Without an argument it loads a built-in default template and does not count; adding slides and saving that new file alone does not count either.

```python
from pptx import Presentation
Presentation(path)
Presentation()
```

The following are not detected: commercial SDK constructors from spire.presentation and aspose.slides; odfpy loading, which covers several formats; pypandoc conversion and Python textract, which invoke external programs; PowerPoint COM automation; and subprocess invocations of LibreOffice, soffice, or pandoc. Hand-written OOXML or ZIP XML parsing is also missed. Access to text, HTML, Markdown, image, PDF, Word, and spreadsheet files alone does not count, including pypdf, PyPDF2, pdfplumber, PyMuPDF, python-docx, openpyxl, and pandas operations.

Mindee, Azure Document Intelligence, Veryfi, Google Document AI, Eden AI, Textract, LlamaParse, Nanonets, Docsumo, ABBYY Vantage, Rossum, Klippa, Parseur, and Extend calls do not count: these are extraction services or use opaque model, processor, skill, queue, template, or workflow identifiers, or configure the document type outside a generic upload call.
