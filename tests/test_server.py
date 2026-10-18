"""Regression tests using only fictional DOCX packages made in temporary storage."""

import hashlib
import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest
from unittest import mock
import zipfile


# Load the actual server file without shadowing the installed mcp SDK package.
SERVER_PATH = Path(__file__).resolve().parents[1] / "mcp" / "server.py"
spec = importlib.util.spec_from_file_location("docdoc_server", SERVER_PATH)
assert spec is not None and spec.loader is not None
server = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = server
spec.loader.exec_module(server)

DOCUMENT = "word/document.xml"
STYLES = "word/styles.xml"
OLD = '<w:sz w:val="30"/>'
NEW = '<w:sz w:val="32"/>'


def fictional_parts(repeats=1):
    """Create a small OPC fixture; no files, templates or private data are read."""
    paragraph = (
        '<w:p><w:r><w:rPr>' + OLD + '</w:rPr>'
        '<w:t>Fictional heading</w:t></w:r></w:p>'
    )
    return {
        "[Content_Types].xml": (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
            '</Types>'
        ).encode(),
        "_rels/.rels": (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships>'
        ).encode(),
        DOCUMENT: (
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body>' + paragraph * repeats + '<w:sectPr/></w:body></w:document>'
        ).encode(),
        "word/_rels/document.xml.rels": (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>'
        ).encode(),
        STYLES: (
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:style w:type="paragraph" w:styleId="Normal">'
            '<w:name w:val="Normal"/><w:rPr><w:color w:val="000000"/></w:rPr>'
            '</w:style></w:styles>'
        ).encode(),
    }


def write_parts(path, parts):
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            archive.writestr(name, content)


def read_parts(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


class ExactEditTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="docdoc-test-")
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)
        self.source = self.directory / "fictional-input.docx"
        self.output = self.directory / "fictional-output.docx"
        self.make_source()
        # This runs after every test, including rejected writes and failed tests.
        self.addCleanup(self.assert_source_unchanged)

    def make_source(self, repeats=1):
        write_parts(self.source, fictional_parts(repeats))
        self.source_bytes = self.source.read_bytes()

    def assert_source_unchanged(self):
        self.assertEqual(self.source.read_bytes(), self.source_bytes)

    def patch(self, output=None, old=OLD):
        return server.patch_docx_part_exact(
            str(self.source), str(output or self.output), DOCUMENT,
            old_fragment=old, new_fragment=NEW, expected_count=1,
        )

    def validate(self):
        candidate_bytes = self.output.read_bytes()
        result = server.validate_docx_edit(
            str(self.source), str(self.output),
            {"exact_replacements": [{
                "part": DOCUMENT, "old_fragment": OLD,
                "new_fragment": NEW, "expected_count": 1,
            }]},
        )
        self.assertEqual(self.output.read_bytes(), candidate_bytes)
        return result

    def assert_no_output(self):
        self.assertFalse(self.output.exists())
        self.assertEqual(list(self.directory.glob(".word-enhanced-*.docx")), [])

    def test_exact_edit_changes_only_the_requested_part(self):
        inspected = server.inspect_docx(str(self.source))
        self.assertTrue(inspected["basic_opc_valid"])
        self.assertEqual(inspected["opc_errors"], [])
        result = self.patch()
        before, after = read_parts(self.source), read_parts(self.output)
        self.assertEqual(set(before), set(after))
        self.assertEqual(after[DOCUMENT], before[DOCUMENT].replace(OLD.encode(), NEW.encode()))
        self.assertEqual(
            {name for name in before if before[name] != after[name]}, {DOCUMENT}
        )
        self.assertEqual(result["replacements"], 1)
        self.assertEqual(result["input_sha256"], hashlib.sha256(self.source_bytes).hexdigest())
        compared = server.compare_docx(str(self.source), str(self.output))
        self.assertEqual(compared["package"]["raw_changed_parts"], [DOCUMENT])
        self.assertTrue(compared["objects"]["styles"]["unchanged"])
        validated = self.validate()
        self.assertEqual(validated["status"], "PASS")
        self.assertTrue(validated["target_completed"])
        self.assertEqual(validated["non_target_change_count"], 0)

    def test_ambiguous_target_is_rejected_without_output(self):
        self.make_source(repeats=2)
        with self.assertRaisesRegex(ValueError, "Expected 1 matches.*found 2"):
            self.patch()
        self.assert_no_output()

    def test_missing_target_is_rejected_without_output(self):
        with self.assertRaisesRegex(ValueError, "Expected 1 matches.*found 0"):
            self.patch(old='<w:sz w:val="99"/>')
        self.assert_no_output()

    def test_existing_output_is_preserved_and_rejected(self):
        write_parts(self.output, fictional_parts(repeats=2))
        existing = self.output.read_bytes()
        with self.assertRaisesRegex(ValueError, "new path distinct from input"):
            self.patch()
        self.assertEqual(self.output.read_bytes(), existing)

    def test_input_cannot_be_used_as_output(self):
        with self.assertRaisesRegex(ValueError, "new path distinct from input"):
            self.patch(output=self.source)
        self.assert_no_output()

    def test_validator_rejects_non_target_part_change(self):
        self.patch()
        parts = read_parts(self.output)
        parts[STYLES] = parts[STYLES].replace(b'000000', b'FF0000')
        write_parts(self.output, parts)
        result = self.validate()
        self.assertEqual(result["status"], "FAIL")
        self.assertFalse(result["pass"])
        self.assertTrue(result["target_completed"])
        self.assertIn(STYLES, result["unexpected_xml_changes"])
        self.assertTrue(any("Unexpected XML changes: " + STYLES in reason
                            for reason in result["non_target_changes"]))

    def test_validator_rejects_extra_change_inside_target_part(self):
        self.patch()
        parts = read_parts(self.output)
        parts[DOCUMENT] = parts[DOCUMENT].replace(b'Fictional heading', b'Other fictional text')
        write_parts(self.output, parts)
        result = self.validate()
        self.assertEqual(result["status"], "FAIL")
        self.assertFalse(result["target_completed"])
        self.assertTrue(any("Actual output differs from exact replacement" in reason
                            for reason in result["non_target_changes"]))

    def test_patch_does_not_publish_non_target_write_corruption(self):
        real_write = zipfile.ZipFile.writestr

        def corrupt_style(archive, info, data, *args, **kwargs):
            # Fault injection at ZIP writing, while all server checks run normally.
            name = info.filename if isinstance(info, zipfile.ZipInfo) else info
            if name == STYLES:
                data = data.replace(b'000000', b'FF0000')
            return real_write(archive, info, data, *args, **kwargs)

        with mock.patch.object(zipfile.ZipFile, "writestr", new=corrupt_style):
            with self.assertRaisesRegex(ValueError, "non-target uncompressed part changed"):
                self.patch()
        self.assert_no_output()


if __name__ == "__main__":
    unittest.main()
