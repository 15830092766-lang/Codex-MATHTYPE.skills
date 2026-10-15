"""Synthetic DOCX tests; no Office application or private files needed."""

import hashlib
import json
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree as E

from reference_math import NS, W, RID, Package, apply_manifest, inventory, xp


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def create_package(path, reference=False):
    doc = E.Element(W + "document", nsmap={k: v for k, v in NS.items() if k not in ("p", "ct")})
    body = E.SubElement(doc, W + "body")
    rels = E.Element("{%s}Relationships" % NS["p"], nsmap={None: NS["p"]})
    styles = E.Element(W + "styles", nsmap={"w": NS["w"]})
    for sid in ("Normal", "MTDisplayEquation"):
        style = E.SubElement(styles, W + "style")
        style.set(W + "type", "paragraph")
        style.set(W + "styleId", sid)
        if sid != "Normal":
            base = E.SubElement(style, W + "basedOn")
            base.set(W + "val", "Normal")
    types = E.Element("{%s}Types" % NS["ct"], nsmap={None: NS["ct"]})
    for extension, content_type in (("bin", "application/vnd.openxmlformats-officedocument.oleObject"),
                                    ("wmf", "image/x-wmf"), ("xml", "application/xml")):
        entry = E.SubElement(types, "{%s}Default" % NS["ct"])
        entry.set("Extension", extension)
        entry.set("ContentType", content_type)
    parts = {}
    def paragraph(*texts):
        p = E.SubElement(body, W + "p")
        for text in texts:
            r = E.SubElement(p, W + "r")
            E.SubElement(r, W + "t").text = text
        return p
    def equation(p, i):
        run = E.SubElement(p, W + "r")
        obj = E.SubElement(run, W + "object")
        shape = E.SubElement(obj, "{%s}shape" % NS["v"])
        shape.set("id", f"_shape{i}")
        shape.set("style", "width:20pt;height:12pt")
        img = E.SubElement(shape, "{%s}imagedata" % NS["v"])
        img.set(RID, f"rImage{i}")
        ole = E.SubElement(obj, "{%s}OLEObject" % NS["o"])
        ole.set("ProgID", "Equation.DSMT4")
        ole.set("ShapeID", f"_shape{i}")
        ole.set("ObjectID", f"_obj{i}")
        ole.set(RID, f"rOle{i}")
        for rid, kind, part in ((f"rImage{i}", "image", f"media/e{i}.wmf"),
                                (f"rOle{i}", "oleObject", f"embeddings/e{i}.bin")):
            rel = E.SubElement(rels, "{%s}Relationship" % NS["p"])
            rel.set("Id", rid)
            rel.set("Type", NS["r"] + "/" + kind)
            rel.set("Target", part)
            parts["word/" + part] = f"synthetic-{kind}-{i}".encode()
    if reference:
        p = paragraph()
        ppr = E.SubElement(p, W + "pPr")
        st = E.SubElement(ppr, W + "pStyle")
        st.set(W + "val", "MTDisplayEquation")
        tabs = E.SubElement(ppr, W + "tabs")
        for align, position in (("center", "4500"), ("right", "9000")):
            t = E.SubElement(tabs, W + "tab")
            t.set(W + "val", align)
            t.set(W + "pos", position)
        equation(p, 1)
        equation(p, 2)
        p = paragraph("参考符号：")
        equation(p, 3)
    else:
        paragraph("FORMULA_A")
        paragraph("其中，", "Δs", "_i", "为点距；再次使用", "Δs_i", "。")
    parts.update({"word/document.xml": E.tostring(doc),
                  "word/_rels/document.xml.rels": E.tostring(rels),
                  "word/styles.xml": E.tostring(styles),
                  "[Content_Types].xml": E.tostring(types)})
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        for key, value in parts.items():
            archive.writestr(key, value)


class ReferenceMathTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.source = self.root / "reference.docx"
        self.input = self.root / "input.docx"
        self.mapping = self.root / "mapping.json"
        create_package(self.source, reference=True)
        create_package(self.input)
        self.manifest = {"display": [{"marker": "FORMULA_A", "source_paragraph": 0, "number": "(1)"}],
                         "inline": [{"text": "Δs_i", "source_paragraph": 1,
                                     "source_object": 0, "expected_count": 2}]}
        self.mapping.write_text(json.dumps(self.manifest), encoding="utf-8")

    def tearDown(self):
        self.temp.cleanup()

    def test_multi_object_formula_inline_split_runs_and_source_immutability(self):
        before = digest(self.source), digest(self.input)
        output = self.root / "result.docx"
        report = apply_manifest(self.input, self.source, self.mapping, output)
        self.assertEqual(report["replaced"], {"display": 1, "inline": 2})
        package = Package(output)
        paragraphs = package.paragraphs()
        self.assertEqual(len(xp(paragraphs[0], ".//o:OLEObject")), 2)
        self.assertEqual(len(xp(paragraphs[1], ".//o:OLEObject")), 2)
        self.assertEqual("".join(xp(paragraphs[1], ".//w:t/text()")), "其中，为点距；再次使用。")
        self.assertEqual("".join(xp(paragraphs[0], ".//w:t/text()")), "(1)")
        self.assertEqual(len(xp(paragraphs[0], "./w:r/w:tab")), 2)
        objects = xp(package.doc, ".//o:OLEObject")
        self.assertEqual(len(set(o.get("ObjectID") for o in objects)), 4)
        for obj, expected in zip(objects, (1, 2, 3, 3)):
            part = package.part_name(package.relationship(obj.get(RID)))
            self.assertEqual(package.parts[part], f"synthetic-oleObject-{expected}".encode())
            shape = xp(obj.getparent(), "./v:shape")[0]
            self.assertEqual(shape.get("id"), obj.get("ShapeID"))
        self.assertEqual(before, (digest(self.source), digest(self.input)))
        self.assertEqual(inventory(output)["object_count"], 4)

    def test_wrong_count_fails_before_output(self):
        self.manifest["inline"][0]["expected_count"] = 99
        self.mapping.write_text(json.dumps(self.manifest), encoding="utf-8")
        output = self.root / "never_written.docx"
        with self.assertRaisesRegex(ValueError, "expected 99"):
            apply_manifest(self.input, self.source, self.mapping, output)
        self.assertFalse(output.exists())

    def test_reference_and_input_cannot_be_overwritten(self):
        for forbidden in (self.source, self.input):
            with self.assertRaisesRegex(ValueError, "must differ"):
                apply_manifest(self.input, self.source, self.mapping, forbidden)


if __name__ == "__main__":
    unittest.main(verbosity=2)
