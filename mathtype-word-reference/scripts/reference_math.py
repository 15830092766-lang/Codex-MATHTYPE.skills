"""Offline MathType inventory, reference-object transplant, and storage audit.

No Office automation or equation generation occurs here. The manifest is the
operator's semantic contract; inspect reports positions, not decoded formulas.
"""

from __future__ import annotations

import argparse
from copy import deepcopy
import io
import json
from pathlib import Path
import posixpath
import re
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree as E


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/package/2006/relationships",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}
W = "{%s}" % NS["w"]
RID = "{%s}id" % NS["r"]
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
OLE = ".//o:OLEObject[@ProgID='Equation.DSMT4']"


def xp(node, query):
    return node.xpath(query, namespaces=NS)


def serialize(node):
    return E.tostring(node, xml_declaration=True, encoding="UTF-8", standalone=True)


def text_of(node):
    return "".join(xp(node, ".//w:t/text()"))


def run_text(text, template=None):
    run = E.Element(W + "r")
    if template is not None and template.find(W + "rPr") is not None:
        run.append(deepcopy(template.find(W + "rPr")))
    value = E.SubElement(run, W + "t")
    value.set(XML_SPACE, "preserve")
    value.text = text
    return run


def tab():
    run = E.Element(W + "r")
    E.SubElement(run, W + "tab")
    return run


class Package:
    def __init__(self, path):
        with ZipFile(path) as archive:
            self.parts = {n: archive.read(n) for n in archive.namelist()}
        self.doc = E.fromstring(self.parts["word/document.xml"])
        self.rels = E.fromstring(self.parts["word/_rels/document.xml.rels"])
        self.styles = E.fromstring(self.parts["word/styles.xml"])
        self.types = E.fromstring(self.parts["[Content_Types].xml"])

    def paragraphs(self):
        return xp(self.doc, "//w:p")

    def relationship(self, rid):
        matches = [r for r in self.rels if r.get("Id") == rid]
        if len(matches) != 1:
            raise ValueError(f"Missing or duplicate relationship: {rid}")
        rel = matches[0]
        if rel.get("TargetMode") == "External":
            raise ValueError(f"External relationship unsupported: {rid}")
        return rel

    def part_name(self, rel):
        target = rel.get("Target", "")
        if target.startswith("/"):
            part = target.lstrip("/")
        else:
            part = posixpath.normpath(posixpath.join("word", target))
        if part not in self.parts:
            raise ValueError(f"Relationship target is absent: {part}")
        return part

    def save(self, path):
        self.parts["word/document.xml"] = serialize(self.doc)
        self.parts["word/_rels/document.xml.rels"] = serialize(self.rels)
        self.parts["word/styles.xml"] = serialize(self.styles)
        self.parts["[Content_Types].xml"] = serialize(self.types)
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_name(path.name + ".tmp")
        with ZipFile(temporary, "w", ZIP_DEFLATED) as archive:
            for name, data in self.parts.items():
                archive.writestr(name, data)
        temporary.replace(path)


def inventory(path):
    package = Package(path)
    items = []
    for index, paragraph in enumerate(package.paragraphs()):
        objects = []
        for oi, obj in enumerate(xp(paragraph, OLE)):
            rel = package.relationship(obj.get(RID))
            objects.append({"source_object": oi, "part": package.part_name(rel),
                            "object_id": obj.get("ObjectID")})
        items.append({"source_paragraph": index, "text": text_of(paragraph),
                      "objects": objects})
    return {"paragraphs": items, "object_count": sum(len(i["objects"]) for i in items)}


class Transplanter:
    def __init__(self, target, reference):
        self.target = target
        self.reference = reference
        self.serial = 0
        self.style_map = {}

    def fresh(self):
        self.serial += 1
        return self.serial

    def import_style(self, style_id):
        if style_id in self.style_map:
            return self.style_map[style_id]
        found = [s for s in self.reference.styles if s.get(W + "styleId") == style_id]
        if not found:
            raise ValueError(f"Referenced style absent: {style_id}")
        new_id = "MR_" + style_id
        existing = {s.get(W + "styleId") for s in self.target.styles}
        while new_id in existing:
            new_id += "_copy"
        self.style_map[style_id] = new_id
        clone = deepcopy(found[0])
        clone.set(W + "styleId", new_id)
        clone.attrib.pop(W + "default", None)
        for link in xp(clone, "./w:basedOn|./w:next|./w:link"):
            link.set(W + "val", self.import_style(link.get(W + "val")))
        self.target.styles.append(clone)
        return new_id

    def copy_content_type(self, source_part, target_part):
        source_types = self.reference.types
        target_types = self.target.types
        extension = posixpath.splitext(source_part)[1].lstrip(".")
        overrides = [t for t in source_types if t.tag.endswith("Override")
                     and t.get("PartName") == "/" + source_part]
        if overrides:
            clone = deepcopy(overrides[0])
            clone.set("PartName", "/" + target_part)
            target_types.append(clone)
            return
        defaults = [t for t in source_types if t.tag.endswith("Default")
                    and t.get("Extension") == extension]
        if not defaults:
            raise ValueError(f"No content type for {source_part}")
        current = [t for t in target_types if t.tag.endswith("Default")
                   and t.get("Extension") == extension]
        if not current:
            target_types.append(deepcopy(defaults[0]))
        elif current[0].get("ContentType") != defaults[0].get("ContentType"):
            item = E.SubElement(target_types, "{%s}Override" % NS["ct"])
            item.set("PartName", "/" + target_part)
            item.set("ContentType", defaults[0].get("ContentType"))

    def clone(self, node):
        clone = deepcopy(node)
        rid_map = {}
        for related in xp(clone, ".//*[@r:id]"):
            old = related.get(RID)
            if old not in rid_map:
                rel = self.reference.relationship(old)
                kind = rel.get("Type", "").rsplit("/", 1)[-1]
                if kind not in ("image", "oleObject"):
                    raise ValueError(f"Unsupported equation relationship: {kind}")
                part = self.reference.part_name(rel)
                directory = "media" if kind == "image" else "embeddings"
                suffix = posixpath.splitext(part)[1]
                while True:
                    token = self.fresh()
                    destination = f"word/{directory}/mr_{token}{suffix}"
                    new_rid = f"rIdMR{token}"
                    if destination not in self.target.parts and not any(
                        r.get("Id") == new_rid for r in self.target.rels
                    ):
                        break
                self.target.parts[destination] = self.reference.parts[part]
                self.copy_content_type(part, destination)
                new_rel = E.SubElement(self.target.rels, "{%s}Relationship" % NS["p"])
                new_rel.set("Id", new_rid)
                new_rel.set("Type", rel.get("Type"))
                new_rel.set("Target", posixpath.relpath(destination, "word"))
                rid_map[old] = new_rid
            related.set(RID, rid_map[old])
        shape_ids = {}
        occupied = {n.get("id") for n in xp(self.target.doc, "//*[@id]")}
        for shape in xp(clone, ".//v:shape|.//v:shapetype"):
            new_id = f"_mr_shape_{self.fresh()}"
            while new_id in occupied:
                new_id = f"_mr_shape_{self.fresh()}"
            occupied.add(new_id)
            shape_ids[shape.get("id")] = new_id
            shape.set("id", new_id)
        existing_objects = {n.get("ObjectID") for n in xp(self.target.doc, ".//o:OLEObject")}
        for obj in xp(clone, ".//o:OLEObject"):
            obj.set("ShapeID", shape_ids.get(obj.get("ShapeID"), obj.get("ShapeID", "")))
            object_id = f"_{2300000000 + self.fresh()}"
            while object_id in existing_objects:
                object_id = f"_{2300000000 + self.fresh()}"
            existing_objects.add(object_id)
            obj.set("ObjectID", object_id)
        for shape in xp(clone, ".//v:shape"):
            old_type = shape.get("type", "").lstrip("#")
            if old_type in shape_ids:
                shape.set("type", "#" + shape_ids[old_type])
            elif old_type:
                # VML shapetypes can live outside the selected equation run.
                declared = xp(self.target.doc, ".//v:shapetype")
                if not any(s.get("id") == old_type for s in declared):
                    source = [s for s in xp(self.reference.doc, ".//v:shapetype")
                              if s.get("id") == old_type]
                    if source:
                        shape.addprevious(deepcopy(source[0]))
        for style in xp(clone, ".//w:pStyle|.//w:rStyle"):
            style.set(W + "val", self.import_style(style.get(W + "val")))
        return clone

    def source_runs(self, index):
        paragraph = self.reference.paragraphs()[index]
        runs = xp(paragraph, "./w:r[w:object/o:OLEObject[@ProgID='Equation.DSMT4']]")
        total = len(xp(paragraph, OLE))
        if not runs or sum(len(xp(r, OLE)) for r in runs) != total:
            raise ValueError(f"Paragraph {index}: unsupported nested equation structure")
        return paragraph, runs


def plain_chunks(paragraph):
    """Group adjacent plain direct runs; never match across objects/fields."""
    chunk = []
    for child in paragraph:
        plain = child.tag == W + "r" and len(xp(child, "./w:t")) > 0
        plain = plain and all(n.tag in (W + "rPr", W + "t") for n in child)
        if plain:
            chunk.append(child)
        else:
            if chunk:
                yield chunk
            chunk = []
    if chunk:
        yield chunk


def apply_manifest(input_path, reference_path, manifest_path, output_path):
    output_path = Path(output_path).resolve()
    if output_path in (Path(input_path).resolve(), Path(reference_path).resolve()):
        raise ValueError("Output must differ from input and reference")
    if output_path.exists():
        raise ValueError("Output already exists; choose a new candidate path")
    mapping = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
    target, reference = Package(input_path), Package(reference_path)
    copier = Transplanter(target, reference)
    changed = {"display": 0, "inline": 0}
    for entry in mapping.get("display", []):
        matches = [p for p in target.paragraphs() if text_of(p) == entry["marker"]]
        if len(matches) != 1 or xp(matches[0], ".//w:object|.//w:fldChar|.//w:fldSimple"):
            raise ValueError(f"Display marker must identify one clean paragraph: {entry['marker']}")
        source, source_runs = copier.source_runs(entry["source_paragraph"])
        paragraph = matches[0]
        for child in list(paragraph):
            paragraph.remove(child)
        source_ppr = source.find(W + "pPr")
        if source_ppr is not None:
            paragraph.append(copier.clone(source_ppr))
        else:
            raise ValueError("Reference formula requires paragraph formatting")
        paragraph.append(tab())
        for run in source_runs:
            paragraph.append(copier.clone(run))
        paragraph.append(tab())
        paragraph.append(run_text(entry["number"]))
        changed["display"] += 1
    for entry in mapping.get("inline", []):
        literal = entry["text"]
        if not literal:
            raise ValueError("Inline text cannot be empty")
        _, source_runs = copier.source_runs(entry["source_paragraph"])
        source = source_runs[entry["source_object"]]
        if len(xp(source, OLE)) != 1:
            raise ValueError("Inline source must contain exactly one equation")
        matches = []
        for paragraph in target.paragraphs():
            for chunk in plain_chunks(paragraph):
                combined = "".join(text_of(r) for r in chunk)
                spans = list(re.finditer(re.escape(literal), combined))
                if spans:
                    matches.append((paragraph, chunk, combined, spans))
        count = sum(len(item[3]) for item in matches)
        if count != entry["expected_count"]:
            raise ValueError(f"Inline {literal!r}: expected {entry['expected_count']}, found {count}")
        for paragraph, chunk, combined, spans in matches:
            positions, offset = [], 0
            for run in chunk:
                end = offset + len(text_of(run))
                positions.append((offset, end, run))
                offset = end
            def text_slice(start, end):
                return [run_text(combined[max(start, a):min(end, b)], run)
                        for a, b, run in positions if max(start, a) < min(end, b)]
            replacement, cursor = [], 0
            for span in spans:
                replacement.extend(text_slice(cursor, span.start()))
                replacement.append(copier.clone(source))
                cursor = span.end()
            replacement.extend(text_slice(cursor, len(combined)))
            insertion = paragraph.index(chunk[0])
            for run in chunk:
                paragraph.remove(run)
            for delta, run in enumerate(replacement):
                paragraph.insert(insertion + delta, run)
        changed["inline"] += count
    target.save(output_path)
    return {"output": str(output_path), "replaced": changed}


def audit(path):
    import olefile
    package = Package(path)
    results, paragraph_count = [], 0
    for pi, paragraph in enumerate(package.paragraphs()):
        objects = xp(paragraph, OLE)
        paragraph_count += bool(objects)
        for oi, obj in enumerate(objects):
            item = {"paragraph": pi, "object": oi, "errors": []}
            try:
                part = package.part_name(package.relationship(obj.get(RID)))
                item["part"] = part
                data = package.parts[part]
                with olefile.OleFileIO(io.BytesIO(data)) as storage:
                    streams = storage.listdir()
                    native = [s for s in streams if s[-1] in ("Equation Native", "MathType EF")]
                    item["native_streams"] = ["/".join(s) for s in native]
                    if not native:
                        item["errors"].append("native_stream_missing")
                    for stream in native:
                        payload = storage.openstream(stream).read()
                        if not payload or payload.lstrip().startswith((b"<math", b"<?xml")):
                            item["errors"].append("native_stream_empty_or_xml")
                shapes = xp(obj.getparent(), "./v:shape")
                shape = [s for s in shapes if s.get("id") == obj.get("ShapeID")]
                if len(shape) != 1:
                    item["errors"].append("shape_missing_or_ambiguous")
                else:
                    previews = xp(shape[0], ".//v:imagedata")
                    if not previews:
                        item["errors"].append("preview_missing")
                    for preview in previews:
                        part = package.part_name(package.relationship(preview.get(RID)))
                        if not package.parts[part]:
                            item["errors"].append("preview_empty")
            except Exception as error:
                item["errors"].append(str(error))
            results.append(item)
    return {"equation_paragraph_count": paragraph_count, "object_count": len(results),
            "error_object_count": sum(bool(i["errors"]) for i in results),
            "scope": "Native storage and relationships only; render and editor checks required.",
            "objects": results}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    for name in ("inspect", "audit"):
        command = commands.add_parser(name)
        command.add_argument("--docx", required=True, type=Path)
        command.add_argument("--json", type=Path)
    command = commands.add_parser("apply")
    for name in ("input", "reference", "manifest", "output"):
        command.add_argument("--" + name, required=True, type=Path)
    args = parser.parse_args()
    if args.command == "apply":
        result = apply_manifest(args.input, args.reference, args.manifest, args.output)
    else:
        result = inventory(args.docx) if args.command == "inspect" else audit(args.docx)
        if args.json:
            args.json.parent.mkdir(parents=True, exist_ok=True)
            args.json.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 2 if args.command == "audit" and result["error_object_count"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
