"""Koppelt het taakvenster van de geregistreerde add-in aan een werkmap, zodat Excel het paneel opent
bij het openen van het bestand (zoals office-addin-dev-settings dat doet met storeType="Registry").

Gebruik: python -I dev/koppel-paneel.py dev/voorbeeld.xlsx
"""
import re
import shutil
import sys
import tempfile
import uuid
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ADDIN_ID = re.search(r"<Id>([^<]+)</Id>", (ROOT / "manifest.xml").read_text(encoding="utf-8")).group(1)

WEBEXTENSION = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{{{str(uuid.uuid4()).upper()}}}">
<we:reference id="{ADDIN_ID}" version="1.0.0.0" store="developer" storeType="Registry"/>
<we:alternateReferences/><we:properties/><we:bindings/>
<we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>"""

TASKPANES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
<wetp:taskpane dockstate="right" visibility="1" width="360" row="4">
<wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
</wetp:taskpane>
</wetp:taskpanes>"""

TASKPANES_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/>
</Relationships>"""

OVERRIDES = (
    '<Override PartName="/xl/webextensions/taskpanes.xml" '
    'ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>'
    '<Override PartName="/xl/webextensions/webextension1.xml" '
    'ContentType="application/vnd.ms-office.webextension+xml"/>'
)

ROOT_REL = (
    '<Relationship Id="rIdBroncellen" '
    'Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" '
    'Target="xl/webextensions/taskpanes.xml"/>'
)


def main(path: Path) -> None:
    with zipfile.ZipFile(path) as src:
        parts = {name: src.read(name) for name in src.namelist()}
    parts = {k: v for k, v in parts.items() if not k.startswith("xl/webextensions/")}

    types = parts["[Content_Types].xml"].decode("utf-8")
    types = re.sub(r'<Override PartName="/xl/webextensions/[^>]*/>', "", types)
    parts["[Content_Types].xml"] = types.replace("</Types>", OVERRIDES + "</Types>").encode("utf-8")

    rels = parts["_rels/.rels"].decode("utf-8")
    rels = re.sub(r'<Relationship [^>]*webextensiontaskpanes[^>]*/>', "", rels)
    parts["_rels/.rels"] = rels.replace("</Relationships>", ROOT_REL + "</Relationships>").encode("utf-8")

    parts["xl/webextensions/taskpanes.xml"] = TASKPANES.encode("utf-8")
    parts["xl/webextensions/_rels/taskpanes.xml.rels"] = TASKPANES_RELS.encode("utf-8")
    parts["xl/webextensions/webextension1.xml"] = WEBEXTENSION.encode("utf-8")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp_path = Path(tmp.name)
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as dst:
        for name, data in parts.items():
            dst.writestr(name, data)
    shutil.move(tmp_path, path)
    print(f"Taakvenster {ADDIN_ID} gekoppeld aan {path}")


if __name__ == "__main__":
    main(Path(sys.argv[1]))
