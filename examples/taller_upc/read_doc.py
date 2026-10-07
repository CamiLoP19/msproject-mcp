import glob
import os
import zipfile
import xml.etree.ElementTree as ET

downloads = os.path.expanduser("~/Downloads")
files = glob.glob(os.path.join(downloads, "*Planeaci*Recursos*.docx"))
if not files:
    print("No file found")
    exit(1)

target = files[0]
print(f"Reading: {target}")

with zipfile.ZipFile(target) as z:
    xml_content = z.read("word/document.xml")
    tree = ET.fromstring(xml_content)
    
    paragraphs = []
    # Iterate through all elements to preserve tables and paragraphs in order
    for elem in tree.iter():
        if elem.tag.endswith('}p'):
            texts = [node.text for node in elem.iter() if node.tag.endswith('}t') and node.text]
            if texts:
                paragraphs.append(''.join(texts))
        elif elem.tag.endswith('}tc'): # table cell
            pass

print("\n".join(paragraphs))
