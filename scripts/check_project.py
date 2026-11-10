#!/usr/bin/env python3
"""Static validation; does not replace testing in actual Word Desktop."""
import json
from pathlib import Path
from html.parser import HTMLParser
from xml.etree import ElementTree
root=Path(__file__).resolve().parent.parent
catalog=json.loads((root/'web/catalog.json').read_text())
assert catalog['schemaVersion']==1
ids=set()
for doc in catalog['documents']:
    assert doc['id'] not in ids
    ids.add(doc['id'])
    assert doc['illustrations']
    for item in doc['illustrations']:
        path=item['path']
        assert path.startswith('/illustrations/') and path.endswith('.html') and '..' not in path
        target=root/'web'/path.lstrip('/')
        assert target.is_file(), target
        print('OK illustration:', target.relative_to(root))
manifest=root/'deploy/manifest.xml'
if manifest.exists():
    tree=ElementTree.parse(manifest)
    assert tree.getroot().tag.endswith('OfficeApp')
    print('OK manifest XML')
print('OK catalog: %d BRD(s)'%len(ids))
