import zipfile
import sys
import xml.etree.ElementTree as ET
from collections import Counter

"""
Usage:
  python tools/inspect_pptx_xml.py path/to/file.pptx

Outputs a per-slide summary listing counts of:
 - <p:sp> (shapes)
 - <p:pic> (pictures)
 - <a:tbl> (tables)
 - chart relationships (via slideRelationships referencing ../charts)
Also writes sample XML fragments for slides that contain non-shape elements.
"""

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'
}


def list_pptx_files(zipf):
    return [name for name in zipf.namelist() if name.startswith('ppt/')]


def find_slide_files(zipf):
    slide_files = sorted([n for n in zipf.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')])
    return slide_files


def parse_slide_xml(xml_bytes):
    try:
        root = ET.fromstring(xml_bytes)
        return root
    except ET.ParseError:
        return None


def count_elements(root):
    counts = Counter()
    if root is None:
        return counts
    # count shapes <p:sp>, pictures <p:pic>, tables <a:tbl>, graphicFrame (charts/tables) <p:graphicFrame>
    counts['p:sp'] = len(root.findall('.//{'+NS['p']+'}sp'))
    counts['p:pic'] = len(root.findall('.//{'+NS['p']+'}pic'))
    counts['a:tbl'] = len(root.findall('.//{'+NS['a']+'}tbl'))
    counts['p:graphicFrame'] = len(root.findall('.//{'+NS['p']+'}graphicFrame'))
    # any drawing chart elements
    counts['c:chart'] = len(root.findall('.//{'+NS['c']+'}chart'))
    return counts


def slide_has_non_shape(counts):
    # treat p:sp as shape; anything else >0 is non-shape
    non_shape_keys = [k for k in counts.keys() if k != 'p:sp' and counts[k] > 0]
    return non_shape_keys


def inspect(pptx_path):
    results = []
    with zipfile.ZipFile(pptx_path, 'r') as z:
        slide_files = find_slide_files(z)
        rel_files = [n for n in z.namelist() if n.startswith('ppt/slides/_rels/')]
        for slide in slide_files:
            xml_bytes = z.read(slide)
            root = parse_slide_xml(xml_bytes)
            counts = count_elements(root)
            non_shape = slide_has_non_shape(counts)
            # check slide relationships for charts or images referenced
            rel_path = 'ppt/slides/_rels/' + slide.split('/')[-1] + '.rels'
            rels = []
            if rel_path in z.namelist():
                rel_xml = ET.fromstring(z.read(rel_path))
                for rel in rel_xml.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                    rtype = rel.get('Type', '')
                    target = rel.get('Target', '')
                    rels.append((rtype, target))
            results.append({
                'slide': slide,
                'counts': dict(counts),
                'non_shape_keys': non_shape,
                'relationships': rels,
                'sample_fragment': None
            })
            if non_shape:
                # capture small fragment for evidence
                fragment = xml_bytes.decode('utf-8', errors='replace')[:2000]
                results[-1]['sample_fragment'] = fragment
    return results


def main():
    if len(sys.argv) < 2:
        print("Usage: python tools/inspect_pptx_xml.py path/to/file.pptx")
        sys.exit(2)
    pptx_path = sys.argv[1]
    res = inspect(pptx_path)
    total_non_shape_slides = 0
    for item in res:
        slide_name = item['slide']
        counts = item['counts']
        non_shape = item['non_shape_keys']
        print(f"Slide: {slide_name}")
        print("  counts:", counts)
        if non_shape:
            total_non_shape_slides += 1
            print("  NON-SHAPE elements found:", non_shape)
            print("  relationships:", item['relationships'])
            print("  sample fragment (truncated):")
            print(item['sample_fragment'].replace('\\n','\\n  '))
        else:
            print("  All detected elements are shapes (p:sp) or none.")
        print()
    print(f"Summary: {len(res)} slides inspected, {total_non_shape_slides} slides contain non-shape elements.")


if __name__ == '__main__':
    main()


