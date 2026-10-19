"""Bounded package, comment and numeric-reference checks; no Word/MathType conversion."""
from __future__ import annotations

from collections import Counter
import posixpath
import re
import unicodedata
from urllib.parse import unquote
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'w': W}


def qn(name):
    return '{' + W + '}' + name


def paragraph_nodes(node):
    """A text box's paragraphs are inventoried separately from their anchor paragraph."""
    yield node
    for child in node:
        if child.tag != qn('p'):
            yield from paragraph_nodes(child)


def text_of(node):
    result = []
    for child in paragraph_nodes(node):
        if child.tag == qn('t'):
            result.append(child.text or '')
        elif child.tag == qn('tab'):
            result.append('\t')
        elif child.tag in {qn('br'), qn('cr')}:
            result.append('\n')
    return ''.join(result)


def audit_package(members, add):
    roots = {}
    relationships = {}
    for name, payload in members.items():
        if name.endswith(('.xml', '.rels')):
            try:
                roots[name] = ET.fromstring(payload)
            except ET.ParseError as exc:
                add('blocker', 'malformed-xml', f'{name}: {exc}')
    for name, root in roots.items():
        if not name.endswith('.rels'):
            continue
        if name == '_rels/.rels':
            owner, directory = '', ''
        elif '/_rels/' in name:
            directory, leaf = name.rsplit('/_rels/', 1)
            owner = directory + '/' + leaf[:-5]
            if owner not in members:
                add('blocker', 'relationship-owner-missing', f'{name} has no owner part')
        else:
            add('blocker', 'relationship-path', f'unsupported relationship part: {name}')
            continue
        entries = {}
        for rel in root.findall('{' + PKG + '}Relationship'):
            rid, target = rel.get('Id'), rel.get('Target', '')
            if not rid or rid in entries:
                add('blocker', 'relationship-id', f'{name}: missing/duplicate relationship id {rid!r}')
            entries[rid] = target
            if rel.get('TargetMode', '').lower() == 'external':
                continue
            target_path = unquote(target.split('#', 1)[0])
            if not target_path:
                continue
            resolved = posixpath.normpath(target_path.lstrip('/') if target_path.startswith('/') else posixpath.join(directory, target_path))
            if resolved.startswith('../') or resolved not in members:
                add('blocker', 'relationship-target-missing', f'{name}: {rid} targets missing {resolved}')
        relationships[owner] = entries
    for name, root in roots.items():
        if name.endswith('.rels'):
            continue
        ids = relationships.get(name, {})
        for node in root.iter():
            for attribute in ('id', 'embed', 'link'):
                rid = node.get('{' + R + '}' + attribute)
                if rid and rid not in ids:
                    add('blocker', 'relationship-reference-missing', f'{name}: undeclared relationship {rid}')
    audit_comments(roots, add)
    return roots


def audit_comments(roots, add):
    definitions = []
    comments = roots.get('word/comments.xml')
    if comments is not None:
        definitions = [item.get(qn('id')) for item in comments.findall('w:comment', NS)]
        if None in definitions or len(set(definitions)) != len(definitions):
            add('blocker', 'comment-definition-id', 'missing or duplicate native comment id')
    defined = set(definitions)
    for name, root in roots.items():
        if not name.startswith('word/') or name.startswith('word/comments') or not name.endswith('.xml'):
            continue
        starts, ends, refs = {}, {}, Counter()
        for position, node in enumerate(root.iter()):
            identifier = node.get(qn('id'))
            if node.tag == qn('commentRangeStart'):
                if identifier in starts:
                    add('major', 'comment-anchor-duplicate', f'{name}: duplicate start {identifier}')
                starts[identifier] = position
            elif node.tag == qn('commentRangeEnd'):
                if identifier in ends:
                    add('major', 'comment-anchor-duplicate', f'{name}: duplicate end {identifier}')
                ends[identifier] = position
            elif node.tag == qn('commentReference'):
                refs[identifier] += 1
        for identifier in set(starts) | set(ends) | set(refs):
            if identifier is None or identifier not in defined:
                add('blocker', 'comment-text-missing', f'{name}: comment {identifier!r} has no definition')
            # Word also permits point comments: a reference without range markers.
            if (identifier in starts) != (identifier in ends):
                add('major', 'comment-anchor-unbalanced', f'{name}: unpaired range for comment {identifier}')
            elif identifier in starts and starts[identifier] >= ends[identifier]:
                add('major', 'comment-anchor-order', f'{name}: reversed range for comment {identifier}')
            if refs[identifier] != 1:
                add('major', 'comment-reference-count', f'{name}: comment {identifier} has {refs[identifier]} references')


def compare_protected_parts(base, current, config, add):
    changed = sorted(name for name in base.keys() & current.keys() if base[name] != current[name])
    added = sorted(current.keys() - base.keys())
    removed = sorted(base.keys() - current.keys())
    allowed = {'word/document.xml'} | set(config.get('allowed_changed_parts', []))
    allowed |= set(config.get('allowed_changed_media', [])) | set(config.get('allowed_changed_embeddings', []))
    for name in changed:
        if name not in allowed:
            add('blocker', 'protected-part-changed', f'part changed without an exact allowance: {name}')
    for name in added:
        if name not in set(config.get('allowed_added_parts', [])):
            add('blocker', 'unexpected-part-added', f'part added without an exact allowance: {name}')
    for name in removed:
        if name not in set(config.get('allowed_removed_parts', [])):
            add('blocker', 'protected-part-removed', f'part removed without an exact allowance: {name}')

    # Allowing comments.xml to grow must not waive protection for existing comments.
    def comments_by_id(parts):
        if 'word/comments.xml' not in parts:
            return {}
        root = ET.fromstring(parts['word/comments.xml'])
        result = {}
        for item in root.findall('w:comment', NS):
            item.tail = None
            result[item.get(qn('id'))] = ET.tostring(item)
        return result

    original_comments, new_comments = comments_by_id(base), comments_by_id(current)
    for identifier, payload in original_comments.items():
        if identifier not in new_comments:
            if identifier not in set(map(str, config.get('allowed_removed_comment_ids', []))):
                add('blocker', 'original-comment-removed', f'original comment {identifier} was removed')
        elif payload != new_comments[identifier] and identifier not in set(map(str, config.get('allowed_changed_comment_ids', []))):
            add('major', 'original-comment-changed', f'original comment {identifier} changed')
    def anchor_inventory(parts, tag, attribute):
        result = set()
        for name, payload in parts.items():
            if name.startswith('word/') and name.endswith('.xml') and not name.startswith('word/comments'):
                root = ET.fromstring(payload)
                result.update(node.get(qn(attribute)) for node in root.findall('.//w:' + tag, NS))
        return result

    removed_comment_anchors = anchor_inventory(base, 'commentReference', 'id') - anchor_inventory(current, 'commentReference', 'id')
    permitted_comment_removals = set(map(str, config.get('allowed_removed_comment_ids', [])))
    for identifier in removed_comment_anchors - permitted_comment_removals:
        add('blocker', 'original-comment-anchor-removed', f'original comment {identifier} lost its anchor')
    removed_bookmarks = anchor_inventory(base, 'bookmarkStart', 'name') - anchor_inventory(current, 'bookmarkStart', 'name')
    for name in removed_bookmarks - set(config.get('allowed_removed_bookmarks', [])):
        add('blocker', 'original-bookmark-removed', f'original bookmark {name!r} was removed')
    return {'changed_parts': changed, 'added_parts': added, 'removed_parts': removed}


NUMERIC_CITATION = re.compile(r'[\[［]\s*(\d+(?:\s*[-–—－,，、]\s*\d+)*)\s*[\]］]')
ENTRY = re.compile(r'^\s*[\[［](\d+)[\]］]\s*(.*)', re.S)


def citation_numbers(value):
    result = []
    for chunk in re.split(r'[,，、]', value):
        pair = re.split(r'[-–—－]', chunk.strip())
        if len(pair) == 1:
            result.append(int(pair[0]))
        else:
            left, right = map(int, pair)
            if not 0 < left <= right or right - left > 10000:
                raise ValueError('invalid or excessively long citation range')
            result.extend(range(left, right + 1))
    return result


def bibliography_inventory(document, styles_root):
    style_names = {}
    heading_styles = set()
    for style in styles_root.findall('w:style', NS):
        sid = style.get(qn('styleId'), '')
        name_node = style.find('w:name', NS)
        name = name_node.get(qn('val'), sid) if name_node is not None else sid
        style_names[sid] = name
        outline = style.find('w:pPr/w:outlineLvl', NS)
        if outline is not None and int(outline.get(qn('val'), '9')) < 9:
            heading_styles.add(sid)
        if '标题' in name or re.match(r'heading\s*[1-9]', name, re.I):
            heading_styles.add(sid)
    entries, ids = [], set()
    active = False
    for paragraph in document.findall('.//w:p', NS):
        value = text_of(paragraph).strip()
        if value.startswith('参考文献') or re.fullmatch(r'references\s*[:：]?', value, re.I):
            active = True
            ids.add(id(paragraph))
            continue
        if not active:
            continue
        match = ENTRY.match(value)
        if match:
            entries.append({'number': int(match.group(1)), 'text': match.group(2)})
            ids.add(id(paragraph))
            continue
        sid_node = paragraph.find('w:pPr/w:pStyle', NS)
        sid = sid_node.get(qn('val'), '') if sid_node is not None else ''
        if sid in heading_styles or re.match(r'^[一二三四五六七八九十]+[、．.]', value) or re.match(r'^\d+(?:\.\d+)*[.、\s]\s*[\u4e00-\u9fff]', value):
            active = False
            continue
        ids.add(id(paragraph))
        if entries and value:
            entries[-1]['text'] += ' ' + value
    return entries, ids


def citation_superscript(paragraph, match, styles_root):
    styles = {s.get(qn('styleId')): s for s in styles_root.findall('w:style', NS)}

    def style_vertical(sid, seen=None):
        seen = set() if seen is None else seen
        if sid not in styles or sid in seen:
            return None
        seen.add(sid)
        vertical = styles[sid].find('w:rPr/w:vertAlign', NS)
        if vertical is not None:
            return vertical.get(qn('val'))
        based = styles[sid].find('w:basedOn', NS)
        return style_vertical(based.get(qn('val')), seen) if based is not None else None

    parents = {child: parent for parent in paragraph.iter() for child in parent}
    pstyle = paragraph.find('w:pPr/w:pStyle', NS)
    cursor = 0
    for node in paragraph_nodes(paragraph):
        if node.tag != qn('t'):
            if node.tag in {qn('tab'), qn('br'), qn('cr')}:
                cursor += 1
            continue
        value = node.text or ''
        left, right = cursor, cursor + len(value)
        cursor = right
        if right <= match.start() or left >= match.end():
            continue
        run = parents.get(node)
        while run is not None and run.tag != qn('r'):
            run = parents.get(run)
        if run is None:
            return False
        direct = run.find('w:rPr/w:vertAlign', NS)
        rstyle = run.find('w:rPr/w:rStyle', NS)
        vertical = direct.get(qn('val')) if direct is not None else style_vertical(rstyle.get(qn('val'))) if rstyle is not None else None
        if vertical is None and pstyle is not None:
            vertical = style_vertical(pstyle.get(qn('val')))
        if vertical != 'superscript':
            return False
    return True


def reference_key(value):
    doi = re.search(r'10\.\d{4,9}/[^\s<>"\x27]+', value, re.I)
    doi_key = doi.group(0).rstrip('.,;，。；').lower() if doi else None
    type_marker = re.search(r'\[(?:J|M|D|R|C|S|EB/OL|J/OL|S/OL|R/OL)\]', value, re.I)
    title = ''
    if type_marker:
        prefix = value[:type_marker.start()]
        pieces = re.split(r'[.．。]\s+', prefix, maxsplit=1)
        if len(pieces) == 2:
            title = re.sub(r'[^\w]', '', unicodedata.normalize('NFKC', pieces[1]).casefold())
    year = re.search(r'\b(?:19|20)\d{2}\b', value)
    return doi_key, (title, year.group(0)) if title and year else None


def audit_citations(document, styles_root, add, config):
    entries, excluded = bibliography_inventory(document, styles_root)
    numbers = [entry['number'] for entry in entries]
    citation_groups = []
    first_use = []
    not_superscript = 0
    for paragraph in document.findall('.//w:p', NS):
        if id(paragraph) in excluded:
            continue
        for match in NUMERIC_CITATION.finditer(text_of(paragraph)):
            try:
                ids = citation_numbers(match.group(1))
            except ValueError as exc:
                add('major', 'citation-range', str(exc), match.group(0))
                continue
            citation_groups.append(ids)
            first_use.extend(number for number in ids if number not in first_use)
            if not citation_superscript(paragraph, match, styles_root):
                not_superscript += 1
                add('major', 'citation-superscript', 'numeric citation is not entirely superscript', match.group(0))
    if config.get('check_references', True):
        duplicates = sorted(number for number, count in Counter(numbers).items() if count > 1)
        if duplicates:
            add('major', 'bibliography-number-duplicate', f'duplicate bibliography numbers: {duplicates}')
        missing = sorted(set(first_use) - set(numbers))
        unused = sorted(set(numbers) - set(first_use))
        if missing:
            add('major', 'citation-entry-missing', f'citations without bibliography entries: {missing}')
        if unused:
            add('major', 'bibliography-unused', f'uncited bibliography entries: {unused}')
        if first_use != sorted(first_use):
            add('major', 'citation-first-use-order', f'first-use order is {first_use}')
        if numbers and numbers != list(range(1, len(numbers) + 1)):
            add('major', 'bibliography-sequence', 'bibliography numbers must occur once in ascending order from 1')
        doi_seen, title_seen = {}, {}
        for entry in entries:
            doi, title = reference_key(entry['text'])
            if doi and doi in doi_seen:
                add('major', 'bibliography-duplicate-doi', f'entries {doi_seen[doi]} and {entry["number"]} share DOI {doi}')
            elif doi:
                doi_seen[doi] = entry['number']
            if title and title in title_seen:
                add('major', 'bibliography-duplicate-title', f'entries {title_seen[title]} and {entry["number"]} share title/year')
            elif title:
                title_seen[title] = entry['number']
    return {
        'bibliography_paragraph_ids': excluded,
        'bibliography_entries': len(entries),
        'citation_groups_checked': len(citation_groups),
        'citation_groups_not_superscript': not_superscript,
        'citation_first_use': first_use,
        'reference_checks_enabled': config.get('check_references', True),
        'metadata_and_claim_verification': 'not_run',
    }
