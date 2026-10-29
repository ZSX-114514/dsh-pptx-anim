#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Inject real PowerPoint entrance animations + slide transitions into a .pptx.

python-pptx has no animation API, so this writes the OOXML `<p:timing>` tree
directly. Only animEffect-based effects are used (fade / wipe / box / circle /
diamond / plus) because those encode as a single `<p:animEffect>` node and are
the most tolerant of hand construction; motion paths (fly-in) are avoided.

Usage:
    python pptx_anim.py <in.pptx> <out.pptx> [--preset auto|title|proposition|image|cards|data|none]
                        [--transition fade|none] [--audit]

Shape naming convention (recommended, produced by my deck builders):
    s01, s02, ... in reading order. Without names, shapes are animated in
    z-order, which for a generated deck is usually the same thing.
"""
import json
import os
import sys

from lxml import etree

from pptx import Presentation
from pptx.oxml.ns import nsdecls, qn

NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'

FILTERS = {
    'fade': 'fade',
    'wipe-up': 'wipe(up)',
    'wipe-right': 'wipe(right)',
    'wipe-down': 'wipe(down)',
    'box': 'box(in)',
    'circle': 'circle(in)',
    'diamond': 'diamond(in)',
    'plus': 'plus(in)',
}

# preset id used by PowerPoint for the matching entrance effect
PRESET_IDS = {
    'fade': 10,
    'wipe-up': 22,
    'wipe-right': 22,
    'wipe-down': 22,
    'box': 4,
    'circle': 5,
    'diamond': 8,
    'plus': 9,
}

PRESETS = {
    # name: (stagger_ms, duration_ms, effect for pictures, effect for text)
    'title': (180, 450, 'fade', 'fade'),
    'proposition': (150, 420, 'fade', 'fade'),
    'image': (200, 500, 'fade', 'wipe-up'),
    'cards': (110, 320, 'wipe-up', 'wipe-up'),
    'data': (90, 320, 'fade', 'fade'),
}


def _el(tag, **attrs):
    e = etree.SubElement if False else None
    node = etree.Element(qn(tag))
    for k, v in attrs.items():
        node.set(k, str(v))
    return node


def _sub(parent, tag, **attrs):
    node = etree.SubElement(parent, qn(tag))
    for k, v in attrs.items():
        node.set(k, str(v))
    return node


def build_timing(entries):
    """entries: [(shape_id, effect, delay_ms, duration_ms), ...] in play order."""
    if not entries:
        return None

    timing = _el('p:timing')
    tnLst = _sub(timing, 'p:tnLst')
    par0 = _sub(tnLst, 'p:par')
    ctn0 = _sub(par0, 'p:cTn', id=1, dur='indefinite', restart='never', nodeType='tmRoot')
    child0 = _sub(ctn0, 'p:childTnLst')
    seq = _sub(child0, 'p:seq', concurrent=1, nextAc='seek')
    ctn1 = _sub(seq, 'p:cTn', id=2, dur='indefinite', nodeType='mainSeq')
    child1 = _sub(ctn1, 'p:childTnLst')

    # one click group; everything plays on the first click, staggered by delay
    parA = _sub(child1, 'p:par')
    ctnA = _sub(parA, 'p:cTn', id=3, fill='hold')
    _sub(_sub(ctnA, 'p:stCondLst'), 'p:cond', delay='indefinite')
    childA = _sub(ctnA, 'p:childTnLst')

    parB = _sub(childA, 'p:par')
    ctnB = _sub(parB, 'p:cTn', id=4, fill='hold')
    _sub(_sub(ctnB, 'p:stCondLst'), 'p:cond', delay=0)
    childB = _sub(ctnB, 'p:childTnLst')

    next_id = 5
    for index, (spid, effect, delay_ms, dur_ms) in enumerate(entries):
        parE = _sub(childB, 'p:par')
        node_type = 'clickEffect' if index == 0 else 'withEffect'
        ctnE = _sub(
            parE, 'p:cTn',
            id=next_id, presetID=PRESET_IDS[effect], presetClass='entr',
            presetSubtype=0, fill='hold', grpId=0, nodeType=node_type,
        )
        next_id += 1
        _sub(_sub(ctnE, 'p:stCondLst'), 'p:cond', delay=max(0, int(delay_ms)))
        childE = _sub(ctnE, 'p:childTnLst')

        # make the shape visible at start
        set_el = _sub(childE, 'p:set')
        cbhvr = _sub(set_el, 'p:cBhvr')
        set_ctn = _sub(cbhvr, 'p:cTn', id=next_id, dur=1, fill='hold')
        next_id += 1
        _sub(_sub(set_ctn, 'p:stCondLst'), 'p:cond', delay=0)
        _sub(_sub(cbhvr, 'p:tgtEl'), 'p:spTgt', spid=spid)
        attr_lst = _sub(cbhvr, 'p:attrNameLst')
        name_el = etree.SubElement(attr_lst, qn('p:attrName'))
        name_el.text = 'style.visibility'
        to_el = _sub(set_el, 'p:to')
        val_el = etree.SubElement(to_el, qn('p:strVal'))
        val_el.set('val', 'visible')

        # entrance effect
        anim = _sub(childE, 'p:animEffect', transition='in', filter=FILTERS[effect])
        beh = _sub(anim, 'p:cBhvr')
        add = _sub(beh, 'p:cTn', id=next_id, dur=int(dur_ms))
        next_id += 1
        _sub(_sub(beh, 'p:tgtEl'), 'p:spTgt', spid=spid)

    # closing condition lists
    _sub(_sub(seq, 'p:prevCondLst'), 'p:cond', evt='onPrev', delay=0)
    _sub(_sub(seq, 'p:nextCondLst'), 'p:cond', evt='onNext', delay=0)
    return timing


def build_transition(kind='fade', speed='med'):
    trans = _el('p:transition', spd=speed)
    _sub(trans, 'p:' + kind)
    return trans


def shapes_in_order(slide):
    """Named s01..sNN first (reading order), then unnamed shapes in z-order."""
    named = []
    unnamed = []
    for shape in slide.shapes:
        name = (shape.name or '').strip()
        sid = getattr(shape, 'shape_id', None)
        if sid is None:
            continue
        if name.lower().startswith('s') and name[1:].isdigit():
            named.append((int(name[1:]), shape, sid))
        else:
            unnamed.append((shape, sid))
    named.sort(key=lambda t: t[0])
    out = [(sh, sid) for _, sh, sid in named]
    out.extend(unnamed)
    return out


def classify(shape):
    """Return 'picture' for images, 'text' otherwise."""
    try:
        if shape.shape_type is not None and int(shape.shape_type) == 13:
            return 'picture'
    except Exception:
        pass
    return 'text'


def has_visible_content(shape):
    try:
        if shape.has_text_frame and shape.text_frame.text.strip():
            return True
    except Exception:
        pass
    try:
        return int(shape.shape_type) == 13
    except Exception:
        return True


def detect_preset(slide):
    """Guess a preset from the slide's own composition."""
    pics = 0
    texts = 0
    shapes = [s for s in slide.shapes if has_visible_content(s)]
    for s in shapes:
        if classify(s) == 'picture':
            pics += 1
        else:
            texts += 1
    total_chars = sum(len(s.text_frame.text) for s in shapes if getattr(s, 'has_text_frame', False))
    if pics == 0 and 0 < texts <= 3:
        return 'proposition'
    if pics >= 2:
        return 'cards'
    if pics == 1:
        return 'image'
    if len(shapes) >= 6:
        return 'data'
    if texts >= 3:
        return 'cards'
    return 'data'


def animate(prs, preset_override=None, transition='fade'):
    report = []
    for index, slide in enumerate(prs.slides, 1):
        sld = slide._element
        for tag in ('p:timing', 'p:transition'):
            for old in sld.findall(qn(tag)):
                sld.remove(old)

        preset = preset_override if preset_override not in (None, 'auto') else detect_preset(slide)
        if preset == 'none':
            report.append((index, 'none', 0))
            continue
        stagger, dur, pic_effect, text_effect = PRESETS[preset]

        entries = []
        delay = 0
        for shape, sid in shapes_in_order(slide):
            if not has_visible_content(shape):
                continue
            effect = pic_effect if classify(shape) == 'picture' else text_effect
            entries.append((sid, effect, delay, dur))
            delay += stagger

        # CT_Slide order: ... clrMapOvr, transition, timing, extLst
        if transition and transition != 'none':
            sld.append(build_transition(transition))
        timing = build_timing(entries)
        if timing is not None:
            sld.append(timing)

        report.append((index, preset, len(entries)))
    return report


def audit(path):
    prs = Presentation(path)
    print(f'  audit: {path}')
    total = 0
    for i, slide in enumerate(prs.slides, 1):
        timing = slide._element.find(qn('p:timing'))
        trans = slide._element.find(qn('p:transition'))
        n = 0 if timing is None else len(timing.findall('.//' + qn('p:animEffect')))
        total += n
        flag = 'ok' if n or trans is not None else 'EMPTY'
        print(f'    slide {i:02d}: anim={n:2d} transition={"yes" if trans is not None else "no ":3s} {flag}')
    print(f'  total animEffect nodes: {total}')
    return total


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    flags = [a for a in sys.argv[1:] if a.startswith('--')]
    if len(args) < 2:
        print(__doc__)
        return 2
    src, dst = args[0], args[1]

    preset = 'auto'
    transition = 'fade'
    for f in flags:
        if f.startswith('--preset='):
            preset = f.split('=', 1)[1]
        elif f.startswith('--transition='):
            transition = f.split('=', 1)[1]

    prs = Presentation(src)
    report = animate(prs, preset_override=preset, transition=transition)
    prs.save(dst)

    for idx, name, n in report:
        print(f'  slide {idx:02d}: preset={name:10s} shapes={n}')
    print(f'  saved -> {dst}')
    if '--audit' in flags:
        audit(dst)
    return 0


if __name__ == '__main__':
    sys.exit(main())
