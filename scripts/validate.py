#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Validate a .pptx that had animations injected.

Checks, in order of how likely they are to make PowerPoint show "需要修复":
  1. package reopens with python-pptx
  2. every slide: CT_Slide child order is ... transition, timing, extLst
  3. no empty <p:stCondLst>
  4. animEffect / transition node counts per slide
  5. optional: LibreOffice render (run separately; see SKILL.md)

Usage: python validate.py <deck.pptx>
Exit code 1 if any hard check fails.
"""
import glob
import sys
import zipfile

from lxml import etree

P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

ORDER = ['cSld', 'clrMapOvr', 'transition', 'timing', 'extLst']


def main(path):
    problems = []

    # 1. reopen
    try:
        from pptx import Presentation
        prs = Presentation(path)
        n_slides = len(prs.slides)
        print(f'  reopen ok: {n_slides} slides')
    except Exception as exc:
        print(f'  REOPEN FAILED: {exc}')
        return 1

    # 2..4 walk the package directly
    with zipfile.ZipFile(path) as zf:
        slides = sorted(n for n in zf.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml'))
        total_anim = 0
        for name in slides:
            root = etree.fromstring(zf.read(name))
            tags = [etree.QName(c).localname for c in root]
            idx = [tags.index(t) for t in ('transition', 'timing') if t in tags]
            if len(idx) == 2 and idx[0] > idx[1]:
                problems.append(f'{name}: transition after timing')
            # other unknown children that break ORDER
            known = [t for t in tags if t in ORDER]
            if known != [t for t in ORDER if t in tags]:
                problems.append(f'{name}: slide children out of order -> {tags}')
            for lst in root.iter(P + 'stCondLst'):
                if len(lst) == 0:
                    problems.append(f'{name}: empty stCondLst')
            n_anim = len(root.findall('.//' + P + 'animEffect'))
            total_anim += n_anim
            has_trans = root.find(P + 'transition') is not None
            print(f'    {name.split("/")[-1]:14s} anim={n_anim:2d} transition={"yes" if has_trans else "no"}')
        print(f'  total animEffect: {total_anim}')

    # 5. every spTgt must point at an existing shape id
    with zipfile.ZipFile(path) as zf:
        for name in slides:
            root = etree.fromstring(zf.read(name))
            ids = {e.get('id') for e in root.iter(P + 'cNvPr')}
            for tgt in root.iter(P + 'spTgt'):
                spid = tgt.get('spid')
                if spid not in ids:
                    problems.append(f'{name}: spTgt spid={spid} has no matching shape')

    if problems:
        print('\n  PROBLEMS:')
        for p in problems[:20]:
            print('   -', p)
        return 1
    print('\n  all structural checks passed')
    return 0


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(sys.argv[1]))
