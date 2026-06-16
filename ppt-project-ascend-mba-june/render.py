#!/usr/bin/env python3
"""Render Ascend MBA 六月会议 deck from content.json."""
import sys, os, json

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, REPO)
from mck_ppt import MckEngine

PROJ = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(PROJ, 'content.json'), encoding='utf-8') as f:
    content = json.load(f)

slides = content['slides']
eng = MckEngine(total_slides=content['meta']['total_slides'])

for s in slides:
    L = s['layout']
    if L == 'cover':
        eng.cover(title=s['title'], subtitle=s.get('subtitle', ''),
                  author=s.get('author', ''), date=s.get('date', ''))
    elif L == 'agenda':
        headers = [tuple(h) for h in s['headers']]
        items = [tuple(i) for i in s['items']]
        eng.agenda(title=s['title'], headers=headers, items=items, source=s.get('source', ''))
    elif L == 'section_divider':
        eng.section_divider(section_label=s.get('section_label', ''),
                            title=s['title'], subtitle=s.get('subtitle', ''))
    elif L == 'executive_summary':
        items = [tuple(i) for i in s['items']]
        eng.executive_summary(title=s['title'], headline=s['headline'],
                              items=items, source=s.get('source', ''))
    elif L == 'before_after':
        eng.before_after(title=s['title'],
                         before_title=s['before_title'], before_points=s['before_points'],
                         after_title=s['after_title'], after_points=s['after_points'],
                         source=s.get('source', ''))
    elif L == 'meet_the_team':
        members = [tuple(m) for m in s['members']]
        eng.meet_the_team(title=s['title'], members=members, source=s.get('source', ''))
    elif L == 'table_insight':
        eng.table_insight(title=s['title'], headers=s['headers'], rows=s['rows'],
                          insights=s['insights'], source=s.get('source', ''))
    elif L == 'vertical_steps':
        steps = [tuple(st) for st in s['steps']]
        eng.vertical_steps(title=s['title'], steps=steps, source=s.get('source', ''))
    elif L == 'side_by_side':
        options = [(o[0], o[1]) for o in s['options']]
        eng.side_by_side(title=s['title'], options=options, source=s.get('source', ''))
    elif L == 'action_items':
        actions = [tuple(a) for a in s['actions']]
        eng.action_items(title=s['title'], actions=actions, source=s.get('source', ''))
    elif L == 'closing':
        eng.closing(title=s['title'], message=s.get('message', ''))
    else:
        raise ValueError(f"Unknown layout: {L}")

out = os.path.join(PROJ, 'output', 'ascend-mba-june-meeting.pptx')
eng.save(out)
print('SAVED', out)
