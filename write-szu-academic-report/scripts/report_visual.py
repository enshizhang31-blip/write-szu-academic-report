"""Conservative line-grid diagnostics. Every finding needs visual review."""
from pathlib import Path
import numpy as np

def _groups(values):
    groups=[]
    for v in values:
        if not groups or v>groups[-1][-1]+1:groups.append([int(v)])
        else:groups[-1].append(int(v))
    return groups

def line_risks(image,dpi=200):
    a=np.asarray(image.convert('L'))<110;h,w=a.shape
    # Find long solid segments without mistaking ordinary text rows for rules.
    horiz=[];vert=[]
    for y in np.flatnonzero(a.sum(axis=1)>w*.20):
        gs=_groups(np.flatnonzero(a[y]));long=[g for g in gs if len(g)>w*.15]
        if long:horiz.append((int(y),min(g[0] for g in long),max(g[-1] for g in long)))
    for x in np.flatnonzero(a.sum(axis=0)>h*.10):
        gs=_groups(np.flatnonzero(a[:,x]));long=[g for g in gs if len(g)>h*.08]
        if long:vert.append((int(x),min(g[0] for g in long),max(g[-1] for g in long)))
    rows=_groups([x[0] for x in horiz]);cols=_groups([x[0] for x in vert]);risks=[]
    rcent=[g[len(g)//2] for g in rows];ccent=[g[len(g)//2] for g in cols]
    for pos,groups,axis in [(rcent,rows,'horizontal'),(ccent,cols,'vertical')]:
        for u,v in zip(pos,pos[1:]):
            if dpi*.008< v-u <dpi*.055:
                bbox=[0,max(0,u-8),w,min(h,v+8)] if axis=='horizontal' else [max(0,u-8),0,min(w,v+8),h]
                risks.append({'rule':'possible_duplicate_border','bbox_px':bbox,'axis':axis})
    # A horizontal rule intersecting both vertical sides should be continuous.
    for y in rcent:
        sides=[x for x in ccent if a[max(0,y-2):min(h,y+3),max(0,x-2):min(w,x+3)].any()]
        for left,right in zip(sides,sides[1:]):
            if right-left<w*.12:continue
            line=a[max(0,y-1):min(h,y+2),left:right+1].any(axis=0)
            if line.mean()<.60:continue
            gaps=[g for g in _groups(np.flatnonzero(~line)) if len(g)>max(3,dpi*.015)]
            if gaps:risks.append({'rule':'possible_border_gap','bbox_px':[left,max(0,y-12),right+1,min(h,y+13)]})
            # Ink immediately beside a rule can indicate text touching/overlapping it.
            # Remove vertical crossings first; this remains a review hint.
            near=a[max(0,y-7):min(h,y+8),left:right+1].copy()
            near[max(0,7-3):min(near.shape[0],7+4),:]=False
            for x in sides:
                if left<=x<=right:near[:,max(0,x-left-4):min(near.shape[1],x-left+5)]=False
            if near.any(axis=0).sum()>(right-left)*.08:
                risks.append({'rule':'possible_table_content_overlap','bbox_px':[left,max(0,y-15),right+1,min(h,y+16)]})
    # Content very near the physical edge is a clipping hint, not proof.
    margin=max(2,int(dpi*.04))
    if a[:margin].any() or a[-margin:].any() or a[:,:margin].any() or a[:,-margin:].any():
        risks.append({'rule':'possible_page_edge_clipping','bbox_px':[0,0,w,h]})
    return risks[:12]

def inspect_pages(images,out,dpi):
    risks=[];out=Path(out)
    for page,image in enumerate(images,1):
        for i,r in enumerate(line_risks(image,dpi),1):
            name=f'risk-{page}-{i}.png';image.crop(tuple(r['bbox_px'])).save(out/name)
            from report_core import digest
            risks.append({**r,'page':page,'evidence_image':name,'evidence_sha256':digest(out/name),'review':'Inspect the actual table/figure region; line heuristics may also flag chart graphics.'})
    return risks
