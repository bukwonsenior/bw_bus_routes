# -*- coding: utf-8 -*-
# 좌표 검증용 미리보기 (배경지도 없음 / 노선 형태·순서 확인 목적)
import json, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm

fm.fontManager.addfont('/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc')
fp=fm.FontProperties(fname='/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc')
d=json.load(open('routes.json',encoding='utf-8'))
conf={s['id']:s['conf'] for s in d['stops']}

fig,axes=plt.subplots(1,3,figsize=(20,7.2))
for ax,r in zip(axes,d['routes']):
    xs=[s['lng'] for s in r['stops']]; ys=[s['lat'] for s in r['stops']]
    ax.plot(xs,ys,'-',color=r['color'],lw=1.8,alpha=.75,zorder=1)
    for i,s in enumerate(r['stops']):
        est = conf[s['id']]=='estimated'
        ax.scatter(s['lng'],s['lat'],s=210,color='#c0392b' if est else r['color'],
                   edgecolor='w',lw=1.4,zorder=3,marker='s' if est else 'o')
        ax.text(s['lng'],s['lat'],str(s['seq']),color='w',fontsize=7.2,fontweight='bold',
                ha='center',va='center',zorder=4)
        ax.annotate(s['name'],(s['lng'],s['lat']),textcoords='offset points',xytext=(8,5),
                    fontsize=6.6,fontproperties=fp,color='#c0392b' if est else '#333',zorder=5)
    ax.set_title('%s  (%s)  %s~%s'%(r['name'],r['sub'],r['stops'][0]['time'],r['stops'][-1]['time']),
                 fontproperties=fp,fontsize=13,fontweight='bold',pad=11)
    ax.set_aspect(1/0.79); ax.grid(alpha=.22,ls=':'); ax.tick_params(labelsize=7)
fig.suptitle('노선 좌표 검증 미리보기  —  전 정류장 BIS 실측 좌표 적용',
             fontproperties=fp,fontsize=14,fontweight='bold',y=.99)
plt.tight_layout(rect=[0,0,1,.96])
plt.savefig('좌표검증_미리보기.png',dpi=125,bbox_inches='tight',facecolor='w')
print('saved')
