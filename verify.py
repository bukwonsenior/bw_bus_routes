# -*- coding: utf-8 -*-
import json, math
d=json.load(open('routes.json',encoding='utf-8'))
def hav(a,b,c,e):
    R=6371.0; p=math.radians
    dlat=p(c-a); dlon=p(e-b)
    x=math.sin(dlat/2)**2+math.cos(p(a))*math.cos(p(c))*math.sin(dlon/2)**2
    return 2*R*math.asin(math.sqrt(x))
def mins(t):
    h,m=t.split(':'); return int(h)*60+int(m)

print('== 좌표 범위 검증 (원주시: 37.20~37.55N, 127.75~128.15E) ==')
bad=[s['name'] for s in d['stops'] if not(37.20<=s['lat']<=37.55 and 127.75<=s['lng']<=128.15)]
print('범위 이탈:', bad if bad else '없음')

for r in d['routes']:
    print('\n===',r['name'],'| 정류장',len(r['stops']),'개 ===')
    tot=0; alerts=[]
    for i in range(len(r['stops'])-1):
        a,b=r['stops'][i],r['stops'][i+1]
        km=hav(a['lat'],a['lng'],b['lat'],b['lng']); tot+=km
        dt=mins(b['time'])-mins(a['time'])
        sp=km/(dt/60) if dt>0 else None
        flag=''
        if dt<0: flag='시간역행!'
        elif dt==0 and km>0.6: flag='동일시간+거리過'
        elif sp and sp>75: flag='속도 %.0fkm/h 과다'%sp
        elif sp and sp<8 and km>0.3: flag='속도 %.0fkm/h 과소'%sp
        if flag: alerts.append('  %s(%s) -> %s(%s): %.2fkm/%d분 %s'%(a['name'],a['time'],b['name'],b['time'],km,dt,flag))
    print('총 주행거리 %.1f km / 총 소요 %d분 / 평균 %.0f km/h'%(tot,mins(r['stops'][-1]['time'])-mins(r['stops'][0]['time']),tot/((mins(r['stops'][-1]['time'])-mins(r['stops'][0]['time']))/60)))
    print('순번 연속성:', '정상' if [s['seq'] for s in r['stops']]==list(range(1,len(r['stops'])+1)) else '오류')
    if alerts:
        print('구간 이상 %d건:'%len(alerts)); [print(a) for a in alerts]
    else: print('구간 이상: 없음')
