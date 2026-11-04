import urllib.request,re,json,time,html,sys
urls=open('colon.txt').read().split()
out=[]
SKIP={'Organization','WebSite','SearchAction','Country','PostalAddress','Place','GeoCoordinates','BreadcrumbList','ListItem','FAQPage','Question','Answer','Offer','OfferCatalog','Service','AggregateRating','Rating','Review','Person','OpeningHoursSpecification','ImageObject','WebPage','ItemList'}
def find(o):
    if isinstance(o,dict):
        t=o.get('@type'); t=t[0] if isinstance(t,list) else t
        if t and t not in SKIP and o.get('name'): return o
        for v in o.values():
            r=find(v)
            if r: return r
    elif isinstance(o,list):
        for v in o:
            r=find(v)
            if r: return r
for i,u in enumerate(urls):
    try:
        s=urllib.request.urlopen(urllib.request.Request(u,headers={'User-Agent':'Mozilla/5.0 (prospectos D.E.K)'}),timeout=40).read().decode('utf-8','ignore')
    except Exception as e:
        out.append({'url':u,'err':str(e)}); continue
    title=html.unescape((re.findall(r'<title>(.*?)</title>',s) or [''])[0])
    biz=None
    for blk in re.findall(r'<script type="application/ld\+json"[^>]*>(.*?)</script>',s,re.S):
        try: biz=find(json.loads(blk)) or biz
        except Exception: pass
        if biz: break
    was=sorted(set(re.findall(r'wa\.me/(\d{8,13})',s)))
    ext=sorted({html.unescape(h) for h in re.findall(r'href="(https?://[^"]+)"',s) if not re.search(r'micita|waze|wa\.me|whatsapp|google|apple\.com|play\.google',h)})
    adr=(biz or {}).get('address') or {}
    out.append({'url':u,'title':title,'name':(biz or {}).get('name') or title.split(' - ')[0],'type':(biz or {}).get('@type'),
                'cat':title.split(' - ')[1].replace(' en Colón Panamá','').split(' en ')[0] if ' - ' in title else '',
                'tel':(biz or {}).get('telephone'),'wa':was,'web':(biz or {}).get('url') if (biz or {}).get('url') and 'micita' not in str((biz or {}).get('url')) else None,
                'sameAs':(biz or {}).get('sameAs'),'ext':ext,'street':adr.get('streetAddress') if isinstance(adr,dict) else str(adr),
                'area':adr.get('addressArea') if isinstance(adr,dict) else '', 'noindex':'noindex' in s[:5000]})
    if i%25==0: print(i,len(urls),flush=True)
    time.sleep(0.8)
json.dump(out,open('perfiles.json','w',encoding='utf-8'),ensure_ascii=False,indent=1)
print('listo',len(out))
