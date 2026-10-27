from pathlib import Path
from urllib.request import Request, urlopen
from concurrent.futures import ThreadPoolExecutor
from PIL import Image
import json, io

out=Path('.pptx-build');out.mkdir(exist_ok=True)
photos=[
('01','https://electrik.info/uploads/posts/2020-05/1588498603_1.jpg','https://electrik.info/main/master/1671-iskrenie-kontaktov-rele-i-puskateley.html'),
('02','https://om-ek.ru/image/cache/catalog/Proizoditely/0002/kktr-800x600.png','https://om-ek.ru/index.php?oct_blog_article_id=79&route=octemplates%2Fblog_article'),
('03','https://prokazan.ru/userfiles/article/Ttp3xXwVNeutkCEZTuKoqWeLcfyf4ODK.jpeg','https://prokazan.ru/interesnoe/view/sekonomili-na-avtomatah-poterali-kvartiru-kakie-osibki-v-novoj-provodke-case-vsego-privodat-k-pozaram'),
('04','https://forum.autoua.net/files/10739370-unnamed_%281%29.jpg','https://forum.autoua.net/showflat.php?Cat=0&Number=10739370&fpart=all&page=..&vc=1'),
('05','https://www.fujielectric.co.jp/fcs/the_switch/box/img/vol16/img01.jpg','https://www.fujielectric.co.jp/fcs/the_switch/vol16/'),
('06','https://vmn-amt.imgix.net/uploads/2018/11/Pico03.png?auto=compress%2Cformat&q=50&w=600','https://www.amt.nl/60939/diagnose-aan-een-epb-module'),
('07','https://picture1.goo-net.com/shop/012/0127634/M/0127634_20190823193958_704_P_Pit_Blog_00.jpg','https://www.goo-net.com/pit/shop/0127634/blog/192352'),
('08','https://d1szwdapq2yuw7.cloudfront.net/photos/business/medium/contactor-de-potencia-quemado_252104.jpg','https://fotos.habitissimo.es/foto/contactor-de-potencia-quemado_252104'),
('09','https://www.elektro.info.pl/media/cache/typical_view/data/202301/zagrozenia-pozarowe-stwarzane-przez-uklady-kompensacji-mocy-biernej-rys6.jpg','https://www.elektro.info.pl/artykul/jakosc-energii-elektrycznej/188083%2Czagrozenia-pozarowe-stwarzane-przez-uklady-kompensacji-mocy-biernej-czesc-1'),
('10','https://www.kreativekiste.de/images/schuetzschaltungen/aufbau_und_funktion/Der%20Schuetzt%20Aufbau%20und%20Funktion%20erklaert%2014.jpg','https://www.kreativekiste.de/component/content/article?id=92%3Aschuetze-aufbau-und-funktion-einfach-erklaert')]
def get(item):
 n,url,source=item
 try:
  with urlopen(Request(url,headers={'User-Agent':'Mozilla/5.0'}),timeout=40) as r:raw=r.read()
  im=Image.open(io.BytesIO(raw));im.load();ext='.png' if im.format=='PNG' else '.jpg';path=out/(n+ext)
  if im.format not in ('PNG','JPEG'):im.convert('RGB').save(path,quality=95)
  else:path.write_bytes(raw)
  return dict(id=n,path=str(path.resolve()),source=source,url=url,width=im.width,height=im.height)
 except Exception as e:return dict(id=n,error=str(e),source=source,url=url)
with ThreadPoolExecutor(max_workers=4) as ex:results=list(ex.map(get,photos))
(out/'photos.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps(results,ensure_ascii=True,indent=2))
