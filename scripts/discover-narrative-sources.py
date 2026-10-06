"""Read public rating indexes; cache evidence without importing unreviewed wording."""
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.parse import urljoin
from concurrent.futures import ThreadPoolExecutor
import html, json, re

root = Path(__file__).resolve().parents[1]
groups = ['administration', 'aviation', 'cryptology', 'construction',
          'special-operations', 'submarine', 'supply', 'surface-operations', 'surface-engineering']
cache = root / 'build/research/narrative-sources'
cache.mkdir(parents=True, exist_ok=True)

def read(name):
    url = 'https://www.navywriter.com/' + name + '.htm'
    content = urlopen(Request(url, headers={'User-Agent': 'Mozilla/5.0'}), timeout=30).read().decode('utf-8', errors='replace')
    (cache / (name + '.html')).write_text(content, encoding='utf-8')
    entries = []
    for href, label in re.findall(r'<a\b[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', content, re.I | re.S):
        label = html.unescape(re.sub('<[^>]+>', '', label)).strip()
        match = re.search(r'\(([A-Z]{2,3})\)', label)
        if match and href.endswith(('.htm', '.shtml')):
            entries.append({'rate': match[1], 'name': label, 'url': urljoin(url, href)})
    return entries

entries = []
with ThreadPoolExecutor(max_workers=5) as pool:
    for rows in pool.map(read, groups):
        entries.extend(rows)
entries.extend([
    {'rate':'IS', 'name':'Intelligence Specialist (IS)', 'url':'https://www.navywriter.com/intelligence.htm'},
    {'rate':'HM', 'name':'Hospital Corpsman (HM)', 'url':'https://www.navywriter.com/medical.htm'},
    {'rate':'MA', 'name':'Master-at-Arms (MA)', 'url':'https://www.navywriter.com/master-at-arms.htm'}])
entries = list({(r['rate'], r['url']): r for r in entries}.values())
entries.sort(key=lambda r: (r['rate'], r['name']))
(cache / 'rating-links.json').write_text(json.dumps(entries, indent=2), encoding='utf-8')
print(json.dumps(entries, indent=2))
