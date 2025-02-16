import requests
import lxml.html

html = requests.get('https://www.eon.dk/privat/strom-til-din-elbil/oversigt-ladepunkter-til-elbil.html')
doc = lxml.html.fromstring(html.content)

new_releases = doc.xpath('//div[@id="evseSection59409"]')[0]
titles = new_releases.xpath('.//div[@ui.grid"]/text()')
print(titles)

#//*[@id="evseSection59409"]/div/div/div[4]/div/span[2]