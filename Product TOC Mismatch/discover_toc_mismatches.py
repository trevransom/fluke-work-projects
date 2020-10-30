# Would need to grab lists from the combine sheet
# would make a list of every unique combination of TOC name and type and then the values that go with it
# then could call a URL from the Start sheet that matches up with that TOC name
# then we'd check that value from the specific TOC name + type (kit, accessory, or product) to what's on the Website
# if it's a match then we return match, if not then we return the names of the values that don't exist in the site


# -*- coding: utf-8 -*-
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import numpy as np
from itertools import islice
import time
import re
import progressbar
import scrapy
from scrapy.crawler import CrawlerProcess
from scrapy.http.request import Request

# list_test = ["Fluke 80PK-22 SureGrip™ Immersion Temperature Probe","Fluke 80PK-24 SureGrip™ Air Temperature Probe","Fluke 80PK-25 SureGrip™ Piercing Temperature Probe","Fluke 80PK-26 SureGrip™ Tapered Temperature Probe","Fluke 80PK-27 SureGrip™ Industrial Surface Temperature Probe","Fluke 80PK-3A Surface Probe"]

# list_test = [w.replace("TM", "").replace("™", "").replace("<sup>&trade;</sup>", "").replace("&trade;", "").replace("&#x02122;", "").replace("&#8482;", "").replace("&TRADE;", "").replace("®", "").replace("&reg;", "").replace("&#x000AE;", "").replace("&#174;", "").replace("&circledR;", "").replace("&REG;", "").replace("<sup>&reg;</sup>", "").replace("&", "").replace("&amp;", "").replace("&#x00026;", "").replace("&#38;", "").replace("&AMP;", "").replace("&nbsp;", "").replace("&#x000A0;", "").replace("&#160;", "").replace("&NonBreakingSpace;", "").replace("»", "").replace("&raquo;", "").replace("&#x000BB;", "").replace("&#187;", "").replace("tm", "") for w in list_test]

# for x in list_test:
#     print (x)
# sys.exit(0)


title = 'Product TOC ECM'

# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

# Find a workbook by name
sheet = client.open("Product TOC ECM")

# Test reporting sheet
values_to_check_sheet = sheet.worksheet("Validation")
urls_sheet = sheet.worksheet("START")

values = values_to_check_sheet.get_all_values()
np_values = np.array(values)

cell_list = values_to_check_sheet.range(1, 5, len(values), 5)

urls = [url for title, url in zip(urls_sheet.col_values(2), urls_sheet.col_values(4)) if 'https' in url and title in np_values[:,0]]
# urls = urls[0]
# print(urls)
key = [(x[0]+x[1]+x[3]) for x in np_values]

pbar = progressbar.ProgressBar(max_value=len(urls))

BASE_URL = 'https://live-fluke-ecm.pantheonsite.io/en-us'
USER_NAME = 'transom'
PASSWORD = '8*U4x%FUA4FF'
PAGES = [*urls]

pbar.start()
count = 0

class TOCMatch(scrapy.Spider):
    name = "TOCMatch"
    start_urls = [BASE_URL]

    def parse(self, response):
        yield scrapy.FormRequest.from_response(
            response,
            formxpath='//*[@id="user-login-form"]',
            formdata={
                'name': USER_NAME,             
                'pass': PASSWORD,             
                'Action':'1',
            },
            callback=self.after_login)

    def after_login(self, response):

        # if "Home" in str(response.body):
        #     print("Login WORKED")
        # else:
        #     print("Login DID NOT WORK")

        for page in PAGES:
            yield Request(
                url=page,
                callback=self.action,
                errback=self.errback_httpbin)

    def action(self, response):
        global count
        count += 1
        pbar.update(count)

        print(f'\nURL: {response.url}')


# okay sooo, could literally go item by item
# okay, but we're working with scrapy. and they need a list of urls up front
# so, maybe let's give them that, and we'll say, for each url - we NEED to know what TOC it belongs to
# we can run a search on the page to find that
# Then once we've disovered the TOC name we can grab all values and types with that name in the Validation sheet
# then we run a loop over those values and if type is... we ccheck those valeus againt what's on the page
# if it can't find the value then we report onto that row + column No match


        url = re.search("productid=(.*)&", response.url)

        title = response.css('#title-field-add-more-wrapper input::attr(value)').extract_first()
        title = re.search("TOC - (.*) for Products", title)
        title = title.group(1).lower()
        
        current_values = [row for row in values if row[0].lower() in title]
        print("Title from Spreadsheet:",current_values[0][0])
        print("Title from ECM:",title,'\n\n')
        
        key_title = current_values[0][0]

        # the problem with looping through all the current values is that there's theoretically the possibility
        # that a value/product could exist twice in current values

        # could we loop through key though? 

        # 421 - NO's

        # for value in current_values:
        #     if key_title+"Products"+value

        products = [row[3] for row in current_values if row[1] == 'Products']
        accessories = [row[3] for row in current_values if row[1] == 'Accessories']
        kits = [row[3] for row in current_values if row[1] == 'Kits']
        top_sellers = [row[3] for row in current_values if row[1] == 'Top Sellers']

        ecm_products = response.css('#field-product-list-sort-solr-values td div input:first-child::attr(value)').extract()
        ecm_accessories = response.css('#field-accessories-list-sort-solr-values td div input:first-child::attr(value)').extract()
        ecm_kits = response.css('#field-kits-list-sort-solr-values td div input:first-child::attr(value)').extract()
        ecm_top_sellers = response.css('#field-top-sellers-solr-values td div input:first-child::attr(value)').extract()
        
        ecm_products = [item.replace("TM", "").replace("™", "").replace("<sup>&trade;</sup>", "").replace("&trade;", "").replace("&#x02122;", "").replace("&#8482;", "").replace("&TRADE;", "").replace("®", "").replace("&reg;", "").replace("&#x000AE;", "").replace("&#174;", "").replace("&circledR;", "").replace("&REG;", "").replace("<sup>&reg;</sup>", "").replace("&", "").replace("&amp;", "").replace("&#x00026;", "").replace("&#38;", "").replace("&AMP;", "").replace("&nbsp;", "").replace("&#x000A0;", "").replace("&#160;", "").replace("&NonBreakingSpace;", "").replace("»", "").replace("&raquo;", "").replace("&#x000BB;", "").replace("&#187;", "").replace("tm", "") for item in ecm_products]
        ecm_accessories = [item.replace("TM", "").replace("™", "").replace("<sup>&trade;</sup>", "").replace("&trade;", "").replace("&#x02122;", "").replace("&#8482;", "").replace("&TRADE;", "").replace("®", "").replace("&reg;", "").replace("&#x000AE;", "").replace("&#174;", "").replace("&circledR;", "").replace("&REG;", "").replace("<sup>&reg;</sup>", "").replace("&", "").replace("&amp;", "").replace("&#x00026;", "").replace("&#38;", "").replace("&AMP;", "").replace("&nbsp;", "").replace("&#x000A0;", "").replace("&#160;", "").replace("&NonBreakingSpace;", "").replace("»", "").replace("&raquo;", "").replace("&#x000BB;", "").replace("&#187;", "").replace("tm", "") for item in ecm_accessories]
        ecm_kits = [item.replace("TM", "").replace("™", "").replace("<sup>&trade;</sup>", "").replace("&trade;", "").replace("&#x02122;", "").replace("&#8482;", "").replace("&TRADE;", "").replace("®", "").replace("&reg;", "").replace("&#x000AE;", "").replace("&#174;", "").replace("&circledR;", "").replace("&REG;", "").replace("<sup>&reg;</sup>", "").replace("&", "").replace("&amp;", "").replace("&#x00026;", "").replace("&#38;", "").replace("&AMP;", "").replace("&nbsp;", "").replace("&#x000A0;", "").replace("&#160;", "").replace("&NonBreakingSpace;", "").replace("»", "").replace("&raquo;", "").replace("&#x000BB;", "").replace("&#187;", "").replace("tm", "") for item in ecm_kits]
        ecm_top_sellers = [item.replace("TM", "").replace("™", "").replace("<sup>&trade;</sup>", "").replace("&trade;", "").replace("&#x02122;", "").replace("&#8482;", "").replace("&TRADE;", "").replace("®", "").replace("&reg;", "").replace("&#x000AE;", "").replace("&#174;", "").replace("&circledR;", "").replace("&REG;", "").replace("<sup>&reg;</sup>", "").replace("&", "").replace("&amp;", "").replace("&#x00026;", "").replace("&#38;", "").replace("&AMP;", "").replace("&nbsp;", "").replace("&#x000A0;", "").replace("&#160;", "").replace("&NonBreakingSpace;", "").replace("»", "").replace("&raquo;", "").replace("&#x000BB;", "").replace("&#187;", "").replace("tm", "") for item in ecm_top_sellers]

        for x, item in enumerate(ecm_products):
            if item:

                # Stripping out the ending (nid) tag
                item = re.search("(.*)\((\d*)", item)
                ecm_products[x] = item.group(1).lower()

        for x, item in enumerate(ecm_accessories):
            if item:
                item = re.search("(.*)\((\d*)", item)
                ecm_accessories[x] = item.group(1).lower()

        for x, item in enumerate(ecm_kits):
            if item:
                item = re.search("(.*)\((\d*)", item)
                ecm_kits[x] = item.group(1).lower()

        for x, item in enumerate(ecm_top_sellers):
            if item:
                item = re.search("(.*)\((\d*)", item)
                ecm_top_sellers[x] = item.group(1).lower()
        
        for product in products:
            if product in ecm_products:
                row = key.index(key_title+"Products"+product)
                cell_list[row].value = "Yes"
            else:
                row = key.index(key_title+"Products"+product)
                cell_list[row].value = "No"

        for accessory in accessories:
            if accessory in ecm_accessories:
                row = key.index(key_title+"Accessories"+accessory)
                cell_list[row].value = "Yes"
            else:
                row = key.index(key_title+"Accessories"+accessory)
                cell_list[row].value = "No"
        
        for kit in kits:
            if kit in ecm_kits:
                row = key.index(key_title+"Kits"+kit)
                cell_list[row].value = "Yes"
            else:
                row = key.index(key_title+"Kits"+kit)
                cell_list[row].value = "No"

        for seller in top_sellers:
            if seller in ecm_top_sellers:
                row = key.index(key_title+"Top Sellers"+seller)
                cell_list[row].value = "Yes"
            else:
                row = key.index(key_title+"Top Sellers"+seller)
                cell_list[row].value = "No"

        values_to_check_sheet.update_cells(cell_list)
        time.sleep(0.1)

    def errback_httpbin(self, failure):
        print("\n\nERROR***********")
        print('URL:',failure.value.response.url, '\n\n')

process = CrawlerProcess({
    'USER_AGENT': 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1)',
    'LOG_LEVEL': 'INFO',
})

process.crawl(TOCMatch)
process.start() # the script will block here until the crawling is finished

pbar.finish()