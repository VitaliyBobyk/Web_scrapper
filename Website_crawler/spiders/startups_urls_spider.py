# -*- coding: utf-8 -*-
import csv
import re
import scrapy
import json
import pandas as pd
import xlsxwriter
import glob
import os

class StartupsUrlsSpider(scrapy.Spider):
    name = 'urls'
    start_urls = ['https://e27.co/api/startups/?tab_name=recentlyupdated&start=0&length=250']

    def parse(self, response):
        try:
            items = json.loads(response.body)
            list_data = []
            file = open('Website_crawler/Result_urls.csv', 'w', newline='')
            for item in range(250):
                title = items['data']['list'][item]['slug']
                url = items['data']['list'][item]['metas']['website']
                company = []
                company.append(title)
                company.append(url)
                list_data.append(company)

            with file:
                header = ['Slug', 'Url']
                writer = csv.DictWriter(file, fieldnames=header)
                writer.writeheader()
                for i in list_data:
                    writer.writerow({'Slug': i[0],
                                     'Url': i[1]
                                     })
        except:
            print("ERORR")


class StartupsDetailSpider(scrapy.Spider):
    name = 'detail'
    try:
        with open('Website_crawler/Result_data.csv', 'w') as file:
            headers = header = ['company_name', 'profile_url', 'company_website', 'location', 'tags', 'founding_date',
                          'urls', 'emails', 'phones', 'description_short', 'description']
            writer = csv.DictWriter(file, fieldnames=header)
            writer.writeheader()
        file = open('Website_crawler/Result_urls.csv', 'r')
        urls_list = [i.split(',') for i in file]
        urls = [i[0] for i in urls_list]
        start_urls = [

        ]
        for i in urls:
            if i == 'Slug' or i == 'Url':
                continue
            else:
                start_urls.append(f'https://e27.co/api/startups/get/?slug={i}&data_type=general&get_badge=true')
    except:
        print('ERORR')
    def parse(self, response):
        try:
            items = json.loads(response.body)
            company_name = str(items['data']['name'] if 'name' in items['data'].keys() else '')
            profile_url = str(f"https://e27.co/startups/{items['data']['slug']}/" if 'slug' in items['data'].keys() else '')
            company_website = str(items['data']['metas']['website'] if 'website' in items['data']['metas'].keys() else '')
            location = str(items['data']['location'][0]['text'] if 'text' in items['data']['location'][0].keys() else '')
            tags = str(re.sub(r'[\[\]\"]', '', str(items['data']['market'] if 'market' in items['data'].keys() else '')))
            founded_year = items['data']['metas']['found_year'] if 'found_year' in items['data']['metas'].keys() else ''
            founded_mount = items['data']['metas']['found_month'] if 'found_month' in items['data'][
                'metas'].keys() else ''
            founding_date = str(f'{founded_year}-{founded_mount}' if founded_year and founded_mount != '' else '')
            url_linked = str(items['data']['metas']['linkedin']+',' if items['data']['metas']['linkedin'] != '' else '')
            url_twitter = str(items['data']['metas']['twitter']+',' if items['data']['metas']['twitter'] != '' else '')
            url_facebook = str(items['data']['metas']['facebook']+',' if items['data']['metas']['facebook'] != '' else '')
            url = str(f"{url_linked} {url_twitter} {url_facebook}")
            emails = str(items['data']['metas']['email'] if 'email' in items['data']['metas'].keys() else '')
            phones = str(items['data']['metas']['phone'] if 'phone' in items['data']['metas'].keys() else '')
            description_short = str(items['data']['metas']['short_description'] if 'short_description' in items['data'][
                'metas'].keys() else '')
            description = str(items['data']['metas']['description'] if 'description' in items['data'][
                'metas'].keys() else '')
            with open('Website_crawler/Result_data.csv', 'a') as file:
                writer = csv.writer(file)
                data = [company_name, profile_url, company_website, location, tags, founding_date,
                        url, emails, phones, description_short, description]
                writer.writerow(data)

            writer = pd.ExcelWriter('multi_sheet.xlsx', engine='xlsxwriter')
            folders = next(os.walk('.'))[1]

            for host in folders:
                Path = os.path.join(os.getcwd(), host)

                for f in glob.glob(os.path.join(Path, "*.csv")):
                    df = pd.read_csv(f)
                    df.to_excel(writer, sheet_name=os.path.basename(f)[:31])

            writer.save()
        except:
            print("ERORR")
