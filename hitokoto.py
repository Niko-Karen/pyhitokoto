# author: Niko
# Email: liuzx0636@outlook.com

import requests
import json
import time
from os.path import exists
import xlwt
from alive_progress import alive_bar
import sys
import getopt
import csv


class Hito(object):
    def __init__(self):
        self.headers = {
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36'
        }

        self.count = 0
        self.excel_name = ''
        self.csv_name = ''

        self.from_list = []
        self.who_list = []
        self.hito_list = []

    def get_argv(self):
        argv = sys.argv[1:]
        if argv == [''] or []:
            return False

        try:
            opts, args = getopt.getopt(argv, "hc:e:", ["help", "count=", "excel="])
        except getopt.GetoptError:
            print('Error: hitokoto.py -c <quantity> -e <excel file name>')
            print('   or: test_arg.py --count=<quantity> --excel=<excel file name>')
            sys.exit(2)

        for opt, arg in opts:
            if opt in ("-h", "--help"):
                print('usage: hitokoto.py -c <quantity> -e <excel file name>')
                print('   or: hitokoto.py --count=<quantity> --excel=<excel file name>')
                sys.exit(0)

            elif opt in ("-c", "--count"):
                self.count = int(arg)

            elif opt in ("-e", "--excel"):
                self.excel_name = arg

    def get_list(self):

        with alive_bar(self.count, title="缓存数据中……") as download:
            for i in range(self.count):

                resp = requests.get(
                    'https://v1.hitokoto.cn/', headers=self.headers
                ).text

                if resp == '':
                    self.count += 1
                    continue
                try:
                    data = json.loads(resp)
                except Exception:
                    self.count += 1
                    continue

                self.hito_list.append(str(data['hitokoto']))
                self.who_list.append(data['from_who'])
                self.from_list.append(str(data['from']))

                download()

        for who in self.who_list:
            if who is None:
                self.who_list[self.who_list.index(who)] = '未知'
        time.sleep(0.1)

        return self.hito_list, self.from_list, self.who_list

    def write_to_excel(self):

        # 创建Workbook
        excel = xlwt.Workbook(encoding='utf-8')
        sheet = excel.add_sheet('审查', cell_overwrite_ok=True)

        # 写入标题
        sheet.write(0, 0, '一言')
        sheet.write(0, 1, '出处')
        sheet.write(0, 2, '作者')

        # 批量写入

        def write_data(y, data_list, name):
            x = 2

            with alive_bar(len(data_list), title='写入{}中……'.format(name)) as save:
                for data in data_list:
                    sheet.write(x, y, data)
                    x = x + 1

                    save()
                    save.text("正在写入: {}".format(data))

        write_data(0, self.hito_list, '一言')
        write_data(1, self.from_list, '出处')
        write_data(2, self.who_list, '作者')

        excel.save(self.excel_name)

    def write_to_csv(self):
        with alive_bar(self.count, title="写入数据中……") as csv_down:
            for i in range(self.count):

                resp = requests.get(
                    'https://v1.hitokoto.cn/', headers=self.headers
                ).text

                if resp == '':
                    self.count += 1
                    continue
                try:
                    data = json.loads(resp)
                except Exception:
                    self.count += 1
                    continue

                hito = str(data['hitokoto'])
                who = data['from_who']
                from_who = str(data['from'])

                if who is None:
                    who = '未知'

                # all_list = list({hito for hito in hito_list : [who for who in
                # who_list, from_who for from_who in from_list]})
                all_list = [hito, who, from_who]

                if exists(self.csv_name):
                    mode = 'a'
                else:
                    mode = 'w'

                with open(self.csv_name, mode=mode) as csv_h:
                    writer = csv.writer(csv_h)
                    writer.writerow(all_list)

                csv_down()

                time.sleep(1)

    def run(self):
        self.write_to_excel()
        self.write_to_csv()
