# author: Niko
# Email: liuzx0636@outlook.com

import requests
import json
import time
import xlwt
from alive_progress import alive_bar
import sys
import getopt


class Hito(object):
    def __init__(self):
        self.headers = {
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36'
        }
        self.from_list = []
        self.who_list = []
        self.hito_list = []
        self.count = 0
        self.excel_name = ''

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
                time.sleep(0.5)
                resp = requests.get('https://v1.hitokoto.cn/', headers=self.headers)
                data = json.loads(resp.text)

                self.hito_list.append(data['hitokoto'])
                self.who_list.append(data['from_who'])
                self.from_list.append(data['from'])

                download()

        for who in self.who_list:
            if who == None:
                self.who_list[self.who_list.index(who)] = '未知'

    def write_to_excel(self, hito_list, who_list, from_list):

        # 创建Workbook
        excel = xlwt.Workbook(encoding='utf-8')
        sheet = excel.add_sheet('审查', cell_overwrite_ok=True)

        # 写入标题
        sheet.write(0, 0, '一言')
        sheet.write(0, 1, '出处')
        sheet.write(0, 2, '作者')

        # 批量写入

        def write_data(y, data_list, name):
            x = 1

            with alive_bar(len(data_list), title='写入{}中……'.format(name)) as save:
                for data in data_list:
                    sheet.write(x, y, data)
                    x = x + 1

                    save()
                    save.text("正在写入: {}".format(data))

        write_data(0, hito_list, '一言')
        write_data(1, from_list, '出处')
        write_data(2, who_list, '作者')

        excel.save(self.excel_name)

    def run(self):
        self.get_list()
        self.write_to_excel(self.hito_list, self.who_list, self.from_list)


