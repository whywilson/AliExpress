#!/usr/bin/env python3
import requests
import json
import time
import sys
import xlwt

cookie = ""

try:
    with open('cookie.txt', 'rt') as f:
        cookie = f.read()
    if cookie != "":
        print("Cookie read success")
    else:
        print("Cookie read failed, Existing")
        sys.exit()
except:
    print("Cookie read failed, please save cookie to cookie.txt")
    sys.exit()
print("Please input product ID: ", end="")
productId = input()

if productId == "":
    print("Please input product ID: ", end="")
    productId = input()

country_code = ["AD", "AE", "AF", "AG", "AI", "AL", "ALA", "AM", "AN", "AO", "AQ", "AR", "AS", "ASC", "AT", "AU", "AW",
                "AZ", "BA", "BB", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BLM", "BM", "BN", "BO", "BQ", "BR", "BS",
                "BT", "BV", "BW", "BY", "BZ", "CA", "CA", "CC", "CF", "CG", "CH", "CI", "CK", "CL", "CL", "CM", "CO",
                "CR", "CV", "CW", "CX", "CY", "CZ", "DE", "DJ", "DK", "DM", "DO", "DZ", "EAZ", "EC", "EE", "EG", "EH",
                "ER", "ES", "ES", "ET", "FI", "FJ", "FK", "FM", "FO", "FR", "FR", "GA", "GBA", "GD", "GE", "GF", "GGY",
                "GH", "GI", "GL", "GM", "GN", "GP", "GQ", "GR", "GT", "GU", "GW", "GY", "HK", "HM", "HN", "HR", "HT",
                "HU", "IC", "ID", "IE", "IL", "IM", "IN", "IO", "IQ", "IS", "IT", "JEY", "JM", "JO", "JP", "KE", "KG",
                "KH", "KI", "KM", "KN", "KR", "KS", "KW", "KY", "KZ", "LA", "LB", "LC", "LI", "LK", "LR", "LS", "LT",
                "LU", "LV", "LY", "MA", "MAF", "MC", "MD", "MG", "MH", "MK", "ML", "MM", "MN", "MNE", "MO", "MP", "MQ",
                "MR", "MS", "MT", "MU", "MV", "MW", "MX", "MY", "MZ", "NA", "NC", "NE", "NF", "NG", "NI", "NL", "NL",
                "NO", "NP", "NR", "NU", "NZ", "OM", "PA", "PE", "PF", "PG", "PH", "PK", "PL", "PL", "PM", "PN", "PR",
                "PS", "PT", "PW", "PY", "QA", "RE", "RO", "RU", "RW", "SA", "SB", "SC", "SE", "SG", "SGS", "SH", "SI",
                "SJ", "SK", "SL", "SM", "SN", "SO", "SR", "SRB", "SS", "ST", "SV", "SX", "SZ", "TC", "TD", "TF", "TG",
                "TH", "TJ", "TK", "TLS", "TM", "TN", "TO", "TR", "TT", "TV", "TW", "TZ", "UA", "UA", "UG", "UK", "UK",
                "UM", "US", "UY", "UZ", "VA", "VC", "VE", "VG", "VI", "VN", "VU", "WF", "WS", "YE", "YT", "ZA", "ZM",
                "ZR", "ZW"]
#country_code = ["US", "UK", "CO", "CA", "XX", "AA", "BB", "HK"]

all_carrier = []
shipping_method = []
count = 0
error_country_code = []

headers = {
    'authority': 'www.aliexpress.com',
    'pragma': 'no-cache',
    'cache-control': 'no-cache',
    'accept': 'application/json, text/plain, */*',
    'dnt': '1',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.88 Safari/537.36',
    'sec-fetch-site': 'same-origin',
    'sec-fetch-mode': 'cors',
    'referer': 'https://www.aliexpress.com/item/'+productId+'.html?spm=a2g0o.productlist.0.0.340a3f05tctG81&algo_pvid=3f12e59e-291f-43ed-9bc6-97d1b623d2f1&algo_expid=3f12e59e-291f-43ed-9bc6-97d1b623d2f1-0&btsid=d04cae72-c48b-4e07-8f97-6321f17912b2&ws_ab_test=searchweb0_0,searchweb201602_9,searchweb201603_53',
    'accept-encoding': 'gzip, deflate, br',
    'accept-language': 'en-US,en;q=0.9,zh-HK;q=0.8,zh-CN;q=0.7,zh;q=0.6',
    'cookie': cookie,
}

print("Start to scrap shipping method of product " + productId)
for i in range(len(country_code)):
    params = (
        ('productId', productId),
        ('count', '1'),
        ('minPrice', '2'),
        ('country', country_code[i]),
        ('provinceCode', ''),
        ('cityCode', ''),
        ('tradeCurrency', 'USD'),
        ('userScene', 'PC_DETAIL_SHIPPING_PANEL'),
    )

    response = requests.get(
        'https://www.aliexpress.com/aeglodetailweb/api/logistics/freight', headers=headers, params=params)
    print(str(count), country_code[i], end=" ")
    try:
        # print response.text
        body = json.loads(response.text)
        shipping_method.append({})
        freightResult = body['body']['freightResult']

        # print(freightResult)
        print("Qty", len(freightResult), end=": ")

        for j in range(len(freightResult)):
            print(freightResult[j]['company'],str(freightResult[j]['freightAmount']['value']), end=", ")
            shipping_method[i].setdefault(
                freightResult[j]['company'], freightResult[j]['freightAmount']['value'])
            if freightResult[j]['company'] not in all_carrier:
                all_carrier.append(freightResult[j]['company'])
                # show carrier and value
        count = count + 1
    except Exception as e:
        error_country_code.append(country_code[i])
        print(" Error " + str(e))
        continue
    time.sleep(0.02)
    print("")

# print all_carrier
print("----------------")
new_workbook = xlwt.Workbook()
work_sheet = new_workbook.add_sheet("Data")
work_sheet.write(0,0,"Country")
for i in range(len(all_carrier)):
    work_sheet.write(0, i+1, all_carrier[i])

result_content = ""
error_count = 0
for i in range(len(country_code)):
    if country_code[i] in error_country_code:
        print(country_code[i] + " Error")
        error_count = error_count+1
        continue
    print("Export " + str(i) + " " + country_code[i])
    work_sheet.write(i+1-error_count, 0, country_code[i])
    try:
        result_content = result_content + country_code[i] + ","
        for j in range(len(all_carrier)):
            if all_carrier[j] in shipping_method[i]:
                work_sheet.write(i+1-error_count, j+1, shipping_method[i][all_carrier[j]])
            else:
                work_sheet.write(i + 1-error_count, j + 1, "null")
        new_workbook.save("Product-" + productId + "-Shipping-Method.xls")
    except:
        print(country_code[i] + " Error")
        result_content = result_content + ",Error"
        continue
    result_content = result_content + "\n"

if len(error_country_code) > 0:
    work_sheet_failed = new_workbook.add_sheet("Failed")
    work_sheet_failed.write(0,0,"Country Code")
    for i in range(len(error_country_code)):
        work_sheet_failed.write(i+1, 0, error_country_code[i])

new_workbook.save("Product-" + productId + "-Shipping-Method.xls")
