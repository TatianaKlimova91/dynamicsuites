import requests
import xlsxwriter


ids = []


def get_data():
    url = f"{domain}/api/v2/testPoints/search?"
    data = {"testPlanIds": [f'{testplanId}']}
    res = requests.post(url, data=f"{data}", headers=auth, verify=False)
    return res


def get_users():
    url = f"{domain}/api/Users/multiple"
    return url


def formater(res, url):
    workbook = xlsxwriter.Workbook('Example.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', 'Status')
    worksheet.write('B1', 'Name')
    worksheet.write('C1', 'Tester')
    worksheet.write('D1', 'Modified Date')

    row = 1
    col = 0

    for i in res.json():
        if i['testerId'] is not None:
            ids.append(i['testerId'])
            users = requests.post(url, data=f'{list(set(ids))}', headers=auth, verify=False)
            for name in users.json():
                full = dict(name=name['displayName'], id=name['id'])
                if i['testerId'] in list(full.values())[1]:
                    worksheet.write(row, col, i['status'])
                    worksheet.write(row, col + 1, i['name'])
                    worksheet.write(row, col + 2, list(full.values())[0])
                    worksheet.write(row, col + 3, i['modifiedDate'])
                    row += 1
        else:
            worksheet.write(row, col, i['status'])
            worksheet.write(row, col + 1, i['name'])
            worksheet.write(row, col + 2, 'Нет тестировщика')
            worksheet.write(row, col + 3, i['modifiedDate'])
            row += 1

    workbook.close()


if __name__ == '__main__':
    domain = ""  #например https://testit.software
    token = ""  #например RTIxa05TaGpva0hHUUpxVzksd
    testplanId = ""  #например 2fa2fbec-7a95-4f81-886a-06cf5f769b69

    auth = {"Authorization": f"PrivateToken {token}", "Content-Type": "application/json"}
    formater(get_data(), get_users())
