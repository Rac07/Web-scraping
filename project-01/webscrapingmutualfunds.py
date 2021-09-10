import requests, webbrowser, openpyxl, bs4, string

wb = openpyxl.load_workbook('mutualfunds.xlsx')
type(wb)

sheet = wb.get_sheet_by_name('Sheet1')
cellnumber = ['4', '6', '8', '10']

def mutual_fund(fundname, cellid):  
    res = requests.get('https://www.google.com/search?q=' + 'money control ' + fundname)
    res.raise_for_status()

    parser = bs4.BeautifulSoup(res.content, 'html.parser')
    elem = parser.select('div#main > div > div > div > a')

    res1 = requests.get('https://www.google.com' + elem[0].get('href'))
    soup = bs4.BeautifulSoup(res1.content, 'html.parser')
    valueElem = soup.select('#mc_content > div > section.clearfix.section_four.laststickysection.graybgtbl > div > div.common_left > div.data_container.returns_table.table-responsive > table > tbody > tr')

    returns = {}
    for t in range(len(valueElem)):
        returns[valueElem[t].select('td.robo_medium')[0].text.strip()]=(valueElem[t].select('td.green_text')[0].text.strip())

    k = list(map(chr, range(ord('C'),ord('M')+1)))

    for ik, ret in zip(k, returns):
        if sheet[ik + '2'].value == ret:
            sheet[ik + cellid] = returns[ret]
        else:
            sheet[ik + cellid] = 0
    wb.save('mutualfunds.xlsx')
    
for i in cellnumber:
    print(i)
    fund = sheet['B' + i].value
    print(fund)
    mutual_fund(fund, i)
    









