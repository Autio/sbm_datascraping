import requests # for HTTP requests
from bs4 import BeautifulSoup # for HTML parsing
import xlwt # for exporting to MS Excel

# Configuration of request variables
url_SBM_TargetVsAchievement = 'http://sbm.gov.in/sbmreport/Report/Physical/SBM_TargetVsAchievement.aspx'
url_SBM_FinanceProgress = 'http://sbm.gov.in/sbmreport/Report/Financial/SBM_StateReleaseAllocationincludingUnapproved.aspx'

stateKey = 'ctl00$ContentPlaceHolder1$ddlState'
stateVal = '26'
districtKey = 'ctl00$ContentPlaceHolder1$ddlDistrict'
districtVal = '577'
blockKey = 'ctl00$ContentPlaceHolder1$ddlBlock'
blockVal = '6498'

submitKey = 'ctl00$ContentPlaceHolder1$btnSubmit'
submitVal = 'View Report'

# __EVENTVALIDATION and __VIEWSTATE are dynamic authentication values which must be freshly updated when making a request. 
eventValKey = '__EVENTVALIDATION' 
eventValVal = '/wEWOwLii9LWDQLq6fiEBwK4qJKGBgL7uLfDBQLMho26CAKkvMv0BAKrp/OzCAKzzOWcCQLfxNm+CQLZ25fbDALc9b7CDAL+s4fwDAL93OEdAvDc1R0C8dytHgLw3K0eAvPcrR4C8tytHgLz3N0dAvDcjR4C9dytHgL03K0eAvfcrR4C5tytHgLp3K0eAvPc0R0C8dztHQLx3OEdAvHc5R0C8dzZHQLx3N0dAvHc0R0C8dzVHQLx3MkdAvHcjR4C89zlHQLx3IEeAvDc7R0C8NzhHQLw3OUdAvPc1R0C8NzZHQLw3N0dAvPc2R0C8NzRHQKkoZaeCQLsrsqjBALvwazOCAKqi5aaDAKOoIjtBgKhysizDgKri+6aDAKatMWmDwLokp/lBgKMsPXEBAKMsIH4DwLP08CPCgKhr/W+CQL40JWiCmOx9253QCz2y/Qah474Zjvc/kkM'
viewStateKey = '__VIEWSTATE'
viewStateVal = '/wEPDwUJNTYyMDA4OTU4D2QWAmYPZBYCAgMPZBYEAh8PDxYCHgRUZXh0BUw8c3BhbiBjbGFzcz0iZ2x5cGhpY29uIGdseXBoaWNvbi1jaXJjbGUtYXJyb3ctbGVmdCI+PC9zcGFuPiBCYWNrIHRvIFByZXZpb3VzFgIeB29uY2xpY2sFKGphdmFzY3JpcHQ6aGlzdG9yeS5iYWNrKCk7IHJldHVybiBmYWxzZTtkAiEPZBYCAgEPZBYCAgMPZBYCAgEPZBYGAgMPEA8WBh4NRGF0YVRleHRGaWVsZAUJU3RhdGVOYW1lHg5EYXRhVmFsdWVGaWVsZAUHU3RhdGVJRB4LXyFEYXRhQm91bmRnZBAVIQlBbGwgU3RhdGUMQSAmIE4gSWxhbmRzDkFuZGhyYSBQcmFkZXNoEUFydW5hY2hhbCBQcmFkZXNoBUFzc2FtBUJpaGFyDENoaGF0dGlzZ2FyaAxEICYgTiBIYXZlbGkDR29hB0d1amFyYXQHSGFyeWFuYRBIaW1hY2hhbCBQcmFkZXNoD0phbW11ICYgS2FzaG1pcglKaGFya2hhbmQJS2FybmF0YWthBktlcmFsYQ5NYWRoeWEgUHJhZGVzaAtNYWhhcmFzaHRyYQdNYW5pcHVyCU1lZ2hhbGF5YQdNaXpvcmFtCE5hZ2FsYW5kBk9kaXNoYQpQdWR1Y2hlcnJ5BlB1bmphYglSYWphc3RoYW4GU2lra2ltClRhbWlsIE5hZHUJVGVsYW5nYW5hB1RyaXB1cmENVXR0YXIgUHJhZGVzaAtVdHRhcmFraGFuZAtXZXN0IEJlbmdhbBUhAi0xAjI2ATEBMgEzATQCMzQCMjgBNQE2ATcBOAE5AjM1AjEwAjExAjEyAjEzAjE0AjE1AjE2AjE3AjE4AjMyAjE5AjIwAjIxAjIyAjM2AjIzAjI0AjMzAjI1FCsDIWdnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZxYBAgFkAgsPEA8WBh8CBQxEaXN0cmljdE5hbWUfAwUKRGlzdHJpY3RJRB8EZ2QQFQUMQWxsIERpc3RyaWN0EiBTVEFURSBIRUFEUVVBUlRFUghOSUNPQkFSUyhOT1JUSCBBTkQgTUlERExFIEFOREFNQU4gICAgICAgICAgICAgICAgDlNPVVRIIEFOREFNQU5TFQUCLTEDNjk2AzU3NwM1NzgDNTc2FCsDBWdnZ2dnFgECAmQCEw8QDxYGHwIFCUJsb2NrTmFtZR8DBQdCbG9ja0lEHwRnZBAVBAlBbGwgQmxvY2syQ0FNUEJFTEwgQkFZICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAyQ0FSIE5JQ09CQVIgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAyTkFOQ09XUklFICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAVBAItMQQ2NDk4BDY0OTkENjUwMBQrAwRnZ2dnZGQYAQUeX19Db250cm9sc1JlcXVpcmVQb3N0QmFja0tleV9fFgMFD2N0bDAwJGljb25fd29yZAUQY3RsMDAkaWNvbl9leGNlbAUSY3RsMDAkaWNvbl9wcmludGVyWYif/lfAnfczRDANjFyFVBDIukQ='

# Host and referer info not required
#hostURL = 'sbm.gov.in'
#originURL = 'http://sbm.gov.in'

postParams = {
    stateKey:stateVal,
    districtKey:districtVal,
    blockKey:blockVal,
    submitKey:submitVal,
    eventValKey:eventValVal,
    viewStateKey:viewStateVal
}

#postHeaders = {'Host':hostURL,'Origin':originURL,'Referer':'http://sbm.gov.in/sbmreport/Report/Physical/SBM_TargetVsAchievement.aspx'}

# Make the request. Form must be POSTed
r = requests.post(url_SBM_FinanceProgress, data=postParams)

# Check that request returns OK and then process HTML
if r.status_code == 200:
    responseHTML = r.content
    responseHTMLParsed = BeautifulSoup(responseHTML, 'html.parser')
    print (responseHTMLParsed)
    responseTable = responseHTMLParsed.find('table')
    responseRows = responseTable.findAll('tr')
    tableArray = []
    rowCount = 0
    cellCount = 0
    wb = xlwt.Workbook()
    ws = wb.add_sheet('SBM Test')
    for tr in responseRows:
        tableRow = []
        cols = tr.findAll('td')
        for td in cols:
            # Tidy the cell content
            cellText = td.text.replace('\*','')
            cellText = cellText.strip()
            # Store the cell data
            tableRow.append(cellText)
            ws.write(rowCount,cellCount,cellText)
            cellCount = cellCount+1
        tableArray.append(tableRow)        
        rowCount = rowCount + 1
        cellCount = 0
    wb.save('C:\Users\Petriau\Documents\sbm_datascraping\SBM_test.xls')
    #print(tableArray)    
        


