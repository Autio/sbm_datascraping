
import ctypes # for popup window
import sys # for exception information

try: # Main exception handler
    
    import requests # for HTTP requests
    from bs4 import BeautifulSoup # for HTML parsing
    import bs4 # for type checking
    import xlsxwriter # for exporting to Excel - need xlsx as over 200k rows of data
    import os # to find user's desktop path
    import time # for adding datestamp to file output
    import re # for regular expressions

    # Timing the script
    startTime = time.time()

    # Configuration of request variables
    #url_SBM_TargetVsAchievement = 'http://sbm.gov.in/sbmreport/Report/Physical/SBM_TargetVsAchievement.aspx'
    url_SBM_FinanceProgress = 'http://sbm.gov.in/sbmreport/Report/Financial/SBM_StateReleaseAllocationincludingUnapproved.aspx'

    # For finance progress
    componentKey = 'ctl00$ContentPlaceHolder1$ddlComponent'
    componentVal = ''
    finYearKey = 'ctl00$ContentPlaceHolder1$ddlFinYear'
    finYearVal = ''

    submitKey = 'ctl00$ContentPlaceHolder1$btnSubmit'
    submitVal = 'Submit'

    targetKey = '__EVENTTARGET'
    targetVal = ''

    # __EVENTVALIDATION and __VIEWSTATE are dynamic authentication values which must be freshly updated when making a request.
    eventValKey = '__EVENTVALIDATION'
    eventVal = ''
    viewStateKey = '__VIEWSTATE'
    viewStateVal = ''

    # Function to return HTML parsed with BeautifulSoup from a POST request URL and parameters.
    def parsePOSTResponse(URL, parameters=''):
        responseHTMLParsed = ''
        r = requests.post(URL, data=parameters)
        if r.status_code == 200:
            responseHTML = r.content
            responseHTMLParsed = BeautifulSoup(responseHTML, 'html.parser')
        return responseHTMLParsed

    # Load the default page and scrape the component, finance year and authentication values
    initPage = parsePOSTResponse(url_SBM_FinanceProgress)
    eventVal = initPage.find('input',{'id':'__EVENTVALIDATION'})['value']
    viewStateVal = initPage.find('input',{'id':'__VIEWSTATE'})['value']
    componentOptions = []
    componentOptionVals = []
    componentSelection = initPage.find('select',{'id':'ctl00_ContentPlaceHolder1_ddlComponent'}) #changed for link 2
    componentOptions = componentSelection.findAll('option',{'contents':''}) # changed from selection to contents
    for componentOption in componentOptions:
        componentOptionVal = componentOption['value']
        componentOptionVals.append(componentOptionVal)
    finyearOptions = []
    finYearOptionVals = []
    finYearSelection = initPage.find('select',{'id':'ctl00_ContentPlaceHolder1_ddlFinYear'})
    finYearOptions = finYearSelection.findAll('option',{'contents':''})
    for finYearOption in finYearOptions:
        finYearOptionVal = finYearOption['value']
        finYearOptionVals.append(finYearOptionVal)

    # Initialise workbook
    todaysDate = time.strftime('%d-%m-%Y')
    desktopFile = os.path.expanduser('~/Desktop/SBM_FinanceProgress_' + todaysDate + '.xlsx')
    wb = xlsxwriter.Workbook(desktopFile)
    ws = wb.add_worksheet('SBM Test')
    ws.set_column('A:AZ', 22)
    rowCount = 1 # Adjust one row for printing table headers after main loop
    cellCount = 0

    # Global variable to store final table data
    lastBlockReportTable = ''

    # Global variable for keeping track of the state
    componentCount = 1

    # MAIN LOOP: loop through STATE values and scrape district and authentication values for each
    for componentOptionVal in componentOptionVals: # For testing, we can limit the states processed due to long runtime
        # params for Financial Progress: __EVENTTARGET, __EVENTARGUMENT,
        # LASTFOCUS, __VIEWSTATE, __EVENTVALIDATION,
        # ctl00$ContentPlaceholder1$ddlComponent, ctl00$ContentPlaceHolder1$ddlFinYear

        postParams = {
            eventValKey:eventVal,
            viewStateKey:viewStateVal,
            componentKey: 'C', # componentOptionVal,
            finYearKey: '2016-2017', #finyearOptionVal,
            submitKey:submitVal
        }
        componentPage = parsePOSTResponse(url_SBM_FinanceProgress, postParams)

        stateOptions = []
        stateOptionVals = []
        stateSelection = componentPage.findAll('a', {'id': re.compile('stName$')})
        stateOptions = stateSelection.findAll('option',{'contents':''})

        state_eventValVal = componentPage.find('input',{'id':'__EVENTVALIDATION'})['value']
        state_viewStateVal = componentPage.find('input',{'id':'__VIEWSTATE'})['value']
        districtOptions = []
        districtOptionVals = []
        districtSelection = componentPage.find('select',{'id':'ctl00_ContentPlaceHolder1_ddlComponent'})
        districtOptions = districtSelection.findAll('option',{'selected':''})
        for districtOption in districtOptions:
            if 'All District' not in districtOption.text and 'STATE HEADQUARTER' not in districtOption.text: # We do not want the top level data for the state or state headquarter data
                districtOptionVal = districtOption['value']
                districtOptionVals.append(districtOptionVal)

                # Process table data and output 
                blockReportTable = blockReport.find('table')
                if isinstance(blockReportTable,bs4.element.Tag): # Check whether data table successfully found on the page. Some blocks have no data.

                    # Store table for writing headers after loop
                    lastBlockReportTable = blockReportTable 

                    # Store state, district, and block information
                    stateNameText = blockReportTable.find('span',{'id':'ctl00_ContentPlaceHolder1_Rpt_data_ctl00_lblstatename'}).text
                    stateNameText = stateNameText.replace('State Name:-','');
                    stateNameText = stateNameText.strip();
                    districtNameText = blockReportTable.find('span',{'id':'ctl00_ContentPlaceHolder1_Rpt_data_ctl00_lbldtname'}).text
                    districtNameText = districtNameText.replace('District Name:-','');
                    districtNameText = districtNameText.strip();
                    blockNameText = blockReportTable.find('span',{'id':'ctl00_ContentPlaceHolder1_Rpt_data_ctl00_lblblname'}).text
                    blockNameText = blockNameText.replace('Block Name:-','');
                    blockNameText = blockNameText.strip();

                    print ('Currently processing: ' + stateNameText + ' (' + str(componentCount) + ' of ' + str(len(componentOptionVals)) + ')' + ' > ' + districtNameText + ' (' + str(districtCount) + ' of ' + str(len(districtOptionVals)) + ')' + ' > ' + blockNameText + ' (' + str(blockCount) + ' of ' + str(len(blockOptionVals)) + ')')

                    # Loop through rows and write data
                    blockReportRows = blockReportTable.find('tbody').findAll('tr') # Only process table body data
                    tableArray = []                
                    for tr in blockReportRows[0:len(blockReportRows)-1]: # Total row (bottom of table) dropped
                        tableRow = []
                        cols = tr.findAll('td')
                        # Write state, district, and block information
                        ws.write(rowCount,cellCount,stateNameText)
                        cellCount = cellCount+1
                        ws.write(rowCount,cellCount,districtNameText)
                        cellCount = cellCount+1
                        ws.write(rowCount,cellCount,blockNameText)
                        cellCount = cellCount+1
                        for td in cols:
                            # Tidy and format the cell content
                            cellText = td.text.replace('\*','')
                            cellText = cellText.strip()
                            try:
                                int(cellText)
                                cellText = int(cellText)
                            except:
                                cellText = cellText
                            # Store the cell data
                            tableRow.append(cellText)
                            ws.write(rowCount,cellCount,cellText)
                            cellCount = cellCount+1
                        tableArray.append(tableRow)        
                        rowCount = rowCount + 1
                        cellCount = 0
                blockCount = blockCount + 1
            districtCount = districtCount + 1
        componentCount = componentCount + 1

    # Write table headers based on final report
    print ('Processing table headers...')
    blockReportHeaderRows = lastBlockReportTable.find('thead').findAll('tr') # Only process table header data
    headerTableArray = []
    rowCount = 0
    cellCount = 0
    headerStyle = wb.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#0A8AD5'})
    for tr in blockReportHeaderRows[len(blockReportHeaderRows)-2:len(blockReportHeaderRows)-1]: # State, district, and block (bottom of table) + other headers dropped
        headerTableRow = []
        headerCols = tr.findAll('th')
        # Write state, district, and block headers
        ws.write(rowCount,cellCount,'State Name',headerStyle)
        cellCount = cellCount+1
        ws.write(rowCount,cellCount,'District Name',headerStyle)
        cellCount = cellCount+1
        ws.write(rowCount,cellCount,'Block Name',headerStyle)
        cellCount = cellCount+1
        for td in headerCols:
            # Tidy the cell content
            cellText = td.text.replace('\*','')
            cellText = cellText.strip()
            # Store the cell data
            headerTableRow.append(cellText)
            ws.write(rowCount,cellCount,cellText,headerStyle)
            cellCount = cellCount+1
        headerTableArray.append(tableRow)        
        rowCount = rowCount + 1
        cellCount = 0
                    
    print ('Done processing.' + ' Script executed in ' + str(int(time.time()-startTime)) + ' seconds.')
    # END MAIN LOOP

    # Finally, save the workbook
    wb.close()

except: # Main exception handler
    print('The program did not complete.')
    e = sys.exc_info()
    ctypes.windll.user32.MessageBoxW(0, "Sorry, there was a problem running this program.\n\nFor developer reference:\n\n" + str(e), "The program did not complete :-/", 1)
