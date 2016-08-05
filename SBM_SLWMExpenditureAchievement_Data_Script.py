### Tool to scrape Swachh Bharat Mission data from the sbm.gov.in government website ###
### Daniel Robertson                                                                 ###
### Petri Autio                                                                      ###
### 2016                                                                             ###
import ctypes  # for popup window
import sys  # for exception information

try:  # Main exception handler

    import requests  # for HTTP requests
    from bs4 import BeautifulSoup  # for HTML parsing
    import bs4  # for type checking
    import xlsxwriter  # for exporting to Excel - need xlsx as over 200k rows of data
    import os  # to find user's desktop path
    import time  # for adding datestamp to file output
    import re  # for regular expressions

    # Timing the script
    startTime = time.time()

    # Configuration of request variables
    url_SBM = 'http://sbm.gov.in/sbmreport/Report/Panchayat/SBM_SLWMExpenditureAchievement.aspx'

    outputArray = []

    # Key / value pairs to pass to ASP.net POST parameters
    stateKey = 'ctl00$ContentPlaceHolder1$ddlState'
    stateVal = ''
    finYearKey = 'ctl00$ContentPlaceHolder1$ddlFinyear'
    finYearVal = ''
    fundTypeKey = 'ctl00$ContentPlaceHolder1$ddlFundType'
    fundTypeVal = ''

    # submit button on first page
    submitKey = 'ctl00$ContentPlaceHolder1$btnGetData'
    submitVal = 'Submit'

    # __EVENTVALIDATION. __VIEWSTATE and __EVENTTARGET are dynamic authentication values which must be freshly updated when making a request.
    eventValKey = '__EVENTVALIDATION'
    eventVal = ''
    eventArgKey = '__EVENTARGUMENT'
    eventArgVal = ''
    viewStateKey = '__VIEWSTATE'
    viewStateVal = ''
    lastFocusKey = '__LASTFOCUS'
    lastFocusVal = ''

    targetKey = '__EVENTTARGET'
    targetVal = ''


    # Function to return HTML parsed with BeautifulSoup from a POST request URL and parameters.
    def parsePOSTResponse(URL, parameters='', pagetype = ''):
        responseHTMLParsed = ''
        attempts = 20
        for i in range(attempts):
            r = requests.post(URL, data=parameters)
            if r.status_code == 200:
                responseHTML = r.content
                responseHTMLParsed = BeautifulSoup(responseHTML, 'html.parser')
            if not responseHTMLParsed == '':
                return responseHTMLParsed
            else:
                print ("    Could not load #s page - attempt %s out of %s" % (pagetype, i+1, attempts))


    # Given two dicts, merge them into a new dict as a shallow copy.
    def merge_two_dicts(x, y):
        z = x.copy()
        z.update(y)
        return z


    # Load the default page and scrape the component, finance year and authentication values
    initPage = parsePOSTResponse(url_SBM)
    eventVal = initPage.find('input', {'id': '__EVENTVALIDATION'})['value']
    viewStateVal = initPage.find('input', {'id': '__VIEWSTATE'})['value']

    stateOptions = []
    stateOptionVals = []
    stateSelection = initPage.find('select', {'id': 'ctl00_ContentPlaceHolder1_ddlState'})
    stateOptions = stateSelection.findAll('option', {'contents': ''})  # changed from selection to contents
    for stateOption in stateOptions:
        if not stateOption['value'] == '-1':
            stateOptionVal = stateOption['value']
            stateOptionVals.append(stateOptionVal)

    finyearOptions = []
    finYearOptionVals = []
    finYearSelection = initPage.find('select', {'id': 'ctl00_ContentPlaceHolder1_ddlFinyear'})
    finYearOptions = finYearSelection.findAll('option', {'contents': ''})
    for finYearOption in finYearOptions:
        # Don't include the -2 option which only indicates nothing has been selected in the dropdown
        if not finYearOption['value'] == '-2':
            finYearOptionVals.append(finYearOption['value'])

    fundTypeOptions = []
    fundTypeOptionVals = []
    fundTypeSelection = initPage.find('select', {'id': 'ctl00_ContentPlaceHolder1_ddlFundType'})
    fundTypeOptions = fundTypeSelection.findAll('option', {'contents': ''})
    for fundTypeOption in fundTypeOptions:
        # Don't include % as an option because it just means "all types"
        if not fundTypeOption['value'] == '%':
            fundTypeOptionVals.append(fundTypeOption['value'])

    # Initialise workbook
    todaysDate = time.strftime('%d-%m-%Y')
    desktopFile = os.path.expanduser('~/Desktop/SBM_SWLMExpenditureAchievement_' + todaysDate + '.xlsx')
    wb = xlsxwriter.Workbook(desktopFile, {'strings_to_numbers': True})
    ws = wb.add_worksheet('SBM')
    ws.set_column('A:AZ', 22)
    rowCount = 1  # Adjust one row for printing table headers after main loop
    cellCount = 0
    stateCount = 1

    # Global variable to store final table data
    lastTable = ''

    # Global variable for keeping track of the state
    componentCount = 1
    componentName = ''

    # Global variable for ensuring headers get added only once
    headerFlag = False

    # MAIN LOOP: loop through States. Scrape link values from page
    # We can merely select All States and and all Financial Years to get to the necessary page to cycle through
    stateOptionVal = '-1'
    finYearOptionVal = '-2'

    # Loop through fund types | Doesn't seem to be working on the site consistently. Maybe can do without?
    # Use static fund type for now
    # for fundTypeOptionVal in fundTypeOptionVals:
    fundTypeOptionVal = '%'

    eventVal = initPage.find('input', {'id': '__EVENTVALIDATION'})['value']
    viewStateVal = initPage.find('input', {'id': '__VIEWSTATE'})['value']

    postParams = {
        eventValKey: eventVal,
        eventArgKey: eventArgVal,
        viewStateKey: viewStateVal,
        lastFocusKey: lastFocusVal,
        targetKey: targetVal,
        stateKey: stateOptionVal,
        finYearKey: finYearOptionVal,
        fundTypeKey: fundTypeOptionVal,
        submitKey: submitVal
    }

    allStatePage = parsePOSTResponse(url_SBM, postParams, 'state')

    # Find GPs
    # By clicking on the Total No. of GP value we get tables at the GP level, if that value isn't 0

    # Find all states in the list
    stateOptions = []
    stateOptionVals = []
    stateSelection = allStatePage.findAll('a', {'id': re.compile('lnkGPRP$')})
    stateNames = allStatePage.findAll('a', {'id': re.compile('lnkStateTotal')})
    sCount = 0
    # Find all states and links to click through
    for s in stateSelection:
        stateOptionVal = s.text
        targetOptionVal = s['id']
        stateOptionVals.append([stateOptionVal, targetOptionVal, stateNames[sCount].text])
        sCount = sCount + 1

    # Find all the parameters required
    linkOptions = []
    linkSelection = allStatePage.findAll('input', {'id': re.compile('hfStateID$')})
    linkIndex = 0
    for link in linkSelection:
        linkId = link['id'].replace('_', '$')
        linkOptions.append([linkId, link['value']])
        linkIndex = linkIndex + 1

    # create dictionary from list
    paramDictionary = {key: str(value) for key, value in linkOptions}

    stateCount = 1
    # Should cycle through all the items in stateOptions
    for s in stateOptionVals:
        # If state has no recorded values for GPs, then can't click to it
        if not s[0] == '0':
            stateLinkVal = s[1]
            stateLinkVal = stateLinkVal.replace('_', '$')  # Tweaking to get $ signs in the right place
            stateLinkVal = stateLinkVal.replace('rptr$cen', 'rptr_cen')
            stateLinkVal = stateLinkVal.replace('lnkbtn$st', 'lnkbtn_st')

            eventVal = allStatePage.find('input', {'id': '__EVENTVALIDATION'})['value']
            viewStateVal = allStatePage.find('input', {'id': '__VIEWSTATE'})['value']
            # Need to call Javascript __doPostBack() on the links
            postParams = {
                '__EVENTARGUMENT': '',
                '__EVENTTARGET': stateLinkVal,
                eventValKey: eventVal,
                viewStateKey: viewStateVal,
            }

            postParams = merge_two_dicts(paramDictionary, postParams)

            GPPage = parsePOSTResponse(url_SBM, postParams, 'GP')

            # Process table data and output
            ReportTable = GPPage.find('table')

            # Write table headers
            if not headerFlag:
                print('Processing table headers...')
                headerRows = ReportTable.find('thead').findAll('tr')  # Only process table header data
                headerTableArray = []
                rowCount = 0

                headerStyle = wb.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#0A8AD5'})

                for tr in headerRows:  # last headeR row only
                    cellCount = 0
                    headerTableRow = []
                    headerCols = tr.findAll('th')
                    # Write state, district, and block headerss
                    for td in headerCols:
                        # Tidy the cell content
                        cellText = td.text.replace('\*', '')
                        cellText = cellText.strip()
                        # Store the cell data
                        headerTableRow.append(cellText)
                        cellCount = cellCount + 1
                    rowCount = rowCount + 1

                headerFlag = True

            # Write table data
            if isinstance(ReportTable,
                          bs4.element.Tag):  # Check whether data table successfully found on the page. Some blocks have no data.
                # Store table for writing headers after loop

                print('Currently processing: ' + s[2] + ' (' + str(
                    stateCount) + ' of ' + str(
                    len(stateOptionVals)) + ')')

                lastReportTable = ReportTable
                ReportRows = ReportTable.findAll(
                    'tr')  # Bring entire table including headers because body isn't specified
                if len(ReportRows) > 1:
                    for tr in ReportRows[2:len(
                            ReportRows) - 2]:  # Start from 2 (body of table) and total row (bottom of table) dropped
                        cellCount = 0
                        tableRow = []
                        cols = tr.findAll('td')

                        for td in cols:
                            # Tidy and format the cell content
                            cellText = td.text.replace('\*', '')
                            cellText = cellText.strip()
                            try:
                                long(cellText)
                                cellText = long(cellText)
                            except:
                                cellText = cellText
                            # Store the cell data
                            tableRow.append(cellText)
                            ws.write(rowCount, cellCount, cellText)
                            cellCount = cellCount + 1
                        rowCount = rowCount + 1

                else:
                    sta = "no data in state"
                stateCount = stateCount + 1

        else:
            print('Currently processing: ' + s[2] + ' (' + str(
                    stateCount) + ' of ' + str(
                    len(stateOptionVals)) + ') - No GP data available')

            stateCount = stateCount + 1

    print('Done processing.' + ' Script executed in ' + str(int(time.time() - startTime)) + ' seconds.')
    # END MAIN LOOP

    # Finally, save the workbook
    wb.close()

except:  # Main exception handler
    print('The program did not complete.')
    e = sys.exc_info()
    print(e)
    ctypes.windll.user32.MessageBoxW(0,
                                     "Sorry, there was a problem running this program.\n\nFor developer reference:\n\n" + str(
                                         e), "The program did not complete :-/", 1)
