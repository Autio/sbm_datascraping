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
        if not finYearOption['value'] == '-2':
            finYearOptionVal = finYearOption['value']
            finYearOptionVals.append(finYearOptionVal)

    # Initialise workbook
    todaysDate = time.strftime('%d-%m-%Y')
    desktopFile = os.path.expanduser('~/Desktop/SBM_FinanceProgress_' + todaysDate + '.xlsx')
    wb = xlsxwriter.Workbook(desktopFile)
    ws = wb.add_worksheet('SBM')
    ws.set_column('A:AZ', 22)
    rowCount = 1 # Adjust one row for printing table headers after main loop
    cellCount = 0

    # Global variable to store final table data
    lastBlockReportTable = ''

    # Global variable for keeping track of the state
    componentCount = 1

    # MAIN LOOP: loop through STATE values and scrape district and authentication values for each
    for componentOptionVal in componentOptionVals: # For testing, we can limit the states processed due to long runtime

        eventVal = initPage.find('input',{'id':'__EVENTVALIDATION'})['value']
        viewStateVal = initPage.find('input',{'id':'__VIEWSTATE'})['value']

        for finYearOptionVal in finYearOptionVals:
            postParams = {
                eventValKey:eventVal,
                viewStateKey:viewStateVal,
                componentKey: componentOptionVal,
                finYearKey: finYearOptionVal,
                submitKey: submitVal
            }
            componentPage = parsePOSTResponse(url_SBM_FinanceProgress, postParams)

            # Should find all event targets too
            stateOptions = []
            stateOptionVals = []

            stateSelection = componentPage.findAll('a', {'id': re.compile('stName$')})
            # Find all states and links to click through
            for s in stateSelection[10:18]:
                stateOptionVal = s.text
                targetOptionVal = s['id']
                stateOptionVals.append([stateOptionVal, targetOptionVal])

            eventVal = componentPage.find('input',{'id':'__EVENTVALIDATION'})['value']
            viewStateVal = componentPage.find('input',{'id':'__VIEWSTATE'})['value']

            info = {'__EVENTARGUMENT' : '', '__EVENTTARGET' : 'ctl00$ContentPlaceHolder1$rptr_cen$ctl01$lnkbtn_stName', eventValKey:eventVal,
                viewStateKey:viewStateVal}

            # create dictionary from list
            paramDictionary = data={key: str(value) for key, value in info.items()}

            # Should cycle through all the items in stateOptions
            for s in stateOptionVals:
                stateLinkVal = s[1]
                stateLinkVal = stateLinkVal.replace('_', '$') # Tweaking to get $ signs in the right place
                stateLinkVal = stateLinkVal.replace('rptr$cen', 'rptr_cen')
                stateLinkVal = stateLinkVal.replace('lnkbtn$st', 'lnkbtn_st')

                # Need to call Javascript __doPostBack() on the links
                postParams = {
                    '__EVENTARGUMENT': '',
                    '__EVENTTARGET': stateLinkVal,
                    eventValKey: eventVal,
                    viewStateKey: viewStateVal,
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl01$hfd_StateId':"26",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl02$hfd_StateId':"1",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl03$hfd_StateId':"2",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl04$hfd_StateId':"3",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl05$hfd_StateId':"4",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl06$hfd_StateId':"34",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl07$hfd_StateId':"28",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl08$hfd_StateId':"5",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl09$hfd_StateId':"6",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl10$hfd_StateId':"7",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl11$hfd_StateId':"8",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl12$hfd_StateId':"9",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl13$hfd_StateId':"35",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl14$hfd_StateId':"10",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl15$hfd_StateId':"11",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl16$hfd_StateId':"12",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl17$hfd_StateId':"13",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl18$hfd_StateId':"14",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl19$hfd_StateId':"15",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl20$hfd_StateId':"16",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl21$hfd_StateId':"17",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl22$hfd_StateId':"18",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl23$hfd_StateId':"32",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl24$hfd_StateId':"19",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl25$hfd_StateId':"20",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl26$hfd_StateId':"21",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl27$hfd_StateId':"22",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl28$hfd_StateId':"36",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl29$hfd_StateId':"23",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl30$hfd_StateId':"24",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl31$hfd_StateId':"33",
                    'ctl00$ContentPlaceHolder1$rptr_cen$ctl32$hfd_StateId':"25"

                }

                componentPage = parsePOSTResponse(url_SBM_FinanceProgress, postParams)

                # Process table data and output
                ReportTable = componentPage.find('table')

                if isinstance(ReportTable,bs4.element.Tag): # Check whether data table successfully found on the page. Some blocks have no data.
                     # Store table for writing headers after loop
                    lastReportTable = ReportTable
                    ReportRows = ReportTable.findAll('tr') # Bring entire table including headers because body isn't specified
                    if len(ReportRows) > 4:
                        for tr in ReportRows[4:len(ReportRows)-1]: # Start from 4 (body of table) and total row (bottom of table) dropped
                            cellCount = 0
                            tableRow = []
                            cols = tr.findAll('td')
                            # Write stored information in columns prior to data: Financial Year, State/Center, Statename
                            ws.write(rowCount,cellCount,componentOptionVal)
                            cellCount = cellCount + 1
                            ws.write(rowCount,cellCount,finYearOptionVal)
                            cellCount = cellCount + 1
                            ws.write(rowCount,cellCount,s[0])
                            cellCount = cellCount + 1
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
                            rowCount = rowCount + 1
                    else:
                        sta = "no data in state"
                componentCount = componentCount + 1

        # Write table headers based on final report
        # print ('Processing table headers...')
        # blockReportHeaderRows = lastBlockReportTable.find('thead').findAll('tr') # Only process table header data
        # headerTableArray = []
        # rowCount = 0
        # cellCount = 0
        # headerStyle = wb.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#0A8AD5'})

        # for tr in blockReportHeaderRows[len(blockReportHeaderRows)-2:len(blockReportHeaderRows)-1]: # State, district, and block (bottom of table) + other headers dropped
        #     headerTableRow = []
        #     headerCols = tr.findAll('th')
        #     # Write state, district, and block headers
        #     ws.write(rowCount,cellCount,'State Name',headerStyle)
        #     cellCount = cellCount+1
        #     ws.write(rowCount,cellCount,'District Name',headerStyle)
        #     cellCount = cellCount+1
        #     ws.write(rowCount,cellCount,'Block Name',headerStyle)
        #     cellCount = cellCount+1
        #     for td in headerCols:
        #         # Tidy the cell content
        #         cellText = td.text.replace('\*','')
        #         cellText = cellText.strip()
        #         # Store the cell data
        #         headerTableRow.append(cellText)
        #         ws.write(rowCount,cellCount,cellText,headerStyle)
        #         cellCount = cellCount+1
        #     headerTableArray.append(tableRow)
        #     rowCount = rowCount + 1
        #     cellCount = 0
                    
    print ('Done processing.' + ' Script executed in ' + str(int(time.time()-startTime)) + ' seconds.')
    # END MAIN LOOP

    # Finally, save the workbook
    wb.close()

except: # Main exception handler
    print('The program did not complete.')
    e = sys.exc_info()
    print (e)
   # ctypes.windll.user32.MessageBoxW(0, "Sorry, there was a problem running this program.\n\nFor developer reference:\n\n" + str(e), "The program did not complete :-/", 1)
