
import ctypes # for popup window
import sys # for exception information

try: # Main exception handler

    import requests # for HTTP requests
    from bs4 import BeautifulSoup # for HTML parsing
    import bs4 # for type checking
    import xlsxwriter # for exporting to Excel - need xlsx as over 200k rows of data
    import os # to find user's desktop path
    import time # for adding datestamp to file output

    # Timing the script
    startTime = time.time()

    # Configuration of request variables
    url_SBM_TargetVsAchievement = 'http://sbm.gov.in/sbmreport/Report/Physical/SBM_TargetVsAchievement.aspx'

    # Store output
    outputArray = []

    stateKey = 'ctl00$ContentPlaceHolder1$ddlState'
    stateVal = ''
    districtKey = 'ctl00$ContentPlaceHolder1$ddlDistrict'
    districtVal = ''
    blockKey = 'ctl00$ContentPlaceHolder1$ddlBlock'
    blockVal = ''

    submitKey = 'ctl00$ContentPlaceHolder1$btnSubmit'
    submitVal = 'View Report'

    targetKey = '__EVENTTARGET'
    targetVal = ''

    # __EVENTVALIDATION and __VIEWSTATE are dynamic authentication values which must be freshly updated when making a request.
    eventValKey = '__EVENTVALIDATION'
    eventValVal = ''
    viewStateKey = '__VIEWSTATE'
    viewStateVal = ''

    headerFlag = False

    # Function to return HTML parsed with BeautifulSoup from a POST request URL and parameters.
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
                print ("    Could not load %s page - attempt %s out of %s" % (pagetype, i+1, attempts))

    # Load the default page and scrape the state and authentication values
    initPage = parsePOSTResponse(url_SBM_TargetVsAchievement, 'initial')
    eventValVal = initPage.find('input',{'id':'__EVENTVALIDATION'})['value']
    viewStateVal = initPage.find('input',{'id':'__VIEWSTATE'})['value']
    stateOptions = []
    stateOptionVals = []
    stateSelection = initPage.find('select',{'id':'ctl00_ContentPlaceHolder1_ddlState'})
    stateOptions = stateSelection.findAll('option',{'selected':''})
    for stateOption in stateOptions:
        if 'All State' not in stateOption.text:
            stateOptionVal = stateOption['value']
            stateOptionVals.append(stateOptionVal)

    # Initialise workbook
    todaysDate = time.strftime('%d-%m-%Y')
    desktopFile = os.path.expanduser('~/Desktop/SBM_TargetVsAchievement_' + todaysDate + '.xlsx')
    wb = xlsxwriter.Workbook(desktopFile)
    ws = wb.add_worksheet('SBM Test')
    ws.set_column('A:AZ', 22)
    rowCount = 1 # Adjust one row for printing table headers after main loop
    cellCount = 0

    # Global variable to store final table data
    lastBlockReportTable = ''

    # Global variable for keeping track of the state
    stateCount = 1

    # MAIN LOOP: loop through STATE values and scrape district and authentication values for each
    for stateOptionVal in stateOptionVals: # For testing, we can limit the states processed due to long runtime
        postParams = {
            eventValKey:eventValVal,
            viewStateKey:viewStateVal,
            stateKey:stateOptionVal,
            districtKey:'-1',
            blockKey:'-1',
            targetKey:'ctl00$ContentPlaceHolder1$ddlState'
        }
        statePage = parsePOSTResponse(url_SBM_TargetVsAchievement, postParams, 'state')
        state_eventValVal = statePage.find('input',{'id':'__EVENTVALIDATION'})['value']
        state_viewStateVal = statePage.find('input',{'id':'__VIEWSTATE'})['value']
        districtOptions = []
        districtOptionVals = []
        districtSelection = statePage.find('select',{'id':'ctl00_ContentPlaceHolder1_ddlDistrict'})
        districtOptions = districtSelection.findAll('option',{'selected':''})
        for districtOption in districtOptions:
            if 'All District' not in districtOption.text and 'STATE HEADQUARTER' not in districtOption.text: # We do not want the top level data for the state or state headquarter data
                districtOptionVal = districtOption['value']
                districtOptionVals.append(districtOptionVal)
        # Loop through the DISTRICT values and scrape block and authentication values for each
        districtCount = 1
        for districtOptionVal in districtOptionVals:
            state_postParams = {
                eventValKey:state_eventValVal,
                viewStateKey:state_viewStateVal,
                stateKey:stateOptionVal,
                districtKey:districtOptionVal,
                blockKey:'-1',
                targetKey:'ctl00$ContentPlaceHolder1$ddlDistrict'
            }
            districtPage = parsePOSTResponse(url_SBM_TargetVsAchievement, state_postParams, 'district')
            district_eventValVal = districtPage.find('input',{'id':'__EVENTVALIDATION'})['value']
            district_viewStateVal = districtPage.find('input',{'id':'__VIEWSTATE'})['value']
            blockOptions = []
            blockOptionVals = []
            blockSelection = districtPage.find('select',{'id':'ctl00_ContentPlaceHolder1_ddlBlock'})
            blockOptions = blockSelection.findAll('option',{'selected':''})
            for blockOption in blockOptions:
                if 'All Block' not in blockOption.text: # We do not want the top level data for the block
                    blockOptionVal = blockOption['value']
                    blockOptionVals.append(blockOptionVal)

            # Loop through the BLOCK values and request the report for each
            blockCount = 1
            for blockOptionVal in blockOptionVals:
                block_postParams = {
                    eventValKey:district_eventValVal,
                    viewStateKey:district_viewStateVal,
                    stateKey:stateOptionVal,
                    districtKey:districtOptionVal,
                    blockKey:blockOptionVal,
                    submitKey:submitVal
                }
                blockReport = parsePOSTResponse(url_SBM_TargetVsAchievement, block_postParams, 'block')

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

                    print ('Currently processing: ' + stateNameText + ' (' + str(stateCount) + ' of ' + str(len(stateOptionVals)) + ')' + ' > ' + districtNameText + ' (' + str(districtCount) + ' of ' + str(len(districtOptionVals)) + ')' + ' > ' + blockNameText + ' (' + str(blockCount) + ' of ' + str(len(blockOptionVals)) + ')')

                    # Loop through rows and write data
                    blockReportRows = blockReportTable.find('tbody').findAll('tr') # Only process table body data

                    if not headerFlag:
                        # Write table headers based on final report
                        headerFlag = True
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
                            headerTableRow.append('State Name')
                            headerTableRow.append('District Name')
                            headerTableRow.append('Block Name')
                            for td in headerCols:
                                # Tidy the cell content
                                cellText = td.text.replace('\*','')
                                cellText = cellText.strip()
                                # Store the cell data
                                headerTableRow.append(cellText)
                                cellCount = cellCount+1
                            outputArray.append(headerTableRow)


                    for tr in blockReportRows[0:len(blockReportRows)-1]: # Total row (bottom of table) dropped
                        tableRow = []
                        cols = tr.findAll('td')
                        # Write state, district, and block information
                        tableRow.append(stateNameText)
                        tableRow.append(districtNameText)
                        tableRow.append(blockNameText)
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

                        outputArray.append(tableRow)
                        # Try writing row at once
                        rowCount = rowCount + 1
                        cellCount = 0
                else:
                    print ('No data for: ' + stateNameText + ' (' + str(stateCount) + ' of ' + str(len(stateOptionVals)) + ')' + ' > ' + districtNameText + ' (' + str(districtCount) + ' of ' + str(len(districtOptionVals)) + ')' + ' > block (' + str(blockCount) + ' of ' + str(len(blockOptionVals)) + ')')
                    # Output rows for where data is not found
                    tableRow.append(stateNameText)
                    tableRow.append(districtNameText)
                    tableRow.append(blockNameText)
                    tableRow.append('1')
                    tableRow.append('No Data')
                    for i in range(15):
                        tableRow.append('0')
                    tableRow.append('No Data')

                blockCount = blockCount + 1
            districtCount = districtCount + 1
        stateCount = stateCount + 1



    print ('Done processing.' + ' Script executed in ' + str(int(time.time()-startTime)) + ' seconds.')
    # END MAIN LOOP

    # Write output
    r = 0
    for entry in outputArray:
        ws.write_row(r, 0, entry)
        r = r + 1
    # Finally, save the workbook
    wb.close()

except: # Main exception handler
    try:
        # Write all data into the file
        r = 0
        for entry in outputArray:
            ws.write_row(r, 0, entry)
            r = r + 1

        print("Error occurred, outputted %s rows of data." % r)
        wb.close()
        e = sys.exc_info()
        print(e)
    except:
        print ("Could not output data.")

        print('The program did not complete.')
        e = sys.exc_info()
        print(e)
        ctypes.windll.user32.MessageBoxW(0,
                                         "Sorry, there was a problem running this program.\n\nFor developer reference:\n\n" + str(
                                             e), "The program did not complete :-/", 1)