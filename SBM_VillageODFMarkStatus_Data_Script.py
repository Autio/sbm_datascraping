# Tool to scrape Swachh Bharat Mission data from the sbm.gov.in government website #
# Daniel Robertson                                                                 #
# Petri Autio                                                                      #
# 2016                                                                             #
__author__ = 'petriau'
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
    url_SBM = 'http://sbm.gov.in/sbmreport/Report/Physical/SBM_VillageODFMarkStatus.aspx'

    # Output held here
    outputArray = []


    # For finance progress
    componentKey = 'ctl00$ContentPlaceHolder1$ddlComponent'
    componentVal = ''

    # submit button on first page
    submitKey = 'ctl00$ContentPlaceHolder1$btnSubmit'
    submitVal = 'Submit'

    # __EVENTVALIDATION. __VIEWSTATE and __EVENTTARGET are dynamic authentication values which must be freshly updated when making a request.
    eventValKey = '__EVENTVALIDATION'
    eventVal = ''
    viewStateKey = '__VIEWSTATE'
    viewStateVal = ''
    targetKey = '__EVENTTARGET'
    targetVal = ''
    argumentKey = '__EVENTARGUMENT'
    argumentVal = ''
    stateKey = 'ctl00$ContentPlaceHolder1$ddlState'

    # Function to return HTML parsed with BeautifulSoup from a POST request URL and parameters.
    def parsePOSTResponse(URL, parameters, pagetype):
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

    # Given two dicts, merge them into a new dict as a shallow copy.
    def merge_two_dicts(x, y):
        z = x.copy()
        z.update(y)
        return z

    # Load the default page and scrape the component, finance year and authentication values
    initPage = parsePOSTResponse(url_SBM, "", "initial")
    eventVal = initPage.find('input', {'id': '__EVENTVALIDATION'})['value']
    viewStateVal = initPage.find('input', {'id': '__VIEWSTATE'})['value']
    componentOptions = []
    componentOptionVals = []
    componentSelection = initPage.find('select', {'id': 'ctl00_ContentPlaceHolder1_ddlComponent'})  # changed for link 2
    componentOptions = componentSelection.findAll('option', {'contents': ''})  # changed from selection to contents
    for componentOption in componentOptions:
        if not componentOption.text == "All State":
            componentOptionVal = componentOption['value']
            componentOptionVals.append(componentOptionVal)

    # Initialise workbook
    todaysDate = time.strftime('%d-%m-%Y')
    filePath = '~/Desktop/SBM_VillageODFMarkStatus_' + todaysDate + '.xlsx'
    desktopFile = os.path.expanduser(filePath)
    wb = xlsxwriter.Workbook(desktopFile, {'strings_to_numbers': True})
    ws = wb.add_worksheet('SBM_ODF')
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

    eventVals = []
    print ("Starting data collection")
    # MAIN LOOP: loop through component values. Scrape link values from page
    for componentOptionVal in componentOptionVals[:1]:  # For testing, we can limit the states processed due to long runtime
        eventVal = initPage.find('input', {'id': '__EVENTVALIDATION'})['value']
        eventVals.append(eventVal)
        viewStateVal = initPage.find('input', {'id': '__VIEWSTATE'})['value']
        postParams = {
            eventValKey: eventVal,
            viewStateKey: viewStateVal,
            '__LASTFOCUS': '',
            argumentKey: argumentVal,
            componentKey: componentOptionVal,
            # submitKey: submitVal
        }

        componentPage = parsePOSTResponse(url_SBM, postParams, "component")

        # Now the states become visible, so read them in
        stateOptions = []
        stateOptionVals = []
        stateSelection = componentPage.find('select', {'id': 'ctl00_ContentPlaceHolder1_ddlState'})
        stateOptions = stateSelection.findAll('option', {'selected': ''})
        # Find all states and links to click through
        for s in stateOptions:
            if 'All State' not in s.text:
                stateOptionVal = s['value']
                stateOptionName = s.text
                stateOptionVals.append([stateOptionVal, stateOptionName])
        stateCount = 1
        # Now cycle through the states and use the stateOptionVal to select the state
        state = []
        for state in stateOptionVals:
            districtCount = 0
            eventVal = componentPage.find('input', {'id': '__EVENTVALIDATION'})['value']
            eventVals.append(eventVal)
            viewStateVal = componentPage.find('input', {'id': '__VIEWSTATE'})['value']

            postParams = {
                argumentKey: argumentVal,
                targetKey: '',
                '__LASTFOCUS': '',
                eventValKey: eventVal,
                viewStateKey: viewStateVal,
                submitKey: submitVal,
                componentKey: componentOptionVal,
                stateKey: state[0]
            }

            statePage = parsePOSTResponse(url_SBM, postParams, "state")

            eventVal = statePage.find('input', {'id': '__EVENTVALIDATION'})['value']
            eventVals.append(eventVal)
            viewStateVal = statePage.find('input', {'id': '__VIEWSTATE'})['value']
            postParams = {
                '__EVENTARGUMENT': '',
                targetKey: "ctl00$ContentPlaceHolder1$Reptdist$ctl01$lbldist",
                eventValKey: eventVal,
                viewStateKey: viewStateVal,
            }

            # Process Districts by using name links
            linkOptions = []
            linkOptions2 = []
            linkOptionVals = []
            linkSelection = statePage.findAll('input', {'id': re.compile('hfCode$')})
            linkSelection2 = statePage.findAll('input', {'id': re.compile('hfdtcode$')})
            districtSelection = statePage.findAll('a', {'id':re.compile('lbldist$')})
            districtOptions = []
            linkIndex = 0
            for link in linkSelection:
                linkId = link['id'].replace('_', '$')
                linkOptions.append([linkId, link['value']])
                linkId = linkSelection2[linkIndex]['id'].replace('_', '$')
                linkOptions2.append([linkId, linkSelection2[linkIndex]['value']])
                districtOptions.append([districtSelection[linkIndex]['id'], districtSelection[linkIndex].text])
                # go through both link lists in parallel
                linkIndex = linkIndex + 1

            # write links into dictionary to be passed into POST params
            paramDict = {key: str(value) for key, value in linkOptions}
            paramDict2 = {key: str(value) for key, value in linkOptions2}
            # and merge the dictionaries
            paramDict = merge_two_dicts(paramDict, paramDict2)

            districtCount = 1
            for district in districtOptions:
                eventVal = statePage.find('input', {'id': '__EVENTVALIDATION'})['value']
                viewStateVal = statePage.find('input', {'id': '__VIEWSTATE'})['value']
                postParams = {
                    '__EVENTARGUMENT': '',
                    targetKey: district[0].replace('_','$'),
                    eventValKey: eventVal,
                    viewStateKey: viewStateVal,
                }
                postParams = merge_two_dicts(paramDict, postParams)

                districtPage = parsePOSTResponse(url_SBM, postParams, "district")

                # Then process the numbers next which are the link to the GP level
                # Also goes down to family head name...

                # Pick out values from the numbers
                eventVal = districtPage.find('input', {'id': '__EVENTVALIDATION'})['value']
                viewStateVal = districtPage.find('input', {'id': '__VIEWSTATE'})['value']

                blockSelection = districtPage.findAll('a', {'id': re.compile('BlockTotalGP$')})

                linkOptions = []
                linkOptions2 = []
                linkOptions3 = []
                blockNames = []
                blockNames = districtPage.findAll('span', {'id': re.compile('lblBlock$')})
                linkSelection = districtPage.findAll('input', {'id': re.compile('hfCode$')})
                linkSelection2 = districtPage.findAll('input', {'id': re.compile('hfdtcode$')})
                linkSelection3 = districtPage.findAll('input', {'id': re.compile('hfBlkcode$')})
                linkIndex = 0
                for link in linkSelection:
                    linkId = link['id'].replace('_', '$')
                    linkOptions.append([linkId, link['value']])
                    linkId = linkSelection2[linkIndex]['id'].replace('_', '$')
                    linkOptions.append([linkId, linkSelection2[linkIndex]['value']])
                    linkId = linkSelection3[linkIndex]['id'].replace('_', '$')
                    linkOptions.append([linkId, linkSelection3[linkIndex]['value']])
                    # go through both link lists in parallel
                    linkIndex = linkIndex + 1

                paramDict3 = {key: str(value) for key, value in linkOptions}
                paramDict3 = merge_two_dicts(paramDict3, {key: str(value) for key, value in linkOptions2})
                paramDict3 = merge_two_dicts(paramDict3, {key: str(value) for key, value in linkOptions3})
                blockCount = 0
                if not headerFlag:
                    print ('Processing table headers...')
                for block in blockSelection:
                    blockName = blockNames[blockCount].text
                    blockCount = blockCount + 1

                    if block.text == '0':
                        print ('Currently processing: ' + state[1].upper() + ' > ' + "" + ' (' + str(stateCount) + ' of ' + str(len(stateOptionVals)) + ')' + ' >  ' + district[1].strip() + ' (' + str(districtCount) + ' of ' + str(len(districtOptions)) + ') >  ' + blockName.strip() + ' (' + str(blockCount) + ' of ' + str(len(blockSelection)) + ') - no GP data')
                        tableRow = []
                        tableRow.append('')
                        tableRow.append(state[1])
                        tableRow.append(district[1])
                        tableRow.append(blockName.strip())
                        tableRow.append('GP')
                        tableRow.append('No data')
                        tableRow.append('No data')
                        for i in range(5):
                            tableRow.append('0')
                        tableRow.append('No data')

                        outputArray.append(tableRow)
                    # Only click into block if the total value of blocks is above 0, otherwise it will not go anywhere
                    else:
                        print ('Currently processing: ' + state[1].upper() + ' > ' + "" + ' (' + str(stateCount) + ' of ' + str(len(stateOptionVals)) + ')' + ' >  ' + district[1].strip() + ' (' + str(districtCount) + ' of ' + str(len(districtOptions)) + ') >  ' + blockName.strip() + ' (' + str(blockCount) + ' of ' + str(len(blockSelection)) + ')')
                        postParams = {
                            argumentKey: argumentVal,
                            targetKey: block['id'].replace('_', '$').replace('lnk$', 'lnk_'),
                            '__LASTFOCUS': '',
                            eventValKey: eventVal,
                            viewStateKey: viewStateVal,
                        }
                        postParamsBlock = merge_two_dicts(paramDict3, postParams)
                        blockPage = parsePOSTResponse(url_SBM, postParamsBlock, "block")

                        # Process table data and output
                        reportTable = blockPage.find('table')

                        # Write table headers
                        if not headerFlag:
                            headerRows = reportTable.find('thead').findAll('tr')  # Only process table header data
                            headerTableArray = []
                            rowCount = 0

                            headerStyle = wb.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#0A8AD5'})

                            for tr in headerRows:  # two header rows
                                cellCount = 0
                                headerTableRow = []
                                headerCols = tr.findAll('th')
                                # Write state, district, and block headers
                                for td in headerCols:
                                    # Tidy the cell content
                                    cellText = td.text.replace('\*', '')
                                    cellText = cellText.strip()
                                    # Store the cell data
                                    headerTableRow.append(cellText)
                                    #ws.write(rowCount, cellCount, cellText, headerStyle)
                                    cellCount = cellCount + 1
                                rowCount = rowCount + 1
                                outputArray.append(headerTableRow)

                            headerFlag = True

                        # Write table data
                        if isinstance(reportTable,
                                      bs4.element.Tag):  # Check whether data table successfully found on the page. Some blocks have no data.
                            # Store table for writing headers after loop

                            lastReportTable = reportTable
                            try:
                                reportRows = reportTable.findAll('tr')  # Bring entire table including headers because body isn't specified
                                if not reportRows == None:
                                    if len(reportRows) > 4:

                                        for tr in reportRows[2:len(reportRows) - 1]:  # Start from 2 (body of table), bottom of table dropped
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
                                                cellCount = cellCount + 1
                                            rowCount = rowCount + 1
                                            outputArray.append(tableRow)

                            except TypeError:
                                print ('No data for ' + block['id'])
                            except AttributeError:
                                print ('No data for ' + block['id'])
                districtCount = districtCount + 1
            stateCount = stateCount + 1
        componentCount = componentCount + 1

    # Write all data into the file
    r = 0
    for entry in outputArray:
        ws.write_row(r, 0, entry)
        r = r + 1

    print ('Done processing.' + ' Script executed in ' + str(int(time.time() - startTime)) + ' seconds.')
    print ('File saved in ' + filePath)
    # END MAIN LOOP

    # Finally, save the workbook
    wb.close()

except:  # Main exception handler
    try:
        # Write all data into the file
        r = 0
        for entry in outputArray:
            ws.write_row(r, 0, entry)
            r = r + 1
            wb.close()
            print ('File saved in ' + filePath)
    except:
        print ("Could not output data.")

    print('The program did not complete.')
    e = sys.exc_info()
    print (e)
    ctypes.windll.user32.MessageBoxW(0,
                                     "Sorry, there was a problem running this program.\n\nFor developer reference:\n\n" + str(
                                         e), "The program did not complete :-/", 1)
