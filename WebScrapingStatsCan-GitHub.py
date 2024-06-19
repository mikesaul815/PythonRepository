import os
import sys


# Check if the libraries are installed, if not, install them
try:
    from bs4 import BeautifulSoup
    from openpyxl import Workbook
except ImportError:
    print("Required libraries not found. Installing...")
    os.system('pip install beautifulsoup4 openpyxl')
    print("Libraries installed successfully. Please run the script again.")
    sys.exit()

html_content = """
<table class="pub-table simpleTable dataTable" id="simpleTable">
                <thead id="simpleTableHeader" class="simpleTableHeader"><tr id="head-row-1"><th class="uom-right bold" colspan="2">Geography</th><th class="uom-center bold" style="" colspan="5" id="1_3" headers="">Canada <a href="https://www150.statcan.gc.ca/g1/datatomap/index.html?action=wf_identify&amp;value={'layers':[{'values':['2016A000011124'],'id':'A0000'}]}" target="_blank" title="Opens Canada map in a new tab window.">(map)</a></th></tr><tr id="head-row-2"><th class="sorting" style="text-align:left;vertical-align:bottom;" id="sort_0">Canadian and foreign direct investment<sup><a href="#Footnote1" onclick="notePopup(this); return false;" aria-controls="centred-popup" aria-label="Footnote 1">1</a></sup></th><th class="sorting" style="text-align:left;vertical-align:bottom;" id="sort_1">Countries or regions<sup><a href="#Footnote2" onclick="notePopup(this); return false;" aria-controls="centred-popup" aria-label="Footnote 2">2</a></sup></th><th class="sorting uom-center bold" style="" colspan="1" id="sort_2" headers="1_3">2019 </th><th class="sorting uom-center bold" style="" colspan="1" id="sort_3" headers="1_3">2020 </th><th class="sorting uom-center bold" style="" colspan="1" id="sort_4" headers="1_3">2021 </th><th class="sorting uom-center bold" style="" colspan="1" id="sort_5" headers="1_3">2022 </th><th class="sorting uom-center bold" style="" colspan="1" id="sort_6" headers="1_3">2023 </th></tr></thead>
                <tbody id="simpleTableBody" class="simpleTableBody"><tr id="bufferRowTop" style="height: 0px"></tr><tr class="highlight-row"><th colindex="0_0" id="0_0" class="align-left stub-indent0" headers="2_1" rowspan="9">Canadian direct investment abroad - total book value </th><th></th><th class="uom-center bold" colspan="5">Dollars</th></tr><tr class="highlight-row"><th colindex="1_1" id="1_1" class="align-left stub-indent0" headers="2_2" rowspan="1">All countries<sup><a href="#Footnote3" onclick="notePopup(this); return false;" aria-controls="centred-popup" aria-label="Footnote 3">3</a></sup> </th><td class="align-right nowrap">1,613,703</td><td class="align-right nowrap">1,666,652</td><td class="align-right nowrap">1,844,423</td><td class="align-right nowrap">2,032,530</td><td class="align-right nowrap">2,171,289</td></tr><tr class="highlight-row"><th colindex="2_1" id="2_1" class="align-left stub-indent1" headers="2_2" rowspan="1">North America </th><td class="align-right nowrap">758,369</td><td class="align-right nowrap">823,220</td><td class="align-right nowrap">926,078</td><td class="align-right nowrap">1,046,693</td><td class="align-right nowrap">1,118,498</td></tr><tr class="highlight-row"><th colindex="3_1" id="3_1" class="align-left stub-indent1" headers="2_2" rowspan="1">South and Central America </th><td class="align-right nowrap">80,019</td><td class="align-right nowrap">70,744</td><td class="align-right nowrap">73,480</td><td class="align-right nowrap">82,587</td><td class="align-right nowrap">90,355</td></tr><tr class="highlight-row"><th colindex="4_1" id="4_1" class="align-left stub-indent1" headers="2_2" rowspan="1">Caribbean </th><td class="align-right nowrap">261,300</td><td class="align-right nowrap">258,060</td><td class="align-right nowrap">291,244</td><td class="align-right nowrap">316,409</td><td class="align-right nowrap">322,364</td></tr><tr class="highlight-row"><th colindex="5_1" id="5_1" class="align-left stub-indent1" headers="2_2" rowspan="1">Europe </th><td class="align-right nowrap">378,942</td><td class="align-right nowrap">366,348</td><td class="align-right nowrap">390,223</td><td class="align-right nowrap">425,650</td><td class="align-right nowrap">459,906</td></tr><tr class="highlight-row"><th colindex="6_1" id="6_1" class="align-left stub-indent1" headers="2_2" rowspan="1">Africa </th><td class="align-right nowrap">11,891</td><td class="align-right nowrap">13,623</td><td class="align-right nowrap">13,859</td><td class="align-right nowrap">11,847</td><td class="align-right nowrap">12,032</td></tr><tr class="highlight-row"><th colindex="7_1" id="7_1" class="align-left stub-indent1" headers="2_2" rowspan="1">Asia/Oceania </th><td class="align-right nowrap">122,209</td><td class="align-right nowrap">133,377</td><td class="align-right nowrap">147,929</td><td class="align-right nowrap">148,626</td><td class="align-right nowrap">167,414</td></tr><tr class="highlight-row"><th colindex="8_1" id="8_1" class="align-left stub-indent1" headers="2_2" rowspan="1">Unallocated countries<sup><a href="#Footnote4" onclick="notePopup(this); return false;" aria-controls="centred-popup" aria-label="Footnote 4">4</a></sup> </th><td class="align-right nowrap">974</td><td class="align-right nowrap">1,282</td><td class="align-right nowrap">1,609</td><td class="align-right nowrap">718</td><td class="align-right nowrap">719</td></tr><tr class="highlight-row"><th colindex="9_0" id="9_0" class="align-left stub-indent0" headers="2_1" rowspan="8">Foreign direct investment in Canada - total book value </th><th colindex="9_1" id="9_1" class="align-left stub-indent0" headers="2_2" rowspan="1">All countries<sup><a href="#Footnote3" onclick="notePopup(this); return false;" aria-controls="centred-popup" aria-label="Footnote 3">3</a></sup> </th><td class="align-right nowrap">1,137,849</td><td class="align-right nowrap">1,079,053</td><td class="align-right nowrap">1,216,525</td><td class="align-right nowrap">1,307,874</td><td class="align-right nowrap">1,360,269</td></tr><tr class="highlight-row"><th colindex="10_1" id="10_1" class="align-left stub-indent1" headers="2_2" rowspan="1">North America </th><td class="align-right nowrap">468,760</td><td class="align-right nowrap">465,335</td><td class="align-right nowrap">547,240</td><td class="align-right nowrap">585,714</td><td class="align-right nowrap">621,336</td></tr><tr class="highlight-row"><th colindex="11_1" id="11_1" class="align-left stub-indent1" headers="2_2" rowspan="1">South and Central America </th><td class="align-right nowrap">3,256</td><td class="align-right nowrap">6,181</td><td class="align-right nowrap">7,393</td><td class="align-right nowrap">7,382</td><td class="align-right nowrap">9,325</td></tr><tr class="highlight-row"><th colindex="12_1" id="12_1" class="align-left stub-indent1" headers="2_2" rowspan="1">Caribbean </th><td class="align-right nowrap">38,046</td><td class="align-right nowrap">28,642</td><td class="align-right nowrap">28,543</td><td class="align-right nowrap">30,961</td><td class="align-right nowrap">31,472</td></tr><tr class="highlight-row"><th colindex="13_1" id="13_1" class="align-left stub-indent1" headers="2_2" rowspan="1">Europe </th><td class="align-right nowrap">425,366</td><td class="align-right nowrap">396,138</td><td class="align-right nowrap">431,699</td><td class="align-right nowrap">462,021</td><td class="align-right nowrap">466,595</td></tr><tr class="highlight-row"><th colindex="14_1" id="14_1" class="align-left stub-indent1" headers="2_2" rowspan="1">Africa </th><td class="align-right nowrap">208</td><td class="align-right nowrap">212</td><td class="align-right nowrap">669</td><td class="align-right nowrap">874</td><td class="align-right nowrap">999</td></tr><tr class="highlight-row"><th colindex="15_1" id="15_1" class="align-left stub-indent1" headers="2_2" rowspan="1">Asia/Oceania </th><td class="align-right nowrap">145,189</td><td class="align-right nowrap">129,300</td><td class="align-right nowrap">138,988</td><td class="align-right nowrap">151,308</td><td class="align-right nowrap">160,931</td></tr><tr class="highlight-row"><th colindex="16_1" id="16_1" class="align-left stub-indent1" headers="2_2" rowspan="1">Unallocated countries<sup><a href="#Footnote4" onclick="notePopup(this); return false;" aria-controls="centred-popup" aria-label="Footnote 4">4</a></sup> </th><td class="align-right nowrap">57,025</td><td class="align-right nowrap">53,245</td><td class="align-right nowrap">61,994</td><td class="align-right nowrap">69,615</td><td class="align-right nowrap">69,612</td></tr><tr id="bufferRowBottom" style="height: 0px"></tr></tbody>
            </table>
"""

# Parse the HTML
soup = BeautifulSoup(html_content, 'html.parser')

# Find the table with id="simpleTable"
table = soup.find('table', id='simpleTable')

# Extract table headers
headers = [th.text.strip() for th in table.find('thead').find_all('th')]

# Extract table rows
data = []
for row in table.find('tbody').find_all('tr'):
    row_data = [td.text.strip() for td in row.find_all('td')]
    if row.find('th', class_='align-left'):
        row_data.insert(0, row.find('th', class_='align-left').text.strip())
    data.append(row_data)

# Create a new Excel workbook
wb = Workbook()
ws = wb.active

# Write headers to the first row
ws.append(headers)

# Write data to subsequent rows
for row in data:
    ws.append(row)

# Save the workbook to the specified path
excel_path = r"filepath\WebScrapingStatsCan.xlsx"
wb.save(excel_path)

print("Data successfully written to Excel file:", excel_path)
