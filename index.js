const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

(async () => {
    const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();

    // Replace the URL with the target login page
    const url = '';

    await page.goto(url);

    // Replace the following with the correct selectors and values
    const username = '';
    const password = '';

    await page.type('#Email', username);
    await page.type('#Password', password);
    await page.click('#btn_submit_login');
    await page.waitForNavigation();

    console.log('Logged in successfully.');
    // Function to select values from select2-enabled dropdowns
    const setSelect2Value = async (page, selector, monthValue) => {
        await page.evaluate(({ selector, monthValue }) => {
            const selectElement = document.querySelector(selector);
            if (!selectElement) {
                return { success: false, message: `Element ${selector} not found.` };
            }
    
            // Ensure the select2 dropdown is initialized
            if (!selectElement.classList.contains('select2-hidden-accessible')) {
                return { success: false, message: `Element ${selector} is not select2-enabled.` };
            }
    
            // Find the span with the 'select2-selection__rendered' class
            const renderedSpan = selectElement
                .closest('.input-div')
                ?.querySelector('.select2-selection__rendered');
    
            if (!renderedSpan) {
                return { success: false, message: `Rendered span not found for selector ${selector}.` };
            }
    
            // Update the span's text content
            renderedSpan.textContent = monthValue;
            renderedSpan.setAttribute('title', monthValue); // Update the title for consistency
    
            // Update the underlying select element's value and trigger events
            const matchingOption = Array.from(selectElement.options).find(
                (option) => option.text.trim() === monthValue
            );
    
            if (matchingOption) {
                selectElement.value = matchingOption.value;
            } else {
                return { success: false, message: `No matching option found for "${monthValue}".` };
            }
    
            // Trigger change and select2-specific events
            const changeEvent = new Event('change', { bubbles: true });
            selectElement.dispatchEvent(changeEvent);
    
            const select2Event = new Event('select2:select', { bubbles: true });
            selectElement.dispatchEvent(select2Event);
        }, { selector, monthValue });
    };
    
    // Usage
    await setSelect2Value(page, '#Period_From_Month', 'December');    
    //await page.select('#Period_From_Year', '2024');
    await setSelect2Value(page, '#Period_To_Month', 'December');
    //await page.select('#Period_To_Year', '2024');
    // await page.type('#Mode','SEA');
    await page.select('#Country', 'Aruba');
    await page.click('#btn_Data_Search');
    await page.waitForSelector('#grid'); // Wait for the table to load

    console.log('Form submitted. Fetching table data...');

    // Save Results to Excel
    const workbook = new ExcelJS.Workbook();

    const getOrCreateWorksheet = (workbook, sheetName) => {
        // List current worksheets for debugging
        console.log('Existing Worksheets:', workbook.worksheets.map(ws => ws.name));
    
        // Try to get the existing worksheet
        let worksheet = workbook.getWorksheet(sheetName);
        if (worksheet) {
            console.log(`Worksheet "${sheetName}" already exists. Clearing rows for update.`);
            // Clear all rows in the worksheet
            worksheet.spliceRows(1, worksheet.rowCount); // Clear all rows starting from the first
        } else {
            console.log(`Creating new worksheet "${sheetName}".`);
            // Add the new worksheet
            worksheet = workbook.addWorksheet(sheetName);
        }
    
        return worksheet;
    };

    // Usage
    const sheetName = 'Aruba_November_2024';
    const worksheet = getOrCreateWorksheet(workbook, sheetName);

    // Extract Header Row (First Page Only)
    const extractHeaderRow = async () => {
        const headerRow = await page.evaluate(() => {
            const headers = Array.from(document.querySelectorAll('#grid thead th')); // Replace '#resultsTable thead th' with the actual selector
            return headers.map(header => header.textContent.trim());
        });
        worksheet.addRow(headerRow);
    };

    await extractHeaderRow();

    // Function to extract data from the table on the current page
    const extractTableData = async () => {
        return await page.evaluate(() => {
            const rows = Array.from(document.querySelectorAll('#grid tbody tr'));

            return rows.map(row => {
                const cells = Array.from(row.querySelectorAll('td'));
                return cells.map(cell => cell.textContent.trim());
            });
        });
    };

    // Function to get the last row's SNo value
    const getLastSNo = async () => {
        return await page.evaluate(() => {
            const rows = document.querySelectorAll('#grid tbody tr');
            const lastRow = rows[rows.length - 1];
            return lastRow ? parseInt(lastRow.querySelector('td:first-child').textContent.trim(), 10) : null;
        });
    };

    // Function to check if the next page has loaded
    const waitForNextPage = async (expectedSNo) => {
        await page.waitForFunction(
            (expected) => {
                const firstRowSNo = document.querySelector('#grid tbody tr td:first-child');
                return firstRowSNo && parseInt(firstRowSNo.textContent.trim(), 10) === expected;
            },
            { polling: 'mutation', timeout: 10000 }, // Poll for changes and set a timeout of 10 seconds
            expectedSNo
        );
    };

    // Function to check if a "Next" button exists and click it
    const goToNextPage = async () => {
        const lastSNo = await getLastSNo();
        console.log('Last SNo on the current page:', lastSNo);

        // const pagination = await page.$('#Table_Pagging'); // Select the pagination container
        // if (!pagination) return false;
    
        // Find the "Next" button's <a> tag inside the pagination container
        const nextButtonExists = await page.evaluate(() => {
            const ul = document.querySelector('#Table_Pagging');
            const nextLi = Array.from(ul.querySelectorAll('li')).find(li => li.textContent.trim() === 'Next');
            if (nextLi) {
                const nextAnchor = nextLi.querySelector('a');
                if (nextAnchor) {
                    nextAnchor.click(); // Perform the click action
                    return true; // Indicate that the Next button was clicked
                }
            }
            return false; // Indicate that the Next button was not found
        });
    
        if (nextButtonExists) {
            console.log('Next button clicked. Waiting for the next page to load...');
            await waitForNextPage(lastSNo + 1);
            return true;
        }
    
        return false; // Return false if the "Next" button doesn't exist
    };

    // Function to scrape data from a paginated table when 'Next' and 'Last' buttons are unavailable
    const scrapeTableData = async () => {
        console.log("Checking for pagination...");
    
        // Check if pagination exists
        const paginationExists = await page.evaluate(() => {
            const pagination = document.querySelector('#Table_Pagging');
            return !!pagination; // Return true if pagination exists
        });
    
        if (!paginationExists) {
            console.log("No pagination found. Scraping data on the current page...");
            const data = await extractDataFromTable();
            return data; // Return data if there's no pagination
        }
    
        console.log("Pagination detected. Calculating total pages...");
        
        // Extract total number of pages
        const totalPages = await page.evaluate(() => {
            const pagination = document.querySelector('#Table_Pagging');
            if (!pagination) return 0;
            return Array.from(pagination.querySelectorAll('li'))
                .map(li => li.textContent.trim())
                .filter(text => !isNaN(Number(text))) // Keep only numeric page numbers
                .map(Number)
                .sort((a, b) => a - b) // Sort page numbers
                .pop(); // Get the last (maximum) page number
        });
    
        console.log(`Total pages detected: ${totalPages}`);
    
        let allData = [];
        let currentPage = 1;
    
        while (currentPage <= totalPages) {
            console.log(`Navigating to page ${currentPage}...`);
            
            // Navigate to the specific page
            await page.evaluate(pageNumber => {
                const pagination = document.querySelector('#Table_Pagging');
                if (!pagination) return;
                const pageLink = Array.from(pagination.querySelectorAll('li'))
                    .find(li => li.textContent.trim() === pageNumber.toString());
                if (pageLink) {
                    const anchor = pageLink.querySelector('a');
                    if (anchor) anchor.click(); // Click the page link
                }
            }, currentPage);
    
            // Wait for the data to load on the current page
            console.log(`Waiting for data on page ${currentPage}...`);
            await page.waitForFunction(
                (expectedPage) => {
                    const pagination = document.querySelector('#Table_Pagging');
                    if (!pagination) return false;
                    const activePage = Array.from(pagination.querySelectorAll('li.active')).find(li => 
                        li.textContent.trim() === expectedPage.toString()
                    );
                    return !!activePage; // Wait until the current page is active
                },
                { polling: 'mutation', timeout: 10000 }, // Adjust timeout as needed
                currentPage
            );
    
            // Extract data from the current page
            const pageData = await extractDataFromTable();
            allData = allData.concat(pageData);
            currentPage++;
        }
        return allData;
    };
    
    const extractDataFromTable = async () => {
        return await page.evaluate(() => {
            const table = document.querySelector('#grid'); // Adjust selector as needed
            if (!table) return [];
    
            const rows = Array.from(table.querySelectorAll('tbody tr'));
            return rows.map(row => {
                const cells = Array.from(row.querySelectorAll('td'));
                return cells.map(cell => cell.textContent.trim());
            });
        });
    };
    
    // Find the "Last" button and get its resource value
    const lastResource = await page.$eval('#Table_Pagging', ul => {
        const lastLi = Array.from(ul.querySelectorAll('li')).find(li => li.textContent.trim() === 'Last');
        return lastLi ? lastLi.getAttribute('resource') : null;
    });

    console.log('Last Resource:', lastResource);

    if(!lastResource) {
        console.log('Last button not found. Unable to determine the total number of pages.');
        const tableData = await scrapeTableData();
        if (tableData.length > 0) {
            tableData.forEach(row => {
                worksheet.addRow(row); // Save data to Excel
            });
        }
    }
    
    const isLastPage = async () => {
        const pagination = await page.$('#Table_Pagging'); // Select the pagination container
        if (!pagination) return true; // If no pagination exists, assume it's the last page
    
        // Find the active page and get its resource value
        const currentResource = await page.$eval('#Table_Pagging', ul => {
            const activeLi = ul.querySelector('li.active');
            return activeLi ? activeLi.getAttribute('resource') : null;
        });
    
        console.log('Current Resource:', currentResource);
    
        // Compare current page resource with the last page resource
        return currentResource === lastResource;
    };

    // Function to navigate directly to a specific page by index
    const goToPageByIndex = async (pageIndex) => {
        console.log(`Navigating directly to page ${pageIndex}...`);
        const lastSNo = await getLastSNo();
    
        // Execute the navigation logic on the browser's context
        const pageClicked = await page.evaluate((index) => {
            const ul = document.querySelector('#Table_Pagging');
            if (!ul) {
                return { success: false, message: 'Pagination container #Table_Pagging not found.' };
            }
    
            const targetLi = Array.from(ul.querySelectorAll('li')).find(li => li.textContent.trim() === index.toString());
            if (!targetLi) {
                return { success: false, message: `Page index ${index} not found in pagination.` };
            }
    
            const pageAnchor = targetLi.querySelector('a');
            if (pageAnchor) {
                pageAnchor.click();
                return { success: true, message: `Page ${index} clicked successfully.` };
            }
    
            return { success: false, message: `Anchor tag for page ${index} not found.` };
        }, pageIndex);
    
        if (pageClicked.success) {
            console.log(pageClicked.message);
            console.log(`Page ${pageIndex} clicked. Waiting for the page to load...`);

            await waitForNextPage(lastSNo + 1);
            return true;
        }
    
        console.error(pageClicked.message);
        return false;
    };    

    // Loop through all pages and extract data
    let pageIndex = 1;
    while (pageIndex <= lastResource) {
        console.log(`Extracting data from page ${pageIndex}...`);
        const tableData = await extractTableData(); // Your function to extract table data
        if (tableData.length > 0) {
            tableData.forEach(row => {
                worksheet.addRow(row); // Save data to Excel
            });
        }
    
        const lastPage = await isLastPage();
        console.log('lastPage:', lastPage);
        if (lastPage) break; // Stop loop if the last page is reached

        const hasNextPage = await goToNextPage(); // Move to the next page
        console.log('hasNextPage:', hasNextPage);
        if (!hasNextPage) {
            console.log('Next button not found. Attempting direct navigation...');
            const navigated = await goToPageByIndex(pageIndex + 1);
            if (!navigated) {
                console.log('Unable to navigate to the next page. Stopping execution.');
                break;
            } else {
                pageIndex++;
            }
        } else {
            pageIndex++;
        }
    }

    // Save Results to Excel
    const fileName = 'paginated-table-data.xlsx';
    await workbook.xlsx.writeFile(fileName);

    console.log(`Table data from all pages saved to ${fileName}`);

    // Close Browser
    await browser.close();
})();
