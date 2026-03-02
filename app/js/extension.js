// Initialize Zoho Extension
let zApp = null;
let organizationID = null;
let invoiceID = null;
let lastInvoiceDataStr = ""; // For change detection
let expectedTaxId = null; // Stores the tax ID we determined should apply

// Standard Debounce Function
function debounce(func, wait) {
    let timeout;
    return function (...args) {
        const context = this;
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(context, args), wait);
    };
}

// Helper logging
function log(msg) {
    console.log(msg);
    const logArea = document.getElementById('logArea');
    logArea.style.display = 'block';
    const div = document.createElement('div');
    div.innerText = `> ${msg}`;
    logArea.appendChild(div);
    logArea.scrollTop = logArea.scrollHeight;
}

function setStatus(msg, type) {
    const el = document.getElementById('statusMessage');
    el.innerText = msg;
    el.className = `status ${type}`; // info, success, error
}

window.onload = function () {
    ZFAPPS.extension.init().then(function (App) {
        zApp = App;
        log("Extension Initialized");
        
        // Try to get organization ID from context if available, otherwise fetch from invoice
        
        ZFAPPS.get('organization').then(function(res){
            organizationID = res.organization.organization_id;
            log("Org ID: " + organizationID);
        }).catch(err => log("Info: Could not get Org ID directly (This is expected in some views)"));

        ZFAPPS.get('invoice').then(function(res){
            invoiceID = res.invoice.invoice_id;
            log("Invoice ID: " + invoiceID);
            
            // Set initial state for polling
            if (res && res.invoice) {
                lastInvoiceDataStr = JSON.stringify(res.invoice);
                initPolling();
            }

        }).catch(err => log("Info: Could not get Invoice ID directly (Run 'Debug' to test without invoice)"));


        document.getElementById('lookupBtn').addEventListener('click', runTaxLookup);
        document.getElementById('debugBtn').addEventListener('click', debugWebhook);
    });
};

async function debugWebhook() {
    log("Starting Debug Webhook...");
    const payload = {
                "destination_address": "100 Cedar Lane",
                "destination_city": "Calpine",
                "destination_zip": "96124",
                "destination_state": "CA",
                "origin_zip": "75001",
                "origin_state": "TX",
                "order_subtotal": 100.00,
                "shipping_fee": 15.00
                };

    try {
     


        let options = {
            api_configuration_key :'ac__com_yabd6a_zohobook_taxes',
            url_param: {orgID: organizationID}, // If your API Config expects URL params
        };
            ZFAPPS.request(options).then(response => {
            log("Connection Response: " + JSON.stringify(response));
        }).catch(err => {
            log("API Config Request Failed.");
            console.error(err);
        });

    } catch (e) {
        log("Debug Error: " + e.message);
        console.error(e);
    }
}

async function runTaxLookup() {
    const btn = document.getElementById('lookupBtn');
    btn.disabled = true;
    setStatus("Fetching Invoice Data...", "info");
    document.getElementById('logArea').innerHTML = ''; // Clear logs

    try {
        // 1. GET INVOICE DATA
        const responseData = await ZFAPPS.get('invoice');
        const invoiceDetails = responseData.invoice;
  
        if(!organizationID) {
            organizationID = invoiceDetails.organization_id; 
        }

        // 2. EXTRACT CUSTOM FIELDS & STANDARD FIELDS
        const customFieldMap = {};
        if (invoiceDetails.custom_fields && invoiceDetails.custom_fields.length > 0) {
            log(`Found ${invoiceDetails.custom_fields.length} Custom Fields`);
            
            // DEBUG: Log the first field structure to help debug
            console.log("First CF Structure:", JSON.stringify(invoiceDetails.custom_fields[0]));

            invoiceDetails.custom_fields.forEach(cf => {
                // Try label, then placeholder, then api_name
                // Trim logic to ensure safe matching
                const label = (cf.label || cf.placeholder || cf.api_name || "").trim();
                
                // Some versions use 'value', some might use 'value_formatted' or just be empty if not committed
                const val = (cf.value !== undefined && cf.value !== null) ? cf.value : "";
                
                customFieldMap[label] = val;
                // Also store by API name just in case
                if(cf.api_name) customFieldMap[cf.api_name] = val;
                
                // Log if we think it's one of our target fields
                if(label.includes("Destination") || label.includes("Origin")) {
                    log(`Field [${label}]: '${val}'`);
                }
            });
        } else {
             log("No Custom Fields found in Invoice Data");
        }
        
        // Debug available keys
        console.log("Available CF Keys: ", Object.keys(customFieldMap));
        
        // Helper to find value by variations (Label, Placeholder, API Name, Fuzzy Match)
        const getValue = (searchKey) => {
            // 1. Direct match (e.g. "Destination Zip")
            if (customFieldMap[searchKey] !== undefined) return customFieldMap[searchKey];

            // 2. Direct match with asterisk (e.g. "Destination Zip*")
            if (customFieldMap[searchKey + "*"] !== undefined) return customFieldMap[searchKey + "*"];

            // 3. API Name Fuzzy Match
            // Convert "Destination Zip" -> "destination_zip"
            // Look for keys ending in "_destination_zip" or equal to "destination_zip"
            const normalized = searchKey.toLowerCase().replace(/ /g, '_');
            const mapKeys = Object.keys(customFieldMap);
            
            for (const key of mapKeys) {
                const lowerKey = key.toLowerCase();
                // Check if key is exactly the normalized version or ends with "_normalizedVersion"
                // e.g. "cf_destination_zip" ends with "_destination_zip" is false, but includes "destination_zip"
                // e.g. "cf__com_yabd6a_destination_zip"
                if (lowerKey === normalized || lowerKey.endsWith("_" + normalized) || lowerKey.includes(normalized)) {
                    log(`Fuzzy match: '${searchKey}' mapped to '${key}'`);
                    return customFieldMap[key];
                }
            }
            
            return "";
        };

        const destAddress = getValue("Destination Address");
        const destCity = getValue("Destination City");
        const destState = getValue("Destination State");
        const destZip = getValue("Destination Zip");
        const originState = getValue("Origin State");
        const originZip = getValue("Origin Zip");
        
        // Log detected values for confirmation
        log(`Dest Address: ${destAddress}`);
        log(`Dest City: ${destCity}`);
        log(`Dest State: ${destState}`);
        log(`Dest Zip: ${destZip}`);
        log(`Origin State: ${originState}`);
        log(`Origin Zip: ${originZip}`);
        
        const subTotal = invoiceDetails.sub_total || 0;
        const shipping = invoiceDetails.shipping_charge || 0;

        // Validation
        if (!destZip) {
            ZFAPPS.invoke('SHOW_MESSAGE', { type: 'warning', content: 'Please enter Destination Zip' });
            throw new Error("Destination Zip is missing.");
        }

        // 3. CALL EXTERNAL WEBHOOK
        setStatus("Calling OnSite Storage Webhook (via API Config)...", "info");
        
        const payload = {
            destination_address: destAddress,
            destination_city: destCity,
            destination_zip: destZip,
            destination_state: destState,
            origin_zip: originZip,
            origin_state: originState,
            order_subtotal: subTotal,
            shipping_fee: shipping
        };

        let data = null;

        try {
            // Using ZFAPPS.request() style for API Configurations
            const options = {
                api_configuration_key: 'ac__com_yabd6a_n8n_webhook',
                body:{
                    mode: 'raw',
                    raw: JSON.stringify(payload)
                }
            };

            

            log("Requesting via Connection: webhook");
            // NOTE: We use ZFAPPS.request, but standard widgets usually expect 'connection_link_name' in options
            // NOT 'api_configuration_key'.
            const connResponse = await ZFAPPS.request(options);
            
            // ZFAPPS.request usually returns the body directly or wrapped in { data: ... }
            let responseBody = connResponse.data || connResponse.body || connResponse;
            
            if (typeof responseBody === 'string') {
                 try {
                     responseBody = JSON.parse(responseBody);
                 } catch(e) { /* keep as string */ }
            }
            
            if(!responseBody) throw new Error("Unexpected response format from Webhook. Missing 'body' field.");
            data = JSON.parse(responseBody.body)[0];

        } catch (connErr) {
            log("API Config Request Failed.");
            throw new Error("Failed to invoke API Config 'webhook'. Check console logs.");
        }

        if (!data) throw new Error("No data received from Webhook");

        const targetCode = data.code;

        // 5. FIND TAX BY CODE (Pagination)
        setStatus(`Searching for Tax Code '${targetCode}'...`, "info");
        const finalTaxId = await findTaxIdByCode(targetCode);

        if (finalTaxId) {
             log(`Found Tax ID: ${finalTaxId}`);
             expectedTaxId = finalTaxId; // Store for consistency checks
             
             console.log(invoiceDetails.line_items);
             // Apply Update to Line Items individually
             setStatus("Updating Invoice Line Items...", "info");
             // await ZFAPPS.set('invoice.tax_id', finalTaxId) <-- Removed global set

             let updatedCount = 0;
             if (invoiceDetails.line_items && Array.isArray(invoiceDetails.line_items)) {
                // Final Strategy: Iterative update per item index.
                // Works around:
                // 1. "Invalid Property" (sub-field access)
                // 2. "Appended Duplicates" (bulk array set)
                // 3. "Stuck" (full invoice set)
                
                for (let i = 0; i < invoiceDetails.line_items.length; i++) {
                     try {
                         // Shallow copy item to avoid mutating original source immediately (though fine here)
                         const item = { ...invoiceDetails.line_items[i] };
                         item.tax_id = finalTaxId;
                         
                         // Set the entire item object at the specific index
                         // This forces replacement of the item at that index
                         ZFAPPS.set(`invoice.line_items[${i}]`, item).then(() => {
                             log(`Updated item ${i} with tax_id ${finalTaxId}`);
                         }).catch(err => {
                             log(`Failed to update item ${i}: ${err.message}`);
                         });
                         updatedCount++;
                     } catch (err) {
                         log(`Failed to update item ${i}: ${err.message}`);
                     }
                }
             }
             
             setStatus(`Success! Applied Tax Code: ${targetCode}`, "success");
             ZFAPPS.invoke('SHOW_MESSAGE', { type: 'success', content: `Tax updated to ${targetCode} for ${updatedCount} items` });

        } else {
             throw new Error(`Tax Code '${targetCode}' not found in Zoho Books settings.`);
        }

    } catch (e) {
        console.error(e);
        const errMsg = e.message || e;
        setStatus("Error: " + errMsg, "error");
        ZFAPPS.invoke('SHOW_MESSAGE', { type: 'error', content: errMsg });
    } finally {
        btn.disabled = false;
    }
}

async function findTaxIdByCode(targetCode) {
    if (!zApp || !organizationID) {
        log("Cannot search taxes: Missing App or OrgID");
        return null;
    }

    let stopPaging = false;
    for (let pageNum = 1; pageNum <= 30; pageNum++) {
        if (stopPaging) break;
        
        log(`Searching Taxes Page ${pageNum}...`);
        
        let taxesList = [];
        let hasMore = false

        try {
            // Using ZFAPPS.request (SDK v2) to call the connection
    

    

        let options = {
            api_configuration_key :'ac__com_yabd6a_zohobook_taxes',
            url_param: {orgID: organizationID,pageIndex: pageNum}, // If your API Config expects URL params

        };
            const response =  await  ZFAPPS.request(options)

            // ZFAPPS.request returns { data: ..., status: ... } or the body directly
            let body = response.data ;
             if(response.code !== 0 && response.code !== "0") {
                log(`API Error: ${body.message}`);
                break;
            }
            if (typeof body.body=== 'string') {
                 try { body = JSON.parse(body.body); } catch(e) {       console.error(e);break;}
            }
            // Check for API errors in body
            if(body.code !== 0 && body.code !== "0") {
                log(`API Error: ${body.message}`);
            } else {
                 taxesList = body.taxes || [];
                 if (body.page_context) {
                     hasMore = body.page_context.has_more_page;
                 }
            }

        } catch (err) {
            console.error(err);
            // In a real scenario we might re-throw, but here we break.
            break;
        }

        // Loop taxes in this page
        for (let tax of taxesList) {
            const taxName = tax.tax_name; // e.g., "Tax Name, Code"
            if (taxName && taxName.includes(",")) {
                const parts = taxName.split(",");
                if (parts.length >= 2) {
                    const existingCode = parts[1].trim(); 
                    if (existingCode.includes(targetCode)) { // Deluge was doing `contains`
                         log(`MATCH FOUND! Name: ${taxName} | ID: ${tax.tax_id}`);
                         return tax.tax_id;
                    }
                }
            }
        }

        if (!hasMore) {
            stopPaging = true;
        }
    }

    return null;
}

/* -------------------------------------------------------------------------
   POLLING & DEBOUNCE LOGIC
   Triggered by initPolling() in window.onload
   ------------------------------------------------------------------------- */

// Debounced Handler: Runs only after 1s of stability
const handleInvoiceChange = debounce((currentInvoice) => {
    log("Debounced Change Processed: Auto-Running Tax Lookup...");
    runTaxLookup(); 
}, 1000);

function initPolling() {
    log("Started polling for field changes (Interval: 2s)");
    setInterval(async () => {
        try {
            const res = await ZFAPPS.get('invoice');
            if(res && res.invoice) {
                checkForChanges(res.invoice);
                // Also validate taxes on every poll
                checkTaxConsistency(res.invoice);
            }
        } catch(e) { /* Silent catch */ }
    }, 2000);
}

function checkForChanges(currentInvoice) {
    // Clone and cleanup to avoid infinite loops and over-triggering
    const checkObj = { ...currentInvoice };
    // User requested: do not trigger on line item changes (prevents loops when tax updates)
    delete checkObj.line_items; 
    
    // Also ignore calculated totals that might change when we apply tax
    delete checkObj.tax_total; 
    delete checkObj.total; 
    delete checkObj.sub_total; 
    delete checkObj.shipping_charge; 
    delete checkObj.adjustment;
    
    const currentStr = JSON.stringify(checkObj);
    
    if (lastInvoiceDataStr && currentStr !== lastInvoiceDataStr) {
        log("Change Detected!");
        lastInvoiceDataStr = currentStr;
        handleInvoiceChange(currentInvoice); 
    } else if (!lastInvoiceDataStr) {
        lastInvoiceDataStr = currentStr;
    }
}

function checkTaxConsistency(invoice) {
    // Only check if we have performed a lookup and have an expected tax ID
    if (!expectedTaxId) {
        document.getElementById('warningArea').style.display = 'none';
        return;
    }
    
    const items = invoice.line_items || [];
    let conflictFound = false;
    
    for (const item of items) {
        // Compare tax_id (be careful with type strings vs numbers)
        // If item doesn't have tax_id, it might be non-taxable or empty.
        // Assuming strict match requirement: item.tax_id must equal expectedTaxId.
        
        // Also safeguard against null/undefined
        const currentTaxId = item.tax_id ? String(item.tax_id) : "";
        const expectedStr = String(expectedTaxId);
        
        if (currentTaxId !== expectedStr) {
            conflictFound = true;
            break;
        }
    }
    
    const warnEl = document.getElementById('warningArea');
    if (conflictFound) {
        warnEl.style.display = 'block';
    } else {
        warnEl.style.display = 'none';
    }
}
