// Initialize Zoho Extension
let zApp = null;
let organizationID = null;
let invoiceID = null;
let lastInvoiceDataStr = ""; // For change detection
let expectedTaxId = null; // Stores the tax ID we determined should apply
let lastCustomerId = null; // Track customer changes to trigger address copy

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

async function fetchContactDetails(contactId) {
    if (!contactId || !organizationID) return null;
    
    // Note: The API config key uses the ID from the user request
    const options = {
        api_configuration_key: 'ac__com_yabd6a_zoho_book_customer_contact',
        url_param: { 
            customerID: contactId,
            orgID: organizationID 
        }
    };
    
    try {
        const response = await ZFAPPS.request(options);
        // ZFAPPS.request returns { data: ..., status: ... } or the body directly
        let body = response.data || response.body; 
        if (typeof body.body === 'string') {
             try { body = JSON.parse(body.body); } catch(e) {}
        }
        
        if (body.code === 0 || body.code === "0") {
             return body.contact;
        } else {
             log(`Contact API Error: ${body.message}`);
        }
    } catch (e) {
        log(`Error fetching contact details: ${e.message}`);
    }
    return null;
}

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

        // --- NEW: Address Autofill Logic ---
        // If Customer Changed OR (Destination Zip is empty AND we have shipping address), autoset.
        // We use shipping address by default, fallback to billing? Usually shipping matters for tax.
        
        let didUpdateAddress = false;
        const currentCustId = invoiceDetails.customer_id;
        
        // Condition: Customer ID Changed from last run (and not initial null run if desired? No, even initial might need fill)
        // OR Fields are empty.
        // 'lastCustomerId' tracks the ID from the LAST SUCCESSFUL RUN.
        // If this is the FIRST run (lastCustomerId is null), we populate if fields are empty.
        // If this is a CHANGE run (lastCustomerId != current), we overwrite fields even if filled?
        // User asked: "when the user select customer... put it to destination fields".
        // This implies overwrite on selection.
        
        const sourceAddress = invoiceDetails.shipping_address || invoiceDetails.billing_address;
        console.log(invoiceDetails)
        // --- CUSTOMER CHECK LOGIC FIX ---
        // If customer ID changed from initial load OR last run.
        // If lastCustomerId is null, it means first run.
        // We should populate if fields are empty OR if customer just changed.
        // However, user said "check customer is not working".
        // Maybe it's not detecting the change properly if we rely on lastCustomerId.
        
        const isNewCustomer = currentCustId && (currentCustId !== lastCustomerId);
        console.log("Customer Check:", { currentCustId, lastCustomerId, isNewCustomer });
        // Also check if we HAVE a source address to copy.
        const canAutofill = sourceAddress && (Object.keys(sourceAddress).length > 0);
        
        console.log("Customer ID Check:", { currentCustId, lastCustomerId, isNewCustomer, canAutofill, sourceAddress });
        if (isNewCustomer) {
             let sourceAddress = null;
             
             log(`Customer Change Detected (${currentCustId}). Fetching details...`);
             
             // Fetch from API
             const contactDetails = await fetchContactDetails(currentCustId);
             
             if (contactDetails) {
                 // Prioritize Shipping Address, then Billing Address, then root address
                 sourceAddress = contactDetails.shipping_address || contactDetails.billing_address || contactDetails;
                 log("Fetched Contact Address from API successfully.");
             } else {
                 // Fallback to invoice object if API fails or returns no data
                 sourceAddress = invoiceDetails.shipping_address || invoiceDetails.billing_address;
                 log("API fetch failed/empty. Using Invoice Address fallback.");
             }
             
             if (sourceAddress && Object.keys(sourceAddress).length > 0) {
                 log("Attempting Address Autofill...");

                 // Helper maps needed for finding index
             const findIndex = (searchKey) => {
                 if (!invoiceDetails.custom_fields) return -1;
                 const normalized = searchKey.toLowerCase().replace(/ /g, '_');
                 for (let i=0; i < invoiceDetails.custom_fields.length; i++) {
                     const cf = invoiceDetails.custom_fields[i];
                     const label = (cf.label || cf.placeholder || cf.api_name || "").trim().toLowerCase();
                     const apiName = (cf.api_name || "").toLowerCase();
                     if (label === normalized || label.includes(normalized) || apiName.includes(normalized)) {
                         return i;
                     }
                 }
                 return -1;
             };
             
             const idxAddr = findIndex("Destination Address");
             const idxCity = findIndex("Destination City");
             const idxState = findIndex("Destination State");
             const idxZip = findIndex("Destination Zip");
             
             // Note: sourceAddress keys are usually lowercase in API (address, city, state, zip/zipcode)
             // Check keys carefully or check values
             // Contact API Shipping Address structure (from screenshots): { address: "", city: "", state: "", zip: "", country: "", ... }
             // Invoice object structure: { address: "...", city: "...", ... }
             // They are compatible.

             const addrVal = sourceAddress.address || sourceAddress.street || "";
             const cityVal = sourceAddress.city || "";
             
             // Note: API returns 'state' or 'state_code'. Prefer code if available.
             // Screenshot shows: "state_code": "CA", "state": "CA", "country": "U.S.A.".
             
             let stateVal = sourceAddress.state_code || sourceAddress.state || "";
             
             // Check if it's a full name in our mapping and convert to abbr if so.
             if (window.ADDRESS_MAPPINGS) {
                 // Try to find the exact state name (case-insensitive done by helper if needed, but simple map lookup first)
                 const abbr = window.getAbbreviation ? window.getAbbreviation('states', stateVal) : stateVal;
                 if (abbr) stateVal = abbr;
             }

             const zipVal = sourceAddress.zip || sourceAddress.zipcode || "";
             
             // Check for country if present and needed
             // For now just focus on state as user requested "country and states"
             let countryVal = sourceAddress.country_code || sourceAddress.country || "";
             if (window.ADDRESS_MAPPINGS && countryVal) {
                 const abbr = window.getAbbreviation ? window.getAbbreviation('countries', countryVal) : countryVal;
                 if (abbr) countryVal = abbr;
             }

             const updates = [];
             if (idxAddr >= 0) updates.push({ i: idxAddr, v: addrVal });
             if (idxCity >= 0) updates.push({ i: idxCity, v: cityVal });
             if (idxState >= 0) updates.push({ i: idxState, v: stateVal });
             if (idxZip >= 0) updates.push({ i: idxZip, v: zipVal });
             
             // If we want to support country, we need to find its index too.
             // But user didn't explicitly ask to map country to a field, just mentioned country is full name.
             // However, maybe destination country is relevant?
             
             // Also look for Country field
             const idxCountry = findIndex("Destination Country");
             
             if (idxCountry >= 0) updates.push({ i: idxCountry, v: countryVal });
             
             if (updates.length > 0) {
                 log(`Autofilling ${updates.length} destination fields from Customer Address...`);
                 
                 // Apply updates
                 for (const u of updates) {
                    
                     await ZFAPPS.set(`invoice.custom_fields.${u.i}.value`, u.v);
                     // Update invoiceDetails locally so subsequent logic uses new values
                     if(invoiceDetails.custom_fields[u.i]) {
                         invoiceDetails.custom_fields[u.i].value = u.v;
                     }
                 }
                 didUpdateAddress = true;
                 setStatus("Autofilled Destination from Customer.", "info");
             } else {
                 log("Address found but no matching destination fields to fill.");
             }
             } // Close sourceAddress block
             
             // Update tracker regardless of success to avoid loop
             lastCustomerId = currentCustId;
        } else if (currentCustId && !lastCustomerId) {
             // Case: Initial Load with Customer already set.
             // We might want to set lastCustomerId here to prevent loop if user clears fields?
             // Or maybe we treat initial load as "processed".
             lastCustomerId = currentCustId;
        }

        // 2. EXTRACT CUSTOM FIELDS & STANDARD FIELDS
        const customFieldMap = {};
        if (invoiceDetails.custom_fields && invoiceDetails.custom_fields.length > 0) {
            log(`Found ${invoiceDetails.custom_fields.length} Custom Fields`);
            
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
            
            // --- SAVE API RESPONSE TO CUSTOM FIELD ---
            try {
                // Find target field index
                let responseFieldIndex = -1;
                const targetApiName = 'cf__com_yabd6a_api_response';
                
                if (invoiceDetails.custom_fields) {
                    log("Searching for response field: " + targetApiName);
                    for (let i = 0; i < invoiceDetails.custom_fields.length; i++) {
                        const cf = invoiceDetails.custom_fields[i];
                        // Match api_name OR placeholder (sometimes api_name is hidden in placeholder)
                        const cfApi = cf.api_name || "";
                        const cfPlaceholder = cf.placeholder || "";
                        const cfLabel = cf.label || "";
                        
                        if (cfApi === targetApiName || cfPlaceholder === targetApiName || cfLabel === targetApiName) {
                            responseFieldIndex = i;
                            log(`Found field match at index ${i} (Label: ${cfLabel})`);
                            break;
                        }
                    }
                }

                if (responseFieldIndex >= 0) {
                    // Update field with raw body content
                    let rawContent = "";
                    if (responseBody && responseBody.body) {
                        rawContent = typeof responseBody.body === 'string' ? responseBody.body : JSON.stringify(responseBody.body);
                    } else {
                        // Fallback if body is structured differently
                        rawContent = JSON.stringify(responseBody);
                    }
                    
                    log("Saving API Response value...");
                    await ZFAPPS.set(`invoice.custom_fields.${responseFieldIndex}.value`, rawContent);
                    // Update local object to reflect change
                    if(invoiceDetails.custom_fields[responseFieldIndex]) {
                        invoiceDetails.custom_fields[responseFieldIndex].value = rawContent;
                    }
                } else {
                    log("Warning: Custom Field 'cf__com_yabd6a_api_response' not found in invoice details.");
                    // Log available fields for debugging
                    if(invoiceDetails.custom_fields) {
                         const av = invoiceDetails.custom_fields.map(c => c.api_name || c.placeholder || c.label).join(", ");
                         log("Available API Names/Placeholders: " + av);
                    }
                }
            } catch (saveErr) {
                log("Error saving API response to field: " + saveErr.message);
            }
            // -----------------------------------------

            data = JSON.parse(responseBody.body)[0];

        } catch (connErr) {
            log("API Config Request Failed.");
            throw new Error("Failed to invoke API Config 'webhook'. Check console logs.");
        }

        if (!data) throw new Error("No data received from Webhook");


        const targetCode = data.code;
        // Make sure jurisdiction is uppercase for consistency
        const jurisdiction = data.jurisdiction ? data.jurisdiction.toUpperCase() : "";
        const state = data.state ? data.state.toUpperCase() : "";

        // Determine Mode: CA/Code vs Non-CA/Jurisdiction
        // Condition: If code is present and not null/empty, use code.
        const useCodeLookup = (targetCode && targetCode !== "null" && targetCode.trim() !== "");

        setStatus(useCodeLookup ? `Searching for Tax Code '${targetCode}'...` : `Resolving Tax for ${jurisdiction}...`, "info");
        
        // Fetch All Taxes once
        const allTaxes = await fetchAllTaxes();
        let finalTaxId = null;
        let appliedTaxName = "";
        let appliedTaxPercentage = 0; // New variable to track percentage

        if (useCodeLookup) {
             log(`Processing as CA/Code-based Tax lookup (Code: ${targetCode})`);
             // CA/Legacy Logic: Find matching code in tax name
             for (const t of allTaxes) {
                if (t.tax_name && t.tax_name.includes(",")) {
                    const parts = t.tax_name.split(",");
                    if (parts.length >= 2) {
                        const existingCode = parts[1].trim(); 
                        if (existingCode.includes(targetCode)) { 
                             log(`MATCH FOUND! Name: ${t.tax_name} | ID: ${t.tax_id}`);
                             finalTaxId = t.tax_id;
                             appliedTaxName = t.tax_name;
                             appliedTaxPercentage = t.tax_percentage; // Capture existing percentage
                             break;
                        }
                    }
                }
             }
             
             if (!finalTaxId) {
                  throw new Error(`Tax Code '${targetCode}' not found in Zoho Books settings.`);
             }

        } else {
             // Non-CA Logic
             log(`Processing as Non-CA/Jurisdiction-based Tax lookup`);
             
             if (!jurisdiction) throw new Error("No Jurisdiction provided for Non-CA Tax calculation.");
             
             const totalRate = data.total_rate; 
             // Group Name: <STATE>_<JURISDICTION>_<TOTAL_RATE>
             const groupName = `${state}_${jurisdiction}_${totalRate}`;
             appliedTaxName = groupName;
             appliedTaxPercentage = parseFloat(totalRate) * 100; // Calculate percentage from rate (0.089 -> 8.9)

             // Check if Group Exists
             const existingGroup = allTaxes.find(t => t.tax_name === groupName);

             if (existingGroup) {
                  finalTaxId = existingGroup.tax_id;
                  appliedTaxPercentage = existingGroup.tax_percentage; // Use system percentage if available
                  log(`Found existing Tax Group: ${groupName} (${finalTaxId})`);
             } else {
                  log(`Tax Group '${groupName}' not found. Creating components...`);
                  
                  // 1. Resolve Tax Authority
                  // The webhook result supposedly contains 'tax_authority'.
                  // If not, we might fall back to constructing one like "<STATE> Department of Revenue" or use default.
                  const authName = data.tax_authority || `${state} ${jurisdiction} Tax Authority`;
                  
                  let authorityId = null;
                  let authorityName = "";
                  
                  if (authName) {
                      const existingAuth = await findTaxAuthorityByName(authName);
                      if (existingAuth) {
                          authorityId = existingAuth.tax_authority_id;
                          authorityName = existingAuth.tax_authority_name;
                          log(`Found Tax Authority: ${authorityName}`);
                      } else {
                          // Create
                          const newAuth = await createTaxAuthority(authName);
                          if (newAuth) {
                              authorityId = newAuth.tax_authority_id;
                              authorityName = newAuth.tax_authority_name;
                              log(`Created Tax Authority: ${authorityName}`);
                          }
                      }
                  }

                  // 2. Component Logic
                  const components = [];
                  // We use raw rates (0.0625) for name and pass to creation (where it becomes percentage)
                  // Allow 0 rates as per user requirement, checking strictly for null/undefined
                  if(data.state_rate != null && data.state_rate >= 0) components.push({ type: 'STATE_RATE', val: data.state_rate });
                  if(data.county_rate != null && data.county_rate >= 0) components.push({ type: 'COUNTY_RATE', val: data.county_rate });
                  if(data.city_rate != null && data.city_rate >= 0) components.push({ type: 'CITY_RATE', val: data.city_rate });
                  if(data.special_rate != null && data.special_rate >= 0) components.push({ type: 'SPECIAL_RATE', val: data.special_rate });
                  
                  const componentIds = [];
                  
                  for (const comp of components) {
                      // Name Format: <STATE>_<JURISDICTION>_<TYPE>_<value>
                      const taxName = `${state}_${jurisdiction}_${comp.type}_${comp.val}`;
                      
                      // Check existence
                      const existingTax = allTaxes.find(t => t.tax_name === taxName);

                      if (existingTax) {
                          log(`Using existing component: ${taxName}`);
                          componentIds.push(existingTax.tax_id);
                      } else {
                          // Create
                          // Pass authority if resolved
                          const newId = await createTaxComponent(taxName, comp.val, authorityId, authorityName);
                          componentIds.push(newId);
                      }
                  }
                  
                  // 3. Create Group
                  if (componentIds.length > 0) {
                      const newGroupId = await createTaxGroup(groupName, componentIds);
                      if(newGroupId) {
                          finalTaxId = newGroupId;
                          // If it returns an object or ID, ensure we have the ID string
                          if(typeof finalTaxId === 'object' && finalTaxId.tax_group_id) finalTaxId = finalTaxId.tax_group_id; 
                          log(`Created new Tax Group: ${groupName} (${finalTaxId})`);
                      }
                  } else {
                      log("No valid tax components found. Cannot create group.");
                  }
             }
        }


        if (finalTaxId) {
             log(`Found/Created Tax ID: ${finalTaxId}`);
             expectedTaxId = finalTaxId; 
             
             setStatus("Updating Invoice Line Items...", "info");

             let updatedCount = 0;
             if (invoiceDetails.line_items && Array.isArray(invoiceDetails.line_items)) {
                for (let i = 0; i < invoiceDetails.line_items.length; i++) {
                     try {
                  
                         const item = invoiceDetails.line_items[i];
                         // Check if item is already set to 'Non-Taxable' - do not override
                         // Logic: If tax_name indicates Non-Taxable OR (exemption code is present AND tax_id is empty)
                         const isExempt = (item.tax_exemption_code && item.tax_exemption_code !== "" && (!item.tax_id || item.tax_id === ""));
                         
                         if (isExempt) {
                             log(`Skipping item ${i} (Non-Taxable/Exempt)`);
                             continue;
                         }

                         // Clone item for modification
                         const newItem = { ...item };
                         newItem.tax_id = finalTaxId;
                         // Force the display name and percentage so the UI shows it even if not in the cached list
                         newItem.tax_name = appliedTaxName;
                         newItem.tax_percentage = appliedTaxPercentage;
                  
                         ZFAPPS.set(`invoice.line_items[${i}]`, newItem).then(() => {
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
             
             setStatus(`Success! Applied Tax: ${appliedTaxName}`, "success");
             ZFAPPS.invoke('SHOW_MESSAGE', { type: 'success', content: `Tax updated to ${appliedTaxName} for ${updatedCount} items` });

        } else {
             throw new Error(`Could not resolve Tax ID.`);
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

/* -------------------------------------------------------------------------
   TAX CREATION & MANAGEMENT HELPERS
   ------------------------------------------------------------------------- */
   
async function fetchAllTaxes() {
    if (!organizationID) return [];
    
    let allTaxes = [];
    let stopPaging = false;
    
    // We'll search up to 30 pages
    for (let pageNum = 1; pageNum <= 30; pageNum++) {
        if (stopPaging) break;
        
        let options = {
            api_configuration_key: 'ac__com_yabd6a_zohobook_taxes',
            url_param: { orgID: organizationID, pageIndex: pageNum }
        };
        
        try {
            const response = await ZFAPPS.request(options);
            // ZFAPPS.request returns { data: ..., status: ... } or the body directly
            let body = response.data || response.body; 
            
            if (typeof body.body === 'string') {
                 try { body = JSON.parse(body.body); } catch(e) {}
            }
            
            if (body.code === 0 || body.code === "0") {
                const pageTaxes = body.taxes || [];
                allTaxes = allTaxes.concat(pageTaxes);
                
                if (body.page_context) {
                   if (!body.page_context.has_more_page) stopPaging = true;
                } else {
                   stopPaging = true; // No context usually means single page or error
                }
            } else {
                log(`Error fetching taxes page ${pageNum}: ${body.message}`);
                break; 
            }
        } catch (e) {
            console.error("Error fetching taxes page " + pageNum, e);
            break;
        }
    }
    return allTaxes;
}

async function findTaxAuthorityByName(authName) {
    if (!authName) return null;
    log(`Searching for Tax Authority: ${authName}`);

    // Since there's no pagination requirement mentioned extensively for authorities or we assume fewer authorities,
    // we can try fetching. However, list APIs are usually paginated.
    // Assuming 'ac__com_yabd6a_zohobook_tax_authorities' lists them.
    
    let options = {
        api_configuration_key: 'ac__com_yabd6a_zohobook_tax_authorities',
        url_param: { orgID: organizationID }
    };
    
    try {
        const response = await ZFAPPS.request(options);
        let body = response.data || response.body;
        if (typeof body.body === 'string') try { body = JSON.parse(body.body); } catch(e) {}
        
        if (body.code === 0 || body.code === "0") {
            const authorities = body.tax_authorities || [];
            // Basic search
            const found = authorities.find(a => a.tax_authority_name.toLowerCase() === authName.toLowerCase());
            if (found) return found;
        }
    } catch (e) {
        log(`Error fetching tax authorities: ${e.message}`);
    }
    return null;
}

async function createTaxAuthority(authName) {
    log(`Creating Tax Authority: ${authName}`);
    
    const payload = {
        "tax_authority_name": authName,
        "description": "Created by Extension",
        "organization_id": organizationID
    };

    const options = {
        api_configuration_key: 'ac__com_yabd6a_zohobook_tax_authorities_po',
        url_param: { orgID: organizationID },
        body: {
            mode: 'raw',
            raw: JSON.stringify(payload)
        }
    };
    
    try {
        const response = await ZFAPPS.request(options);
        let body = response.data || response.body;
        if(typeof body.body === 'string') try { body = JSON.parse(body.body); } catch(e){}
        
        if((body.code === 0 || body.code === "0") && body.tax_authority_name) {
            return body.tax_authority_name;
        } else {
            throw new Error(body.message || "Unknown error creating tax authority");
        }
    } catch (e) {
        log(`API Create Authority Error: ${e.message}`);
        throw e;
    }
}

async function createTaxComponent(name, rateRaw, authorityId, authorityName) {
    // Rate raw is like 0.0625 -> 6.25
    const ratePercentage = parseFloat(rateRaw) * 100;
    
    log(`Creating Tax Component: ${name} (${ratePercentage}%) [Auth: ${authorityName}]`);
    
    const payload = {
        "tax_name": name,
        "tax_percentage": ratePercentage,
        "tax_type": "tax",
        "organization_id": organizationID
    };
    
    // Add authority if provided
    if (authorityId) {
        payload.tax_authority_id = authorityId;
        payload.tax_authority_name = authorityName;
    }
    
    // Using provided API Config for Tax Creation
    const options = {
        api_configuration_key: 'ac__com_yabd6a_zohobook_taxes_post',
        url_param: { orgID: organizationID },
        body: {
            mode: 'raw',
            raw: JSON.stringify(payload)
        }
    };
    
    try {
        const response = await ZFAPPS.request(options);
        let body = response.data || response.body;

        if(typeof body.body === 'string') try { body = JSON.parse(body.body); } catch(e){}
        
        if((body.code === 0 || body.code === "0") && body.tax) {
            return body.tax.tax_id;
        } else {
            throw new Error(body.message || "Unknown error creating tax");
        }
    } catch (e) {
        log(`API Create Tax Error: ${e.message}`);
        throw e;
    }
}

async function createTaxGroup(name, taxIds) {
    log(`Creating Tax Group: ${name} with IDs: ${taxIds.join(', ')}`);
    
    const payload = {
        "tax_group_name": name, // API doc says 'tax_group_name' for groups usually, checking context... 
        // Docs provided say: 'tax_group_name' for Create a tax group.
        // Wait, standard Create Tax endpoint supports 'tax_type': 'compound_tax' or similar?
        // Ah, the user provided a specific connection for 'create tax group' and screenshot says 'Create a tax group'.
        // It uses 'tax_group_name' and 'taxes' (list of IDs string).
        
        "taxes": taxIds, // Screenshot says "taxes": "98200...", array or CSV?
        // Screenshot example: "taxes": "982000000566009" (string)
        // Description says "Comma Separated list of tax IDs"
        "organization_id": organizationID
    };
    
    // Use the CSV string if screenshot implies simple string for single ID, but usually for group it's multiple.
    // Argument 'taxes' description: "Comma Seperated list of tax IDs".
    payload.taxes = taxIds.join(',');

    const options = {
        api_configuration_key: 'ac__com_yabd6a_zohobook_tax_group',
        url_param: { orgID: organizationID },
        body: {
            mode: 'raw',
            raw: JSON.stringify(payload)
        }
    };
    
    try {
        const response = await ZFAPPS.request(options);
        let body = response.data || response.body;
        if(typeof body.body === 'string') try { body = JSON.parse(body.body); } catch(e){}
        
        if((body.code === 0 || body.code === "0") && body.tax_group) {
             return body.tax_group.tax_group_id; 
        } else {
             // Fallback: sometimes it returns 'tax' object?
             if(body.tax && body.tax.tax_id) return body.tax.tax_id;
             throw new Error(body.message || "Unknown error creating tax group");
        }
    } catch (e) {
        log(`API Create Tax Group Error: ${e.message}`);
        throw e;
    }
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
    // However, we DO want to trigger on line item COUNT changes (adding/removing items)
    // So we manually store the length before deleting the full array.
    if (currentInvoice.line_items) {
        checkObj._lineItemCount = currentInvoice.line_items.length;
    }

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
        
        // If currently marked as Non-Taxable, don't consider it a conflict
        const isNonTaxable = (item.tax_name === "Non-Taxable" || item.tax_name === "Non Taxable");
        const isExempt = (item.tax_exemption_code && item.tax_exemption_code !== "" && (!item.tax_id || item.tax_id === ""));
        
        if (isNonTaxable || isExempt) {
            continue;
        }

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
