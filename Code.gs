/**
 * THREAD TRACK - Google Apps Script Backend
 * This script runs as a Web App bound to the central Admin Spreadsheet.
 * It manages user authentication, audit logs, and handles reading/writing data
 * from/to individual user spreadsheets.
 */

// Simple SHA-256 function for password hashing
function sha256(input) {
  if (!input) {
    throw new Error("Do not run the 'sha256' function directly. Select 'initializeAdminSpreadsheet' from the dropdown at the top of the Apps Script editor, then click 'Run' to initialize the database tables.");
  }
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  var output = '';
  for (var i = 0; i < rawHash.length; i++) {
    var byteValue = rawHash[i];
    if (byteValue < 0) byteValue += 256;
    var byteString = byteValue.toString(16);
    if (byteString.length == 1) byteString = '0' + byteString;
    output += byteString;
  }
  return output;
}

// Open and verify the Admin Spreadsheet (throws structured warning if script is unbound)
function getAdminSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("This Google Apps Script is not bound to a Google Spreadsheet. " +
                     "Please make sure you create this script by opening your central 'Thread Track Admin DB' Google Sheet, " +
                     "and clicking 'Extensions' -> 'Apps Script' on the top menu bar, then pasting the code there. " +
                     "If you created a standalone script, it will not be able to find the active sheet!");
  }
  return ss;
}

// Open a spreadsheet by ID, or fall back to the active one (Admin Spreadsheet)
function getSpreadsheet(spreadsheetId) {
  if (spreadsheetId && spreadsheetId !== "null" && spreadsheetId !== "undefined") {
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      Logger.log("Error opening spreadsheet by ID " + spreadsheetId + ": " + e.message);
    }
  }
  return getAdminSpreadsheet();
}

// Convert sheet data to array of objects using headers
function getSheetData(sheet) {
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  
  return data.map(function(row) {
    var obj = {};
    headers.forEach(function(header, index) {
      obj[header] = row[index];
    });
    return obj;
  });
}

// Write objects array back to sheet (overwriting data below headers)
function writeSheetData(sheet, data, headers) {
  if (!sheet) return;
  
  // Clear existing content below header
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  
  if (data.length === 0) return;
  
  var values = data.map(function(rowObj) {
    return headers.map(function(header) {
      return rowObj[header] !== undefined ? rowObj[header] : "";
    });
  });
  
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

// Append a single row as an object
function appendRow(sheet, rowObj) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var newRow = headers.map(function(header) {
    return rowObj[header] !== undefined ? rowObj[header] : "";
  });
  sheet.appendRow(newRow);
}

// Get standard response with CORS support
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Ensure the sheets exist inside a spreadsheet (checks and creates them if needed)
function initializeUserSpreadsheet(ss) {
  if (!ss) {
    throw new Error("Do not run the 'initializeUserSpreadsheet' function directly. Select 'initializeAdminSpreadsheet' from the dropdown at the top of the Apps Script editor, then click 'Run' to initialize the database tables.");
  }
  var sheetsNeeded = {
    "vendors": ["vendor_id", "vendor_name", "vendor_type", "mobile", "address", "gst_number", "notes", "created_at"],
    "sarees": ["saree_id", "lot_number", "design_name", "quantity", "current_stage", "current_vendor_id", "status", "remarks", "created_at"],
    "workflow_history": ["history_id", "saree_id", "stage_name", "vendor_id", "sent_date", "received_date", "work_cost", "remarks", "updated_by", "created_at"],
    "payments": ["payment_id", "vendor_id", "payment_date", "amount", "payment_method", "remarks", "created_by", "created_at"],
    "settings": ["key", "value"]
  };
  
  for (var sheetName in sheetsNeeded) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(sheetsNeeded[sheetName]);
      // Format headers
      sheet.getRange(1, 1, 1, sheetsNeeded[sheetName].length).setFontWeight("bold").setBackground("#f3f3f3");
    }
  }
}

// GET handler
function doGet(e) {
  var action = e.parameter.action;
  var spreadsheetId = e.parameter.spreadsheetId;
  
  var adminSS = getAdminSpreadsheet();
  var userSS = getSpreadsheet(spreadsheetId);
  
  try {
    if (action === "getUsers") {
      var usersSheet = adminSS.getSheetByName("users");
      var users = getSheetData(usersSheet);
      // Remove password hash from response
      users.forEach(function(u) { delete u.password_hash; });
      return jsonResponse({ success: true, users: users });
      
    } else if (action === "getAuditLogs") {
      var logsSheet = adminSS.getSheetByName("audit_logs");
      var logs = getSheetData(logsSheet);
      // Return newest first
      logs.reverse();
      return jsonResponse({ success: true, logs: logs.slice(0, 200) });
      
    } else if (action === "getVendors") {
      var type = e.parameter.type;
      var vendorsSheet = userSS.getSheetByName("vendors");
      var vendors = getSheetData(vendorsSheet);
      if (type) {
        vendors = vendors.filter(function(v) { return v.vendor_type === type; });
      }
      return jsonResponse({ success: true, vendors: vendors });
      
    } else if (action === "getSarees") {
      var stage = e.parameter.stage;
      var status = e.parameter.status;
      var vendorId = e.parameter.vendorId;
      
      var sareesSheet = userSS.getSheetByName("sarees");
      var vendorsSheet = userSS.getSheetByName("vendors");
      
      var sarees = getSheetData(sareesSheet);
      var vendors = getSheetData(vendorsSheet);
      
      // Build vendor lookup
      var vendorMap = {};
      vendors.forEach(function(v) { vendorMap[v.vendor_id] = v; });
      
      sarees.forEach(function(s) {
        var v = vendorMap[s.current_vendor_id];
        s.current_vendor_name = v ? v.vendor_name : "";
        s.current_vendor_type = v ? v.vendor_type : "";
      });
      
      if (stage && stage !== "All") {
        sarees = sarees.filter(function(s) { return s.current_stage === stage; });
      }
      if (status && status !== "All") {
        sarees = sarees.filter(function(s) { return s.status === status; });
      }
      if (vendorId && vendorId !== "All") {
        sarees = sarees.filter(function(s) { return String(s.current_vendor_id) === String(vendorId); });
      }
      
      return jsonResponse({ success: true, sarees: sarees });
      
    } else if (action === "getNextLot") {
      var sareesSheet = userSS.getSheetByName("sarees");
      var sarees = getSheetData(sareesSheet);
      var maxLot = 1000; // default start
      sarees.forEach(function(s) {
        var lot = parseInt(s.lot_number);
        if (lot > maxLot) maxLot = lot;
      });
      return jsonResponse({ success: true, nextLot: maxLot + 1 });
      
    } else if (action === "getSareeDetails") {
      var sareeId = e.parameter.sareeId;
      var sareesSheet = userSS.getSheetByName("sarees");
      var historySheet = userSS.getSheetByName("workflow_history");
      var vendorsSheet = userSS.getSheetByName("vendors");
      
      var sarees = getSheetData(sareesSheet);
      var history = getSheetData(historySheet);
      var vendors = getSheetData(vendorsSheet);
      
      var saree = sarees.find(function(s) { return String(s.saree_id) === String(sareeId); });
      if (!saree) {
        return jsonResponse({ success: false, error: "Saree lot not found" });
      }
      
      var vendorMap = {};
      vendors.forEach(function(v) { vendorMap[v.vendor_id] = v; });
      
      var sareeHistory = history
        .filter(function(h) { return String(h.saree_id) === String(sareeId); })
        .map(function(h) {
          var v = vendorMap[h.vendor_id];
          h.vendor_name = v ? v.vendor_name : "";
          h.vendor_type = v ? v.vendor_type : "";
          return h;
        });
        
      // Sort chronologically
      sareeHistory.sort(function(a, b) { return new Date(a.sent_date) - new Date(b.sent_date); });
      
      return jsonResponse({ success: true, saree: saree, history: sareeHistory });
      
    } else if (action === "getHisab") {
      var vendorsSheet = userSS.getSheetByName("vendors");
      var historySheet = userSS.getSheetByName("workflow_history");
      var paymentsSheet = userSS.getSheetByName("payments");
      
      var vendors = getSheetData(vendorsSheet);
      var history = getSheetData(historySheet);
      var payments = getSheetData(paymentsSheet);
      
      // Calculate work values (only where received_date is completed)
      var workMap = {};
      var completedCountMap = {};
      var pendingCountMap = {};
      
      history.forEach(function(h) {
        var vId = h.vendor_id;
        if (h.received_date) {
          workMap[vId] = (workMap[vId] || 0) + parseFloat(h.work_cost || 0);
          completedCountMap[vId] = (completedCountMap[vId] || 0) + 1;
        } else {
          pendingCountMap[vId] = (pendingCountMap[vId] || 0) + 1;
        }
      });
      
      // Calculate total paid
      var paidMap = {};
      payments.forEach(function(p) {
        var vId = p.vendor_id;
        paidMap[vId] = (paidMap[vId] || 0) + parseFloat(p.amount || 0);
      });
      
      var hisabSummary = vendors.map(function(v) {
        var vId = v.vendor_id;
        var totalWork = workMap[vId] || 0;
        var totalPaid = paidMap[vId] || 0;
        return {
          vendor_id: v.vendor_id,
          vendor_name: v.vendor_name,
          vendor_type: v.vendor_type,
          mobile: v.mobile,
          total_work: totalWork,
          total_paid: totalPaid,
          pending_balance: totalWork - totalPaid,
          completed_sarees_count: completedCountMap[vId] || 0,
          pending_sarees_count: pendingCountMap[vId] || 0
        };
      });
      
      return jsonResponse({ success: true, hisab: hisabSummary });
      
    } else if (action === "getVendorLedger") {
      var vendorId = e.parameter.vendorId;
      var vendorsSheet = userSS.getSheetByName("vendors");
      var historySheet = userSS.getSheetByName("workflow_history");
      var paymentsSheet = userSS.getSheetByName("payments");
      var sareesSheet = userSS.getSheetByName("sarees");
      
      var vendors = getSheetData(vendorsSheet);
      var vendor = vendors.find(function(v) { return String(v.vendor_id) === String(vendorId); });
      if (!vendor) return jsonResponse({ success: false, error: "Vendor not found" });
      
      var history = getSheetData(historySheet).filter(function(h) { return String(h.vendor_id) === String(vendorId) && h.received_date; });
      var payments = getSheetData(paymentsSheet).filter(function(p) { return String(p.vendor_id) === String(vendorId); });
      var sarees = getSheetData(sareesSheet);
      
      var sareeMap = {};
      serees.forEach(function(s) { sareeMap[s.saree_id] = s; });
      
      var ledger = [];
      
      // Work credits
      history.forEach(function(w) {
        var s = sareeMap[w.saree_id] || { lot_number: "Unknown", design_name: "Unknown", quantity: 0 };
        ledger.push({
          id: "work_" + w.history_id,
          date: w.received_date,
          type: "work",
          description: w.stage_name + " work completed on Lot #" + s.lot_number + " (" + s.design_name + ", Qty: " + s.quantity + ")",
          amount: parseFloat(w.work_cost || 0),
          remarks: w.remarks
        });
      });
      
      // Payment debits
      payments.forEach(function(p) {
        ledger.push({
          id: "payment_" + p.payment_id,
          date: p.payment_date,
          type: "payment",
          description: "Payment made via " + p.payment_method,
          amount: parseFloat(p.amount || 0),
          remarks: p.remarks
        });
      });
      
      // Sort chronologically
      ledger.sort(function(a, b) { return new Date(a.date) - new Date(b.date); });
      
      var runningBalance = 0;
      var ledgerWithBalance = ledger.map(function(entry) {
        if (entry.type === "work") {
          runningBalance += entry.amount;
        } else {
          runningBalance -= entry.amount;
        }
        entry.running_balance = runningBalance;
        return entry;
      });
      
      // Reverse to show newest transactions first
      ledgerWithBalance.reverse();
      
      var totalWork = history.reduce(function(sum, w) { return sum + parseFloat(w.work_cost || 0); }, 0);
      var totalPaid = payments.reduce(function(sum, p) { return sum + parseFloat(p.amount || 0); }, 0);
      
      return jsonResponse({
        success: true,
        vendor: vendor,
        summary: {
          total_work: totalWork,
          total_paid: totalPaid,
          pending_balance: totalWork - totalPaid
        },
        ledger: ledgerWithBalance
      });
      
    } else if (action === "getPayments") {
      var paymentsSheet = userSS.getSheetByName("payments");
      var vendorsSheet = userSS.getSheetByName("vendors");
      var payments = getSheetData(paymentsSheet);
      var vendors = getSheetData(vendorsSheet);
      
      var vendorMap = {};
      vendors.forEach(function(v) { vendorMap[v.vendor_id] = v; });
      
      payments.forEach(function(p) {
        var v = vendorMap[p.vendor_id];
        p.vendor_name = v ? v.vendor_name : "Deleted Vendor";
        p.vendor_type = v ? v.vendor_type : "";
      });
      
      payments.reverse(); // Newest first
      return jsonResponse({ success: true, payments: payments });
      
    } else if (action === "getDashboardStats") {
      var sareesSheet = userSS.getSheetByName("sarees");
      var vendorsSheet = userSS.getSheetByName("vendors");
      var historySheet = userSS.getSheetByName("workflow_history");
      var paymentsSheet = userSS.getSheetByName("payments");
      
      var sarees = getSheetData(sareesSheet);
      var vendors = getSheetData(vendorsSheet);
      var history = getSheetData(historySheet);
      var payments = getSheetData(paymentsSheet);
      
      // Calculate pipeline stats
      var pipeline = {
        "Dyed": { count: 0, qty: 0 },
        "Embroidery": { count: 0, qty: 0 },
        "Stitching": { count: 0, qty: 0 },
        "Diamond": { count: 0, qty: 0 },
        "Folding": { count: 0, qty: 0 },
        "Completed": { count: 0, qty: 0 }
      };
      
      sarees.forEach(function(s) {
        if (pipeline[s.current_stage]) {
          pipeline[s.current_stage].count++;
          pipeline[s.current_stage].qty += parseInt(s.quantity || 0);
        }
      });
      
      // Outstanding balance calculations
      var totalWorkCost = history.reduce(function(sum, h) {
        return sum + (h.received_date ? parseFloat(h.work_cost || 0) : 0);
      }, 0);
      
      var totalPaidAmount = payments.reduce(function(sum, p) {
        return sum + parseFloat(p.amount || 0);
      }, 0);
      
      var summary = {
        total_sarees_lots: sarees.length,
        total_sarees_quantity: sarees.reduce(function(sum, s) { return sum + parseInt(s.quantity || 0); }, 0),
        total_manufacturing_cost: totalWorkCost,
        total_outstanding_payable: totalWorkCost - totalPaidAmount,
        total_vendors: vendors.length
      };
      
      return jsonResponse({ success: true, summary: summary, pipeline: pipeline });
      
    } else if (action === "getReports") {
      var reportType = e.parameter.reportType;
      var filterVendor = e.parameter.vendorId;
      var filterType = e.parameter.vendorType;
      var filterStage = e.parameter.stage;
      var startDate = e.parameter.startDate;
      var endDate = e.parameter.endDate;
      
      if (reportType === "outstanding") {
        // Output outstanding liabilities
        var hisabRes = doGet({ parameter: { action: "getHisab", spreadsheetId: spreadsheetId } });
        var hisab = JSON.parse(hisabRes.getContent()).hisab;
        
        if (filterVendor && filterVendor !== "All") {
          hisab = hisab.filter(function(h) { return String(h.vendor_id) === String(filterVendor); });
        }
        if (filterType && filterType !== "All") {
          hisab = hisab.filter(function(h) { return h.vendor_type === filterType; });
        }
        hisab = hisab.filter(function(h) { return h.pending_balance > 0; });
        
        return jsonResponse({ success: true, outstanding: hisab });
        
      } else if (reportType === "production") {
        var sareesSheet = userSS.getSheetByName("sarees");
        var vendorsSheet = userSS.getSheetByName("vendors");
        var sarees = getSheetData(sareesSheet);
        var vendors = getSheetData(vendorsSheet);
        var vendorMap = {};
        vendors.forEach(function(v) { vendorMap[v.vendor_id] = v; });
        
        sarees.forEach(function(s) {
          var v = vendorMap[s.current_vendor_id];
          s.vendor_name = v ? v.vendor_name : "";
        });
        
        if (filterVendor && filterVendor !== "All") {
          sarees = sarees.filter(function(s) { return String(s.current_vendor_id) === String(filterVendor); });
        }
        if (filterStage && filterStage !== "All") {
          sarees = sarees.filter(function(s) { return s.current_stage === filterStage; });
        }
        if (startDate) {
          sarees = sarees.filter(function(s) { return new Date(s.created_at) >= new Date(startDate); });
        }
        if (endDate) {
          sarees = sarees.filter(function(s) { return new Date(s.created_at) <= new Date(endDate + "T23:59:59.999Z"); });
        }
        
        return jsonResponse({ success: true, production: sarees });
        
      } else if (reportType === "payments") {
        var paymentsRes = doGet({ parameter: { action: "getPayments", spreadsheetId: spreadsheetId } });
        var payments = JSON.parse(paymentsRes.getContent()).payments;
        
        if (filterVendor && filterVendor !== "All") {
          payments = payments.filter(function(p) { return String(p.vendor_id) === String(filterVendor); });
        }
        if (filterType && filterType !== "All") {
          payments = payments.filter(function(p) { return p.vendor_type === filterType; });
        }
        if (startDate) {
          payments = payments.filter(function(p) { return new Date(p.payment_date) >= new Date(startDate); });
        }
        if (endDate) {
          payments = payments.filter(function(p) { return new Date(p.payment_date) <= new Date(endDate + "T23:59:59.999Z"); });
        }
        
        return jsonResponse({ success: true, payments: payments });
      }
    }
    
    return jsonResponse({ success: false, error: "Invalid action" });
    
  } catch (error) {
    return jsonResponse({ success: false, error: error.toString() });
  }
}

// POST handler (handles writes, creates, updates, and deletes)
function doPost(e) {
  var postData = JSON.parse(e.postData.contents);
  var action = postData.action;
  var spreadsheetId = postData.spreadsheetId;
  
  var adminSS = getAdminSpreadsheet();
  var userSS = getSpreadsheet(spreadsheetId);
  
  try {
    if (action === "login") {
      var email = postData.email;
      var password = postData.password;
      
      var usersSheet = adminSS.getSheetByName("users");
      var users = getSheetData(usersSheet);
      
      var user = users.find(function(u) { return u.email === email || u.mobile === email; });
      if (!user) return jsonResponse({ success: false, error: "User not found" });
      
      if (!user.active || String(user.active) === "false") {
        return jsonResponse({ success: false, error: "Account deactivated. Contact Admin." });
      }
      
      var pwdHash = sha256(password);
      if (user.password_hash !== pwdHash) {
        return jsonResponse({ success: false, error: "Incorrect password" });
      }
      
      // Login successful - return profile details
      var logObj = {
        log_id: new Date().getTime(),
        user_id: user.id,
        user_name: user.name,
        action: "LOGIN",
        details: "User logged in successfully",
        ip_address: postData.ipAddress || "N/A",
        created_at: new Date().toISOString()
      };
      appendRow(adminSS.getSheetByName("audit_logs"), logObj);
      
      return jsonResponse({
        success: true,
        user: {
          id: user.id,
          name: user.name,
          email: user.email,
          mobile: user.mobile,
          role: user.role,
          spreadsheetId: user.spreadsheet_id
        }
      });
      
    } else if (action === "addUser") {
      // Create new user (Admin only)
      var usersSheet = adminSS.getSheetByName("users");
      var users = getSheetData(usersSheet);
      
      var email = postData.email;
      var mobile = postData.mobile;
      
      var userExists = users.some(function(u) { return u.email === email || u.mobile === mobile; });
      if (userExists) return jsonResponse({ success: false, error: "Email or Mobile already registered" });
      
      var newUserId = users.length > 0 ? Math.max.apply(null, users.map(function(u) { return parseInt(u.id); })) + 1 : 1;
      
      // Auto-clone spreadsheet template if requested, or create new empty spreadsheet
      var newSSId = "";
      try {
        var templateId = postData.templateId || adminSS.getSheetByName("settings").getRange("B1").getValue() || ""; // Assume setting or key exists
        var newSSFile;
        if (templateId) {
          newSSFile = DriveApp.getFileById(templateId).makeCopy("Thread Track DB - " + postData.name);
        } else {
          var newSS = SpreadsheetApp.create("Thread Track DB - " + postData.name);
          initializeUserSpreadsheet(newSS);
          newSSId = newSS.getId();
        }
        if (newSSFile) {
          newSSId = newSSFile.getId();
        }
      } catch (err) {
        Logger.log("Drive copy failed: " + err.toString());
        // Fallback: create fresh Spreadsheet
        var newSS = SpreadsheetApp.create("Thread Track DB - " + postData.name);
        initializeUserSpreadsheet(newSS);
        newSSId = newSS.getId();
      }
      
      var newUser = {
        id: newUserId,
        name: postData.name,
        mobile: mobile,
        email: email,
        password_hash: sha256(postData.password),
        role: postData.role || "User",
        active: true,
        spreadsheet_id: newSSId,
        created_at: new Date().toISOString()
      };
      
      appendRow(usersSheet, newUser);
      
      // Log Action
      appendRow(adminSS.getSheetByName("audit_logs"), {
        log_id: new Date().getTime(),
        user_id: postData.adminId,
        user_name: postData.adminName,
        action: "ADD_USER",
        details: "Created user: " + postData.name + " (" + email + ") with sheet " + newSSId,
        ip_address: postData.ipAddress || "N/A",
        created_at: new Date().toISOString()
      });
      
      return jsonResponse({ success: true, message: "User created successfully", user: newUser });
      
    } else if (action === "updateUser") {
      var usersSheet = adminSS.getSheetByName("users");
      var users = getSheetData(usersSheet);
      var userId = postData.userId;
      
      var userIdx = users.findIndex(function(u) { return String(u.id) === String(userId); });
      if (userIdx === -1) return jsonResponse({ success: false, error: "User not found" });
      
      var targetUser = users[userIdx];
      if (postData.name) targetUser.name = postData.name;
      if (postData.role) targetUser.role = postData.role;
      if (postData.spreadsheetId) targetUser.spreadsheet_id = postData.spreadsheetId;
      if (postData.password) targetUser.password_hash = sha256(postData.password);
      if (postData.active !== undefined) targetUser.active = postData.active;
      
      writeSheetData(usersSheet, users, ["id", "name", "mobile", "email", "password_hash", "role", "active", "spreadsheet_id", "created_at"]);
      
      // Log Action
      appendRow(adminSS.getSheetByName("audit_logs"), {
        log_id: new Date().getTime(),
        user_id: postData.adminId,
        user_name: postData.adminName,
        action: "UPDATE_USER",
        details: "Updated user ID " + userId + " properties: " + JSON.stringify(postData),
        ip_address: postData.ipAddress || "N/A",
        created_at: new Date().toISOString()
      });
      
      return jsonResponse({ success: true, message: "User updated successfully" });
      
    } else if (action === "addVendor") {
      var vendorsSheet = userSS.getSheetByName("vendors");
      var vendors = getSheetData(vendorsSheet);
      var vendorId = vendors.length > 0 ? Math.max.apply(null, vendors.map(function(v) { return parseInt(v.vendor_id); })) + 1 : 1;
      
      var newVendor = {
        vendor_id: vendorId,
        vendor_name: postData.vendor_name,
        vendor_type: postData.vendor_type,
        mobile: postData.mobile,
        address: postData.address || "",
        gst_number: postData.gst_number || "",
        notes: postData.notes || "",
        created_at: new Date().toISOString()
      };
      
      appendRow(vendorsSheet, newVendor);
      return jsonResponse({ success: true, vendor: newVendor });
      
    } else if (action === "updateVendor") {
      var vendorsSheet = userSS.getSheetByName("vendors");
      var vendors = getSheetData(vendorsSheet);
      var vendorId = postData.vendor_id;
      
      var idx = vendors.findIndex(function(v) { return String(v.vendor_id) === String(vendorId); });
      if (idx === -1) return jsonResponse({ success: false, error: "Vendor not found" });
      
      vendors[idx].vendor_name = postData.vendor_name;
      vendors[idx].vendor_type = postData.vendor_type;
      vendors[idx].mobile = postData.mobile;
      vendors[idx].address = postData.address || "";
      vendors[idx].gst_number = postData.gst_number || "";
      vendors[idx].notes = postData.notes || "";
      
      writeSheetData(vendorsSheet, vendors, ["vendor_id", "vendor_name", "vendor_type", "mobile", "address", "gst_number", "notes", "created_at"]);
      return jsonResponse({ success: true, message: "Vendor updated" });
      
    } else if (action === "deleteVendor") {
      var vendorsSheet = userSS.getSheetByName("vendors");
      var vendors = getSheetData(vendorsSheet);
      var vendorId = postData.vendorId;
      
      var filtered = vendors.filter(function(v) { return String(v.vendor_id) !== String(vendorId); });
      writeSheetData(vendorsSheet, filtered, ["vendor_id", "vendor_name", "vendor_type", "mobile", "address", "gst_number", "notes", "created_at"]);
      return jsonResponse({ success: true, message: "Vendor deleted" });
      
    } else if (action === "addSaree") {
      var sareesSheet = userSS.getSheetByName("sarees");
      var sarees = getSheetData(sareesSheet);
      
      // Check lot number uniqueness
      var lotExists = sarees.some(function(s) { return parseInt(s.lot_number) === parseInt(postData.lot_number); });
      if (lotExists) return jsonResponse({ success: false, error: "Lot number already exists" });
      
      var sareeId = sarees.length > 0 ? Math.max.apply(null, sarees.map(function(s) { return parseInt(s.saree_id); })) + 1 : 1;
      
      var newSaree = {
        saree_id: sareeId,
        lot_number: postData.lot_number,
        design_name: postData.design_name,
        quantity: postData.quantity,
        current_stage: "Dyed", // Starts at Dyed stage
        current_vendor_id: "",
        status: "In Process",
        remarks: postData.remarks || "",
        created_at: new Date().toISOString()
      };
      
      appendRow(sareesSheet, newSaree);
      return jsonResponse({ success: true, saree: newSaree });
      
    } else if (action === "sendSareeToStage") {
      var sareeId = postData.sareeId;
      var stageName = postData.stage_name;
      var vendorId = postData.vendor_id;
      var remarks = postData.remarks || "";
      var estimatedCost = parseFloat(postData.estimated_cost || 0);
      
      var sareesSheet = userSS.getSheetByName("sarees");
      var sarees = getSheetData(sareesSheet);
      var sIdx = sarees.findIndex(function(s) { return String(s.saree_id) === String(sareeId); });
      if (sIdx === -1) return jsonResponse({ success: false, error: "Saree lot not found" });
      
      var saree = sarees[sIdx];
      saree.current_stage = stageName;
      saree.current_vendor_id = vendorId;
      saree.status = "In Process";
      
      writeSheetData(sareesSheet, sarees, ["saree_id", "lot_number", "design_name", "quantity", "current_stage", "current_vendor_id", "status", "remarks", "created_at"]);
      
      var historySheet = userSS.getSheetByName("workflow_history");
      var history = getSheetData(historySheet);
      var historyId = history.length > 0 ? Math.max.apply(null, history.map(function(h) { return parseInt(h.history_id); })) + 1 : 1;
      
      var newHistory = {
        history_id: historyId,
        saree_id: sareeId,
        stage_name: stageName,
        vendor_id: vendorId,
        sent_date: new Date().toISOString(),
        received_date: "",
        work_cost: estimatedCost, // estimated at send time
        remarks: remarks,
        updated_by: postData.updatedBy || "",
        created_at: new Date().toISOString()
      };
      
      appendRow(historySheet, newHistory);
      return jsonResponse({ success: true, message: "Saree sent to stage " + stageName });
      
    } else if (action === "receiveSareeFromStage") {
      var sareeId = postData.sareeId;
      var finalCost = parseFloat(postData.work_cost || 0);
      var remarks = postData.remarks || "";
      var isCompleted = postData.isCompleted; // true if marking the entire Lot as completed
      
      var sareesSheet = userSS.getSheetByName("sarees");
      var sarees = getSheetData(sareesSheet);
      var sIdx = sarees.findIndex(function(s) { return String(s.saree_id) === String(sareeId); });
      if (sIdx === -1) return jsonResponse({ success: false, error: "Saree lot not found" });
      
      var saree = sarees[sIdx];
      
      // Update history record (find the active pending record for this saree & stage)
      var historySheet = userSS.getSheetByName("workflow_history");
      var history = getSheetData(historySheet);
      
      var hIdx = history.findIndex(function(h) { 
        return String(h.saree_id) === String(sareeId) && 
               h.stage_name === saree.current_stage && 
               String(h.vendor_id) === String(saree.current_vendor_id) && 
               !h.received_date;
      });
      
      if (hIdx === -1) return jsonResponse({ success: false, error: "Active job slip record not found" });
      
      history[hIdx].received_date = new Date().toISOString();
      history[hIdx].work_cost = finalCost;
      if (remarks) history[hIdx].remarks = remarks;
      
      writeSheetData(historySheet, history, ["history_id", "saree_id", "stage_name", "vendor_id", "sent_date", "received_date", "work_cost", "remarks", "updated_by", "created_at"]);
      
      // Update saree state
      if (isCompleted) {
        saree.current_stage = "Completed";
        saree.current_vendor_id = "";
        saree.status = "Completed";
      } else {
        // Returned to workshop, wait to assign next stage
        saree.current_vendor_id = "";
        saree.status = "In Process";
      }
      
      writeSheetData(sareesSheet, sarees, ["saree_id", "lot_number", "design_name", "quantity", "current_stage", "current_vendor_id", "status", "remarks", "created_at"]);
      return jsonResponse({ success: true, message: "Saree received from stage successfully" });
      
    } else if (action === "rollbackSareeStage") {
      var sareeId = postData.sareeId;
      var sareesSheet = userSS.getSheetByName("sarees");
      var sarees = getSheetData(sareesSheet);
      var sIdx = sarees.findIndex(function(s) { return String(s.saree_id) === String(sareeId); });
      if (sIdx === -1) return jsonResponse({ success: false, error: "Saree lot not found" });
      
      var saree = sarees[sIdx];
      
      var historySheet = userSS.getSheetByName("workflow_history");
      var history = getSheetData(historySheet);
      
      // Find the last history item for this saree
      var sareeHistory = history.filter(function(h) { return String(h.saree_id) === String(sareeId); });
      if (sareeHistory.length === 0) {
        // If there's no history, we can't rollback
        return jsonResponse({ success: false, error: "No history to rollback" });
      }
      
      sareeHistory.sort(function(a, b) { return new Date(b.created_at) - new Date(a.created_at); });
      var lastRecord = sareeHistory[0];
      
      // Delete this history record
      var filteredHistory = history.filter(function(h) { return String(h.history_id) !== String(lastRecord.history_id); });
      writeSheetData(historySheet, filteredHistory, ["history_id", "saree_id", "stage_name", "vendor_id", "sent_date", "received_date", "work_cost", "remarks", "updated_by", "created_at"]);
      
      // Revert saree state to previous
      if (lastRecord.received_date) {
        // If it was already received, rollback makes it "Sent to this vendor/stage" again
        saree.current_stage = lastRecord.stage_name;
        saree.current_vendor_id = lastRecord.vendor_id;
        saree.status = "In Process";
      } else {
        // If it was not yet received (still at vendor), rollback means we revert to the stage *before* this one
        var prevHistory = sareeHistory[1];
        if (prevHistory) {
          saree.current_stage = prevHistory.stage_name;
          saree.current_vendor_id = ""; // In workshop, since it was received in previous stage
          saree.status = "In Process";
        } else {
          // If this was the first stage, revert to initial state (Workshop)
          saree.current_stage = "Dyed";
          saree.current_vendor_id = "";
          saree.status = "In Process";
        }
      }
      
      writeSheetData(sareesSheet, sarees, ["saree_id", "lot_number", "design_name", "quantity", "current_stage", "current_vendor_id", "status", "remarks", "created_at"]);
      return jsonResponse({ success: true, message: "Workflow rolled back successfully" });
      
    } else if (action === "deleteSaree") {
      var sareeId = postData.sareeId;
      var sareesSheet = userSS.getSheetByName("sarees");
      var sarees = getSheetData(sareesSheet);
      
      var filtered = sarees.filter(function(s) { return String(s.saree_id) !== String(sareeId); });
      writeSheetData(sareesSheet, filtered, ["saree_id", "lot_number", "design_name", "quantity", "current_stage", "current_vendor_id", "status", "remarks", "created_at"]);
      
      // Also delete history for this saree
      var historySheet = userSS.getSheetByName("workflow_history");
      var history = getSheetData(historySheet);
      var filteredHistory = history.filter(function(h) { return String(h.saree_id) !== String(sareeId); });
      writeSheetData(historySheet, filteredHistory, ["history_id", "saree_id", "stage_name", "vendor_id", "sent_date", "received_date", "work_cost", "remarks", "updated_by", "created_at"]);
      
      return jsonResponse({ success: true, message: "Saree lot and history deleted" });
      
    } else if (action === "addPayment") {
      var paymentsSheet = userSS.getSheetByName("payments");
      var payments = getSheetData(paymentsSheet);
      var paymentId = payments.length > 0 ? Math.max.apply(null, payments.map(function(p) { return parseInt(p.payment_id); })) + 1 : 1;
      
      var newPayment = {
        payment_id: paymentId,
        vendor_id: postData.vendor_id,
        payment_date: postData.payment_date || new Date().toISOString(),
        amount: parseFloat(postData.amount),
        payment_method: postData.payment_method,
        remarks: postData.remarks || "",
        created_by: postData.createdBy || "",
        created_at: new Date().toISOString()
      };
      
      appendRow(paymentsSheet, newPayment);
      return jsonResponse({ success: true, payment: newPayment });
      
    } else if (action === "deletePayment") {
      var paymentId = postData.paymentId;
      var paymentsSheet = userSS.getSheetByName("payments");
      var payments = getSheetData(paymentsSheet);
      
      var filtered = payments.filter(function(p) { return String(p.payment_id) !== String(paymentId); });
      writeSheetData(paymentsSheet, filtered, ["payment_id", "vendor_id", "payment_date", "amount", "payment_method", "remarks", "created_by", "created_at"]);
      return jsonResponse({ success: true, message: "Payment deleted" });
    }
    
    return jsonResponse({ success: false, error: "Invalid action" });
    
  } catch (error) {
    return jsonResponse({ success: false, error: error.toString() });
  }
}

// Admin function to initialize the main Admin Spreadsheet (manually run once or automatically)
function initializeAdminSpreadsheet() {
  var ss = getAdminSpreadsheet();
  
  var usersSheet = ss.getSheetByName("users");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("users");
    usersSheet.appendRow(["id", "name", "mobile", "email", "password_hash", "role", "active", "spreadsheet_id", "created_at"]);
    usersSheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#d9ead3");
    
    // Insert a default Admin user (password: admin123)
    var defaultAdmin = [
      1,
      "Super Admin",
      "9999999999",
      "admin@threadtrack.com",
      sha256("admin123"),
      "Admin",
      true,
      "",
      new Date().toISOString()
    ];
    usersSheet.appendRow(defaultAdmin);
  }
  
  var logsSheet = ss.getSheetByName("audit_logs");
  if (!logsSheet) {
    logsSheet = ss.insertSheet("audit_logs");
    logsSheet.appendRow(["log_id", "user_id", "user_name", "action", "details", "ip_address", "created_at"]);
    logsSheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#f3ecf9");
  }
  
  var settingsSheet = ss.getSheetByName("settings");
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("settings");
    settingsSheet.appendRow(["key", "value"]);
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#fff2cc");
    
    // Add default settings
    settingsSheet.appendRow(["template_spreadsheet_id", ""]);
  }
}
