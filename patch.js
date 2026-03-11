const fs = require('fs');

let codeGs = fs.readFileSync('Code.gs', 'utf8');

// Update Rent_Collection headers
codeGs = codeGs.replace(
  "['Bill ID', 'TenantID', 'Name', 'Month', 'Rent Amount', 'EB Amount', 'Total Amount', 'Status', 'Payment Mode', 'Payment Date']",
  "['Bill ID', 'TenantID', 'Name', 'Month', 'Rent Amount', 'EB Amount', 'Total Amount', 'Status', 'Payment Mode', 'Payment Date', 'Previous Reading', 'Current Reading', 'Units']"
);

// Update recordMeterReading to save reading details
codeGs = codeGs.replace(
  "      'Unpaid',\n      '',\n      ''\n    ];",
  "      'Unpaid',\n      '',\n      '',\n      tenantInfo.previousReading,\n      currReadingNum,\n      unitsConsumed\n    ];"
);

fs.writeFileSync('Code.gs', codeGs);

let indexHtml = fs.readFileSync('Index.html', 'utf8');

// Fix global variable leakage
indexHtml = indexHtml.replace(/tr = document\.createElement\('tr'\);/g, "const tr = document.createElement('tr');");

// Add Subcategory and Account to Expenses Form
indexHtml = indexHtml.replace(
  '<div class="form-group"><label class="form-label">Purpose</label><input type="text" class="custom-input" id="e_purpose" required></div>',
  '<div class="form-group"><label class="form-label">Subcategory</label><input type="text" class="custom-input" id="e_subcat"></div>\n                <div class="form-group"><label class="form-label">Purpose</label><input type="text" class="custom-input" id="e_purpose" required></div>'
);

indexHtml = indexHtml.replace(
  '<div class="form-group"><label class="form-label">Mode of Payment</label><input type="text" class="custom-input" id="e_mop"></div>',
  '<div class="form-group"><label class="form-label">Mode of Payment</label><input type="text" class="custom-input" id="e_mop"></div>\n                <div class="form-group"><label class="form-label">Account</label><input type="text" class="custom-input" id="e_account"></div>'
);

// Update expense submission payload
indexHtml = indexHtml.replace(
  "      subcategory: '',",
  "      subcategory: document.getElementById('e_subcat') ? document.getElementById('e_subcat').value : '',"
);
indexHtml = indexHtml.replace(
  "      account: ''",
  "      account: document.getElementById('e_account') ? document.getElementById('e_account').value : ''"
);

// Update WhatsApp message format to include requested details
indexHtml = indexHtml.replace(
  /const message = `\*BILL SUMMARY - \$\{bill.Month\}\*[\s\S]*?Please pay at your earliest convenience. Thank you!`;/,
  `const message = \`*BILL SUMMARY - \${bill.Month}*
Bill ID: \${bill['Bill ID']}
Name: \${bill.Name}

Previous Reading: \${bill['Previous Reading'] || 0}
Current Reading: \${bill['Current Reading'] || 0}
Units Consumed: \${bill['Units'] || 0}

Rent: ₹\${bill['Rent Amount']}
Electricity: ₹\${bill['EB Amount']}
-------------------------
*Total Due: ₹\${bill['Total Amount']}*
-------------------------
Please pay at your earliest convenience. Thank you!\`;`
);

fs.writeFileSync('Index.html', indexHtml);
console.log("Patch applied successfully");
