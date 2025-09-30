/* src/taskpane/taskpane.ts */

/* global console, document, Excel, Office, localStorage, fetch */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    init();
  }
});

// Custom message display function for Office Add-ins
function showMessage(message: string, type: "success" | "error" | "info" = "info") {
  const messageDiv = document.getElementById("message") as HTMLDivElement;
  if (messageDiv) {
    messageDiv.textContent = message;
    messageDiv.className = `message message-${type}`;
    messageDiv.style.display = "block";
    
    // Auto-hide after 5 seconds for success/info messages
    if (type === "success" || type === "info") {
      setTimeout(() => {
        messageDiv.style.display = "none";
      }, 5000);
    }
  }
}

// Data transformation layer
interface BillData {
  id: string;
  vendorName: string;
  amount: number;
  dueAmount: number;
  dueDate: string;
  invoiceNumber: string;
  invoiceDate?: string;
  // Add any other fields you want to track
}

interface TransformationRule {
  name: string;
  condition: (bill: any) => boolean;
  transform: (bill: any) => Partial<BillData>;
}

// Define your transformation rules here
const transformationRules: TransformationRule[] = [
  {
    name: "Format Vendor Names",
    condition: (bill) => true, // Apply to all bills
    transform: (bill) => ({
      vendorName: bill.vendorName?.toUpperCase() || "UNKNOWN VENDOR"
    })
  },
  {
    name: "Format Currency",
    condition: (bill) => true,
    transform: (bill) => ({
      amount: Math.round((bill.amount || 0) * 100) / 100, // Round to 2 decimal places
      dueAmount: Math.round((bill.dueAmount || 0) * 100) / 100
    })
  },
  {
    name: "Format Dates",
    condition: (bill) => bill.dueDate,
    transform: (bill) => ({
      dueDate: new Date(bill.dueDate).toLocaleDateString() // Convert to readable format
    })
  },
  {
    name: "High Value Bills Alert",
    condition: (bill) => (bill.amount || 0) > 1000,
    transform: (bill) => ({
      vendorName: `🚨 ${bill.vendorName} (HIGH VALUE)`
    })
  },
  {
    name: "Overdue Bills Alert",
    condition: (bill) => {
      if (!bill.dueDate) return false;
      const dueDate = new Date(bill.dueDate);
      const today = new Date();
      return dueDate < today;
    },
    transform: (bill) => ({
      vendorName: `⚠️ ${bill.vendorName} (OVERDUE)`
    })
  }
];

// Apply transformation rules to a single bill
function transformBill(originalBill: any): BillData {
  let transformedBill: BillData = {
    id: originalBill.id || "",
    vendorName: originalBill.vendorName || "",
    amount: originalBill.amount || 0,
    dueAmount: originalBill.dueAmount || 0,
    dueDate: originalBill.dueDate || "",
    invoiceNumber: originalBill.invoice?.invoiceNumber || "",
    invoiceDate: originalBill.invoice?.invoiceDate || ""
  };

  // Apply each transformation rule
  transformationRules.forEach(rule => {
    if (rule.condition(originalBill)) {
      const changes = rule.transform(originalBill);
      transformedBill = { ...transformedBill, ...changes };
    }
  });

  return transformedBill;
}

// Apply transformation rules to all bills
function transformBills(originalBills: any[]): BillData[] {
  console.log("Applying transformation rules to bills...");
  return originalBills.map(transformBill);
}

function init() {
  const callButton = document.getElementById("callApi") as HTMLButtonElement;
  const sessionIdInput = document.getElementById("sessionId") as HTMLInputElement;
  const devKeyInput = document.getElementById("devKey") as HTMLInputElement;
  
  const maxValue = 20;
  const maxPages = 10; // Safety limit to prevent infinite loops

  callButton.onclick = async () => {
    // Validate inputs
    const sessionId = sessionIdInput.value.trim();
    const devKey = devKeyInput.value.trim();
    
    if (!sessionId || !devKey) {
      showMessage("Please enter both Session ID and Developer Key.", "error");
      return;
    }

    const headers = {
      "Accept": "application/json",
      "Content-Type": "application/json",
      "sessionId": sessionId,
      "devKey": devKey
    };

    try {
      let billsUrl = "https://gateway.stage.bill.com/connect/v3/bills";
      let allBills: any[] = [];
      let pageCount = 0;

      // First API call
      console.log("Fetching bills from Bill.com API...");
      let response = await fetch(billsUrl, {
        method: 'GET',
        headers: headers
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      let result = await response.json();
      console.log("Initial API response:", result);

      if (result.results && result.results.length > 0) {
        console.log(`Results found: ${result.results.length}`);
        allBills.push(...result.results);

        let nextPage = result.nextPage;
        pageCount = 1;

        // Pagination loop with safety checks
        while (nextPage && pageCount < maxPages) {
          const nextBillsUrl = `https://gateway.stage.bill.com/connect/v3/bills?max=${maxValue}&page=${nextPage}`;
          console.log(`Fetching page ${pageCount + 1}...`);
          
          const nextResponse = await fetch(nextBillsUrl, {
            method: 'GET',
            headers: headers
          });

          if (!nextResponse.ok) {
            throw new Error(`HTTP error! status: ${nextResponse.status}`);
          }

          const nextResult = await nextResponse.json();
          console.log(`Page ${pageCount + 1} response:`, nextResult);

          if (nextResult.results) {
            allBills.push(...nextResult.results);
          }

          nextPage = nextResult.nextPage || null;
          pageCount++;
        }

        if (pageCount >= maxPages) {
          console.warn(`Reached maximum page limit (${maxPages}). Stopping pagination.`);
        }
      } else {
        console.log("No results found in the initial API request.");
      }

      // Output results to Excel
      if (allBills.length > 0) {
        console.log(`Total bills fetched: ${allBills.length}`);
        
        // Apply transformation rules
        const transformedBills = transformBills(allBills);
        console.log("Transformation complete. Sample transformed bill:", transformedBills[0]);
        
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          
          // Write headers
          sheet.getRange("A1:F1").values = [["Bill ID", "Vendor Name", "Amount", "Due Amount", "Due Date", "Invoice Number"]];
          
          // Write transformed data (limit to first 100 rows for performance)
          const displayBills = transformedBills.slice(0, 100);
          const excelData = displayBills.map(bill => [
            bill.id,
            bill.vendorName,
            bill.amount,
            bill.dueAmount,
            bill.dueDate,
            bill.invoiceNumber
          ]);
          
          if (excelData.length > 0) {
            sheet.getRange(`A2:F${excelData.length + 1}`).values = excelData;
          }
          
          await context.sync();
        });
        
        showMessage(`Successfully fetched ${allBills.length} bills, applied transformations, and wrote to Excel!`, "success");
      } else {
        showMessage("No bills found.", "info");
      }

    } catch (err) {
      console.error("API call failed", err);
      showMessage(`API call failed: ${err instanceof Error ? err.message : 'Unknown error'}`, "error");
    }
  };
}
