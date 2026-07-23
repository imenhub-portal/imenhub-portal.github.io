# Web App Integration: Printing Cost & Allocation Logic

This document outlines the calculation logic and financial routing structure to be integrated into your "Tempahan Printing" web application. When a payment (e.g., RM 35.00) is received, the system will automatically parse the revenue into 5 distinct pools (channels) to ensure cash flow sustainability.

## 1. System Constants & Variables

To make your web app flexible, store these as configurable variables in your database rather than hardcoding them, allowing you to update them if supplier prices change.

*   **Selling Price (Per A1):** RM 35.00
*   **Paper Cost (Per A1):** RM 6.67 *(Calculation: RM 300 / 45 prints)*
*   **Ink Cost (Per A1):** RM 20.00 *(Calculation: 2.0ml usage at RM 10/ml)*
*   **Printhead Depreciation (Per A1):** RM 0.50 *(Calculation: RM 2500 / 5000 prints)*
*   **Maintenance Buffer (Per A1):** RM 2.83
*   **Net Profit (Per A1):** RM 5.00

[Unverified] The ink cost of RM 20.00 per print is based on an estimated 2.0ml usage for heavy color coverage; you should allow the web app admin to adjust this constant after monitoring actual yields.

---

## 2. The 5 Allocation Channels

For every successful A1 poster transaction, the system should log the split into the following ledger channels:

### Channel 1: Paper Fund
*   **Allocation:** RM 6.67 per print
*   **Target Cycle:** RM 300.00 (Every ~45 prints)
*   **Purpose:** Immediate replenishment reserve. Since paper is the most frequent consumable to run out, this channel ensures you always have exact funds ready to purchase the next 30-meter roll without touching your profits.

### Channel 2: Ink Fund
*   **Allocation:** RM 20.00 per print
*   **Target Cycle:** RM 3,000.00 (Every ~150 prints)
*   **Purpose:** Mid-term consumable reserve. This incrementally sets aside money so that when the cartridges run empty, the substantial RM 3,000 replacement cost is fully funded, preventing sudden out-of-pocket cash flow gaps.

### Channel 3: Maintenance & Wear-and-Tear
*   **Allocation:** RM 2.83 per print
*   **Target Cycle:** Ongoing continuous reserve
*   **Purpose:** Operational overhead buffer. This pool covers miscellaneous expenses such as printer cleaning supplies, minor mechanical wear, periodic servicing, electricity overhead, or the cost of unexpected waste/misprints.

### Channel 4: Printhead Fund
*   **Allocation:** RM 0.50 per print
*   **Target Cycle:** RM 2,500.00 (Every 5,000 prints)
*   **Purpose:** Long-term hardware depreciation. It safely tucks away a micro-amount per print over a long period. By the time the printer signals a required printhead replacement, the RM 2,500 cost is already available.

### Channel 5: Net Profit Margin
*   **Allocation:** The remainder (RM 5.00 per print at current pricing)
*   **Target Cycle:** Immediate / Monthly withdrawal
*   **Purpose:** Actual business earnings. This is your take-home pay or capital for business expansion, realized *after* all consumables, hardware depreciation, and maintenance buffers are fully covered.

---

## 3. Example JavaScript Implementation

If you are building this logic in JavaScript, here is a simple function to handle dynamic pricing and routing into the 5 pools:

```javascript
function calculateAllocation(sellingPrice, quantity = 1) {
    // System Constants (Fetch from your DB/Settings in production)
    const costPaper = 6.67;
    const costInk = 20.00;
    const costPrinthead = 0.50;
    const costMaintenance = 2.83;

    // Total base cost
    const totalBaseCost = costPaper + costInk + costPrinthead + costMaintenance;

    // Calculate Profit
    const profit = sellingPrice - totalBaseCost;

    // Return allocation object for the specific quantity
    return {
        totalRevenue: (sellingPrice * quantity).toFixed(2),
        allocations: {
            paperFund: (costPaper * quantity).toFixed(2),
            inkFund: (costInk * quantity).toFixed(2),
            maintenanceFund: (costMaintenance * quantity).toFixed(2),
            printheadFund: (costPrinthead * quantity).toFixed(2),
            netProfit: (profit * quantity).toFixed(2)
        }
    };
}

// Example usage for an order of 3 A1 posters at RM 35 each:
// console.log(calculateAllocation(35.00, 3));
```
