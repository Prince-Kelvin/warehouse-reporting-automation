# Warehouse Reporting Automation Using Excel VBA

## Project Overview

**Role:** Warehouse Business Analyst  
**Organization:** Al Khayyat Investments  
**Tools:** Microsoft Excel, VBA Macros, Warehouse Management System (WMS) exports  
**Focus:** Process Automation, Operational Reporting, Data Cleaning, Fulfilment Monitoring

## 1. Business Context

Warehouse operations depended on daily reports extracted from the Warehouse Management System (WMS) to support order picking, dispatch preparation, and next-day delivery planning.

The warehouse operated with a 4:00 PM order cut-off. Orders received after the cut-off were treated as next-day orders, requiring the warehouse team to prepare outbound orders in advance.

Before the automation, the coordinator manually cleaned, consolidated, and structured multiple Excel exports to produce the daily operational report. This repetitive process took approximately 1–2 hours per reporting cycle.

## 2. Problem Statement

The manual reporting process created several operational challenges:

- Significant time spent cleaning and consolidating exported data.
- Repetitive preparation of the same report fields every working day.
- Delayed access to a standardized view of order fulfilment progress.
- Additional manual effort to organize completed, pending, and in-picking orders.
- Less time available for reviewing operational information and supporting fulfilment decisions.

## 3. My Role and Contribution

I designed and implemented a macro-enabled Excel solution using VBA to automate the daily reporting process.

My contribution included:

- Developing a VBA macro to process downloaded WMS Excel exports.
- Automating the cleaning, consolidation, and structuring of operational data.
- Standardizing the report layout and formatting.
- Incorporating source order statuses and workbook rules to support fulfilment-status tracking.
- Organizing order information into a consistent format for operational review.

## 4. Solution and Workflow

The automated workflow was designed to simplify report preparation:

1. **Export:** Download the daily order data from the WMS.
2. **Upload:** Load the exported Excel data into the automation workbook.
3. **Process:** Run the VBA macro to clean, consolidate, and structure the data.
4. **Standardize:** Generate the report using a consistent layout and formatting.
5. **Review:** Use the resulting report to monitor order fulfilment progress and prepare for next-day delivery operations.

### Key report fields

The standardized report included:

- Order number
- Customer name
- Delivery area and route
- Invoice number and invoice generation time
- Handover time to the transport team
- Quantity of items picked
- Wave number and the person who created the wave
- Order type
- Order fulfilment status

## 5. Business Impact

The solution improved the efficiency and consistency of daily warehouse reporting.

| Measure | Before automation | After automation |
|---|---|---|
| Report preparation time | Approximately 1–2 hours | Under 15 minutes |
| Data preparation | Manual cleaning and consolidation | VBA-assisted processing |
| Report structure | Manually prepared | Standardized layout |
| Fulfilment visibility | Manual report review | Structured status tracking |

**Estimated time reduction:** At least approximately 75%, based on the reported preparation times.

The automation also helped make completed, pending, and in-picking orders easier to review, supporting operational visibility and next-day delivery preparation.

*The time reduction is an estimate based on the reported before-and-after preparation times, not a formally measured time study.*

## 6. Skills Demonstrated

- Business process analysis
- Excel VBA automation
- Data cleaning and consolidation
- Operational reporting
- Requirements-to-solution translation
- Fulfilment process understanding
- Process improvement

## 7. Confidentiality and Portfolio Approach

This case study describes the solution at a high level. Any demonstration workbook, screenshots, or sample reports published with the project should use fictional data and exclude confidential company information, customer details, internal system exports, and proprietary code.

## Conclusion

This project demonstrates how I applied business analysis and Excel automation to simplify a recurring operational process, standardize warehouse reporting, and improve visibility into fulfilment activities.
