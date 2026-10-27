# Warehouse Order Tracking & Reporting Automation

**Excel VBA | Microsoft Excel | Operational Reporting | Data Consolidation**

## Project Overview
**Role:** Warehouse Business Analyst  
**Organization:** Al Khayyat Investments  
**Tools:** Microsoft Excel, VBA Macros, Mnahattan Warehouse Management System (WMS) exports, Paragon Exports
**Focus:** Process Automation, Operational Reporting, Data Cleaning, Fulfilment Monitoring

During my time supporting warehouse operations at Al Khayyat Investments, I worked on an Excel VBA-based reporting process designed to consolidate operational data from multiple warehouse and transportation sources into a standardized order-tracking report.

The process brought together order information, transportation plans, picking progress, quality assurance checks, and invoice handover records to provide a consolidated view of order fulfilment.

The report was refreshed every 30 minutes to reflect updates in the underlying operational files and support timely monitoring of warehouse activities.

**Portfolio note:** This repository contains a reconstructed demonstration based on my recollection of the original workflow. The original macro source code and restricted operational input files are no longer available to me. All demonstration records are fictional, and the code developed for this portfolio is an independent recreation rather than the original company code.

## Business Problem

Warehouse order tracking involved information maintained across multiple operational files and teams. Preparing a consolidated report manually required collecting, matching, and reviewing information from these different sources.

This created several challenges:

- Repetitive report preparation and data consolidation.
- Difficulty obtaining a unified view of order progress across operational stages.
- The need to match records consistently across separate workbooks.
- Delays in producing an updated report for operational monitoring.

The objective was to reduce repetitive reporting effort and provide a more consistent, consolidated view of order status.

## Data Sources

The original reporting process used five primary operational sources:

| Source | Information used |
|---|---|
| Manhattan WMS | Order, shipment, customer, route, and order-detail information |
| Paragon transportation file | Refined route and shipping information, including next-day delivery planning |
| Outbound picking file | Picking status, picker information, picking times, and quantities |
| Quality Assurance (QA) file | QA completion status and checks on picked orders |
| Invoice reconciliation file | Reconciliation between Outbound and Transportation, including invoice handover status |

A separate cut-off report provided the relevant 4:00 p.m. cut-off classification. The reconstructed demonstration models this as a separate input rather than treating it as a field sourced directly from Manhattan WMS.

## Operational Workflow

The reporting process followed the warehouse order-fulfilment workflow:

1. **Order information:** Order and shipment details were obtained from Manhattan WMS.
2. **Transportation planning:** The transportation team's Paragon file supplied refined route and delivery-planning information.
3. **Picking:** Outbound staff picked the required items and recorded picking progress.
4. **Quality assurance:** QA staff checked whether the picked order matched the request. Where an order was incomplete, QA could review the picking location and help complete it.
5. **Invoice reconciliation and handover:** The invoice reconciliation process tracked the movement of invoices between Outbound and Transportation.
6. **Consolidated reporting:** The reporting process brought the available records together into a detailed order-tracking report and summary.

For this reporting process, invoice handover to Transportation served as the indicator that an order had reached the relevant fulfilment handover stage. A picked order was not automatically treated as fulfilled.

## Automation Approach

The original solution used an Excel VBA macro-enabled workbook to automate repetitive reporting tasks.

The reconstructed portfolio implementation is designed around the following approach:

- **Data consolidation:** Combine records from the five operational inputs and the separate cut-off report.
- **Record matching:** Use Order Number as the primary key for matching corresponding records across sources.
- **Shipment identifier parsing:** Extract the Order Number from the first seven characters of a shipment identifier, while retaining the department code, six-digit date segment, and order-type suffix as separate fields.
- **Order-type classification:** Recognize suffixes such as `SO` for Sales Order and `IR` for Van Internal Order.
- **Status visibility:** Preserve distinctions between picking progress, QA completion, invoice generation, and invoice handover.
- **Standardized output:** Produce a consolidated order-tracking report and a summary for operational review.
- **Fully automated refresh:** The original VBA reporting solution refreshed automatically every 30 minutes, reducing the need for manual report-refresh actions..

The Paragon file was an input to the reporting process; route planning itself was performed by the transportation team, not by the VBA automation.

## Reporting Outputs

The consolidated report supported review of order-level information such as:

- Order number and customer
- Area and route
- Shipment and invoice details
- Picking quantities and progress
- Wave information and order type
- Invoice generation and handover information
- Relevant order status indicators
- The `Picked & IDT Completed` metric represented orders for which all three conditions were satisfied: picking was completed, QA was marked Done, and the invoice had been handed over to Transportation. An order being picked alone was not sufficient to qualify.

The accompanying summary provided a more concise view of operational progress and outstanding activities.

Management also used a separate Power BI dashboard to monitor warehouse performance, including daily order volumes, picking progress, invoice generation and handover, pending invoices, and orders by route. That dashboard was a separate reporting solution and was not solely developed by me as part of this VBA project.

## Business Impact

Based on my experience with the original process:

- **Before automation:** Report preparation typically took approximately 1–2 hours.
- **After automation:** Report preparation took under 15 minutes.
- **Estimated time reduction:** At least approximately 75%, depending on the starting preparation time.

These figures are estimates based on my recollection of the operational process, not the results of a formal time-and-motion study.

The automation reduced repetitive preparation work and made it easier to produce a consolidated view of order progress for operational monitoring.

## Reconstructed Demonstration

The portfolio version is intended to demonstrate the underlying analytical and automation concepts using synthetic data.

It will include:

- Fictional examples of the operational source files.
- A documented order-matching and shipment-parsing approach.
- Explicit status rules and assumptions.
- Recreated VBA logic for consolidating data.
- An expected-output sheet for validating the results.

Where the original business rules cannot be verified from the available materials, they will be identified as assumptions rather than presented as confirmed historical logic.

## Skills Demonstrated

- Excel reporting and data consolidation
- VBA automation design
- Multi-source data integration
- Record matching and data validation
- Operational process analysis
- Order-fulfilment reporting
- Requirements interpretation
- Business-focused documentation

## Confidentiality

This portfolio demonstration does not reproduce the original company's restricted source files, customer records, or proprietary VBA source code. Company-specific information is included only to describe my professional experience at a high level; the demonstration data and recreated implementation are intended for portfolio and learning purposes.
