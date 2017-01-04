
## Setup and prep working

* set working directory

* set path to file directory

* find excel files

* compare to extracted file log

* output list of file to extract

## Extract info from Excel reports

* Extract info

...

* Store and clean in ExcelDataset

* Dataset stucture

Inspection date | Report ref | ReportType | MO ref | Actual quantity | PCB Side | A1defectLoc | A1defectQty | A2defect | ... | DotQty |

* log in extracted file log

## Cross-check and enhanced with SAP database information

* match excel info with SAP datbase info (planned MO quantity, part number and product)

* Store and clean in FinalDataset

* ... list variables, description and class in CodeBook

* Final Dataset structure

Part Number | PCB name | Product name | Planned MO Qty | Inspection date | Report ref | ReportType | MO ref | Actual quantity | PCB Side | A1defectLoc | A1defectQty | A2defect | ... | DotQty |

## Computation and summaries

* Mutate to calculate PPM
* Group by month, product, PCB side
* Summarize_all by summation

## Plotting

... need to define what to see
