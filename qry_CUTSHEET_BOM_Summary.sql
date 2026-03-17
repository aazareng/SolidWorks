/* qry_CUTSHEET_BOM_Summary
   Aggregates tbl_tmp_CUTSHEET_BOM (bom_MECH equivalent) for a single serial.
   Total Cost = RawCost * QTY (computed inline).
   Purch derived as Department IN ('STOCK','SPECIAL ORDER').
   PART_NUMBER starting with 'A' = assembly row, excluded from cost/qty totals.
   StockedPartNumber starting with '8' or '9' = internal MFG stock. */
SELECT
    SUM(IIF(b.Department = 'STOCK'
        AND Left(Trim(Nz(b.StockedPartNumber, '')), 1) NOT IN ('8', '9'),
        Nz(b.QTY, 0), 0))                                     AS STOCK_PURCH_QTY,
    SUM(IIF(b.Department = 'STOCK'
        AND Left(Trim(Nz(b.StockedPartNumber, '')), 1) NOT IN ('8', '9'),
        Nz(b.RawCost, 0) * Nz(b.QTY, 0), 0))                 AS STOCK_PURCH_COST,
    SUM(IIF(b.Department = 'SPECIAL ORDER',
        Nz(b.QTY, 0), 0))                                     AS SP_QTY,
    SUM(IIF(b.Department = 'SPECIAL ORDER',
        Nz(b.RawCost, 0) * Nz(b.QTY, 0), 0))                 AS SP_COST,
    SUM(IIF(b.Department NOT IN ('STOCK', 'SPECIAL ORDER')
        AND Left(Nz(b.PART_NUMBER, ''), 1) <> 'A',
        Nz(b.RawCost, 0) * Nz(b.QTY, 0), 0))                 AS MECH_COST,
    SUM(IIF(b.Department = 'LATHE',
        Nz(b.QTY, 0), 0))                                     AS LATHE_QTY,
    SUM(IIF(b.Department = 'LATHE'
        AND Left(Nz(b.PART_NUMBER, ''), 1) <> 'A',
        Nz(b.RawCost, 0) * Nz(b.QTY, 0), 0))                 AS LATHE_COST,
    SUM(IIF(b.Department = 'MILL',
        Nz(b.QTY, 0), 0))                                     AS MILL_QTY,
    SUM(IIF(b.Department = 'MILL'
        AND Left(Nz(b.PART_NUMBER, ''), 1) <> 'A',
        Nz(b.RawCost, 0) * Nz(b.QTY, 0), 0))                 AS MILL_COST,
    SUM(IIF(b.Department = 'PUNCH',
        Nz(b.QTY, 0), 0))                                     AS PUNCH_QTY,
    SUM(IIF(b.Department = 'PUNCH'
        AND Left(Nz(b.PART_NUMBER, ''), 1) <> 'A',
        Nz(b.RawCost, 0) * Nz(b.QTY, 0), 0))                 AS PUNCH_COST,
    SUM(IIF(b.Department = 'ROUTER',
        Nz(b.QTY, 0), 0))                                     AS ROUTER_QTY,
    SUM(IIF(b.Department = 'ROUTER'
        AND Left(Nz(b.PART_NUMBER, ''), 1) <> 'A',
        Nz(b.RawCost, 0) * Nz(b.QTY, 0), 0))                 AS ROUTER_COST,
    SUM(IIF(b.Department = 'STOCK'
        AND Left(Trim(Nz(b.StockedPartNumber, '')), 1) IN ('8', '9'),
        Nz(b.QTY, 0), 0))                                     AS STOCK_MFG_QTY,
    SUM(IIF(b.Department = 'STOCK'
        AND Left(Trim(Nz(b.StockedPartNumber, '')), 1) IN ('8', '9'),
        Nz(b.RawCost, 0) * Nz(b.QTY, 0), 0))                 AS STOCK_MFG_COST,
    SUM(IIF(Left(Nz(b.PART_NUMBER, ''), 1) <> 'A',
        Nz(b.QTY, 0), 0))                                     AS TOTAL_QTY,
    SUM(IIF(Nz(b.RawCost, 0) = 0
        AND b.Department IN ('PUNCH', 'LATHE', 'MILL', 'ROUTER', 'SHEAR'),
        Nz(b.QTY, 0), 0))                                     AS MFG_ZERO_QTY
FROM tbl_tmp_CUTSHEET_BOM AS b;


/* qry_CUTSHEET_BOM_Derived
   Builds on qry_CUTSHEET_BOM_Summary to add rolled-up columns.
   Save qry_CUTSHEET_BOM_Summary as a query in Access first, then save
   this as qry_CUTSHEET_BOM_Derived referencing it.
   BOM_MATERIAL_COST excludes ELE_COST - add that from bom_ELE separately. */
SELECT
    s.STOCK_PURCH_QTY,
    s.STOCK_PURCH_COST,
    s.SP_QTY,
    s.SP_COST,
    s.MECH_COST,
    s.LATHE_QTY,
    s.LATHE_COST,
    s.MILL_QTY,
    s.MILL_COST,
    s.PUNCH_QTY,
    s.PUNCH_COST,
    s.ROUTER_QTY,
    s.ROUTER_COST,
    s.STOCK_MFG_QTY,
    s.STOCK_MFG_COST,
    s.TOTAL_QTY,
    s.MFG_ZERO_QTY,
    s.STOCK_PURCH_COST + s.SP_COST
        AS TOTAL_PURCH_COST,
    s.LATHE_COST + s.MILL_COST + s.PUNCH_COST + s.ROUTER_COST + s.STOCK_MFG_COST
        AS MFG_COST,
    s.LATHE_QTY + s.MILL_QTY + s.PUNCH_QTY + s.ROUTER_QTY + s.STOCK_MFG_QTY
        AS MFG_QTY,
    IIF(
        (s.LATHE_QTY + s.MILL_QTY + s.PUNCH_QTY + s.ROUTER_QTY + s.STOCK_MFG_QTY) = 0,
        0,
        (s.LATHE_COST + s.MILL_COST + s.PUNCH_COST + s.ROUTER_COST + s.STOCK_MFG_COST)
        / (s.LATHE_QTY + s.MILL_QTY + s.PUNCH_QTY + s.ROUTER_QTY + s.STOCK_MFG_QTY)
    )                                                           AS AVG_MFG_COST,
    IIF(
        (s.LATHE_QTY + s.MILL_QTY + s.PUNCH_QTY + s.ROUTER_QTY + s.STOCK_MFG_QTY) = 0,
        0,
        s.MFG_ZERO_QTY
        * (s.LATHE_COST + s.MILL_COST + s.PUNCH_COST + s.ROUTER_COST + s.STOCK_MFG_COST)
        / (s.LATHE_QTY + s.MILL_QTY + s.PUNCH_QTY + s.ROUTER_QTY + s.STOCK_MFG_QTY)
    )                                                           AS EST_ZERO_COST,
    (s.STOCK_PURCH_COST + s.SP_COST)
    + s.MECH_COST
    + (s.LATHE_COST + s.MILL_COST + s.PUNCH_COST + s.ROUTER_COST + s.STOCK_MFG_COST)
    + IIF(
        (s.LATHE_QTY + s.MILL_QTY + s.PUNCH_QTY + s.ROUTER_QTY + s.STOCK_MFG_QTY) = 0,
        0,
        s.MFG_ZERO_QTY
        * (s.LATHE_COST + s.MILL_COST + s.PUNCH_COST + s.ROUTER_COST + s.STOCK_MFG_COST)
        / (s.LATHE_QTY + s.MILL_QTY + s.PUNCH_QTY + s.ROUTER_QTY + s.STOCK_MFG_QTY)
    )                                                           AS BOM_MATERIAL_COST
FROM qry_CUTSHEET_BOM_Summary AS s;
