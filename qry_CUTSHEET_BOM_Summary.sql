-- =============================================================================
-- qry_CUTSHEET_BOM_Summary
-- =============================================================================
-- Aggregates tbl_tmp_CUTSHEET_BOM (bom_MECH) per Serial.
--
-- ASSUMPTIONS / PREREQUISITES
--   1. tbl_tmp_CUTSHEET_BOM must have a [Total Cost] field (Numeric).
--      Add it as a computed/stored column, or populate it via a look-up query
--      before running this query.
--   2. [Purch] is DERIVED here as Department IN ('STOCK','SPECIAL ORDER').
--      If a separate Purch boolean column is added later, replace the
--      Department-based condition with  b.Purch = TRUE / FALSE.
--   3. Left(PART_NUMBER, 1) = 'A'  identifies assembly rows (excluded from
--      cost/qty totals, matching the Excel  "<>A*"  filter).
--   4. StockedPartNumber starting with '8' or '9'  → internal MFG stock.
--      All other STOCK rows → purchased (external) stock.
--
-- COLUMNS PRODUCED
--   STOCK_PURCH_QTY / COST   – purchased stock (excl. 8x/9x internal parts)
--   SP_QTY / SP_COST         – Special Order items
--   MECH_COST                – all non-purchased, non-assembly parts
--   LATHE / MILL / PUNCH / ROUTER  _QTY and _COST
--   STOCK_MFG_QTY / COST     – internal mfg stock (StockedPartNumber 8x/9x)
--   TOTAL_QTY                – all non-assembly part quantities
--   MFG_ZERO_QTY             – mfg-dept parts where Total Cost = 0 (need est.)
--   AVG_MFG_COST             – average cost across mfg depts (for $0 estimate)
-- =============================================================================

SELECT
    b.Serial,

    -- ── STOCK (PURCH) ──────────────────────────────────────────────────────────
    -- Department = 'STOCK', StockedPartNumber does NOT start with '8' or '9'
    SUM(IIF(b.Department = 'STOCK'
        AND Left(Trim(Nz(b.StockedPartNumber, '')), 1) NOT IN ('8', '9'),
        Nz(b.QTY, 0), 0))                                     AS STOCK_PURCH_QTY,

    SUM(IIF(b.Department = 'STOCK'
        AND Left(Trim(Nz(b.StockedPartNumber, '')), 1) NOT IN ('8', '9'),
        Nz(b.[Total Cost], 0), 0))                            AS STOCK_PURCH_COST,

    -- ── SPECIAL ORDER ──────────────────────────────────────────────────────────
    SUM(IIF(b.Department = 'SPECIAL ORDER',
        Nz(b.QTY, 0), 0))                                     AS SP_QTY,

    SUM(IIF(b.Department = 'SPECIAL ORDER',
        Nz(b.[Total Cost], 0), 0))                            AS SP_COST,

    -- ── MECH COST ──────────────────────────────────────────────────────────────
    -- Non-purchased (not STOCK / SPECIAL ORDER), non-assembly (part# not 'A*')
    SUM(IIF(b.Department NOT IN ('STOCK', 'SPECIAL ORDER')
        AND Left(Nz(b.PART_NUMBER, ''), 1) <> 'A',
        Nz(b.[Total Cost], 0), 0))                            AS MECH_COST,

    -- ── LATHE ──────────────────────────────────────────────────────────────────
    SUM(IIF(b.Department = 'LATHE',
        Nz(b.QTY, 0), 0))                                     AS LATHE_QTY,

    SUM(IIF(b.Department = 'LATHE'
        AND Left(Nz(b.PART_NUMBER, ''), 1) <> 'A',
        Nz(b.[Total Cost], 0), 0))                            AS LATHE_COST,

    -- ── MILL ───────────────────────────────────────────────────────────────────
    SUM(IIF(b.Department = 'MILL',
        Nz(b.QTY, 0), 0))                                     AS MILL_QTY,

    SUM(IIF(b.Department = 'MILL'
        AND Left(Nz(b.PART_NUMBER, ''), 1) <> 'A',
        Nz(b.[Total Cost], 0), 0))                            AS MILL_COST,

    -- ── PUNCH ──────────────────────────────────────────────────────────────────
    SUM(IIF(b.Department = 'PUNCH',
        Nz(b.QTY, 0), 0))                                     AS PUNCH_QTY,

    SUM(IIF(b.Department = 'PUNCH'
        AND Left(Nz(b.PART_NUMBER, ''), 1) <> 'A',
        Nz(b.[Total Cost], 0), 0))                            AS PUNCH_COST,

    -- ── ROUTER ─────────────────────────────────────────────────────────────────
    SUM(IIF(b.Department = 'ROUTER',
        Nz(b.QTY, 0), 0))                                     AS ROUTER_QTY,

    SUM(IIF(b.Department = 'ROUTER'
        AND Left(Nz(b.PART_NUMBER, ''), 1) <> 'A',
        Nz(b.[Total Cost], 0), 0))                            AS ROUTER_COST,

    -- ── STOCK (MFG) ────────────────────────────────────────────────────────────
    -- Department = 'STOCK', StockedPartNumber starts with '8' or '9'
    SUM(IIF(b.Department = 'STOCK'
        AND Left(Trim(Nz(b.StockedPartNumber, '')), 1) IN ('8', '9'),
        Nz(b.QTY, 0), 0))                                     AS STOCK_MFG_QTY,

    SUM(IIF(b.Department = 'STOCK'
        AND Left(Trim(Nz(b.StockedPartNumber, '')), 1) IN ('8', '9'),
        Nz(b.[Total Cost], 0), 0))                            AS STOCK_MFG_COST,

    -- ── TOTAL QTY (non-assembly) ────────────────────────────────────────────────
    SUM(IIF(Left(Nz(b.PART_NUMBER, ''), 1) <> 'A',
        Nz(b.QTY, 0), 0))                                     AS TOTAL_QTY,

    -- ── MFG $0 QTY ─────────────────────────────────────────────────────────────
    -- Mfg-dept rows where Total Cost = 0; used to estimate missing costs
    SUM(IIF(Nz(b.[Total Cost], 0) = 0
        AND b.Department IN ('PUNCH', 'LATHE', 'MILL', 'ROUTER', 'SHEAR'),
        Nz(b.QTY, 0), 0))                                     AS MFG_ZERO_QTY

FROM tbl_tmp_CUTSHEET_BOM AS b
GROUP BY b.Serial;


-- =============================================================================
-- qry_CUTSHEET_BOM_Derived
-- =============================================================================
-- Builds on qry_CUTSHEET_BOM_Summary to add derived/rolled-up columns.
-- Save qry_CUTSHEET_BOM_Summary as a query in Access first, then save this
-- as qry_CUTSHEET_BOM_Derived referencing it.
--
-- COLUMNS ADDED
--   TOTAL_PURCH_COST   = STOCK_PURCH_COST + SP_COST
--   MFG_COST           = LATHE + MILL + PUNCH + ROUTER + STOCK_MFG costs
--   MFG_QTY            = LATHE + MILL + PUNCH + ROUTER + STOCK_MFG qtys
--   AVG_MFG_COST       = MFG_COST / MFG_QTY  (0 when MFG_QTY = 0)
--   EST_ZERO_COST      = MFG_ZERO_QTY * AVG_MFG_COST
--   BOM_MATERIAL_COST  = TOTAL_PURCH_COST + MECH_COST + MFG_COST + EST_ZERO_COST
--                        (excludes ELE_COST — add that from bom_ELE separately)
-- =============================================================================

SELECT
    s.Serial,

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

    -- ── Rolled-up totals ───────────────────────────────────────────────────────
    s.STOCK_PURCH_COST + s.SP_COST
        AS TOTAL_PURCH_COST,

    s.LATHE_COST + s.MILL_COST + s.PUNCH_COST + s.ROUTER_COST + s.STOCK_MFG_COST
        AS MFG_COST,

    s.LATHE_QTY + s.MILL_QTY + s.PUNCH_QTY + s.ROUTER_QTY + s.STOCK_MFG_QTY
        AS MFG_QTY,

    -- ── AVG_MFG_COST ───────────────────────────────────────────────────────────
    IIF(
        (s.LATHE_QTY + s.MILL_QTY + s.PUNCH_QTY + s.ROUTER_QTY + s.STOCK_MFG_QTY) = 0,
        0,
        (s.LATHE_COST + s.MILL_COST + s.PUNCH_COST + s.ROUTER_COST + s.STOCK_MFG_COST)
        / (s.LATHE_QTY + s.MILL_QTY + s.PUNCH_QTY + s.ROUTER_QTY + s.STOCK_MFG_QTY)
    )                                                           AS AVG_MFG_COST,

    -- ── EST_ZERO_COST = MFG_ZERO_QTY * AVG_MFG_COST ───────────────────────────
    IIF(
        (s.LATHE_QTY + s.MILL_QTY + s.PUNCH_QTY + s.ROUTER_QTY + s.STOCK_MFG_QTY) = 0,
        0,
        s.MFG_ZERO_QTY
        * (s.LATHE_COST + s.MILL_COST + s.PUNCH_COST + s.ROUTER_COST + s.STOCK_MFG_COST)
        / (s.LATHE_QTY + s.MILL_QTY + s.PUNCH_QTY + s.ROUTER_QTY + s.STOCK_MFG_QTY)
    )                                                           AS EST_ZERO_COST,

    -- ── BOM_MATERIAL_COST (partial — excludes ELE_COST from bom_ELE) ───────────
    -- Full MATERIAL_COST = BOM_MATERIAL_COST + ELE_COST  (add from bom_ELE)
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
