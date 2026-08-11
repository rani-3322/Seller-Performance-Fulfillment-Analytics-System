-- ============================================================================
-- Flipkart Fulfillment & Seller Performance Analytics — SQL Analysis Layer
-- SQLite dialect. Mirrors the logic in the Excel workbook's derived sheets
-- (Seller Scorecard, Warehouse Performance, Delivery SLA, Inventory Summary,
-- Dashboard) so the two artifacts can be cross-checked against each other.
--
-- Parameters used below (kept as literals here; in the Excel workbook these
-- live as named ranges on the README sheet so they can be changed in one place):
--   SLA target                 = 0.95
--   Scorecard weights          = On-Time Delivery 30%, Cancellation 20% (inverted),
--                                 Return Rate 20% (inverted), Inventory Accuracy 15%,
--                                 Customer Rating 15%
--   Tier cutoffs                = Gold >= 85, Silver >= 70, Bronze >= 55, else Watchlist
--   Safety-stock service Z      = 1.65  (95% service level)
--   Holding cost rate           = 1.5% of inventory value per month
--   On-time dispatch threshold  = dispatch within 1 day of order date
-- ============================================================================


-- ----------------------------------------------------------------------------
-- 1. SELLER SCORECARD
-- Weighted composite score per seller -> Gold / Silver / Bronze / Watchlist.
-- Mirrors the Seller Scorecard sheet exactly (same weights, same tier cutoffs).
-- ----------------------------------------------------------------------------
WITH seller_orders AS (
    SELECT
        seller_id,
        COUNT(*)                                                    AS total_lines,
        SUM(CASE WHEN order_status = 'Delivered' THEN 1 ELSE 0 END) AS delivered_lines,
        SUM(CASE WHEN order_status = 'Delivered'
                  AND delivery_date <= expected_delivery_date THEN 1 ELSE 0 END) AS ontime_count,
        SUM(CASE WHEN order_status = 'Cancelled' THEN 1 ELSE 0 END) AS cancelled_lines,
        SUM(CASE WHEN julianday(dispatch_date) - julianday(order_date) <= 1
                  THEN 1 ELSE 0 END)                                AS ontime_dispatch_count
    FROM orders
    GROUP BY seller_id
),
seller_returns AS (
    SELECT seller_id, COUNT(*) AS returns_count
    FROM returns
    GROUP BY seller_id
),
seller_metrics AS (
    SELECT
        s.seller_id,
        s.seller_name,
        s.seller_category,
        s.warehouse_assigned,
        s.seller_rating,
        COALESCE(so.total_lines, 0)      AS total_lines,
        COALESCE(so.delivered_lines, 0)  AS delivered_lines,
        CAST(COALESCE(so.ontime_count, 0) AS REAL)
            / NULLIF(so.delivered_lines, 0)                          AS ontime_rate,
        CAST(COALESCE(so.cancelled_lines, 0) AS REAL)
            / NULLIF(so.total_lines, 0)                               AS cancellation_rate,
        CAST(COALESCE(sr.returns_count, 0) AS REAL)
            / NULLIF(so.delivered_lines, 0)                           AS return_rate,
        CAST(COALESCE(so.ontime_dispatch_count, 0) AS REAL)
            / NULLIF(so.total_lines, 0)                                AS ontime_dispatch_rate,
        w.inventory_accuracy_audit
    FROM sellers s
    LEFT JOIN seller_orders  so ON so.seller_id = s.seller_id
    LEFT JOIN seller_returns sr ON sr.seller_id = s.seller_id
    JOIN warehouses w ON w.warehouse_id = s.warehouse_assigned
)
SELECT
    seller_id,
    seller_name,
    seller_category,
    warehouse_assigned,
    ROUND(COALESCE(ontime_rate, 0), 4)        AS ontime_rate,
    ROUND(COALESCE(cancellation_rate, 0), 4)  AS cancellation_rate,
    ROUND(COALESCE(return_rate, 0), 4)        AS return_rate,
    inventory_accuracy_audit,
    seller_rating,
    ROUND(COALESCE(ontime_dispatch_rate, 0), 4) AS ontime_dispatch_rate,
    ROUND(
        (COALESCE(ontime_rate, 0)              * 0.30
       + (1 - COALESCE(cancellation_rate, 0))  * 0.20
       + (1 - COALESCE(return_rate, 0))        * 0.20
       + inventory_accuracy_audit              * 0.15
       + (seller_rating / 5.0)                 * 0.15
        ) * 100, 1)                             AS seller_score,
    CASE
        WHEN (COALESCE(ontime_rate, 0)             * 0.30
            + (1 - COALESCE(cancellation_rate, 0)) * 0.20
            + (1 - COALESCE(return_rate, 0))       * 0.20
            + inventory_accuracy_audit             * 0.15
            + (seller_rating / 5.0)                * 0.15) * 100 >= 85 THEN 'Gold'
        WHEN (COALESCE(ontime_rate, 0)             * 0.30
            + (1 - COALESCE(cancellation_rate, 0)) * 0.20
            + (1 - COALESCE(return_rate, 0))       * 0.20
            + inventory_accuracy_audit             * 0.15
            + (seller_rating / 5.0)                * 0.15) * 100 >= 70 THEN 'Silver'
        WHEN (COALESCE(ontime_rate, 0)             * 0.30
            + (1 - COALESCE(cancellation_rate, 0)) * 0.20
            + (1 - COALESCE(return_rate, 0))       * 0.20
            + inventory_accuracy_audit             * 0.15
            + (seller_rating / 5.0)                * 0.15) * 100 >= 55 THEN 'Bronze'
        ELSE 'Watchlist'
    END AS seller_tier
FROM seller_metrics
ORDER BY seller_score DESC;


-- ----------------------------------------------------------------------------
-- 2. WAREHOUSE PERFORMANCE
-- Utilization, throughput, dispatch speed and SLA% per fulfillment centre.
-- ----------------------------------------------------------------------------
WITH occupied AS (
    SELECT
        p.seller_id,
        s.warehouse_assigned AS warehouse_id,
        it.sku,
        SUM(CASE WHEN it.txn_type = 'Inbound'  THEN it.quantity ELSE 0 END)
      - SUM(CASE WHEN it.txn_type = 'Outbound' THEN it.quantity ELSE 0 END) AS current_stock,
        p.volume_cbm
    FROM inventory_txn it
    JOIN products p ON p.sku = it.sku
    JOIN sellers  s ON s.seller_id = p.seller_id
    GROUP BY it.sku
),
wh_occupied AS (
    SELECT warehouse_id, SUM(current_stock * volume_cbm) AS occupied_volume
    FROM occupied
    GROUP BY warehouse_id
),
wh_orders AS (
    SELECT warehouse_id, COUNT(*) AS orders_processed,
           AVG(julianday(dispatch_date) - julianday(order_date)) AS avg_dispatch_days
    FROM orders
    GROUP BY warehouse_id
),
wh_sla AS (
    SELECT warehouse_id,
           COUNT(*) AS deliveries_from_wh,
           CAST(SUM(CASE WHEN sla_met = 'Yes' THEN 1 ELSE 0 END) AS REAL) / COUNT(*) AS sla_met_pct
    FROM deliveries
    GROUP BY warehouse_id
)
SELECT
    w.warehouse_id, w.location, w.storage_capacity_units,
    ROUND(wo.occupied_volume / w.storage_capacity_units, 4) AS current_utilization,
    w.picking_accuracy, w.packing_accuracy, w.inventory_accuracy_audit,
    wor.orders_processed,
    ROUND(wor.avg_dispatch_days, 2) AS avg_dispatch_days,
    ws.deliveries_from_wh,
    ROUND(ws.sla_met_pct, 4) AS sla_met_pct,
    CASE
        WHEN wo.occupied_volume / w.storage_capacity_units > 0.9 THEN 'Overloaded'
        WHEN wo.occupied_volume / w.storage_capacity_units < 0.6 THEN 'Underutilized'
        ELSE 'Healthy'
    END AS capacity_status
FROM warehouses w
JOIN wh_occupied wo ON wo.warehouse_id = w.warehouse_id
JOIN wh_orders   wor ON wor.warehouse_id = w.warehouse_id
JOIN wh_sla      ws  ON ws.warehouse_id  = w.warehouse_id
ORDER BY current_utilization DESC;


-- ----------------------------------------------------------------------------
-- 3. DELIVERY SLA — by city, by warehouse, by month
-- ----------------------------------------------------------------------------
-- 3a. By city
SELECT
    city,
    COUNT(*) AS total_deliveries,
    SUM(CASE WHEN sla_met = 'Yes' THEN 1 ELSE 0 END) AS sla_met_count,
    ROUND(CAST(SUM(CASE WHEN sla_met = 'Yes' THEN 1 ELSE 0 END) AS REAL) / COUNT(*), 4) AS sla_pct,
    COUNT(*) - SUM(CASE WHEN sla_met = 'Yes' THEN 1 ELSE 0 END) AS delayed_orders,
    ROUND(AVG(delay_days), 2) AS avg_delay_days
FROM deliveries
GROUP BY city
ORDER BY sla_pct ASC;

-- 3b. By warehouse
SELECT
    warehouse_id,
    COUNT(*) AS total_deliveries,
    ROUND(CAST(SUM(CASE WHEN sla_met = 'Yes' THEN 1 ELSE 0 END) AS REAL) / COUNT(*), 4) AS sla_pct,
    ROUND(AVG(delay_days), 2) AS avg_delay_days
FROM deliveries
GROUP BY warehouse_id
ORDER BY sla_pct ASC;

-- 3c. By month
SELECT
    delivery_month,
    COUNT(*) AS total_deliveries,
    ROUND(CAST(SUM(CASE WHEN sla_met = 'Yes' THEN 1 ELSE 0 END) AS REAL) / COUNT(*), 4) AS sla_pct,
    ROUND(AVG(delay_days), 2) AS avg_delay_days
FROM deliveries
GROUP BY delivery_month
ORDER BY delivery_month;


-- ----------------------------------------------------------------------------
-- 4. INVENTORY SUMMARY — per SKU: current stock, safety stock, reorder point,
--    days of inventory, stockout risk, movement class. Mirrors the Inventory
--    sheet's derived table in the workbook.
-- ----------------------------------------------------------------------------
WITH stock AS (
    SELECT
        sku,
        SUM(CASE WHEN txn_type = 'Inbound'  THEN quantity ELSE 0 END)
      - SUM(CASE WHEN txn_type = 'Outbound' THEN quantity ELSE 0 END) AS current_stock,
        CAST(SUM(CASE WHEN txn_type = 'Outbound' THEN quantity ELSE 0 END) AS REAL) / 180 AS avg_daily_outbound
    FROM inventory_txn
    GROUP BY sku
)
SELECT
    p.sku, p.product_name, p.category, s.warehouse_assigned AS warehouse_id, p.abc_category,
    st.current_stock,
    ROUND(st.avg_daily_outbound, 2) AS avg_daily_outbound,
    p.lead_time_days,
    ROUND(st.avg_daily_outbound * 1.65 * SQRT(p.lead_time_days), 0) AS safety_stock,
    ROUND(st.avg_daily_outbound * p.lead_time_days
          + st.avg_daily_outbound * 1.65 * SQRT(p.lead_time_days), 0) AS reorder_point,
    CASE WHEN st.avg_daily_outbound = 0 THEN NULL
         ELSE ROUND(st.current_stock / st.avg_daily_outbound, 1) END  AS days_of_inventory,
    CASE
        WHEN st.current_stock < (st.avg_daily_outbound * p.lead_time_days
                                  + st.avg_daily_outbound * 1.65 * SQRT(p.lead_time_days))
        THEN 'At Risk' ELSE 'OK'
    END AS stockout_risk,
    ROUND(st.current_stock * p.cost_price, 0) AS inventory_value,
    CASE
        WHEN st.avg_daily_outbound = 0 THEN 'Dead Stock'
        WHEN (st.current_stock / NULLIF(st.avg_daily_outbound,0)) < 15 THEN 'Fast Moving'
        WHEN (st.current_stock / NULLIF(st.avg_daily_outbound,0)) > 45 THEN 'Slow Moving'
        ELSE 'Normal'
    END AS movement_class
FROM products p
JOIN stock st ON st.sku = p.sku
JOIN sellers s ON s.seller_id = p.seller_id
ORDER BY stockout_risk DESC, days_of_inventory ASC;


-- ----------------------------------------------------------------------------
-- 5. DASHBOARD KPIs — the same numbers behind the Excel executive dashboard.
-- ----------------------------------------------------------------------------
-- 5a. Headline cards
SELECT
    (SELECT COUNT(DISTINCT order_id) FROM orders)                                    AS total_orders_distinct,
    (SELECT ROUND(SUM(order_value), 2) FROM orders WHERE order_status = 'Delivered') AS revenue_delivered,
    (SELECT ROUND(CAST(SUM(CASE WHEN sla_met='Yes' THEN 1 ELSE 0 END) AS REAL) / COUNT(*), 4) FROM deliveries) AS sla_pct,
    (SELECT ROUND(AVG(delivery_time_hr), 2) FROM deliveries)                          AS avg_delivery_time_hr,
    (SELECT ROUND(CAST(COUNT(*) AS REAL) /
            (SELECT COUNT(*) FROM orders WHERE order_status = 'Delivered'), 4)
     FROM returns)                                                                    AS return_rate,
    (SELECT ROUND(CAST(SUM(CASE WHEN order_status='Cancelled' THEN 1 ELSE 0 END) AS REAL)
            / COUNT(*), 4) FROM orders)                                               AS cancellation_rate,
    (SELECT ROUND(AVG(inventory_accuracy_audit), 4) FROM warehouses)                  AS network_inventory_accuracy;

-- 5b. Top seller by score (reuses query 1's logic via a view — see below)
-- 5c. Revenue by category
SELECT category,
       ROUND(SUM(CASE WHEN order_status='Delivered' THEN order_value ELSE 0 END), 2) AS revenue,
       COUNT(*) AS order_lines
FROM orders o
JOIN products p ON p.sku = o.sku
GROUP BY category
ORDER BY revenue DESC;

-- 5d. Return reasons breakdown
SELECT return_reason, COUNT(*) AS return_count,
       ROUND(100.0 * COUNT(*) / (SELECT COUNT(*) FROM returns), 1) AS pct_of_returns
FROM returns
GROUP BY return_reason
ORDER BY return_count DESC;


-- ----------------------------------------------------------------------------
-- 6. REUSABLE VIEW — Seller Scorecard as a view, so query 5b (and anything else)
--    can just SELECT from it instead of repeating the CTE.
-- ----------------------------------------------------------------------------
DROP VIEW IF EXISTS vw_seller_scorecard;
CREATE VIEW vw_seller_scorecard AS
WITH seller_orders AS (
    SELECT
        seller_id,
        COUNT(*)                                                    AS total_lines,
        SUM(CASE WHEN order_status = 'Delivered' THEN 1 ELSE 0 END) AS delivered_lines,
        SUM(CASE WHEN order_status = 'Delivered'
                  AND delivery_date <= expected_delivery_date THEN 1 ELSE 0 END) AS ontime_count,
        SUM(CASE WHEN order_status = 'Cancelled' THEN 1 ELSE 0 END) AS cancelled_lines,
        SUM(CASE WHEN julianday(dispatch_date) - julianday(order_date) <= 1
                  THEN 1 ELSE 0 END)                                AS ontime_dispatch_count
    FROM orders
    GROUP BY seller_id
),
seller_returns AS (
    SELECT seller_id, COUNT(*) AS returns_count FROM returns GROUP BY seller_id
)
SELECT
    s.seller_id, s.seller_name, s.seller_category, s.warehouse_assigned,
    CAST(COALESCE(so.ontime_count, 0) AS REAL) / NULLIF(so.delivered_lines, 0)     AS ontime_rate,
    CAST(COALESCE(so.cancelled_lines, 0) AS REAL) / NULLIF(so.total_lines, 0)      AS cancellation_rate,
    CAST(COALESCE(sr.returns_count, 0) AS REAL) / NULLIF(so.delivered_lines, 0)    AS return_rate,
    CAST(COALESCE(so.ontime_dispatch_count, 0) AS REAL) / NULLIF(so.total_lines,0) AS ontime_dispatch_rate,
    w.inventory_accuracy_audit, s.seller_rating,
    ROUND((CAST(COALESCE(so.ontime_count,0) AS REAL)/NULLIF(so.delivered_lines,0) * 0.30
      + (1 - CAST(COALESCE(so.cancelled_lines,0) AS REAL)/NULLIF(so.total_lines,0)) * 0.20
      + (1 - CAST(COALESCE(sr.returns_count,0) AS REAL)/NULLIF(so.delivered_lines,0)) * 0.20
      + w.inventory_accuracy_audit * 0.15
      + (s.seller_rating/5.0) * 0.15) * 100, 1) AS seller_score
FROM sellers s
LEFT JOIN seller_orders  so ON so.seller_id = s.seller_id
LEFT JOIN seller_returns sr ON sr.seller_id = s.seller_id
JOIN warehouses w ON w.warehouse_id = s.warehouse_assigned;

-- Example use of the view — top seller overall:
SELECT seller_id, seller_name, seller_score
FROM vw_seller_scorecard
ORDER BY seller_score DESC
LIMIT 1;

-- Example use of the view — sellers who would drop to Watchlist if SLA target tightened
-- (shows the same "what-if" flexibility the README parameters give the Excel workbook):
SELECT seller_id, seller_name, ontime_rate, seller_score
FROM vw_seller_scorecard
WHERE ontime_rate < 0.95 AND seller_score >= 55
ORDER BY ontime_rate ASC
LIMIT 10;


-- ----------------------------------------------------------------------------
-- 7. SELLER INVENTORY AVAILABILITY (stockout rate) — distinct from the
--    warehouse-level Inventory_Accuracy_Audit figure already in the scorecard.
--    "Availability" here means: of this seller's listed SKUs, what share are
--    currently below their reorder point (At Risk of stocking out)?
--    Mirrors the Sellers sheet's SKU_Count / AtRisk_SKU_Count / Stockout_Rate
--    columns and the Seller Scorecard's Stockout_Rate column in the workbook.
-- ----------------------------------------------------------------------------
DROP VIEW IF EXISTS vw_seller_availability;
CREATE VIEW vw_seller_availability AS
WITH stock AS (
    SELECT
        sku,
        SUM(CASE WHEN txn_type = 'Inbound'  THEN quantity ELSE 0 END)
      - SUM(CASE WHEN txn_type = 'Outbound' THEN quantity ELSE 0 END) AS current_stock,
        CAST(SUM(CASE WHEN txn_type = 'Outbound' THEN quantity ELSE 0 END) AS REAL) / 180 AS avg_daily_outbound
    FROM inventory_txn
    GROUP BY sku
),
sku_risk AS (
    SELECT
        p.sku, p.seller_id,
        CASE
            WHEN st.current_stock < (st.avg_daily_outbound * p.lead_time_days
                                      + st.avg_daily_outbound * 1.65 * SQRT(p.lead_time_days))
            THEN 1 ELSE 0
        END AS at_risk
    FROM products p
    JOIN stock st ON st.sku = p.sku
)
SELECT
    seller_id,
    COUNT(*)                      AS sku_count,
    SUM(at_risk)                  AS at_risk_sku_count,
    ROUND(CAST(SUM(at_risk) AS REAL) / COUNT(*), 4) AS stockout_rate
FROM sku_risk
GROUP BY seller_id;

-- Example use — sellers with the worst inventory availability (highest stockout rate):
SELECT sa.seller_id, s.seller_name, sa.sku_count, sa.at_risk_sku_count, sa.stockout_rate
FROM vw_seller_availability sa
JOIN sellers s ON s.seller_id = sa.seller_id
ORDER BY sa.stockout_rate DESC
LIMIT 10;

-- Combined view — full scorecard + on-time dispatch + inventory availability in one row,
-- exactly the shape of the Seller Scorecard sheet in the Excel workbook:
SELECT
    v.seller_id, v.seller_name, v.seller_category, v.warehouse_assigned,
    v.ontime_rate, v.cancellation_rate, v.return_rate, v.inventory_accuracy_audit,
    v.seller_rating, v.ontime_dispatch_rate,
    a.stockout_rate,
    v.seller_score,
    CASE WHEN v.seller_score >= 85 THEN 'Gold'
         WHEN v.seller_score >= 70 THEN 'Silver'
         WHEN v.seller_score >= 55 THEN 'Bronze'
         ELSE 'Watchlist' END AS seller_tier
FROM vw_seller_scorecard v
LEFT JOIN vw_seller_availability a ON a.seller_id = v.seller_id
ORDER BY v.seller_score DESC;
