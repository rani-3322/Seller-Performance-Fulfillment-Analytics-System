-- Flipkart Fulfillment & Seller Performance Analytics — Database Schema
-- SQLite dialect (portable to MySQL/Postgres with minor type tweaks: TEXT->VARCHAR, DATE stays DATE)

CREATE TABLE deliveries (
    delivery_id     TEXT PRIMARY KEY,
    order_id        TEXT NOT NULL,
    city            TEXT NOT NULL,
    warehouse_id    TEXT NOT NULL REFERENCES warehouses(warehouse_id),
    rider_id        TEXT NOT NULL REFERENCES riders(rider_id),
    distance_km     REAL,
    delivery_time_hr REAL,
    delay_days      INTEGER,
    sla_met         TEXT NOT NULL CHECK (sla_met IN ('Yes','No')),
    fuel_cost       REAL,
    traffic_level   TEXT,
    delivery_status TEXT,
    delivery_month  TEXT
);

CREATE TABLE exceptions (
    exception_id    TEXT PRIMARY KEY,
    order_id        TEXT NOT NULL,
    exception_type  TEXT NOT NULL,
    exception_date  DATE NOT NULL,
    owner           TEXT NOT NULL,
    status          TEXT NOT NULL,
    action_taken    TEXT
);

CREATE TABLE inventory_txn (
    txn_id      TEXT PRIMARY KEY,
    txn_date    DATE NOT NULL,
    sku         TEXT NOT NULL REFERENCES products(sku),
    warehouse_id TEXT NOT NULL REFERENCES warehouses(warehouse_id),
    txn_type    TEXT NOT NULL CHECK (txn_type IN ('Inbound','Outbound')),
    quantity    INTEGER NOT NULL,
    unit_cost   REAL NOT NULL
);

CREATE TABLE orders (
    order_line_id           TEXT PRIMARY KEY,
    order_id                TEXT NOT NULL,
    customer_id             TEXT NOT NULL,
    city                    TEXT NOT NULL,
    warehouse_id            TEXT NOT NULL REFERENCES warehouses(warehouse_id),
    seller_id               TEXT NOT NULL REFERENCES sellers(seller_id),
    sku                     TEXT NOT NULL REFERENCES products(sku),
    order_date              DATE NOT NULL,
    dispatch_date           DATE,
    expected_delivery_date  DATE,
    delivery_date           DATE,
    order_status            TEXT NOT NULL CHECK (order_status IN ('Delivered','Cancelled','RTO','Processing')),
    payment_mode            TEXT NOT NULL,
    qty                     INTEGER NOT NULL,
    order_value             REAL NOT NULL,
    delivery_distance_km    REAL
);

CREATE TABLE products (
    sku             TEXT PRIMARY KEY,
    product_name    TEXT NOT NULL,
    category        TEXT NOT NULL,
    seller_id       TEXT NOT NULL REFERENCES sellers(seller_id),
    selling_price   REAL NOT NULL,
    cost_price      REAL NOT NULL,
    weight_kg       REAL,
    volume_cbm      REAL,
    abc_category    TEXT,
    supplier        TEXT,
    lead_time_days  INTEGER NOT NULL
);

CREATE TABLE returns (
    return_id           TEXT PRIMARY KEY,
    order_id             TEXT NOT NULL,
    sku                  TEXT NOT NULL REFERENCES products(sku),
    seller_id            TEXT NOT NULL REFERENCES sellers(seller_id),
    warehouse_id         TEXT NOT NULL REFERENCES warehouses(warehouse_id),
    return_reason         TEXT NOT NULL,
    return_date           DATE NOT NULL,
    refund_amount         REAL NOT NULL,
    resolution_time_hrs   INTEGER,
    return_category       TEXT
);

CREATE TABLE riders (
    rider_id        TEXT PRIMARY KEY,
    rider_name      TEXT NOT NULL,
    hub_warehouse   TEXT NOT NULL REFERENCES warehouses(warehouse_id),
    availability    TEXT NOT NULL,
    daily_capacity  INTEGER NOT NULL
);

CREATE TABLE sellers (
    seller_id           TEXT PRIMARY KEY,
    seller_name         TEXT NOT NULL,
    seller_category     TEXT NOT NULL,
    seller_city         TEXT NOT NULL,
    fulfillment_type    TEXT NOT NULL,
    seller_rating       REAL NOT NULL,
    monthly_order_target INTEGER,
    warehouse_assigned  TEXT NOT NULL REFERENCES warehouses(warehouse_id),
    onboarded_on        DATE
);

CREATE TABLE warehouses (
    warehouse_id            TEXT PRIMARY KEY,
    location                TEXT NOT NULL,
    storage_capacity_units  REAL NOT NULL,
    target_utilization      REAL NOT NULL,
    inventory_accuracy_audit REAL NOT NULL,
    picking_accuracy        REAL NOT NULL,
    packing_accuracy        REAL NOT NULL,
    dock_utilization         REAL NOT NULL
);

CREATE INDEX idx_deliv_city      ON deliveries(city);

CREATE INDEX idx_deliv_month     ON deliveries(delivery_month);

CREATE INDEX idx_deliv_wh        ON deliveries(warehouse_id);

CREATE INDEX idx_invtxn_sku      ON inventory_txn(sku);

CREATE INDEX idx_invtxn_wh       ON inventory_txn(warehouse_id);

CREATE INDEX idx_orders_orderid  ON orders(order_id);

CREATE INDEX idx_orders_seller   ON orders(seller_id);

CREATE INDEX idx_orders_sku      ON orders(sku);

CREATE INDEX idx_orders_status   ON orders(order_status);

CREATE INDEX idx_orders_wh       ON orders(warehouse_id);

CREATE INDEX idx_returns_seller  ON returns(seller_id);

CREATE INDEX idx_returns_sku     ON returns(sku);

