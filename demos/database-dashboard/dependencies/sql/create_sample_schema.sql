-- =============================================================================
-- create_sample_schema.sql
-- Sample Database Schema for Excel-PostgreSQL-Dashboard
  -- Author: Shiv Iyer | MinervaDB | ChistaDATA
  -- License: MIT
  -- Usage: psql -U your_username -d your_database -f sql/setup/create_sample_schema.sql
  -- =============================================================================

  -- Create schema
  CREATE SCHEMA IF NOT EXISTS dashboard;
  SET search_path TO dashboard, public;

  -- =============================================================================
  -- CUSTOMERS
  -- =============================================================================
  CREATE TABLE IF NOT EXISTS customers (
        customer_id     SERIAL PRIMARY KEY,
    customer_name   VARCHAR(200) NOT NULL,
    email           VARCHAR(200) UNIQUE,
    phone           VARCHAR(50),
    company         VARCHAR(200),
    industry        VARCHAR(100),
    country         VARCHAR(100),
    region          VARCHAR(100),
    customer_tier   VARCHAR(50)  DEFAULT 'Standard' CHECK (customer_tier IN ('Standard','Silver','Gold','Platinum')),
      created_at      TIMESTAMPTZ  NOT NULL DEFAULT NOW(),
      updated_at      TIMESTAMPTZ  NOT NULL DEFAULT NOW()
    );

    -- =============================================================================
    -- PRODUCTS
    -- =============================================================================
    CREATE TABLE IF NOT EXISTS products (
          product_id      SERIAL PRIMARY KEY,
      product_name    VARCHAR(200) NOT NULL,
      sku             VARCHAR(100) UNIQUE NOT NULL,
      category        VARCHAR(100),
      subcategory     VARCHAR(100),
      unit_cost       NUMERIC(12,2) NOT NULL DEFAULT 0,
      list_price      NUMERIC(12,2) NOT NULL DEFAULT 0,
          is_active       BOOLEAN NOT NULL DEFAULT TRUE,
      created_at      TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );

    -- =============================================================================
    -- SALES REPS
    -- =============================================================================
    CREATE TABLE IF NOT EXISTS sales_reps (
          rep_id          SERIAL PRIMARY KEY,
      rep_name        VARCHAR(200) NOT NULL,
      email           VARCHAR(200) UNIQUE,
      team            VARCHAR(100),
      region          VARCHAR(100),
      manager_id      INTEGER REFERENCES sales_reps(rep_id),
      quarterly_quota NUMERIC(15,2) DEFAULT 0,
          hire_date       DATE,
          is_active       BOOLEAN NOT NULL DEFAULT TRUE
    );

    -- =============================================================================
    -- ORDERS
    -- =============================================================================
    CREATE TABLE IF NOT EXISTS orders (
          order_id        SERIAL PRIMARY KEY,
      customer_id     INTEGER NOT NULL REFERENCES customers(customer_id),
      rep_id          INTEGER REFERENCES sales_reps(rep_id),
          order_date      DATE NOT NULL,
          ship_date       DATE,
      order_status    VARCHAR(50) NOT NULL DEFAULT 'completed'
      CHECK (order_status IN ('pending','processing','completed','cancelled','refunded')),
        order_amount    NUMERIC(15,2) NOT NULL DEFAULT 0,
        discount_amount NUMERIC(15,2) NOT NULL DEFAULT 0,
            days_to_close   INTEGER,
        payment_method  VARCHAR(50),
            notes           TEXT,
        created_at      TIMESTAMPTZ NOT NULL DEFAULT NOW()
      );

      -- =============================================================================
      -- ORDER ITEMS
      -- =============================================================================
      CREATE TABLE IF NOT EXISTS order_items (
            order_item_id   SERIAL PRIMARY KEY,
        order_id        INTEGER NOT NULL REFERENCES orders(order_id) ON DELETE CASCADE,
        product_id      INTEGER NOT NULL REFERENCES products(product_id),
            quantity        INTEGER NOT NULL DEFAULT 1,
        unit_price      NUMERIC(12,2) NOT NULL,
        unit_cost       NUMERIC(12,2) NOT NULL,
        discount_pct    NUMERIC(5,2) NOT NULL DEFAULT 0,
        line_amount     NUMERIC(15,2) GENERATED ALWAYS AS
        (quantity * unit_price * (1 - discount_pct / 100)) STORED
      );

      -- =============================================================================
      -- SALES PIPELINE
      -- =============================================================================
      CREATE TABLE IF NOT EXISTS sales_pipeline (
            opportunity_id  SERIAL PRIMARY KEY,
        opportunity_name VARCHAR(300) NOT NULL,
        customer_id     INTEGER REFERENCES customers(customer_id),
        rep_id          INTEGER REFERENCES sales_reps(rep_id),
        stage_name      VARCHAR(100) NOT NULL,
            stage_order     INTEGER NOT NULL DEFAULT 1,
        win_probability INTEGER NOT NULL DEFAULT 0 CHECK (win_probability BETWEEN 0 AND 100),
        deal_value      NUMERIC(15,2) NOT NULL DEFAULT 0,
        lead_source     VARCHAR(100),
            close_date      DATE,
            created_date    DATE NOT NULL DEFAULT CURRENT_DATE,
            days_in_pipeline INTEGER GENERATED ALWAYS AS
        (CURRENT_DATE - created_date) STORED,
            is_active       BOOLEAN NOT NULL DEFAULT TRUE,
            notes           TEXT
      );

      -- =============================================================================
      -- EMPLOYEES (for HR Dashboard)
        -- =============================================================================
        CREATE TABLE IF NOT EXISTS departments (
              dept_id         SERIAL PRIMARY KEY,
          dept_name       VARCHAR(200) NOT NULL,
          parent_dept_id  INTEGER REFERENCES departments(dept_id),
          cost_center     VARCHAR(50),
          location        VARCHAR(100)
        );

        CREATE TABLE IF NOT EXISTS employees (
              employee_id     SERIAL PRIMARY KEY,
          employee_name   VARCHAR(200) NOT NULL,
          email           VARCHAR(200) UNIQUE,
          dept_id         INTEGER REFERENCES departments(dept_id),
          job_title       VARCHAR(200),
          manager_id      INTEGER REFERENCES employees(employee_id),
              hire_date       DATE NOT NULL,
              termination_date DATE,
          status          VARCHAR(20) NOT NULL DEFAULT 'Active'
          CHECK (status IN ('Active','Inactive','Leave')),
            base_salary     NUMERIC(12,2),
            location        VARCHAR(100),
            employment_type VARCHAR(50) DEFAULT 'Full-time'
          );

          -- =============================================================================
          -- DASHBOARD VIEWS
          -- =============================================================================

          CREATE OR REPLACE VIEW vw_sales_dashboard AS
          SELECT
              o.order_id,
              o.order_date,
          DATE_TRUNC('month', o.order_date)               AS order_month,
          DATE_TRUNC('quarter', o.order_date)             AS order_quarter,
          EXTRACT(YEAR FROM o.order_date)                 AS order_year,
              c.customer_name,
              c.industry,
              c.region                                         AS customer_region,
              sr.rep_name,
              sr.team,
              sr.region                                        AS sales_region,
              o.order_amount,
              o.discount_amount,
              o.order_amount - o.discount_amount              AS net_revenue,
              o.order_status
          FROM orders o
          JOIN customers c  ON c.customer_id = o.customer_id
          LEFT JOIN sales_reps sr ON sr.rep_id = o.rep_id
          WHERE o.order_status NOT IN ('cancelled', 'refunded');


            CREATE OR REPLACE VIEW vw_executive_kpis AS
            WITH monthly_revenue AS (
                  SELECT
              DATE_TRUNC('month', order_date) AS month,
              SUM(order_amount)               AS revenue
                  FROM orders
              WHERE order_status NOT IN ('cancelled', 'refunded')
                    GROUP BY 1
              ),
              current_month AS (
                    SELECT revenue FROM monthly_revenue
                WHERE month = DATE_TRUNC('month', CURRENT_DATE)
              ),
              prior_month AS (
                    SELECT revenue FROM monthly_revenue
                WHERE month = DATE_TRUNC('month', CURRENT_DATE) - INTERVAL '1 month'
              )
              SELECT
                  cm.revenue                                       AS current_month_revenue,
                  pm.revenue                                       AS prior_month_revenue,
              ROUND((cm.revenue - pm.revenue) / NULLIF(pm.revenue, 0) * 100, 2)
                                                                   AS mom_revenue_growth_pct,
              (SELECT COUNT(DISTINCT customer_id) FROM orders
                WHERE DATE_TRUNC('month', order_date) = DATE_TRUNC('month', CURRENT_DATE))
                                                                   AS active_customers_this_month,
              (SELECT COUNT(*) FROM employees WHERE status = 'Active')
                                                                     AS active_employee_count
                FROM current_month cm, prior_month pm;

                -- =============================================================================
                -- INDEXES FOR DASHBOARD PERFORMANCE
                -- =============================================================================
                CREATE INDEX IF NOT EXISTS idx_orders_order_date    ON orders(order_date);
                CREATE INDEX IF NOT EXISTS idx_orders_customer_id   ON orders(customer_id);
                CREATE INDEX IF NOT EXISTS idx_orders_rep_id        ON orders(rep_id);
                CREATE INDEX IF NOT EXISTS idx_orders_status        ON orders(order_status);
                CREATE INDEX IF NOT EXISTS idx_order_items_order_id ON order_items(order_id);
                CREATE INDEX IF NOT EXISTS idx_pipeline_stage       ON sales_pipeline(stage_name);
                CREATE INDEX IF NOT EXISTS idx_pipeline_close_date  ON sales_pipeline(close_date);
                CREATE INDEX IF NOT EXISTS idx_employees_dept       ON employees(dept_id);
                CREATE INDEX IF NOT EXISTS idx_employees_status     ON employees(status);

                -- =============================================================================
                -- READ-ONLY DASHBOARD ROLE
                -- =============================================================================
                -- Provision the login and its credential outside this redistributable bundle,
                -- then run these grants as a database administrator:
                -- GRANT CONNECT ON DATABASE <database_name> TO excel_dashboard_ro;
                -- GRANT USAGE ON SCHEMA dashboard TO excel_dashboard_ro;
                -- GRANT SELECT ON ALL TABLES IN SCHEMA dashboard TO excel_dashboard_ro;
                -- ALTER DEFAULT PRIVILEGES IN SCHEMA dashboard
                --     GRANT SELECT ON TABLES TO excel_dashboard_ro;

                  -- Done
                  SELECT 'Schema created successfully. Tables: customers, products, sales_reps, orders, order_items, sales_pipeline, departments, employees' AS status;
