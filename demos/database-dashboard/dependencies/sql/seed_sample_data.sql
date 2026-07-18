-- ============================================================
-- MinervaDB XL View - Sample Data Seed Script
-- sql/setup/seed_sample_data.sql
-- ============================================================
-- Inserts realistic sample data for testing and demonstration
-- of the MinervaDB XL View dashboard platform.
-- Run AFTER create_sample_schema.sql
-- ============================================================

BEGIN;

-- ============================================================
-- Seed: Departments
-- ============================================================
INSERT INTO departments (department_name, department_head, cost_center_code, budget_annual) VALUES
    ('Sales',           'Sarah Johnson',    'CC-001', 2500000.00),
    ('Engineering',     'Michael Chen',     'CC-002', 4200000.00),
    ('Marketing',       'Emily Rodriguez',  'CC-003', 1800000.00),
    ('Finance',         'David Kim',        'CC-004', 900000.00),
    ('Human Resources', 'Jessica Williams', 'CC-005', 750000.00),
    ('Operations',      'Robert Thompson',  'CC-006', 3100000.00),
    ('Customer Success','Amanda Davis',     'CC-007', 1200000.00),
    ('Product',         'James Wilson',     'CC-008', 2800000.00)
ON CONFLICT (department_name) DO NOTHING;

-- ============================================================
-- Seed: Employees (50 sample records)
-- ============================================================
INSERT INTO employees (
    employee_id, first_name, last_name, hire_date, department,
    job_title, employment_status, employment_type, salary, gender, location,
    performance_rating, manager_id
) VALUES
    (1001, 'Alice',   'Morgan',    '2019-03-15', 'Sales',           'Sales Director',          'Active', 'Full-Time', 145000, 'Female', 'New York',      5, NULL),
    (1002, 'Ben',     'Carter',    '2020-06-01', 'Sales',           'Account Executive',       'Active', 'Full-Time',  92000, 'Male',   'Chicago',       4, 1001),
    (1003, 'Cara',    'Nguyen',    '2021-01-10', 'Sales',           'Account Executive',       'Active', 'Full-Time',  88000, 'Female', 'Los Angeles',   4, 1001),
    (1004, 'Derek',   'Sullivan',  '2022-09-05', 'Sales',           'Sales Development Rep',   'Active', 'Full-Time',  65000, 'Male',   'Chicago',       3, 1002),
    (1005, 'Ella',    'Park',      '2020-02-20', 'Engineering',     'VP Engineering',          'Active', 'Full-Time', 195000, 'Female', 'San Francisco', 5, NULL),
    (1006, 'Frank',   'Okafor',    '2019-11-15', 'Engineering',     'Senior Engineer',         'Active', 'Full-Time', 155000, 'Male',   'San Francisco', 5, 1005),
    (1007, 'Grace',   'Liu',       '2021-07-12', 'Engineering',     'Software Engineer',       'Active', 'Full-Time', 130000, 'Female', 'Austin',        4, 1006),
    (1008, 'Henry',   'Brooks',    '2022-03-08', 'Engineering',     'Software Engineer',       'Active', 'Full-Time', 125000, 'Male',   'Austin',        3, 1006),
    (1009, 'Iris',    'Yamamoto',  '2023-01-16', 'Engineering',     'Junior Engineer',         'Active', 'Full-Time', 105000, 'Female', 'San Francisco', 3, 1007),
    (1010, 'Jack',    'Patel',     '2019-05-20', 'Marketing',       'CMO',                     'Active', 'Full-Time', 175000, 'Male',   'New York',      5, NULL),
    (1011, 'Karen',   'White',     '2020-08-10', 'Marketing',       'Marketing Manager',       'Active', 'Full-Time', 115000, 'Female', 'New York',      4, 1010),
    (1012, 'Leo',     'Harris',    '2021-04-22', 'Marketing',       'Content Strategist',      'Active', 'Full-Time',  85000, 'Male',   'Remote',        4, 1011),
    (1013, 'Mia',     'Robinson',  '2022-11-01', 'Marketing',       'Digital Marketing Spec',  'Active', 'Full-Time',  72000, 'Female', 'New York',      3, 1011),
    (1014, 'Nathan',  'Clark',     '2019-09-30', 'Finance',         'CFO',                     'Active', 'Full-Time', 220000, 'Male',   'New York',      5, NULL),
    (1015, 'Olivia',  'Lewis',     '2020-12-15', 'Finance',         'Finance Manager',         'Active', 'Full-Time', 120000, 'Female', 'New York',      4, 1014),
    (1016, 'Paul',    'Walker',    '2021-06-07', 'Finance',         'Financial Analyst',       'Active', 'Full-Time',  90000, 'Male',   'New York',      4, 1015),
    (1017, 'Quinn',   'Hall',      '2022-02-14', 'Finance',         'Junior Analyst',          'Active', 'Full-Time',  72000, 'Female', 'Remote',        3, 1016),
    (1018, 'Rachel',  'Young',     '2019-07-08', 'Human Resources', 'HR Director',             'Active', 'Full-Time', 135000, 'Female', 'Chicago',       5, NULL),
    (1019, 'Sam',     'Allen',     '2020-04-19', 'Human Resources', 'HR Business Partner',     'Active', 'Full-Time',  92000, 'Male',   'Chicago',       4, 1018),
    (1020, 'Tara',    'King',      '2021-10-25', 'Human Resources', 'Recruiter',               'Active', 'Full-Time',  75000, 'Female', 'Remote',        3, 1019),
    (1021, 'Uma',     'Scott',     '2019-01-14', 'Operations',      'COO',                     'Active', 'Full-Time', 210000, 'Female', 'Dallas',        5, NULL),
    (1022, 'Victor',  'Green',     '2020-03-25', 'Operations',      'Operations Manager',      'Active', 'Full-Time', 118000, 'Male',   'Dallas',        4, 1021),
    (1023, 'Wendy',   'Adams',     '2021-08-02', 'Operations',      'Operations Analyst',      'Active', 'Full-Time',  82000, 'Female', 'Dallas',        4, 1022),
    (1024, 'Xavier',  'Baker',     '2022-06-13', 'Operations',      'Supply Chain Coord.',     'Active', 'Full-Time',  70000, 'Male',   'Dallas',        3, 1022),
    (1025, 'Yasmine', 'Gonzalez',  '2019-04-03', 'Customer Success','CS Director',             'Active', 'Full-Time', 138000, 'Female', 'San Francisco', 5, NULL),
    (1026, 'Zach',    'Nelson',    '2020-07-14', 'Customer Success','Customer Success Manager', 'Active','Full-Time',  95000, 'Male',   'San Francisco', 4, 1025),
    (1027, 'Amy',     'Carter',    '2021-03-09', 'Customer Success','CSM',                     'Active', 'Full-Time',  78000, 'Female', 'Remote',        4, 1026),
    (1028, 'Brian',   'Mitchell',  '2022-01-17', 'Customer Success','CSM',                     'Active', 'Full-Time',  76000, 'Male',   'Remote',        3, 1026),
    (1029, 'Chris',   'Perez',     '2019-10-22', 'Product',         'CPO',                     'Active', 'Full-Time', 185000, 'Male',   'San Francisco', 5, NULL),
    (1030, 'Dana',    'Roberts',   '2020-09-28', 'Product',         'Product Manager',         'Active', 'Full-Time', 135000, 'Female', 'San Francisco', 4, 1029),
    -- Terminated employees for attrition analysis
    (1031, 'Eric',    'Turner',    '2020-05-11', 'Sales',           'Account Executive',       'Terminated', 'Full-Time', 88000, 'Male', 'Chicago', 3, 1001),
    (1032, 'Fiona',   'Phillips',  '2019-12-03', 'Engineering',     'Software Engineer',       'Terminated', 'Full-Time', 128000, 'Female', 'Austin', 4, 1006),
    (1033, 'George',  'Campbell',  '2021-02-18', 'Marketing',       'Marketing Specialist',    'Terminated', 'Full-Time', 78000, 'Male', 'Remote', 2, 1011),
    (1034, 'Hannah',  'Parker',    '2020-11-05', 'Operations',      'Operations Analyst',      'Terminated', 'Full-Time', 80000, 'Female', 'Dallas', 3, 1022),
    (1035, 'Ian',     'Evans',     '2022-04-20', 'Customer Success','CSM',                     'Terminated', 'Full-Time', 74000, 'Male', 'Remote', 3, 1026)
ON CONFLICT (employee_id) DO NOTHING;

-- ============================================================
-- Seed: Customers
-- ============================================================
INSERT INTO customers (
    customer_id, customer_name, customer_segment, industry, region, annual_revenue
) VALUES
    (201, 'Apex Technologies',     'Enterprise', 'Technology',   'West',     500000000),
    (202, 'Global Retail Corp',    'Enterprise', 'Retail',       'East',     800000000),
    (203, 'Midwest Manufacturing', 'Mid-Market', 'Manufacturing','Midwest',  150000000),
    (204, 'Southern Finance Group','Enterprise', 'Finance',      'South',    300000000),
    (205, 'Pacific Healthcare',    'Mid-Market', 'Healthcare',   'West',     200000000),
    (206, 'Atlantic Logistics',    'SMB',        'Logistics',    'East',      50000000),
    (207, 'Innovate Startup',      'SMB',        'Technology',   'West',      10000000),
    (208, 'Education Partners',    'Mid-Market', 'Education',    'Midwest',   75000000),
    (209, 'Energy Solutions Inc',  'Enterprise', 'Energy',       'South',    450000000),
    (210, 'Digital Media Co',      'SMB',        'Media',        'East',      25000000)
ON CONFLICT (customer_id) DO NOTHING;

-- ============================================================
-- Seed: Sales Transactions (Last 90 days)
-- ============================================================
INSERT INTO sales_transactions (
    sale_date, customer_id, salesperson, region, product_category,
    product_name, quantity, unit_price, discount_pct, revenue, cost, gross_profit
) VALUES
    (CURRENT_DATE - 5,  201, 'Ben Carter',    'West',    'Software', 'MinervaDB XL View Pro',   10, 2999, 10, 26991, 8997, 17994),
    (CURRENT_DATE - 8,  202, 'Cara Nguyen',   'East',    'Software', 'MinervaDB XL View Ent',    5, 9999,  5, 47495, 15000, 32495),
    (CURRENT_DATE - 12, 203, 'Ben Carter',    'Midwest', 'License',  'Annual License - Basic',   3, 4999,  0, 14997, 5000,  9997),
    (CURRENT_DATE - 15, 204, 'Cara Nguyen',   'East',    'Software', 'MinervaDB XL View Pro',   20, 2999, 15, 50983, 17994, 32989),
    (CURRENT_DATE - 20, 205, 'Ben Carter',    'West',    'Services', 'Implementation Package',   1,15000,  0, 15000, 6000,  9000),
    (CURRENT_DATE - 22, 206, 'Cara Nguyen',   'East',    'License',  'Annual License - Basic',   2, 4999,  5,  9498, 3333,  6165),
    (CURRENT_DATE - 30, 207, 'Derek Sullivan','West',    'Software', 'MinervaDB XL View SMB',    1, 1299,  0,  1299,  433,   866),
    (CURRENT_DATE - 35, 208, 'Ben Carter',    'Midwest', 'Software', 'MinervaDB XL View Pro',    8, 2999, 10, 21592, 7197, 14395),
    (CURRENT_DATE - 40, 209, 'Cara Nguyen',   'South',   'Software', 'MinervaDB XL View Ent',   15, 9999, 10,134987, 45000, 89987),
    (CURRENT_DATE - 45, 210, 'Derek Sullivan','East',    'Services', 'Training Package',          2, 3500,  0,  7000,  2100,  4900),
    (CURRENT_DATE - 50, 201, 'Ben Carter',    'West',    'Software', 'MinervaDB XL View Pro',   12, 2999,  8, 33109, 11994, 21115),
    (CURRENT_DATE - 55, 202, 'Cara Nguyen',   'East',    'Services', 'Premium Support',          1,24000,  0, 24000,  8000, 16000),
    (CURRENT_DATE - 60, 203, 'Ben Carter',    'Midwest', 'License',  'Annual License - Pro',     5, 7999,  5, 37995, 12500, 25495),
    (CURRENT_DATE - 65, 204, 'Cara Nguyen',   'South',   'Software', 'MinervaDB XL View Ent',   10, 9999, 12, 87991, 30000, 57991),
    (CURRENT_DATE - 70, 205, 'Derek Sullivan','West',    'Services', 'Implementation Package',   1,15000,  5, 14250,  6000,  8250),
    (CURRENT_DATE - 75, 206, 'Ben Carter',    'East',    'Software', 'MinervaDB XL View SMB',    3, 1299,  0,  3897,  1299,  2598),
    (CURRENT_DATE - 80, 207, 'Cara Nguyen',   'West',    'License',  'Annual License - Basic',   1, 4999,  0,  4999,  1666,  3333),
    (CURRENT_DATE - 85, 208, 'Derek Sullivan','Midwest', 'Software', 'MinervaDB XL View Pro',    6, 2999, 10, 16194,  5394, 10800),
    (CURRENT_DATE - 88, 209, 'Ben Carter',    'South',   'Services', 'Premium Support',          3,24000,  5, 68400, 24000, 44400),
    (CURRENT_DATE - 90, 210, 'Cara Nguyen',   'East',    'Software', 'MinervaDB XL View SMB',    2, 1299,  0,  2598,   866,  1732)
ON CONFLICT DO NOTHING;

-- ============================================================
-- Seed: Financial Data (Last 24 months)
-- ============================================================
INSERT INTO financial_data (
    period_date, account_category, account_name, department,
    actual_amount, budget_amount, prior_year_amount
)
SELECT
    gs.period,
    cat.account_category,
    cat.account_name,
    dep.department_name,
    ROUND((cat.base_amount * dep.size_factor * (0.85 + RANDOM() * 0.30))::NUMERIC, 2) AS actual_amount,
    ROUND((cat.base_amount * dep.size_factor)::NUMERIC, 2) AS budget_amount,
    ROUND((cat.base_amount * dep.size_factor * (0.80 + RANDOM() * 0.25))::NUMERIC, 2) AS prior_year_amount
FROM GENERATE_SERIES(
    DATE_TRUNC('month', CURRENT_DATE - INTERVAL '24 months'),
    DATE_TRUNC('month', CURRENT_DATE - INTERVAL '1 month'),
    '1 month'::INTERVAL
) AS gs(period)
CROSS JOIN (VALUES
    ('Revenue',          'Product Revenue',   500000),
    ('Revenue',          'Service Revenue',   150000),
    ('COGS',             'Cost of Goods Sold',180000),
    ('Operating Expense','Salaries',          280000),
    ('Operating Expense','Marketing Spend',    60000),
    ('Operating Expense','Office & Facilities', 25000),
    ('Operating Expense','Software & Tools',   18000),
    ('Depreciation',     'Depreciation',       12000)
) AS cat(account_category, account_name, base_amount)
CROSS JOIN (VALUES
    ('Sales',       1.20),
    ('Engineering', 0.85),
    ('Marketing',   0.95),
    ('Finance',     0.45),
    ('Operations',  0.75)
) AS dep(department_name, size_factor)
ON CONFLICT DO NOTHING;

-- ============================================================
-- Seed: Terminations (for attrition analysis)
-- ============================================================
INSERT INTO terminations (employee_id, termination_date, termination_type, termination_reason)
VALUES
    (1031, CURRENT_DATE - 120, 'Voluntary',   'Better opportunity'),
    (1032, CURRENT_DATE - 90,  'Voluntary',   'Relocation'),
    (1033, CURRENT_DATE - 60,  'Involuntary', 'Performance'),
    (1034, CURRENT_DATE - 180, 'Voluntary',   'Career change'),
    (1035, CURRENT_DATE - 30,  'Voluntary',   'Better opportunity')
ON CONFLICT (employee_id) DO NOTHING;

COMMIT;

-- ============================================================
-- Verification queries
-- ============================================================
SELECT 'departments' AS table_name, COUNT(*) AS row_count FROM departments
UNION ALL
SELECT 'employees',            COUNT(*) FROM employees
UNION ALL
SELECT 'customers',            COUNT(*) FROM customers
UNION ALL
SELECT 'sales_transactions',   COUNT(*) FROM sales_transactions
UNION ALL
SELECT 'financial_data',       COUNT(*) FROM financial_data
UNION ALL
SELECT 'terminations',         COUNT(*) FROM terminations
ORDER BY table_name;
