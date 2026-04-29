"""
为 company2 和 company3 生成模拟 CSV 数据。
数据格式与 company1 完全一致，但人名/地址/金额等做合理变化。
"""
import csv
import os
import random
import math

random.seed(42)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
COMPANY1 = os.path.join(SCRIPT_DIR, "company1")


def read_csv(filepath):
    with open(filepath, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        headers = next(reader)
        data = [row for row in reader]
    return headers, data


def write_csv(filepath, headers, data):
    os.makedirs(os.path.dirname(filepath), exist_ok=True)
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(data)
    print(f"  -> {filepath} ({len(data)} rows)")


# ── 数据池 ──────────────────────────────────────────

FIRST_NAMES_C2 = [
    "Liam", "Olivia", "Noah", "Emma", "Oliver", "Ava", "William", "Sophia",
    "Elijah", "Isabella", "James", "Mia", "Benjamin", "Charlotte", "Lucas",
]

FIRST_NAMES_C3 = [
    "Henry", "Amelia", "Alexander", "Harper", "Sebastian", "Evelyn", "Jack",
    "Abigail", "Owen", "Ella", "Aiden", "Scarlett", "Samuel", "Grace", "Joseph",
]

LAST_NAMES = [
    "Roberts", "Turner", "Phillips", "Campbell", "Parker", "Evans", "Edwards",
    "Collins", "Stewart", "Morris", "Nguyen", "Murphy", "Rivera", "Cook", "Morgan",
]

DEPARTMENT_POSITIONS = {
    "Engineering": [
        ("Senior Engineer", 17000, 22000),
        ("Frontend Developer", 13000, 16000),
        ("Backend Developer", 14000, 17500),
        ("QA Engineer", 11000, 14000),
        ("DevOps Engineer", 13500, 16500),
        ("Architect", 22000, 28000),
    ],
    "Marketing": [
        ("Marketing Manager", 14000, 16500),
        ("Brand Strategist", 11000, 13500),
        ("Sales Representative", 8000, 11000),
    ],
    "Finance": [
        ("Finance Supervisor", 15000, 17500),
        ("Accountant", 10000, 12500),
    ],
    "HR": [
        ("Recruiter", 9000, 11500),
        ("Training Supervisor", 12000, 14000),
    ],
    "Operations": [
        ("Operations Director", 20000, 24000),
        ("Content Editor", 8500, 10500),
    ],
}

CITIES = [
    ("Beijing", ["Haidian", "Chaoyang", "Dongcheng"]),
    ("Shanghai", ["Pudong", "Jingan", "Xuhui"]),
    ("Guangzhou", ["Tianhe", "Haizhu", "Liwan"]),
    ("Shenzhen", ["Nanshan", "Futian", "Luohu"]),
    ("Hangzhou", ["Xihu", "Binjiang", "Gongshu"]),
    ("Chengdu", ["Wuhou", "Gaoxin", "Jinniu"]),
    ("Nanjing", ["Gulou", "Xuanwu", "Jianye"]),
    ("Wuhan", ["Hongshan", "Wuchang", "Hankou"]),
    ("Xian", ["Yanta", "Beilin", "Weiyang"]),
    ("Chongqing", ["Yuzhong", "Jiangbei", "Shapingba"]),
    ("Tianjin", ["Nankai", "Heping", "Hexi"]),
    ("Suzhou", ["Gusu", "Huqiu", "Xiangcheng"]),
]

STREETS = [
    "Science Park Road", "Central Avenue", "East Lake Road", "West Ring Road",
    "Technology Street", "Harbour Boulevard", "Forest Lane", "Sunshine Avenue",
    "Fortune Plaza", "Dragon Street", "Phoenix Road", "Golden Boulevard",
    "Moonlight Road", "Star Avenue", "Crystal Street",
]

PHONES = []
for _ in range(100):
    PHONES.append("1" + str(random.randint(30, 39)) + "".join(str(random.randint(0, 9)) for _ in range(8)))

CUSTOMER_NAMES_C2 = [
    "Tyler Moore", "Zoe Reed", "Dylan Cooper", "Hannah Bell", "Caleb Cox",
    "Aria Ward", "Mason Hughes", "Lily Torres", "Logan Peterson", "Chloe Ramirez",
    "Jackson Long", "Sofia Flores", "Aiden Jenkins", "Riley Sullivan", "Carter Myers",
]

CUSTOMER_NAMES_C3 = [
    "Gabriel West", "Penelope Fisher", "Julian Hamilton", "Layla Tucker", "Isaac Butler",
    "Stella Gibson", "Levi Gordon", "Violet Hansen", "Dylan Larson", "Naomi Kennedy",
    "Nathaniel Shaw", "Eleanor Warren", "Carson Spencer", "Madelyn Carpenter", "Dominic Cole",
]

# 通用的20个产品（与company1相同产品名，但价格可能不同）
PRODUCTS_MASTER = [
    ("P001", "Mechanical Keyboard K1",  "Peripherals", 220),
    ("P002", "Wireless Mouse M200",     "Peripherals", 65),
    ("P003", "27-inch Monitor",          "Monitors",    1050),
    ("P004", "USB-C Hub Dongle",         "Accessories", 120),
    ("P005", "Noise-Cancelling Headphones N50", "Audio", 320),
    ("P006", "Portable HDD 2TB",        "Storage",     280),
    ("P007", "HD Webcam",               "Peripherals", 85),
    ("P008", "Bluetooth Speaker S30",   "Audio",       160),
    ("P009", "Ergonomic Chair",          "Furniture",   880),
    ("P010", "Laptop Stand",            "Accessories", 35),
    ("P011", "4K HDMI Cable 2m",        "Cables",      12),
    ("P012", "Desk Lamp L100",          "Lighting",    72),
    ("P013", "Mini Projector",           "Projectors",  1800),
    ("P014", "Gamepad G1",              "Gaming",      135),
    ("P015", "SSD 1TB",                 "Storage",     310),
    ("P016", "Laptop Sleeve",           "Accessories", 18),
    ("P017", "Wireless Charger Pad",    "Accessories", 28),
    ("P018", "Monitor Arm Mount",        "Accessories", 95),
    ("P019", "Earbuds E20",             "Audio",       55),
    ("P020", "Smart Plug",              "Smart Home",  25),
]


# ── 生成函数 ──────────────────────────────────────────

def gen_salary(base, perf_ratio):
    """生成 BaseSalary 和 PerformanceBonus，并计算社保公积金和净工资。"""
    bonus = round(base * perf_ratio)
    si = round(base * 0.12)
    hf = round(base * 0.10)
    net = base + bonus - si - hf
    return base, bonus, si, hf, net


def gen_employees(company_num, first_names):
    """生成 employees.csv"""
    random.shuffle(first_names)
    employees = []
    dept_pool = list(DEPARTMENT_POSITIONS.items())
    emp_id = 1
    for dept, positions in dept_pool:
        for pos_name, min_sal, max_sal in positions:
            base = random.randint(min_sal, max_sal)
            base = round(base / 100) * 100  # 四舍五入到百
            perf_ratio = random.uniform(0.15, 0.40)
            bonus_perf = round(perf_ratio * 100)
            # 根据职位级别调整绩效比例
            if "Director" in pos_name or "Supervisor" in pos_name or "Manager" in pos_name:
                perf_ratio = max(perf_ratio, 0.28)
            elif "Senior" in pos_name or "Architect" in pos_name:
                perf_ratio = max(perf_ratio, 0.22)

            _, bonus, si, hf, net = gen_salary(base, perf_ratio)

            hire_year = random.choice([2017, 2018, 2019, 2020, 2021, 2022])
            hire_month = random.randint(1, 12)
            hire_day = random.randint(1, 28)
            hire_date = f"{hire_year}-{hire_month:02d}-{hire_day:02d}"

            first = first_names[emp_id - 1]
            last = random.choice(LAST_NAMES)
            eid = f"E{emp_id:03d}"

            employees.append([
                eid, first, last, dept, pos_name, hire_date,
                base, bonus, si, hf, net,
            ])
            emp_id += 1

    return employees


def gen_products(company_num):
    """生成 products.csv，基于通用产品列表但价格有变动。"""
    random.seed(company_num * 100 + 1)
    suppliers = {
        "Peripherals": ["Shenzhen Precision Electronics", "Dongguan Optoelectronics", "Hangzhou VisionLink"],
        "Monitors":    ["Suzhou Vision Tech", "Shenzhen Display Co"],
        "Accessories": ["Guangzhou Connector Tech", "Yiwu ProHardware", "Ningbo ComfortSeat", "Shenzhen Precision Electronics"],
        "Audio":       ["Huizhou Acoustics", "Foshan SoundJoy"],
        "Storage":     ["Shanghai Storage Star", "Shenzhen MemTech"],
        "Furniture":   ["Ningbo ComfortSeat", "Guangzhou ErgoPlus"],
        "Cables":      ["Dongguan Xinda Electronics", "Shenzhen LinkPro"],
        "Lighting":    ["Zhongshan BrightView", "Foshan Illumination"],
        "Projectors":  ["Shenzhen ShadowTech", "Guangzhou OptiView"],
        "Gaming":      ["Guangzhou JoyPlay", "Shenzhen GamePower"],
        "Smart Home":  ["Hangzhou SmartLink", "Beijing IoT Solutions"],
    }
    warehouses = {
        "Peripherals": "A", "Monitors": "B", "Accessories": "A",
        "Audio": "C", "Storage": "D", "Furniture": "E",
        "Cables": "F", "Lighting": "G", "Projectors": "B",
        "Gaming": "A", "Smart Home": "G",
    }
    statuses = ["Active", "Active", "Active", "Active", "Promo"]

    products = []
    for pid, name, cat, cost in PRODUCTS_MASTER:
        price_variance = random.uniform(-0.15, 0.20)
        unit_price = max(int(cost * (1 + price_variance) / 5) * 5, cost + 5)

        stock = random.randint(100, 6000)
        sold = random.randint(300, 15000)
        profit_margin = round((unit_price - cost) / unit_price, 2)

        list_year = random.choice([2022, 2023])
        list_month = random.randint(1, 9)
        list_day = random.randint(1, 25)
        list_date = f"{list_year}-{list_month:02d}-{list_day:02d}"

        status = random.choice(statuses)
        wh = warehouses.get(cat, "A")
        wh_loc = f"{wh}-{random.randint(1,4):02d}-{random.randint(1,6):02d}"

        supplier = random.choice(suppliers.get(cat, ["Generic Supplier"]))
        products.append([
            pid, name, cat, supplier, unit_price, stock, sold, cost,
            profit_margin, list_date, status, wh_loc,
        ])

    return products


def gen_orders(company_num, month_offset, customer_names, price_mult_map=None):
    """
    生成 orders.csv 或 orders2.csv。
    month_offset: 0=第一个月, 1=第二个月
    price_mult_map: 若提供，对每个产品的单价做调整（用于第二个月份的价格变动）
    """
    random.seed(company_num * 100 + 10 + month_offset)
    orders = []

    customer_pool = customer_names[:]
    random.shuffle(customer_pool)

    used_names = []
    # 12个订单，复用客户池
    for i in range(12):
        cid = f"ORD2024{str(i+1).zfill(3)}" if i < 9 else f"ORD2004{str(i+1).zfill(3)}"

        if not customer_pool:
            customer_pool = used_names[:]
            random.shuffle(customer_pool)
        customer = customer_pool.pop()
        used_names.append(customer)

        phone = random.choice(PHONES)
        product_idx = i % 20
        pid, name, cat, cost = PRODUCTS_MASTER[product_idx]

        # 价格决定
        if price_mult_map and name in price_mult_map:
            unit_price = price_mult_map[name]
        else:
            # 用这个公司生成 products.csv 时的定价逻辑对齐
            random.seed(company_num * 100 + 1 + product_idx)
            price_variance = random.uniform(-0.15, 0.20)
            unit_price = int(cost * (1 + price_variance) / 5) * 5
            if unit_price <= cost:
                unit_price = cost + 5
            random.seed(company_num * 100 + 10 + month_offset)

        qty = random.choice([1, 1, 1, 2, 2, 3, 4, 5, 10])
        total_amount = unit_price * qty

        month = 1 + month_offset
        day = random.randint(1, 28)
        hour = random.randint(8, 17)
        minute = random.randint(0, 59)
        date_str = f"2024-{month:02d}-{day:02d} {hour}:{minute:02d}"

        city, districts = random.choice(CITIES)
        district = random.choice(districts)
        street_num = random.randint(8, 200)
        street = random.choice(STREETS)
        address = f"{street_num} {street} {district} {city}"

        method = random.choice(["WeChat", "Alipay", "WeChat", "Alipay", "Credit Card", "WeChat"])
        status = random.choices(
            ["Completed", "Shipped", "Pending", "Cancelled"],
            weights=[5, 2, 1, 1],
        )[0]

        orders.append([cid, customer, phone, name, qty, unit_price, total_amount, date_str, method, address, status])

    return orders


def gen_price_map_from_products(company_num):
    """从 products.csv 中提取产品名->单价 映射。"""
    _, prod_data = read_csv(os.path.join(COMPANY1, "products.csv"))
    # 但对于 company2/3，我们按照相同 seed 生成
    random.seed(company_num * 100 + 1)
    price_map = {}
    for pid, name, cat, cost in PRODUCTS_MASTER:
        price_variance = random.uniform(-0.15, 0.20)
        unit_price = int(cost * (1 + price_variance) / 5) * 5
        if unit_price <= cost:
            unit_price = cost + 5
        price_map[name] = unit_price
    return price_map


def gen_orders2_price_map(company_num, base_price_map):
    """基于第一个月的价格生成第二个月的价格（部分调整）。"""
    price_map = base_price_map.copy()
    # 对 4-6 个产品做价格调整
    names = list(price_map.keys())
    rng = random.Random(company_num * 100 + 77)
    to_adjust = rng.sample(names, rng.randint(4, 6))
    for name in to_adjust:
        adj_type = rng.choices(
            ["increase_big", "increase_small", "decrease", "same"],
            weights=[2, 2, 2, 1],
        )[0]
        base = base_price_map[name]
        if adj_type == "increase_big":
            price_map[name] = base + rng.randint(100, 450)
        elif adj_type == "increase_small":
            price_map[name] = base + rng.randint(6, 99)
        elif adj_type == "decrease":
            price_map[name] = max(base - rng.randint(50, 300), base // 2)
        # else: same, keep as is
    return price_map


# ── 主流程 ──────────────────────────────────────────

def main():
    for company_num in [2, 3]:
        print(f"\n=== Generating company{company_num} ===")
        out_dir = os.path.join(SCRIPT_DIR, f"company{company_num}")

        first_names = FIRST_NAMES_C2 if company_num == 2 else FIRST_NAMES_C3
        customer_names = CUSTOMER_NAMES_C2 if company_num == 2 else CUSTOMER_NAMES_C3

        # employees
        headers = ["EmployeeID", "FirstName", "LastName", "Department", "Position",
                    "HireDate", "BaseSalary", "PerformanceBonus", "SocialInsurance",
                    "HousingFund", "NetSalary"]
        emp_data = gen_employees(company_num, list(first_names))
        write_csv(os.path.join(out_dir, "employees.csv"), headers, emp_data)

        # products
        headers_p = ["ProductID", "ProductName", "Category", "Supplier", "UnitPrice",
                      "StockQuantity", "SoldQuantity", "CostPrice", "ProfitMargin",
                      "ListDate", "Status", "WarehouseLocation"]
        prod_data = gen_products(company_num)
        write_csv(os.path.join(out_dir, "products.csv"), headers_p, prod_data)

        # products2 (company's updated product catalog)
        # Re-seed to get slightly different data from products.csv
        random.seed(company_num * 100 + 50)
        prod2_data = gen_products(company_num + 10)  # different seed
        write_csv(os.path.join(out_dir, "products2.csv"), headers_p, prod2_data)

        # orders (month 1)
        headers_o = ["OrderID", "CustomerName", "Phone", "ProductName", "Quantity",
                      "UnitPrice", "TotalAmount", "OrderDate", "PaymentMethod",
                      "ShippingAddress", "OrderStatus"]
        base_price_map = gen_price_map_from_products(company_num)
        ord_data = gen_orders(company_num, 0, list(customer_names), base_price_map)
        write_csv(os.path.join(out_dir, "orders.csv"), headers_o, ord_data)

        # orders2 (month 2, with price adjustments)
        ord2_price_map = gen_orders2_price_map(company_num, base_price_map)
        ord2_data = gen_orders(company_num, 1, list(customer_names), ord2_price_map)
        write_csv(os.path.join(out_dir, "orders2.csv"), headers_o, ord2_data)

    print("\nDone! Generated company2/ and company3/ data.")


if __name__ == "__main__":
    main()
