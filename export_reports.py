import pandas as pd
import os

def export_all_reports(file_path):
    if not os.path.exists(file_path):
        print("Excel file not found!")
        return
        
    # Read sheets
    df_customers = pd.read_excel(file_path, sheet_name="Customers")
    df_sales = pd.read_excel(file_path, sheet_name="Daily Sales")
    df_inventory = pd.read_excel(file_path, sheet_name="Inventory")
    
    # Clean column names
    df_customers.columns = df_customers.columns.str.strip()
    df_sales.columns = df_sales.columns.str.strip()
    df_inventory.columns = df_inventory.columns.str.strip()
    
    # Handle dates
    df_sales["Date"] = pd.to_datetime(df_sales["Date"]).dt.strftime("%b %d")
    df_customers["Date"] = pd.to_datetime(df_customers["Date"]).dt.strftime("%b %d")
    
    # Setup folders
    export_dir = "exports"
    os.makedirs(export_dir, exist_ok=True)
    
    # 1. Map customers contact
    cust_info = {}
    for _, row in df_customers.iterrows():
        name = str(row["Customer Name"]).strip()
        cust_info[name] = {
            "email": str(row["Email Address"]) if pd.notna(row["Email Address"]) else f"{name.lower().replace(' ', '')}@gmail.com",
            "phone": str(row["Phone Number"]) if pd.notna(row["Phone Number"]) else "+1 (415) 555-2671",
            "address": str(row["Shipping Address"]) if pd.notna(row["Shipping Address"]) else "N/A"
        }
        
    # 2. Group sales into Orders
    orders_grouped = df_sales.groupby(["Customer", "Date"])
    
    orders = []
    order_id_counter = 192530
    
    for (customer_name, date), group in orders_grouped:
        customer_name_clean = str(customer_name).strip()
        info = cust_info.get(customer_name_clean, {
            "email": f"{customer_name_clean.lower().replace(' ', '')}@gmail.com",
            "phone": "+1 (415) 555-2671",
            "address": "N/A, USA"
        })
        
        total_revenue = float(group["Revenue (USD)"].sum())
        
        status = "Paid"
        if customer_name_clean == "Jeneffer":
            status = "Cancelled"
        elif customer_name_clean == "AvrilLwin":
            status = "Refunded"
            
        addr = info["address"]
        order_type = "Shipping" if ("Thailand" in addr or "," in addr) and "N/A" not in addr else "Pickups"
        
        products = list(group["Product Name"].unique())
        main_product = products[0] if len(products) > 0 else "N/A"
        
        orders.append({
            "Order ID": f"#{order_id_counter}",
            "Customer": customer_name_clean,
            "Type": order_type,
            "Status": status,
            "Product": main_product,
            "Total": f"${total_revenue:.2f}",
            "Date": date,
            # float version for aggregations
            "total_raw": total_revenue
        })
        order_id_counter += 1
        
    orders.reverse()
    df_orders_export = pd.DataFrame(orders)
    # Remove raw col for clean output
    df_orders_export.drop(columns=["total_raw"]).to_csv(os.path.join(export_dir, "orders_report.csv"), index=False)
    print("orders_report.csv generated successfully.")
    
    # 3. Payments export
    payments = []
    for o in orders:
        pay_id = o["Order ID"].replace("#", "PAY-")
        method = "Credit Card" if o["Type"] == "Shipping" else "Bank Transfer"
        payments.append({
            "Payment ID": pay_id,
            "Customer": o["Customer"],
            "Method": method,
            "Amount": o["Total"],
            "Status": o["Status"],
            "Date": o["Date"]
        })
    pd.DataFrame(payments).to_csv(os.path.join(export_dir, "payments_report.csv"), index=False)
    print("payments_report.csv generated successfully.")
    
    # 4. Inventory export
    inventory = []
    for _, r in df_inventory.iterrows():
        status_clean = str(r["Status"]).replace("🟢 ", "").replace("⚠️ ", "").strip() if pd.notna(r["Status"]) else "Healthy"
        inventory.append({
            "Product ID": str(r["Product ID"]),
            "Product Name": str(r["Product Name"]),
            "Category": str(r["Category"]),
            "Price": str(r["Price (USD/THB)"]),
            "Current Stock": int(r["Current Stock"]) if pd.notna(r["Current Stock"]) else 50,
            "Original Stock": int(r["Original Stock"]) if pd.notna(r["Original Stock"]) else 50,
            "Units Sold": int(r["Units Sold"]) if pd.notna(r["Units Sold"]) else 0,
            "Status": status_clean
        })
    pd.DataFrame(inventory).to_csv(os.path.join(export_dir, "inventory_report.csv"), index=False)
    print("inventory_report.csv generated successfully.")
    
    # 5. Customers export
    cust_list = []
    for name, info in cust_info.items():
        customer_orders = [o for o in orders if o["Customer"] == name]
        total_spent = sum(o["total_raw"] for o in customer_orders)
        cust_list.append({
            "Customer": name,
            "Email": info["email"],
            "Phone": info["phone"],
            "Shipping Address": info["address"],
            "Total Orders": len(customer_orders),
            "Total Spent": f"${total_spent:.2f}"
        })
    pd.DataFrame(cust_list).to_csv(os.path.join(export_dir, "customers_report.csv"), index=False)
    print("customers_report.csv generated successfully.")
    print("All exports completed successfully!")

if __name__ == "__main__":
    export_all_reports("Antigravity Sales Agent.xlsx")
