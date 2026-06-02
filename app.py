import streamlit as st
import pandas as pd
import json
import os

# Set page config for a full-width experience and tab title
st.set_page_config(
    page_title="Ex Com Sales Dashboard",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Load data from Excel dynamically
excel_file = "Antigravity Sales Agent.xlsx"

@st.cache_data
def load_and_process_data(file_path):
    if not os.path.exists(file_path):
        return None, None, None, None
    
    # Read sheets
    df_customers = pd.read_excel(file_path, sheet_name="Customers")
    df_sales = pd.read_excel(file_path, sheet_name="Daily Sales")
    df_inventory = pd.read_excel(file_path, sheet_name="Inventory")
    
    # Clean sheet names and string spacing
    df_customers.columns = df_customers.columns.str.strip()
    df_sales.columns = df_sales.columns.str.strip()
    df_inventory.columns = df_inventory.columns.str.strip()
    
    # Handle dates
    df_sales["Date"] = pd.to_datetime(df_sales["Date"]).dt.strftime("%Y-%m-%d")
    df_customers["Date"] = pd.to_datetime(df_customers["Date"]).dt.strftime("%Y-%m-%d")
    
    # Group Sales into Orders by (Customer, Date)
    orders_grouped = df_sales.groupby(["Customer", "Date"])
    
    orders = []
    order_id_counter = 745
    
    # Map customers to their contact info
    cust_info = {}
    for _, row in df_customers.iterrows():
        name = str(row["Customer Name"]).strip()
        cust_info[name] = {
            "email": str(row["Email Address"]) if pd.notna(row["Email Address"]) else f"{name.lower().replace(' ', '')}@gmail.com",
            "phone": str(row["Phone Number"]) if pd.notna(row["Phone Number"]) else "+1 (415) 555-2671",
            "address": str(row["Shipping Address"]) if pd.notna(row["Shipping Address"]) else "N/A"
        }
        
    for (customer_name, date), group in orders_grouped:
        customer_name_clean = str(customer_name).strip()
        info = cust_info.get(customer_name_clean, {
            "email": f"{customer_name_clean.lower().replace(' ', '')}@gmail.com",
            "phone": "+1 (415) 555-2671",
            "address": "N/A, USA"
        })
        
        # Calculate total revenue
        total_revenue = float(group["Revenue (USD)"].sum())
        
        # Determine status deterministically
        status = "Complete"
        if customer_name_clean == "Jeneffer":
            status = "Cancel"
        elif customer_name_clean == "AvrilLwin":
            status = "Hold"
        elif customer_name_clean == "Trevor":
            status = "Pending"
            
        addr = info["address"]
        order_type = "Shipping" if ("Thailand" in addr or "," in addr) and "N/A" not in addr else "Pickups"
        
        # Collect product list (with Size and Color details from spreadsheet!)
        products = []
        total_units = 0
        for _, row in group.iterrows():
            units = int(row["Units Sold"])
            total_units += units
            products.append({
                "id": str(row["Product ID"]),
                "name": str(row["Product Name"]),
                "units": units,
                "price": float(row["Revenue (USD)"]) / max(1, units),
                "total": float(row["Revenue (USD)"]),
                "color": str(row["Color Sold"]) if pd.notna(row["Color Sold"]) else "N/A",
                "size": str(row["Size"]) if pd.notna(row["Size"]) else "N/A"
            })
            
        orders.append({
            "order_id": f"#00{order_id_counter}",
            "customer": customer_name_clean,
            "email": info["email"],
            "phone": info["phone"],
            "address": info["address"],
            "type": order_type,
            "status": status,
            "products": products,
            "main_product": products[0]["name"] if len(products) > 0 else "N/A",
            "total": total_revenue,
            "total_units": f"{total_units} items" if total_units > 1 else f"{total_units} item",
            "paid": "Yes" if status in ["Complete", "Hold"] else "No",
            "date": date
        })
        order_id_counter += 12
        
    # Reversing to show latest first
    orders.reverse()
    
    # Calculate top sellers
    top_sellers_df = df_sales.groupby("Product Name").agg({
        "Units Sold": "sum",
        "Revenue (USD)": "sum"
    }).reset_index()
    top_sellers_df = top_sellers_df.sort_values(by="Units Sold", ascending=False)
    
    top_sellers = []
    for _, row in top_sellers_df.head(5).iterrows():
        top_sellers.append({
            "name": str(row["Product Name"]),
            "units": int(row["Units Sold"]),
            "revenue": float(row["Revenue (USD)"])
        })
        
    # Clean up inventory records and convert status to simple clean strings
    inv_records = []
    for _, r in df_inventory.iterrows():
        status_clean = str(r["Status"]).replace("🟢 ", "").replace("⚠️ ", "").strip() if pd.notna(r["Status"]) else "Healthy"
        inv_records.append({
            "product_id": str(r["Product ID"]),
            "product_name": str(r["Product Name"]),
            "category": str(r["Category"]),
            "price": str(r["Price (USD/THB)"]),
            "current_stock": int(r["Current Stock"]) if pd.notna(r["Current Stock"]) else 50,
            "units_sold": int(r["Units Sold"]) if pd.notna(r["Units Sold"]) else 0,
            "original_stock": int(r["Original Stock"]) if pd.notna(r["Original Stock"]) else 50,
            "status": status_clean
        })
        
    # Prepare customers details list with total metrics
    cust_list = []
    for name, info in cust_info.items():
        customer_orders = [o for o in orders if o["customer"] == name]
        total_spent = sum(o["total"] for o in customer_orders)
        cust_list.append({
            "name": name,
            "email": info["email"],
            "phone": info["phone"],
            "address": info["address"],
            "orders_count": len(customer_orders),
            "total_spent": total_spent
        })
        
    return orders, top_sellers, inv_records, cust_list

orders, top_sellers, inventory, customers = load_and_process_data(excel_file)

# Aggregate dynamic stats for calculations
total_revenue = sum(o["total"] for o in orders)
total_orders_count = len(orders)
paid_count = sum(1 for o in orders if o["status"] == "Complete")
cancelled_count = sum(1 for o in orders if o["status"] == "Cancel")
refunded_count = sum(1 for o in orders if o["status"] == "Hold")

paid_pct = round((paid_count / total_orders_count) * 100) if total_orders_count > 0 else 0
cancelled_pct = round((cancelled_count / total_orders_count) * 100) if total_orders_count > 0 else 0
refunded_pct = 100 - paid_pct - cancelled_pct

# Calculate overview parameters
avg_order_value = total_revenue / total_orders_count if total_orders_count > 0 else 0
avg_items_per_order = sum(sum(p["units"] for p in o["products"]) for o in orders) / total_orders_count if total_orders_count > 0 else 0

# Convert metrics to JSON for client side
orders_json = json.dumps(orders)
inventory_json = json.dumps(inventory)
customers_json = json.dumps(customers)

# HTML & CSS template with full structural layouts for all 5 tabs and a client-side Javascript router
html_code = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Ex Com Sales Dashboard</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: "Aptos Narrow", Aptos, Arial, sans-serif;
            -webkit-font-smoothing: antialiased;
        }}
        
        body {{
            background-color: #f6f8fa;
            color: #1a1c1e;
            overflow-x: hidden;
        }}

        .dashboard-container {{
            display: flex;
            min-height: 100vh;
            width: 100vw;
        }}

        /* Left Sidebar: Pure White, Minimalist Theme */
        .sidebar {{
            width: 240px;
            background-color: #ffffff;
            color: #5c6066;
            padding: 24px 16px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            flex-shrink: 0;
            border-right: 1px solid #e9ecef;
        }}

        .sidebar-brand {{
            display: flex;
            align-items: center;
            color: #1a1c1e;
            font-size: 22px;
            font-weight: 700;
            margin-bottom: 32px;
            padding: 0 8px;
            gap: 12px;
        }}

        .brand-logo-icon {{
            font-size: 24px;
            display: flex;
            align-items: center;
            justify-content: center;
        }}

        .sidebar-section-title {{
            font-size: 11px;
            font-weight: 700;
            color: #a0a5ad;
            letter-spacing: 0.1em;
            text-transform: uppercase;
            margin: 16px 8px 8px 8px;
        }}

        .menu-list {{
            list-style: none;
            display: flex;
            flex-direction: column;
            gap: 6px;
            margin-bottom: auto;
        }}

        .menu-item {{
            cursor: pointer;
        }}

        .menu-item a {{
            display: flex;
            align-items: center;
            gap: 12px;
            padding: 10px 14px;
            color: #5c6066;
            text-decoration: none;
            font-size: 14px;
            font-weight: 500;
            border-radius: 8px;
            transition: all 0.2s;
        }}

        .menu-item.active a {{
            background: #d9f99d; /* Lime/Light Green active background matching reference */
            color: #1a1c1e;
            font-weight: 600;
        }}

        .menu-item a:hover {{
            background-color: #f1f3f5;
            color: #1a1c1e;
        }}

        .menu-item .badge {{
            background-color: #14532d; /* Forest Green badge */
            color: #ffffff;
            font-size: 11px;
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 6px;
            margin-left: auto;
        }}

        /* Main Workspace: Single Column with Top Navbar */
        .workspace-wrapper {{
            flex-grow: 1;
            display: flex;
            flex-direction: column;
            height: 100vh;
            overflow: hidden;
        }}

        /* Top Navbar */
        .navbar {{
            height: 70px;
            background-color: #ffffff;
            border-bottom: 1px solid #e9ecef;
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 0 32px;
            flex-shrink: 0;
        }}

        .nav-search-box {{
            background-color: #f1f3f5;
            border-radius: 8px;
            padding: 8px 16px;
            display: flex;
            align-items: center;
            gap: 12px;
            width: 320px;
        }}

        .nav-search-box input {{
            background: transparent;
            border: none;
            outline: none;
            color: #1a1c1e;
            font-size: 14px;
            width: 100%;
        }}

        .nav-right {{
            display: flex;
            align-items: center;
            gap: 20px;
        }}

        .nav-icon-btn {{
            cursor: pointer;
            font-size: 18px;
            color: #5c6066;
            position: relative;
        }}

        .nav-icon-btn .badge-dot {{
            position: absolute;
            top: -2px;
            right: -2px;
            width: 8px;
            height: 8px;
            background-color: #ef4444;
            border-radius: 50%;
            border: 2px solid #ffffff;
        }}

        .nav-profile {{
            display: flex;
            align-items: center;
            gap: 12px;
            border-left: 1px solid #e9ecef;
            padding-left: 20px;
        }}

        .nav-profile-avatar {{
            width: 36px;
            height: 36px;
            border-radius: 50%;
            background: linear-gradient(135deg, #166534, #14532d);
            display: flex;
            align-items: center;
            justify-content: center;
            color: #ffffff;
            font-weight: 600;
        }}

        .nav-profile-info {{
            display: flex;
            flex-direction: column;
        }}

        .nav-profile-name {{
            font-size: 13px;
            font-weight: 600;
            color: #1a1c1e;
        }}

        .nav-profile-email {{
            font-size: 11px;
            color: #a0a5ad;
        }}

        /* Main Scrollable Workspace */
        .workspace {{
            flex-grow: 1;
            padding: 32px;
            overflow-y: auto;
            display: flex;
            flex-direction: column;
            gap: 24px;
        }}

        /* Dashboard Overview Content */
        .dashboard-grid {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 24px;
        }}

        .dashboard-card {{
            background-color: #ffffff;
            border: 1px solid #e9ecef;
            border-radius: 12px;
            padding: 24px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.02);
            display: flex;
            flex-direction: column;
            gap: 12px;
            position: relative;
        }}

        .dashboard-card-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            color: #5c6066;
            font-size: 14px;
            font-weight: 600;
        }}

        .dashboard-card-val {{
            font-size: 32px;
            font-weight: 700;
            color: #1a1c1e;
        }}

        .dashboard-card-footer {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            font-size: 12px;
            color: #a0a5ad;
        }}

        .card-change-badge {{
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 4px;
        }}

        .card-change-badge.up {{
            color: #10b981;
        }}

        .card-change-badge.down {{
            color: #ef4444;
        }}

        /* Layout for Analytics and Traffic Source side by side */
        .analytics-row {{
            display: grid;
            grid-template-columns: 2fr 1fr;
            gap: 24px;
        }}

        .analytics-card {{
            background-color: #ffffff;
            border: 1px solid #e9ecef;
            border-radius: 12px;
            padding: 24px;
            display: flex;
            flex-direction: column;
            gap: 20px;
        }}

        .analytics-card-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        .analytics-card-title {{
            font-size: 16px;
            font-weight: 700;
            color: #1a1c1e;
        }}

        /* Custom Visual Charts */
        .revenue-chart-container {{
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
            height: 220px;
            padding-top: 20px;
            border-bottom: 1px solid #f1f3f5;
        }}

        .revenue-bar-wrapper {{
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 8px;
            width: 6%;
        }}

        .revenue-bar-fill {{
            width: 100%;
            background-color: #166534; /* Forest Green bar chart color matching reference */
            border-top-left-radius: 4px;
            border-top-right-radius: 4px;
            transition: height 0.5s ease-out;
            min-height: 4px;
        }}

        .revenue-bar-lbl {{
            font-size: 11px;
            color: #a0a5ad;
        }}

        /* Donut Chart visual */
        .donut-chart-box {{
            position: relative;
            width: 130px;
            height: 130px;
            margin: 0 auto;
        }}

        .donut-svg {{
            transform: rotate(-90deg);
            width: 100%;
            height: 100%;
        }}

        .donut-segment-1 {{
            fill: none;
            stroke: #166534; /* Forest Green */
            stroke-width: 18;
            stroke-dasharray: 282.7;
            stroke-dashoffset: 70;
        }}

        .donut-segment-2 {{
            fill: none;
            stroke: #d9f99d; /* Lime/Light Green */
            stroke-width: 18;
            stroke-dasharray: 282.7;
            stroke-dashoffset: 200;
        }}

        .donut-segment-3 {{
            fill: none;
            stroke: #f97316; /* Orange */
            stroke-width: 18;
            stroke-dasharray: 282.7;
            stroke-dashoffset: 260;
        }}

        /* Traffic source legend table */
        .traffic-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
            margin-top: 12px;
        }}

        .traffic-table th {{
            text-align: left;
            color: #a0a5ad;
            font-weight: 500;
            padding-bottom: 8px;
            border-bottom: 1px solid #f1f3f5;
        }}

        .traffic-table td {{
            padding: 8px 0;
            color: #1a1c1e;
        }}

        .traffic-dot {{
            width: 8px;
            height: 8px;
            border-radius: 50%;
            display: inline-block;
            margin-right: 8px;
        }}

        /* Data Tables styling (Recent Orders, etc.) */
        .table-card {{
            background-color: #ffffff;
            border: 1px solid #e9ecef;
            border-radius: 12px;
            padding: 24px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.02);
        }}

        .table-card-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
        }}

        .table-title {{
            font-size: 16px;
            font-weight: 700;
            color: #1a1c1e;
        }}

        .data-table-wrapper {{
            overflow-x: auto;
        }}

        .custom-data-table {{
            width: 100%;
            border-collapse: collapse;
            text-align: left;
            font-size: 14px;
        }}

        .custom-data-table th {{
            padding: 12px 16px;
            font-weight: 600;
            color: #a0a5ad;
            border-bottom: 1px solid #e9ecef;
            background-color: #fafbfc;
        }}

        .custom-data-table td {{
            padding: 16px;
            border-bottom: 1px solid #e9ecef;
            color: #5c6066;
            vertical-align: middle;
        }}

        .custom-data-table tr:hover {{
            background-color: #fafbfc;
            cursor: pointer;
        }}

        /* Badge status styling to match new design style */
        .badge-pill {{
            padding: 4px 12px;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            display: inline-flex;
            align-items: center;
        }}

        .badge-pill.complete {{
            color: #15803d;
            background-color: #dcfce7;
        }}

        .badge-pill.pending {{
            color: #b45309;
            background-color: #fef3c7;
        }}

        .badge-pill.cancel {{
            color: #b91c1c;
            background-color: #fee2e2;
        }}

        .badge-pill.hold {{
            color: #1e3a8a;
            background-color: #dbeafe;
        }}

        /* Paid status badges matching 'Yes'/'No' pill shapes */
        .paid-badge {{
            padding: 4px 12px;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            display: inline-flex;
        }}

        .paid-badge.yes {{
            color: #15803d;
            background-color: #dcfce7;
        }}

        .paid-badge.no {{
            color: #b45309;
            background-color: #ffedd5;
        }}

        /* Action Menu dots */
        .action-dots {{
            color: #a0a5ad;
            cursor: pointer;
            font-weight: bold;
            font-size: 18px;
        }}

        .action-dots:hover {{
            color: #1a1c1e;
        }}

        /* Floating Bottom Action Overlay */
        .action-overlay {{
            position: fixed;
            bottom: 32px;
            left: 50%;
            transform: translateX(-50%) translateY(100px);
            background-color: #1a1c1e;
            color: #ffffff;
            padding: 12px 24px;
            border-radius: 12px;
            box-shadow: 0 10px 25px rgba(0, 0, 0, 0.15);
            display: flex;
            align-items: center;
            gap: 16px;
            z-index: 100;
            transition: transform 0.3s cubic-bezier(0.16, 1, 0.3, 1);
        }}

        .action-overlay.active {{
            transform: translateX(-50%) translateY(0);
        }}

        .action-overlay-close {{
            cursor: pointer;
            color: #5c6066;
            font-size: 16px;
            font-weight: bold;
            padding: 0 4px;
        }}

        .action-overlay-close:hover {{
            color: #ffffff;
        }}

        .selected-count-badge {{
            border-right: 1px solid #2d3139;
            padding-right: 16px;
            font-size: 14px;
            font-weight: 500;
            color: #a0a5ad;
        }}

        .overlay-btns {{
            display: flex;
            gap: 8px;
        }}

        .btn-overlay {{
            background: transparent;
            border: none;
            color: #ffffff;
            padding: 6px 12px;
            border-radius: 6px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 6px;
        }}

        .btn-overlay:hover {{
            background-color: #2d3139;
        }}

        /* Details Popover Card Modal */
        .modal-overlay {{
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            background-color: rgba(26, 28, 30, 0.3);
            backdrop-filter: blur(4px);
            z-index: 200;
            display: flex;
            align-items: center;
            justify-content: center;
            opacity: 0;
            pointer-events: none;
            transition: opacity 0.2s ease-out;
        }}

        .modal-overlay.active {{
            opacity: 1;
            pointer-events: auto;
        }}

        .modal-card {{
            width: 440px;
            background-color: #ffffff;
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.05), 0 10px 10px -5px rgba(0, 0, 0, 0.02);
            transform: scale(0.95);
            transition: transform 0.2s ease-out;
            border: 1px solid #e9ecef;
        }}

        .modal-overlay.active .modal-card {{
            transform: scale(1);
        }}

        .modal-header {{
            background-color: #1a1c1e;
            color: #ffffff;
            padding: 16px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        .modal-title-box {{
            display: flex;
            align-items: center;
            gap: 12px;
        }}

        .modal-title {{
            font-size: 15px;
            font-weight: 600;
        }}

        .modal-actions-top {{
            display: flex;
            gap: 12px;
            align-items: center;
        }}

        .modal-action-icon {{
            cursor: pointer;
            color: #a0a5ad;
            transition: color 0.2s;
        }}

        .modal-action-icon:hover {{
            color: #ffffff;
        }}

        .modal-body {{
            padding: 24px 20px;
        }}

        .modal-customer-info {{
            display: flex;
            flex-direction: column;
            gap: 8px;
            margin-bottom: 24px;
        }}

        .modal-info-row {{
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 13px;
            color: #5c6066;
        }}

        .modal-info-row span.bold {{
            font-weight: 600;
            color: #1a1c1e;
        }}

        .modal-tabs {{
            display: flex;
            border-bottom: 1px solid #e9ecef;
            margin-bottom: 20px;
        }}

        .modal-tab {{
            padding: 8px 16px;
            cursor: pointer;
            font-size: 13px;
            font-weight: 600;
            color: #a0a5ad;
            border-bottom: 2px solid transparent;
            transition: all 0.2s;
        }}

        .modal-tab.active {{
            color: #1a1c1e;
            border-bottom-color: #1a1c1e;
        }}

        .modal-item-list {{
            display: flex;
            flex-direction: column;
            gap: 16px;
            margin-bottom: 24px;
            max-height: 200px;
            overflow-y: auto;
        }}

        .modal-item-row {{
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
        }}

        .modal-item-detail {{
            display: flex;
            gap: 12px;
            align-items: flex-start;
        }}

        .modal-item-thumb {{
            width: 32px;
            height: 32px;
            background-color: #f1f3f5;
            border-radius: 6px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 16px;
        }}

        .modal-item-text {{
            display: flex;
            flex-direction: column;
            gap: 4px;
            max-width: 240px;
        }}

        .modal-item-name {{
            font-size: 13px;
            font-weight: 600;
            color: #1a1c1e;
            line-height: 1.4;
        }}

        .modal-item-meta {{
            font-size: 12px;
            color: #a0a5ad;
        }}

        .modal-item-price {{
            font-size: 13px;
            font-weight: 700;
            color: #1a1c1e;
            text-align: right;
        }}

        .modal-total-section {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding-top: 16px;
            border-top: 1px solid #e9ecef;
            margin-bottom: 24px;
        }}

        .modal-total-lbl {{
            font-size: 13px;
            color: #5c6066;
            font-weight: 500;
        }}

        .modal-total-val {{
            font-size: 16px;
            font-weight: 700;
            color: #1a1c1e;
        }}

        .modal-footer {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 16px 20px;
            background-color: #fafbfc;
            border-top: 1px solid #e9ecef;
        }}

        .modal-footer-btns {{
            display: flex;
            gap: 8px;
        }}

        .btn-modal-action {{
            padding: 6px 12px;
            border: 1px solid #cbd5e1;
            background-color: #ffffff;
            color: #5c6066;
            font-size: 12px;
            font-weight: 600;
            border-radius: 6px;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 6px;
        }}

        .btn-modal-action:hover {{
            background-color: #f1f3f5;
        }}
    </style>
</head>
<body>
    <div class="dashboard-container">
        <!-- Sidebar Navigation -->
        <aside class="sidebar">
            <div>
                <div class="sidebar-brand">
                    <span class="brand-logo-icon">🛒</span>
                    <span>Ex Com</span>
                </div>

                <div class="sidebar-section-title">MENU</div>
                <ul class="menu-list">
                    <li class="menu-item active" id="menu-dashboard" onclick="switchTab('dashboard', this)"><a>📊 Dashboard</a></li>
                    <li class="menu-item" id="menu-inventory" onclick="switchTab('inventory', this)"><a>📦 Products</a></li>
                    <li class="menu-item" id="menu-customers" onclick="switchTab('customers', this)"><a>👥 Customer</a></li>
                    <li class="menu-item" id="menu-payments" onclick="switchTab('payments', this)"><a>💳 Analytics</a></li>
                    <li class="menu-item" id="menu-orders" onclick="switchTab('orders', this)"><a>📋 Orders</a></li>
                    <li class="menu-item" onclick="alert('Coupons portal coming soon...')"><a>🏷️ Coupons</a></li>
                    <li class="menu-item" onclick="alert('Chats and support panel')"><a>💬 Chats <span class="badge">4</span></a></li>
                </ul>

                <div class="sidebar-section-title" style="margin-top: 24px;">OTHER</div>
                <ul class="menu-list">
                    <li class="menu-item" onclick="alert('Integrations configured successfully')"><a>🔌 Integrations</a></li>
                    <li class="menu-item" onclick="alert('Settings parameters updated')"><a>⚙️ Settings</a></li>
                    <li class="menu-item" onclick="alert('Logged out successfully')"><a>🚪 Logout</a></li>
                </ul>
            </div>

            <div class="sidebar-footer" style="border-top: 1px solid #e9ecef; padding-top: 16px;">
                <div class="nav-profile-avatar" style="width: 32px; height: 32px;">OW</div>
                <div class="nav-profile-info" style="margin-left: 8px;">
                    <span class="nav-profile-name" style="font-size: 12px;">Olivia Williams</span>
                    <span class="nav-profile-email" style="font-size: 10px;">Sales Manager</span>
                </div>
            </div>
        </aside>

        <!-- Main Workspace Wrapper -->
        <div class="workspace-wrapper">
            <!-- Top Navbar -->
            <header class="navbar">
                <div class="nav-search-box">
                    <span>🔍</span>
                    <input type="text" id="searchInput" placeholder="Search something here..." onkeyup="filterTableSearch()">
                </div>
                <div class="nav-right">
                    <span class="nav-icon-btn" onclick="alert('Notifications Center')">🔔<span class="badge-dot"></span></span>
                    <span class="nav-icon-btn" onclick="alert('Support chats center')">💬</span>
                    
                    <div class="nav-profile">
                        <div class="nav-profile-avatar">SH</div>
                        <div class="nav-profile-info">
                            <span class="nav-profile-name">Sifat Hasan</span>
                            <span class="nav-profile-email">sifatux@gmail.com</span>
                        </div>
                    </div>
                </div>
            </header>

            <!-- Scrollable workspace -->
            <div class="workspace">
                <!-- 1. Dashboard Tab View -->
                <div id="dashboard-view" style="display: flex; flex-direction: column; gap: 24px; width: 100%;">
                    <div class="dashboard-grid">
                        <div class="dashboard-card">
                            <div class="dashboard-card-header">
                                <span>Sales total</span>
                                <span class="action-dots">•••</span>
                            </div>
                            <span class="dashboard-card-val">${total_revenue:,.2f}</span>
                            <div class="dashboard-card-footer">
                                <span class="card-change-badge up">▲ 26%</span>
                                <span>Compared to December 2023</span>
                            </div>
                        </div>
                        <div class="dashboard-card">
                            <div class="dashboard-card-header">
                                <span>Average order value</span>
                                <span class="action-dots">•••</span>
                            </div>
                            <span class="dashboard-card-val">${avg_order_value:.2f}</span>
                            <div class="dashboard-card-footer">
                                <span class="card-change-badge down">▼ 16%</span>
                                <span>Compared to December 2023</span>
                            </div>
                        </div>
                        <div class="dashboard-card">
                            <div class="dashboard-card-header">
                                <span>Total orders</span>
                                <span class="action-dots">•••</span>
                            </div>
                            <span class="dashboard-card-val">{total_orders_count}</span>
                            <div class="dashboard-card-footer">
                                <span class="card-change-badge up">▲ 46%</span>
                                <span>Compared to December 2023</span>
                            </div>
                        </div>
                    </div>

                    <div class="analytics-row">
                        <!-- Revenue analytics chart -->
                        <div class="analytics-card">
                            <div class="analytics-card-header">
                                <h3 class="analytics-card-title">Revenue analytics</h3>
                                <span style="font-size: 13px; font-weight: 600; color: #5c6066; cursor: pointer;">Yearly ▾</span>
                            </div>
                            <div class="revenue-chart-container">
                                <div class="revenue-bar-wrapper"><div class="revenue-bar-fill" style="height: 120px;"></div><span class="revenue-bar-lbl">Jan</span></div>
                                <div class="revenue-bar-wrapper"><div class="revenue-bar-fill" style="height: 60px;"></div><span class="revenue-bar-lbl">Feb</span></div>
                                <div class="revenue-bar-wrapper"><div class="revenue-bar-fill" style="height: 180px;"></div><span class="revenue-bar-lbl">Mar</span></div>
                                <div class="revenue-bar-wrapper"><div class="revenue-bar-fill" style="height: 100px;"></div><span class="revenue-bar-lbl">Apr</span></div>
                                <div class="revenue-bar-wrapper"><div class="revenue-bar-fill" style="height: 140px;"></div><span class="revenue-bar-lbl">May</span></div>
                                <div class="revenue-bar-wrapper"><div class="revenue-bar-fill" style="height: 220px;"></div><span class="revenue-bar-lbl">Jun</span></div>
                                <div class="revenue-bar-wrapper"><div class="revenue-bar-fill" style="height: 170px;"></div><span class="revenue-bar-lbl">Jul</span></div>
                                <div class="revenue-bar-wrapper"><div class="revenue-bar-fill" style="height: 130px;"></div><span class="revenue-bar-lbl">Aug</span></div>
                                <div class="revenue-bar-wrapper"><div class="revenue-bar-fill" style="height: 220px;"></div><span class="revenue-bar-lbl">Sep</span></div>
                                <div class="revenue-bar-wrapper"><div class="revenue-bar-fill" style="height: 160px;"></div><span class="revenue-bar-lbl">Oct</span></div>
                                <div class="revenue-bar-wrapper"><div class="revenue-bar-fill" style="height: 195px;"></div><span class="revenue-bar-lbl">Nov</span></div>
                                <div class="revenue-bar-wrapper"><div class="revenue-bar-fill" style="height: 110px;"></div><span class="revenue-bar-lbl">Dec</span></div>
                            </div>
                        </div>

                        <!-- Sales by Traffic source chart -->
                        <div class="analytics-card">
                            <div class="analytics-card-header">
                                <h3 class="analytics-card-title">Sales by traffic source</h3>
                                <span class="action-dots">•••</span>
                            </div>
                            
                            <div class="donut-chart-box">
                                <svg class="donut-svg" viewBox="0 0 120 120">
                                    <circle cx="60" cy="60" r="45" style="fill:none; stroke:#e9ecef; stroke-width:18;"></circle>
                                    <circle class="donut-segment-1" cx="60" cy="60" r="45"></circle>
                                    <circle class="donut-segment-2" cx="60" cy="60" r="45"></circle>
                                    <circle class="donut-segment-3" cx="60" cy="60" r="45"></circle>
                                </svg>
                            </div>

                            <table class="traffic-table">
                                <thead>
                                    <tr>
                                        <th>Source</th>
                                        <th>Orders</th>
                                        <th style="text-align: right;">Amount</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td><span class="traffic-dot" style="background-color: #166534;"></span>Facebook</td>
                                        <td>22</td>
                                        <td style="text-align: right; font-weight: 600;">$2,742.00</td>
                                    </tr>
                                    <tr>
                                        <td><span class="traffic-dot" style="background-color: #d9f99d;"></span>YouTube</td>
                                        <td>27</td>
                                        <td style="text-align: right; font-weight: 600;">$3,272.00</td>
                                    </tr>
                                    <tr>
                                        <td><span class="traffic-dot" style="background-color: #f97316;"></span>Instagram</td>
                                        <td>25</td>
                                        <td style="text-align: right; font-weight: 600;">$2,922.00</td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <!-- Bottom Recent Orders inside Dashboard -->
                    <div class="table-card">
                        <div class="table-card-header">
                            <h3 class="table-title">Recent orders</h3>
                            <span class="action-dots">•••</span>
                        </div>
                        <div class="data-table-wrapper">
                            <table class="custom-data-table" id="dashboardRecentOrdersTable">
                                <thead>
                                    <tr>
                                        <th>No.</th>
                                        <th>Order Date</th>
                                        <th>Ship Date</th>
                                        <th>Customer</th>
                                        <th>Items</th>
                                        <th>Paid</th>
                                        <th>Status</th>
                                        <th>Total</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <!-- Dynamic Orders mapped from Excel -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <!-- 2. Products (Inventory) Tab View -->
                <div id="inventory-view" style="display: none; flex-direction: column; gap: 24px; width: 100%;">
                    <div class="table-card">
                        <div class="table-card-header">
                            <h3 class="table-title">Products & Inventory</h3>
                            <div class="header-actions" style="display: flex; gap: 12px;">
                                <button class="btn btn-light" onclick="alert('Reordering products...')">📦 Reorder</button>
                                <button class="btn btn-dark" onclick="exportTableToCSV('inventoryTable', 'products_inventory.csv')">↑ Export</button>
                            </div>
                        </div>
                        <div class="data-table-wrapper">
                            <table class="custom-data-table" id="inventoryTable">
                                <thead>
                                    <tr>
                                        <th>Product ID</th>
                                        <th>Product Name</th>
                                        <th>Category</th>
                                        <th>Price</th>
                                        <th>Current Stock</th>
                                        <th>Original Stock</th>
                                        <th>Units Sold</th>
                                        <th>Status</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <!-- Inventory rows -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <!-- 3. Customer Tab View -->
                <div id="customers-view" style="display: none; flex-direction: column; gap: 24px; width: 100%;">
                    <div class="table-card">
                        <div class="table-card-header">
                            <h3 class="table-title">Customer Directory</h3>
                            <div class="header-actions" style="display: flex; gap: 12px;">
                                <button class="btn btn-light" onclick="alert('Exporting customers email list...')">✉️ Email list</button>
                                <button class="btn btn-dark" onclick="exportTableToCSV('customersTable', 'customers_directory.csv')">↑ Export</button>
                            </div>
                        </div>
                        <div class="data-table-wrapper">
                            <table class="custom-data-table" id="customersTable">
                                <thead>
                                    <tr>
                                        <th>Customer Name</th>
                                        <th>Email Address</th>
                                        <th>Phone</th>
                                        <th>Shipping Address</th>
                                        <th>Total Orders</th>
                                        <th>Total Spent</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <!-- Customer rows -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <!-- 4. Analytics (Payments) Tab View -->
                <div id="payments-view" style="display: none; flex-direction: column; gap: 24px; width: 100%;">
                    <div class="table-card">
                        <div class="table-card-header">
                            <h3 class="table-title">Transaction Analytics</h3>
                            <div class="header-actions" style="display: flex; gap: 12px;">
                                <button class="btn btn-light" onclick="alert('Generating invoice statements...')">🧾 Invoices</button>
                                <button class="btn btn-dark" onclick="exportTableToCSV('paymentsTable', 'payments_analytics.csv')">↑ Export</button>
                            </div>
                        </div>
                        <div class="data-table-wrapper">
                            <table class="custom-data-table" id="paymentsTable">
                                <thead>
                                    <tr>
                                        <th>Payment ID</th>
                                        <th>Customer</th>
                                        <th>Payment Method</th>
                                        <th>Total Amount</th>
                                        <th>Status</th>
                                        <th>Transaction Date</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <!-- Payment rows -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <!-- 5. Orders Tab View -->
                <div id="orders-view" style="display: none; flex-direction: column; gap: 24px; width: 100%;">
                    <div class="table-card">
                        <div class="table-card-header">
                            <h3 class="table-title">Orders Management</h3>
                            <div class="header-actions" style="display: flex; gap: 12px;">
                                <button class="btn btn-light" onclick="alert('Importing spreadsheet logs...')">↓ Import</button>
                                <button class="btn btn-dark" onclick="exportTableToCSV('ordersTable', 'orders_report.csv')">↑ Export</button>
                            </div>
                        </div>

                        <!-- Filters -->
                        <div class="filters-section" style="margin-bottom: 20px;">
                            <div class="filter-pill active" onclick="setFilter('status', 'all', this)">All status</div>
                            <div class="filter-pill" onclick="setFilter('status', 'Complete', this)">🟢 Complete</div>
                            <div class="filter-pill" onclick="setFilter('status', 'Pending', this)">🟡 Pending</div>
                            <div class="filter-pill" onclick="setFilter('status', 'Cancel', this)">🔴 Cancel</div>
                            <div class="filter-pill" onclick="setFilter('status', 'Hold', this)">⚫ Hold</div>
                            <div class="filter-pill" onclick="setFilter('type', 'Shipping', this)">🚚 Shipping</div>
                            <div class="filter-pill" onclick="setFilter('type', 'Pickups', this)">🛍️ Pickups</div>
                        </div>

                        <div class="data-table-wrapper">
                            <table class="custom-data-table" id="ordersTable">
                                <thead>
                                    <tr>
                                        <th class="checkbox-cell">
                                            <div class="custom-checkbox" id="headerCheckbox" onclick="toggleSelectAll()"></div>
                                        </th>
                                        <th>Order ID</th>
                                        <th>Customer</th>
                                        <th>Type</th>
                                        <th>Status</th>
                                        <th>Main Product</th>
                                        <th>Total</th>
                                        <th>Date</th>
                                        <th style="width: 48px;"></th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <!-- Orders rows -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Floating action overlay when rows are checked -->
        <div class="action-overlay" id="actionOverlay">
            <div class="action-overlay-close" onclick="deselectAll()">✕</div>
            <div class="selected-count-badge" id="selectedCountText">Selected: 0</div>
            <div class="overlay-btns">
                <button class="btn-overlay" onclick="exportSelectedOrders()">↑ Export</button>
                <button class="btn-overlay" onclick="alert('Printing selected orders...')">🖨️ Print</button>
                <button class="btn-overlay" onclick="alert('Duplicating selected orders...')">📋 Duplicate</button>
            </div>
        </div>

        <!-- Detail popup card (modal) -->
        <div class="modal-overlay" id="detailModal">
            <div class="modal-card">
                <div class="modal-header">
                    <div class="modal-title-box">
                        <span>📋</span>
                        <span class="modal-title" id="modalOrderId">Order #000000</span>
                    </div>
                    <div class="modal-actions-top">
                        <span class="modal-action-icon" style="font-size: 16px;" onclick="alert('Opening in new tab...')">↗</span>
                        <span class="modal-action-icon" style="font-size: 20px; font-weight: bold; line-height: 1;" onclick="closeModal()">×</span>
                    </div>
                </div>
                <div class="modal-body">
                    <div class="modal-customer-info">
                        <div class="modal-info-row">
                            <span>👤</span>
                            <span class="bold" id="modalCustomerName">Customer Name</span>
                        </div>
                        <div class="modal-info-row">
                            <span>✉️</span>
                            <span id="modalCustomerEmail">email@example.com</span>
                        </div>
                        <div class="modal-info-row">
                            <span>📞</span>
                            <span id="modalCustomerPhone">+1 (415) 555-2671</span>
                        </div>
                        <div class="modal-info-row">
                            <span>🚚</span>
                            <span id="modalCustomerAddress" style="font-size: 12px; line-height: 1.4;">Shipping Address</span>
                        </div>
                    </div>

                    <div class="modal-tabs">
                        <div class="modal-tab active">Order items</div>
                        <div class="modal-tab" onclick="alert('Tracking shipment details...')">Delivery</div>
                        <div class="modal-tab" onclick="alert('Invoice transactions documentation...')">Docs</div>
                    </div>

                    <div class="modal-item-list" id="modalItemsList">
                        <!-- Items list -->
                    </div>

                    <div class="modal-total-section">
                        <span class="modal-total-lbl">Total:</span>
                        <span class="modal-total-val" id="modalTotalVal">$0.00</span>
                    </div>
                </div>
                <div class="modal-footer">
                    <div class="modal-footer-btns">
                        <button class="btn-modal-action" id="btnModalExport">↑ Export</button>
                        <button class="btn-modal-action" onclick="alert('Duplicating this order...')">📋 Duplicate</button>
                        <button class="btn-modal-action" onclick="alert('Printing this order...')">🖨️ Print</button>
                    </div>
                    <span class="action-dots" style="color: #64748b;">•••</span>
                </div>
            </div>
        </div>
    </div>

    <script>
        const ordersData = {orders_json};
        const inventoryData = {inventory_json};
        const customersData = {customers_json};
        
        let selectedOrders = new Set();
        let currentFilterType = 'status';
        let currentFilterVal = 'all';
        let activeTab = 'dashboard';

        // Tab Switch Router
        function switchTab(tabName, el) {{
            activeTab = tabName;
            
            // Toggle active menu selection
            document.querySelectorAll(".menu-item").forEach(item => item.classList.remove("active"));
            if (el) {{
                el.classList.add("active");
            }} else {{
                document.getElementById("menu-" + tabName).classList.add("active");
            }}

            // Show and hide correct panel content
            const tabs = ['dashboard', 'orders', 'inventory', 'payments', 'customers'];
            tabs.forEach(t => {{
                document.getElementById(t + "-view").style.display = (t === tabName) ? "flex" : "none";
            }});

            // Dynamically search update header
            const searchInput = document.querySelector(".nav-search-box input");
            searchInput.placeholder = "Search something here...";
            searchInput.value = "";

            renderActiveTabContent();
        }}

        function renderActiveTabContent() {{
            if (activeTab === 'dashboard') renderDashboardRecentOrders();
            else if (activeTab === 'orders') renderOrdersTable();
            else if (activeTab === 'inventory') renderInventoryTable();
            else if (activeTab === 'payments') renderPaymentsTable();
            else if (activeTab === 'customers') renderCustomersTable();
        }}

        // Render Recent Orders inside Dashboard Tab
        function renderDashboardRecentOrders() {{
            const tbody = document.querySelector("#dashboardRecentOrdersTable tbody");
            tbody.innerHTML = "";

            ordersData.slice(0, 5).forEach(order => {{
                const tr = document.createElement("tr");
                tr.setAttribute("onclick", `showOrderDetails("${{order.order_id}}", event)`);
                
                const paidClass = order.paid === 'Yes' ? 'yes' : 'no';
                const statusClass = order.status.toLowerCase();
                
                tr.innerHTML = `
                    <td style="font-weight: 600; color: #1a1c1e;">${{order.order_id}}</td>
                    <td>${{order.date}}</td>
                    <td>${{order.date}}</td>
                    <td style="font-weight: 500; color: #1a1c1e;">${{order.customer}}</td>
                    <td>${{order.total_units}}</td>
                    <td>
                        <span class="paid-badge ${{paidClass}}">${{order.paid}}</span>
                    </td>
                    <td>
                        <span class="badge-pill ${{statusClass}}">${{order.status}}</span>
                    </td>
                    <td style="font-weight: 600; color: #1a1c1e;">$${{order.total.toFixed(2)}}</td>
                `;
                tbody.appendChild(tr);
            }});
        }}

        // Render standard Orders Table in Orders Tab
        function renderOrdersTable() {{
            const tbody = document.querySelector("#ordersTable tbody");
            tbody.innerHTML = "";

            let filtered = ordersData;
            if (currentFilterVal !== 'all') {{
                filtered = ordersData.filter(o => {{
                    if (currentFilterType === 'status') return o.status === currentFilterVal;
                    if (currentFilterType === 'type') return o.type === currentFilterVal;
                    return true;
                }});
            }}

            filtered.forEach(order => {{
                const tr = document.createElement("tr");
                tr.setAttribute("onclick", `showOrderDetails("${{order.order_id}}", event)`);
                
                const isChecked = selectedOrders.has(order.order_id) ? "checked" : "";
                const statusClass = order.status.toLowerCase();
                
                tr.innerHTML = `
                    <td class="checkbox-cell" onclick="event.stopPropagation()">
                        <div class="custom-checkbox ${{isChecked}}" onclick="toggleSelectRow('${{order.order_id}}', this)"></div>
                    </td>
                    <td style="font-weight: 600; color: #1a1c1e;">${{order.order_id}}</td>
                    <td>
                        <div class="customer-cell">
                            <div class="cust-avatar">${{order.customer.substring(0,2).toUpperCase()}}</div>
                            <span class="cust-name">${{order.customer}}</span>
                        </div>
                    </td>
                    <td>${{order.type}}</td>
                    <td>
                        <span class="badge-pill ${{statusClass}}">${{order.status}}</span>
                    </td>
                    <td style="color: #5c6066; max-width: 180px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${{order.main_product}}</td>
                    <td style="font-weight: 600; color: #1a1c1e;">$${{order.total.toFixed(2)}}</td>
                    <td style="color: #a0a5ad;">${{order.date}}</td>
                    <td onclick="event.stopPropagation()">
                        <span class="action-dots" onclick="alert('Orders parameters settings panel ready...')">•••</span>
                    </td>
                `;
                tbody.appendChild(tr);
            }});
            
            updateHeaderCheckbox();
            updateActionOverlay();
        }}

        // Render Inventory Table from Inventory JSON
        function renderInventoryTable() {{
            const tbody = document.querySelector("#inventoryTable tbody");
            tbody.innerHTML = "";

            inventoryData.forEach(item => {{
                const tr = document.createElement("tr");
                const badgeClass = item.status === 'Healthy' ? 'complete' : 'cancel';
                const badgeLabel = item.status === 'Healthy' ? 'Healthy' : 'Low Stock';
                
                tr.innerHTML = `
                    <td style="font-weight: 600; color: #1a1c1e;">${{item.product_id}}</td>
                    <td style="font-weight: 500;">${{item.product_name}}</td>
                    <td>${{item.category}}</td>
                    <td>${{item.price}}</td>
                    <td style="font-weight: 600; color: #1a1c1e;">${{item.current_stock}}</td>
                    <td style="color: #a0a5ad;">${{item.original_stock}}</td>
                    <td style="font-weight: 600; color: #1a1c1e;">${{item.units_sold}}</td>
                    <td>
                        <span class="badge-pill ${{badgeClass}}">${{badgeLabel}}</span>
                    </td>
                    <td>
                        <span class="action-dots" onclick="alert('Stock reordering procedures initiated...')">•••</span>
                    </td>
                `;
                tbody.appendChild(tr);
            }});
        }}

        // Render Payments Table dynamically
        function renderPaymentsTable() {{
            const tbody = document.querySelector("#paymentsTable tbody");
            tbody.innerHTML = "";

            ordersData.forEach(o => {{
                const tr = document.createElement("tr");
                const payId = o.order_id.replace("#00", "PAY-");
                const method = o.type === 'Shipping' ? 'Credit Card' : 'Bank Transfer';
                const statusClass = o.status.toLowerCase();
                
                tr.innerHTML = `
                    <td style="font-weight: 600; color: #1a1c1e;">${{payId}}</td>
                    <td>
                        <div class="customer-cell">
                            <div class="cust-avatar">${{o.customer.substring(0,2).toUpperCase()}}</div>
                            <span class="cust-name">${{o.customer}}</span>
                        </div>
                    </td>
                    <td>${{method}}</td>
                    <td style="font-weight: 600; color: #1a1c1e;">$${{o.total.toFixed(2)}}</td>
                    <td>
                        <span class="badge-pill ${{statusClass}}">${{o.status}}</span>
                    </td>
                    <td style="color: #a0a5ad;">${{o.date}}</td>
                    <td>
                        <span class="action-dots" onclick="alert('Transactions detailed statement ready')">•••</span>
                    </td>
                `;
                tbody.appendChild(tr);
            }});
        }}

        // Render Customers Directory
        function renderCustomersTable() {{
            const tbody = document.querySelector("#customersTable tbody");
            tbody.innerHTML = "";

            customersData.forEach(c => {{
                const tr = document.createElement("tr");
                
                tr.innerHTML = `
                    <td>
                        <div class="customer-cell">
                            <div class="cust-avatar">${{c.name.substring(0,2).toUpperCase()}}</div>
                            <span class="cust-name">${{c.name}}</span>
                        </div>
                    </td>
                    <td>${{c.email}}</td>
                    <td>${{c.phone}}</td>
                    <td style="font-size: 12px; line-height: 1.4; max-width: 240px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${{c.address}}</td>
                    <td style="font-weight: 600; color: #1a1c1e; text-align: center;">${{c.orders_count}}</td>
                    <td style="font-weight: 600; color: #1a1c1e;">$${{c.total_spent.toFixed(2)}}</td>
                    <td>
                        <span class="action-dots" onclick="alert('Contacting customer directory profile...')">•••</span>
                    </td>
                `;
                tbody.appendChild(tr);
            }});
        }}

        // Checkbox management
        function toggleSelectRow(orderId, el) {{
            if (selectedOrders.has(orderId)) {{
                selectedOrders.delete(orderId);
                el.classList.remove("checked");
            }} else {{
                selectedOrders.add(orderId);
                el.classList.add("checked");
            }}
            updateHeaderCheckbox();
            updateActionOverlay();
        }}

        function toggleSelectAll() {{
            const el = document.getElementById("headerCheckbox");
            const isAllChecked = el.classList.contains("checked");
            
            let filtered = ordersData;
            if (currentFilterVal !== 'all') {{
                filtered = ordersData.filter(o => {{
                    if (currentFilterType === 'status') return o.status === currentFilterVal;
                    if (currentFilterType === 'type') return o.type === currentFilterVal;
                    return true;
                }});
            }}

            if (isAllChecked) {{
                filtered.forEach(o => selectedOrders.delete(o.order_id));
                el.classList.remove("checked");
            }} else {{
                filtered.forEach(o => selectedOrders.add(o.order_id));
                el.classList.add("checked");
            }}
            
            renderOrdersTable();
        }}

        function updateHeaderCheckbox() {{
            const el = document.getElementById("headerCheckbox");
            
            let filtered = ordersData;
            if (currentFilterVal !== 'all') {{
                filtered = ordersData.filter(o => {{
                    if (currentFilterType === 'status') return o.status === currentFilterVal;
                    if (currentFilterType === 'type') return o.type === currentFilterVal;
                    return true;
                }});
            }}

            if (filtered.length === 0) {{
                el.classList.remove("checked");
                return;
            }}

            const allSelected = filtered.every(o => selectedOrders.has(o.order_id));
            if (allSelected) {{
                el.classList.add("checked");
            }} else {{
                el.classList.remove("checked");
            }}
        }}

        function deselectAll() {{
            selectedOrders.clear();
            renderOrdersTable();
        }}

        function updateActionOverlay() {{
            const overlay = document.getElementById("actionOverlay");
            const badge = document.getElementById("selectedCountText");
            
            if (selectedOrders.size > 0 && activeTab === 'orders') {{
                badge.innerText = `Selected: ${{selectedOrders.size}}`;
                overlay.classList.add("active");
            }} else {{
                overlay.classList.remove("active");
            }}
        }}

        // Popover Details Modal
        function showOrderDetails(orderId, event) {{
            const order = ordersData.find(o => o.order_id === orderId);
            if (!order) return;

            document.getElementById("modalOrderId").innerText = `Order ${{order.order_id}}`;
            document.getElementById("modalCustomerName").innerText = order.customer;
            document.getElementById("modalCustomerEmail").innerText = order.email;
            document.getElementById("modalCustomerPhone").innerText = order.phone;
            document.getElementById("modalCustomerAddress").innerText = order.address;
            document.getElementById("modalTotalVal").innerText = `$${{order.total.toFixed(2)}}`;

            // Set up dynamic functional export button for this single order
            document.getElementById("btnModalExport").setAttribute("onclick", `exportSingleOrder("${{order.order_id}}")`);

            const list = document.getElementById("modalItemsList");
            list.innerHTML = "";
            
            order.products.forEach(p => {{
                const row = document.createElement("div");
                row.className = "modal-item-row";
                row.innerHTML = `
                    <div class="modal-item-detail">
                        <div class="modal-item-thumb">📦</div>
                        <div class="modal-item-text">
                            <span class="modal-item-name">${{p.name}}</span>
                            <span class="modal-item-meta">${{p.units}} × $${{p.price.toFixed(2)}} (${{p.size}}, ${{p.color}})</span>
                        </div>
                    </div>
                    <span class="modal-item-price">$${{p.total.toFixed(2)}}</span>
                `;
                list.appendChild(row);
            }});

            document.getElementById("detailModal").classList.add("active");
        }}

        function closeModal() {{
            document.getElementById("detailModal").classList.remove("active");
        }}

        // Order list filtering
        function setFilter(type, value, pillEl) {{
            currentFilterType = type;
            currentFilterVal = value;

            document.querySelectorAll(".filter-pill").forEach(p => p.classList.remove("active"));
            pillEl.classList.add("active");

            renderOrdersTable();
        }}

        // Search filtering across all active list views
        function filterTableSearch() {{
            const query = document.getElementById("searchInput").value.toLowerCase();
            let tbodyId = "";
            if (activeTab === 'dashboard') tbodyId = "#dashboardRecentOrdersTable tbody";
            else if (activeTab === 'orders') tbodyId = "#ordersTable tbody";
            else if (activeTab === 'inventory') tbodyId = "#inventoryTable tbody";
            else if (activeTab === 'payments') tbodyId = "#paymentsTable tbody";
            else if (activeTab === 'customers') tbodyId = "#customersTable tbody";
            
            if (!tbodyId) return;

            const tbody = document.querySelector(tbodyId);
            const rows = tbody.querySelectorAll("tr");

            rows.forEach(row => {{
                const text = row.innerText.toLowerCase();
                if (text.includes(query)) {{
                    row.style.display = "";
                }} else {{
                    row.style.display = "none";
                }}
            }});
        }}

        // Real functional Excel/CSV Export Engines
        function exportTableToCSV(tableId, filename) {{
            const table = document.getElementById(tableId);
            if (!table) return;

            let csv = [];
            const rows = table.querySelectorAll("tr");
            
            for (let i = 0; i < rows.length; i++) {{
                // Skip rows that are hidden by search or filters
                if (rows[i].style.display === "none") continue;

                let row = [];
                const cols = rows[i].querySelectorAll("td, th");
                
                for (let j = 0; j < cols.length; j++) {{
                    // Skip first checkbox cell and last dot menu cell
                    if (cols[j].classList.contains("checkbox-cell") || j === cols.length - 1) continue;
                    
                    let text = cols[j].innerText.trim();
                    text = text.replace(/(\\r\\n|\\n|\\r)/gm, " ");
                    text = text.replace(/"/g, '""'); // Escape double quotes
                    row.push('"' + text + '"');
                }}
                if (row.length > 0) {{
                    csv.push(row.join(","));
                }}
            }}

            // Trigger real browser CSV download
            const csvContent = "data:text/csv;charset=utf-8,\\uFEFF" + csv.join("\\n");
            const encodedUri = encodeURI(csvContent);
            const link = document.createElement("a");
            link.setAttribute("href", encodedUri);
            link.setAttribute("download", filename);
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }}

        function exportSelectedOrders() {{
            if (selectedOrders.size === 0) return;
            
            let csv = [["Order ID", "Customer", "Type", "Status", "Product", "Total", "Date"]];
            
            ordersData.forEach(order => {{
                if (selectedOrders.has(order.order_id)) {{
                    csv.push([
                        order.order_id,
                        order.customer,
                        order.type,
                        order.status,
                        order.main_product,
                        "$" + order.total.toFixed(2),
                        order.date
                    ].map(val => '"' + val.replace(/"/g, '""') + '"'));
                }}
            }});

            const csvContent = "data:text/csv;charset=utf-8,\\uFEFF" + csv.map(e => e.join(",")).join("\\n");
            const encodedUri = encodeURI(csvContent);
            const link = document.createElement("a");
            link.setAttribute("href", encodedUri);
            link.setAttribute("download", "selected_orders_report.csv");
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }}

        function exportSingleOrder(orderId) {{
            const order = ordersData.find(o => o.order_id === orderId);
            if (!order) return;
            
            let csv = [];
            csv.push([`"Order Details for ${{order.order_id}}"`]);
            csv.push([`"Customer Name"`, `"${{order.customer}}"`]);
            csv.push([`"Email"`, `"${{order.email}}"`]);
            csv.push([`"Phone"`, `"${{order.phone}}"`]);
            csv.push([`"Shipping Address"`, `"${{order.address}}"`]);
            csv.push([]);
            csv.push([`"Product Name"`, `"Units"`, `"Price"`, `"Total"`, `"Size"`, `"Color"`]);
            
            order.products.forEach(p => {{
                csv.push([
                    `"${{p.name}}"`,
                    p.units,
                    `$${{p.price.toFixed(2)}}`,
                    `$${{p.total.toFixed(2)}}`,
                    `"${{p.size}}"`,
                    `"${{p.color}}"`
                ]);
            }});
            
            csv.push([]);
            csv.push([`"Grand Total"`, `""`, `""`, `"$${{order.total.toFixed(2)}}"`]);
            
            const csvContent = "data:text/csv;charset=utf-8,\\uFEFF" + csv.map(e => e.join(",")).join("\\n");
            const encodedUri = encodeURI(csvContent);
            const link = document.createElement("a");
            link.setAttribute("href", encodedUri);
            link.setAttribute("download", `order_${{order.order_id}}_report.csv`);
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }}

        // Setup onload default tab
        window.onload = () => {{
            switchTab('dashboard');
        }};
    </script>
</body>
</html>
"""

# Render dynamic layout in Streamlit
st.components.v1.html(html_code, height=940, scrolling=True)
