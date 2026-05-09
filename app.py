import streamlit as st
import pandas as pd
import os
import zipfile
import base64
import re
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib import colors
from playwright.sync_api import sync_playwright
import subprocess

subprocess.run(["playwright", "install", "chromium"])
st.set_page_config(page_title="Go-Green Payroll", layout="wide")

# --------------------------------------------------
# UI DESIGN
# --------------------------------------------------
st.markdown(""" 
<style>

/* GLOBAL BACKGROUND */
[data-testid="stAppViewContainer"]{
    background: #f4f6f9;
}

/* SIDEBAR */
[data-testid="stSidebar"]{
    background: #111827;
}

[data-testid="stSidebar"] * {
    color: white !important;
}

/* HEADINGS */
h1,h2,h3 {
    color:#111827;
    font-weight:600;
}

/* CARD UI */
.card {
    background: white;
    padding: 20px;
    border-radius: 14px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    margin-bottom: 20px;
}

/* BUTTONS */
div.stButton > button {
    background: linear-gradient(135deg,#2563eb,#3b82f6);
    color:white;
    border-radius:8px;
    padding:10px 20px;
    border:none;
    font-weight:600;
    transition:0.3s;
}

div.stButton > button:hover {
    background: linear-gradient(135deg,#1d4ed8,#2563eb);
}

/* INPUT */
input, .stTextInput input, .stNumberInput input {
    border-radius:8px !important;
    border:1px solid #e5e7eb !important;
    padding:10px !important;
}

/* METRIC CARDS */
.metric {
    background: white;
    padding:15px;
    border-radius:12px;
    text-align:center;
    box-shadow:0 2px 6px rgba(0,0,0,0.05);
}

</style>
""", unsafe_allow_html=True)
# --------------------------------------------------
# TITLE
# --------------------------------------------------
st.title("Go-Green Payslip Generator")

# --------------------------------------------------
# EXCEL DATABASE
# --------------------------------------------------
         
EMP_FILE = "employees.xlsx"


def load_employees():
    if os.path.exists(EMP_FILE):
        return pd.read_excel(EMP_FILE)

    return pd.DataFrame(columns=[
        "emp_id","name","category","basic","da","uan","esic","wc_no","doj","bank","account"
    ])


def save_employee(data):

    df = load_employees()

    df = pd.concat([df,pd.DataFrame([data])],ignore_index=True)

    df.to_excel(EMP_FILE,index=False)


def delete_employee(emp_id):

    df = load_employees()

    df = df[df["emp_id"].astype(str)!=str(emp_id)]

    df.to_excel(EMP_FILE,index=False)


# --------------------------------------------------
# COMPANY HEADER
# --------------------------------------------------

# -------- Convert image to base64 --------

def get_base64_image(image_path):
    with open(image_path, "rb") as img:
        return base64.b64encode(img.read()).decode()

if os.path.exists("logo.png"):
    logo_base64 = get_base64_image("logo.png")
else:
    logo_base64 = ""

# -------- HEADER LAYOUT --------

col1, col2,col3 = st.columns([1,3,1])

with col2:

    st.markdown(f"""
    <div style="text-align:center;padding:20px;background:white;border-radius:10px;
    box-shadow:0px 4px 15px rgba(0,0,0,0.08);">

    <img src="data:image/png;base64,{logo_base64}" width="120">

    <h1>GO-GREEN</h1>

    <p>
    No:820/9-A-1, Ramaiah Complex<br>
    Perandapalli, Hosur<br>
    Tamil Nadu - 635109
    </p>

    </div>
    """, unsafe_allow_html=True)

with col3:

    selected_date = st.date_input(
        "Payslip Month",
        value=datetime.today()
    )

    pay_month = selected_date.strftime("%B %Y")

    st.markdown(f"""
    <div style="font-size:16px;color:#1b5e20;">
    <b>Selected Payslip Month:</b> {pay_month}
    </div>
    """, unsafe_allow_html=True)
# --------------------------------------------------
# --------------------------------------------------
# EMPLOYEE MANAGEMENT (PROFESSIONAL UI)
# --------------------------------------------------

st.markdown("## 👨‍💼 Employee Management")

# Session state
if "emp_page" not in st.session_state:
    st.session_state.emp_page = ""

# ---- BUTTON UI ----
col1, col2, col3, col4 = st.columns(4)

with col1:
    if st.button("➕ Add"):
        st.session_state.emp_page = "add"

with col2:
    if st.button("✏️ Edit"):
        st.session_state.emp_page = "edit"

with col3:
    if st.button("🗑️ Delete"):
        st.session_state.emp_page = "delete"

with col4:
    if st.button("🔍 Search"):
        st.session_state.emp_page = "search"

page = st.session_state.emp_page

st.divider()

# ---------------- ADD ----------------
import re

if page == "add":

    st.markdown("### ➕ Add Employee")

    category = st.selectbox(
        "Employee Category",
        ["Unskilled", "Semi-Skilled", "Skilled", "Manager"]
    )

    category_salary = {
        "Unskilled": {"basic": 5579, "da": 6805.2},
        "Semi-Skilled": {"basic": 6579, "da": 8052.0},
        "Skilled": {"basic": 8079, "da": 9871.0},
        "Manager": {"basic": 10079, "da": 12312.4}
    }

    basic_salary = category_salary[category]["basic"]
    da_salary = category_salary[category]["da"]

    # 👉 MOVE EVERYTHING INSIDE FORM
    with st.form("employee_form", clear_on_submit=True):

        col1, col2 = st.columns(2)
        col1.text_input("Basic Salary", value=basic_salary, disabled=True)
        col2.text_input("DA", value=f"{da_salary:.2f}", disabled=True)

        st.divider()

        c1, c2 = st.columns(2)

        emp_id_master = c1.text_input("Employee ID")
        emp_name_master = c2.text_input("Employee Name")

        uan = c1.text_input("UAN Number")

        if category == "Manager":
            wc_no = c2.text_input("WC Number")
            esic = ""
        else:
            esic = c2.text_input("ESIC Number")
            wc_no = ""

        doj = c1.date_input("Date of Joining")
        bank_name = c2.text_input("Bank Name")

        bank_account = st.text_input("Bank Account Number")

        submitted = st.form_submit_button("Add Employee")

        # ---------------- VALIDATION ----------------
        if submitted:

            errors = []

            if not emp_id_master.strip():
                errors.append("Employee ID is required")

            if not emp_name_master.strip():
                errors.append("Employee Name is required")

            if not uan.strip():
                errors.append("UAN Number is required")

            if category != "Manager" and not esic.strip():
                errors.append("ESIC Number is required")

            if category == "Manager" and not wc_no.strip():
                errors.append("WC Number is required")

            if not bank_name.strip():
                errors.append("Bank Name is required")

            if not bank_account.strip():
                errors.append("Bank Account Number is required")

            # Name validation
            if emp_name_master and not re.fullmatch(r"[A-Za-z ]+", emp_name_master):
                errors.append("Employee Name should contain only letters")

            # Numeric validation
            if uan and not uan.isdigit():
                errors.append("UAN must be numeric")

            if category != "Manager" and esic and not esic.isdigit():
                errors.append("ESIC must be numeric")

            if bank_account and not bank_account.isdigit():
                errors.append("Bank Account must be numeric")

            # ---------------- RESULT ----------------
            if errors:
                st.error("⚠️ Please fix the following errors:")
                for err in errors:
                    st.write(f"• {err}")
            else:
                employee = {
                    "emp_id": emp_id_master.strip(),
                    "name": emp_name_master.strip(),
                    "category": category,
                    "basic": basic_salary,
                    "da": float(da_salary),
                    "uan": int(uan),
                    "esic": int(esic) if esic else "",
                    "wc_no": wc_no.strip(),
                    "doj": str(doj),
                    "bank": bank_name.strip(),
                    "account": int(bank_account)
                }

                save_employee(employee)
                st.success("✅ Employee Saved Successfully")
    
    
                

# ---------------- EDIT ----------------
elif page == "edit":

    st.markdown("### ✏️ Edit Employee")

    edit_id = st.text_input("Enter Employee ID")

    if edit_id:

        df = load_employees()
        emp = df[df["emp_id"].astype(str) == edit_id]

        if not emp.empty:
            emp = emp.iloc[0]

            name = st.text_input("Name", emp["name"])

            category = st.selectbox(
                "Category",
                ["Unskilled","Semi-Skilled","Skilled","Manager"],
                index=["Unskilled","Semi-Skilled","Skilled","Manager"].index(emp["category"])
            )

            uan = st.text_input("UAN", emp["uan"])

            if category == "Manager":
                wc_no = st.text_input("WC No", emp["wc_no"])
                esic = ""
            else:
                esic = st.text_input("ESIC", emp["esic"])
                wc_no = ""

            if st.button("Update Employee"):

                category_salary = {
                    "Unskilled": {"basic": 5579, "da": 6805.2},
                    "Semi-Skilled": {"basic": 6579, "da": 8052.0},
                    "Skilled": {"basic": 8079, "da": 9871.0},
                    "Manager": {"basic": 10079, "da": 12312.4}
                }

                new_basic = category_salary[category]["basic"]
                new_da = category_salary[category]["da"]

                df.loc[df["emp_id"].astype(str) == edit_id, "name"] = name
                df.loc[df["emp_id"].astype(str) == edit_id, "category"] = category
                df.loc[df["emp_id"].astype(str) == edit_id, "basic"] = new_basic
                df.loc[df["emp_id"].astype(str) == edit_id, "da"] = new_da
                df.loc[df["emp_id"].astype(str) == edit_id, "uan"] = uan
                df.loc[df["emp_id"].astype(str) == edit_id, "esic"] = esic
                df.loc[df["emp_id"].astype(str) == edit_id, "wc_no"] = wc_no

                df.to_excel(EMP_FILE, index=False)

                st.success("✅ Updated Successfully")

               
        else:
            st.error("❌ Employee Not Found")


# ---------------- DELETE ----------------
elif page == "delete":

    st.markdown("### 🗑️ Delete Employee")

    delete_id = st.text_input("Employee ID")

    if st.button("Delete Employee"):

        df = load_employees()

        if delete_id in df["emp_id"].astype(str).values:

            df = df[df["emp_id"].astype(str) != delete_id]
            df.to_excel(EMP_FILE, index=False)

            st.success("✅ Deleted Successfully")

        else:
            st.error("❌ Employee Not Found")


# ---------------- SEARCH ----------------
elif page == "search":

    st.markdown("### 🔍 Search Employee")

    search_id = st.text_input("Employee ID")

    if search_id:

        df = load_employees()
        emp = df[df["emp_id"].astype(str) == search_id]

        if not emp.empty:

            emp = emp.iloc[0]

            st.success("Employee Found")

            st.write("**Name:**", emp["name"])
            st.write("**Category:**", emp["category"])
            st.write("**Basic:** ₹", emp["basic"])
            st.write("**DA:** ₹", emp["da"])

        else:
            st.error("❌ Employee Not Found")
# --------------------------------------------------
# --------------------------------------------------
# EMPLOYEE PAY SUMMARY
# --------------------------------------------------

st.subheader("Employee Pay Summary")

emp_id_input = st.text_input("Enter Employee ID", key="pay_emp_id")

employee_name = ""
employee_id = ""
employee_category = ""
basic_salary = 0
da_salary = 0
uan = ""
esic = ""
wc_no = ""
bank_name = ""
bank_account = ""
doj = ""

if emp_id_input:

    df = load_employees()

    emp = df[df["emp_id"].astype(str) == emp_id_input]

    if not emp.empty:

        emp = emp.iloc[0]

        employee_name = emp["name"]
        employee_id = emp["emp_id"]
        employee_category = emp["category"]
        basic_salary = emp["basic"]
        da_salary = emp["da"]
        uan = emp["uan"]
        esic = emp["esic"]
        wc_no = emp["wc_no"]
        bank_name = emp["bank"]
        bank_account = emp["account"]
        doj = emp["doj"]

        st.success("Employee Found")

        st.write(f"Name : {employee_name}")
        st.write(f"Category : {employee_category}")
        st.write(f"Basic : ₹ {basic_salary}")
        st.write(f"DA : ₹ {da_salary:.2f}")

    else:
        st.error("Employee Not Found")
# --------------------------------------------------
# PAY DETAILS (FIX - REQUIRED)
# --------------------------------------------------

col1, col2 = st.columns(2)

with col1:
    pay_period_date = st.date_input(
    "Pay Period",
    value=datetime.today(),
    key="pay_period_date"
)

    pay_period = pay_period_date.strftime("%B %Y")

with col2:
    total_days = st.number_input(
        "Total Working Days (Month)",
        min_value=1,
        max_value=31,
        value=26
    )

    paid_days = st.number_input(
        "Paid Days",
        min_value=0.0,
        max_value=40.0,
        value=float(total_days),
        step=0.5
    )

    pay_date = st.date_input("Pay Date", key="pay_date")
# --------------------------------------------------
# INCOME DETAILS (PROFESSIONAL STRUCTURE)
# --------------------------------------------------

st.subheader("Income Details")

def r2(val):
    return round(val, 2)   

def r0(val):
    return int(round(val)) 

# -------- PRO-RATA CALCULATION --------
# -------- LOP CALCULATION FIX --------
working_days =paid_days

# safety (avoid negative)
if working_days < 0:
    working_days = 0

basic_current = r2((basic_salary / total_days) * working_days)
original_da = r2((da_salary / total_days) * working_days)


# Only initialize when employee exists
if employee_id != "":
    
    if ("da_current" not in st.session_state) or (st.session_state.get("last_emp_da") != employee_id):
        st.session_state.da_current = original_da
        st.session_state.last_emp_da = employee_id

    da_current = st.session_state.da_current
else:
    da_current = 0

# -------- DISPLAY EARNINGS --------
st.markdown("### Earnings")
default_nh_h = r2((basic_current + da_current) * 0.0288)
# ---------- NH&H EDITABLE ----------
if "nh_h_value" not in st.session_state:
    st.session_state.nh_h_value = default_nh_h

# Reset when employee changes
if st.session_state.get("last_emp_nhh") != employee_id:
    st.session_state.nh_h_value = default_nh_h
    st.session_state.last_emp_nhh = employee_id

col1, col2 = st.columns([3,1])
col1.write("Basic")
col2.write(f"₹ {basic_current:.2f}")

col1, col2 = st.columns([3,1])

col1.write("Dearness Allowance")

if employee_id != "":
    da_current = col2.number_input(
        "DA",
        min_value=0.0,
        value=float(st.session_state.da_current),
        step=10.0,
        key="da_input",
        label_visibility="collapsed"
    )

    st.session_state.da_current = da_current
else:
    col2.write("₹ 0.00")
# Recalculate after editable DA change
earned_leave = r2((basic_current + da_current) * 0.0481)

# BONUS 8.33% using updated DA
bonus = r2((basic_current + da_current) * 0.0833)

# NH&H default using updated DA

col1, col2 = st.columns([3,1])
col1.write("Earned Leave (4.81%)")
col2.write(f"₹ {earned_leave:.2f}")

# BONUS
col1, col2 = st.columns([3,1])
col1.write("Bonus (8.33%)")
col2.write(f"₹ {bonus:.2f}")

# NH&H Editable
col1, col2 = st.columns([3,1])

col1.write("NH & H")

nh_h = col2.number_input(
    "NH&H",
    value=float(st.session_state.nh_h_value),
    step=0.01,
    key="nh_h_input",
    label_visibility="collapsed"
)

st.session_state.nh_h_value = nh_h

extra_earnings = []

# ➕ Add Extra Earnings
if "extra_earnings" not in st.session_state:
    st.session_state.extra_earnings = []

if st.button("+ Add Extra Earnings"):
    st.session_state.extra_earnings.append({"name":"Other Allowance","amount":0})

for i, item in enumerate(st.session_state.extra_earnings):
    c1,c2 = st.columns([2,1])
    item["name"] = c1.text_input("Name", item["name"], key=f"extra_e_name{i}")
    item["amount"] = c2.number_input("Amount", value=item["amount"], key=f"extra_e_amt{i}")

# --------------------------------------------------
# DEDUCTIONS
# --------------------------------------------------
st.markdown("### Deductions")

# ---------- PF CALCULATION ----------
default_pf = r2((basic_current + da_current) * 0.12)

if "pf_value" not in st.session_state:
    st.session_state.pf_value = default_pf

# Reset PF when employee changes
if st.session_state.get("last_emp") != employee_id:
    st.session_state.pf_value = default_pf
    st.session_state.last_emp = employee_id

col1, col2 = st.columns([3,1])
col1.write("PF")

# Default PF
default_pf = r2((basic_current + da_current) * 0.12)

# Reset when employee changes
if st.session_state.get("last_emp") != employee_id:
    st.session_state.pf_value = default_pf
    st.session_state.last_emp = employee_id

if employee_name != "":

    if employee_category in ["Skilled", "Manager"]:
        pf = col2.number_input(
            "PF",
            value=float(st.session_state.pf_value),
            step=10.0,
            key="pf_input",
            label_visibility="collapsed"
        )
        st.session_state.pf_value = pf
    else:
        pf = default_pf
        col2.write(f"₹ {pf:.2f}")

    # ✅ deductions calculation
    if employee_category != "Manager":
        esic_calc = r2((basic_current + da_current) * 0.0075)
    else:
        esic_calc = 0

    if employee_category == "Unskilled":
        pt = 0
    elif employee_category in ["Skilled", "Semi-Skilled"]:
        pt = 150
    elif employee_category == "Manager":
        pt = 200
    else:
        pt = 0

    canteen = r2(paid_days * 10)

else:
    pf = 0
    esic_calc = 0
    pt = 0
    canteen = 0
deductions_data = [
    ("PF(0.12%)", pf),
    ("ESIC (0.75%)", esic_calc),
]

# Add PT only if applicable
if employee_category != "Unskilled":
    deductions_data.append(("Professional Tax", pt))

deductions_data.append(("Canteen (₹10/day)", canteen))

for name, amount in deductions_data:
    col1, col2 = st.columns([3,1])
    col1.write(name)
    col2.write(f"₹ {amount:.2f}")

# ➕ Add Extra Deductions
if "extra_deductions" not in st.session_state:
    st.session_state.extra_deductions = []

if st.button("+ Add Extra Deductions"):
    st.session_state.extra_deductions.append({"name":"Other Deduction","amount":0})

for i, item in enumerate(st.session_state.extra_deductions):
    c1,c2 = st.columns([2,1])
    item["name"] = c1.text_input("Name", item["name"], key=f"extra_d_name{i}")
    item["amount"] = c2.number_input("Amount", value=item["amount"], key=f"extra_d_amt{i}")

# --------------------------------------------------
# FINAL CALCULATION
# --------------------------------------------------

gross = r2(
    basic_current +
    da_current +
    earned_leave +
    bonus +
    nh_h +
    sum(e["amount"] for e in st.session_state.extra_earnings)
)

total_deduction =r2(
    pf +
    esic_calc +
    pt +
    canteen +
    sum(d["amount"] for d in st.session_state.extra_deductions)
)

net_salary = r2(gross - total_deduction)

st.divider()

st.write(f"### Gross Earnings : ₹ {gross:.2f}")
st.write(f"### Total Deductions : ₹ {total_deduction:.2f}")
st.success(f"Net Salary : ₹ {round(net_salary)}") 


from reportlab.platypus import Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from num2words import num2words

def amount_in_words(amount):
    words = num2words(amount, lang='en_IN').title()
    return words + " Rupees Only"
def r(val):
    return f"{val:.2f}"   

def format_currency(val):
    return f"Rs. {val:,.2f}"   
from datetime import datetime

def create_pdf(data):
    
    dt = datetime.strptime(data["pay_period"], "%B %Y")
    formatted_date = dt.strftime("%b_%Y")

    safe_name = re.sub(r'[^A-Za-z0-9 ]', '', data["employee_name"])  # remove special chars
    safe_name = safe_name.replace(" ", "_")  # replace spaces

    file_name = f"{safe_name}_{data['employee_id']}_{formatted_date}.pdf"
    deduction_rows = ""
    # ---------------- EXTRA EARNINGS ----------------
    extra_earnings_rows = ""
    for item in data.get("extra_earnings", []):
        extra_earnings_rows += f"""
        <tr>
            <td>{item["name"]}</td>
            <td>₹ {r(item["amount"])}</td>
            <td></td>
            <td></td>
        </tr>
        """

    # ---------------- EXTRA DEDUCTIONS ----------------
    extra_deduction_rows = ""
    for item in data.get("extra_deductions", []):
        extra_deduction_rows += f"""
        <tr>
            <td></td>
            <td></td>
            <td>{item["name"]}</td>
            <td>₹ {r(item["amount"])}</td>
        </tr>
        """

    deduction_rows = ""

    # Row 1
    deduction_rows += f"""
    <tr>
        <td>Basic</td>
        <td>₹ {r(data["basic_current"])}</td>
        <td>PF</td>
        <td>₹ {r(data["pf"])}</td>
    </tr>
    """

    # -------- DYNAMIC ROW BUILDING --------

    deduction_items = []

    # ESIC (only if not Manager)
    if data["employee_category"] != "Manager":
        deduction_items.append(("ESIC", data["esic_calc"]))

    # PT (only if not Unskilled)
    if data["employee_category"] != "Unskilled":
        deduction_items.append(("PT", data["pt"]))

    # Canteen (always)
    deduction_items.append(("Canteen", data["canteen"]))
    
    # -------- ROW 2 --------
    if len(deduction_items) > 0:
        name, val = deduction_items.pop(0)
        deduction_rows += f"""
        <tr>
            <td>DA</td>
            <td>₹ {r(data["da_current"])}</td>
            <td>{name}</td>
            <td>₹ {r(val)}</td>
        </tr>
        """

    # -------- ROW 3 --------
    if len(deduction_items) > 0:
        name, val = deduction_items.pop(0)
        deduction_rows += f"""
        <tr>
            <td>Earned Leave</td>
            <td>₹ {r(data["earned_leave"])}</td>
            <td>{name}</td>
            <td>₹ {r(val)}</td>
        </tr>
        """
   # -------- BONUS ROW --------
    if len(deduction_items) > 0:

        name, val = deduction_items.pop(0)

        deduction_rows += f"""
        <tr>
            <td>Bonus</td>
            <td>₹ {r(data["bonus"])}</td>
            <td>{name}</td>
            <td>₹ {r(val)}</td>
        </tr>
        """
    else:

        deduction_rows += f"""
        <tr>
            <td>Bonus</td>
            <td>₹ {r(data["bonus"])}</td>
            <td></td>
            <td></td>
        </tr>
        """
    # -------- NH&H ROW --------
    if data["nh_h"] > 0:

        if len(deduction_items) > 0:
            name, val = deduction_items.pop(0)

            deduction_rows += f"""
            <tr>
                <td>NH & H</td>
                <td>₹ {r(data["nh_h"])}</td>
                <td>{name}</td>
                <td>₹ {r(val)}</td>
            </tr>
            """
        else:
            deduction_rows += f"""
            <tr>
                <td>NH & H</td>
                <td>₹ {r(data["nh_h"])}</td>
                <td></td>
                <td></td>
            </tr>
            """


    html = f"""
    <html>
    <head>
    <meta charset="UTF-8">

    <style>
    body {{
        font-family: 'Segoe UI', sans-serif;
        background: #f4f6f9;
        padding: 40px;
    }}

    .container {{
        background: white;
        padding: 30px;
        border-radius: 12px;
        width: 800px;
        margin: auto;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
    }}

    .header {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 1px solid #ddd;
        padding-bottom: 15px;
    }}

    .company {{
        display: flex;
        align-items: center;
        gap: 15px;
    }}

    .company img {{
        width: 60px;
    }}

    .company-details {{
        font-size: 13px;
        color: #555;
    }}

    .title {{
        text-align: right;
    }}

    .net-box {{
        width: 25%;
    }}

    .right-details p {{
        margin: 4px 0;
    }}

    .summary-left p {{
        margin: 6px 0;
    }}

    .net-box {{
        background: #e8f5e9;
        border-left: 5px solid #2e7d32;
        padding: 20px;
        border-radius: 10px;
        width: 220px;
        text-align: center;
    }}

    .net-box h1 {{
        margin: 0;
        color: #1b5e20;
    }}
    
    .top-table {{
    width: 100%;
    margin-top: 5px;
    }}

    .left-cell {{
        width: 60%;
        vertical-align: top;
        font-size: 14px;
    }}

    .right-cell {{
        width: 40%;
        text-align: right;
        vertical-align: top;
        font-size: 14px;
    }}

    .top-table p {{
        margin: 4px 0;
    }}

    .net-container {{
        display: flex;
        justify-content: center;
        margin-top: 20px;
    }}
   
    table {{
        width: 100%;
        border-collapse: collapse;
        margin-top: 25px;
    }}

    th {{
        text-align: left;
        border-bottom: 2px solid #ddd;
        padding: 10px;
        font-size: 13px;
        color: #555;
    }}

    td {{
        padding: 10px;
        border-bottom: 1px solid #eee;
        font-size: 14px;
    }}

    .total-row {{
        font-weight: bold;
        background: #f9f9f9;
    }}

    .final {{
        margin-top: 20px;
        padding: 15px;
        background: #e8f5e9;
        border-radius: 8px;
        display: flex;
        justify-content: space-between;
        font-weight: bold;
    }}

    .footer {{
        margin-top: 30px;
        text-align: center;
        font-size: 12px;
        color: gray;
    }}
    </style>
    </head>

    <body>

    <div class="container">

        <div class="header">

            <div class="company">
                <img src="data:image/png;base64,{logo_base64}" />
                <div>
                    <b style="font-size:18px;">GO-GREEN</b><br>
                    <div class="company-details">
                    No:820/9-A-1, Ramaiah Complex<br>
                    Perandapalli, Hosur<br>
                    Tamil Nadu - 635109
                    </div>
                </div>
            </div>

            <div class="title">
                <b>Payslip For the Month</b><br>
                {data["pay_period"]}
            </div>

        </div>

       <table class="top-table">
        <tr>
            <!-- LEFT -->
            <td class="left-cell">
                <p><b>Employee Name</b> : {data["employee_name"]}</p>
                <p><b>Employee ID</b> : {data["employee_id"]}</p>
                <p><b>Pay Period</b> : {data["pay_period"]}</p>
                <p><b>Pay Date</b> : {data.get("pay_date","")}</p>
                <p><b>Paid Days</b> : {data["paid_days"]}</p>
            </td>

            <!-- RIGHT -->
            <td class="right-cell">
                <p><b>Bank Name</b> : {data["bank_name"]}</p>
                <p><b>Account No</b> : {data["bank_account"]}</p>
                <p><b>UAN</b> : {data["uan"]}</p>
                {f'<p><b>WC No</b> : {data["wc_no"]}</p>'
                if data["employee_category"] == "Manager"
                else f'<p><b>ESIC</b> : {int(data["esic"]) if data["esic"] else "-"} </p>'
                }
            </td>
        </tr>
    </table>

            <!-- NET PAY CENTER -->
            <div class="net-container">
                <div class="net-box">
                    <div style="font-size:13px;">Total Net Pay</div>
                    <h1>₹ {r0(data["net_salary"])}</h1>
                </div>
            </div>
            <table>
                <tr>
                    <th>EARNINGS</th>
                    <th>AMOUNT</th>
                    <th>DEDUCTIONS</th>
                    <th>AMOUNT</th>
                </tr>
                {deduction_rows}
                {extra_earnings_rows}
                {extra_deduction_rows}

                <tr class="total-row">
                    <td>Gross Earnings</td>
                    <td>₹ {r(data["gross"])}</td>
                    <td>Total Deductions</td>
                    <td>₹ {r(data["total_deduction"])}</td>
                </tr>

            </table>

        <div class="final">
            <div>
                TOTAL NET PAYABLE<br>
                <span style="font-size:12px; color:#555;">
                    ({amount_in_words(r0(data["net_salary"]))})
                </span>
            </div>
            <div>₹ {r0(data["net_salary"])}</div>
        </div>

        <div class="footer">
            -- This is a system-generated Payslip-No signature required --
        </div>

    </div>

    </body>
    </html>
    """
    
    try:
        with sync_playwright() as p:

            browser = p.chromium.launch(
                headless=True,
                args=[
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-gpu",
                    "--single-process"
                ]
            )

            page = browser.new_page()

            page.set_content(html, wait_until="networkidle")

            page.pdf(
                path=file_name,
                format="A4",
                print_background=True,
                margin={
                    "top": "10mm",
                    "bottom": "10mm",
                    "left": "10mm",
                    "right": "10mm"
                }
            )

            browser.close()

    except Exception as e:
        st.error(str(e))
        raise
    return file_name
    # ---------------- BUILD PDF ----------------
if st.button("Generate Payslip"):

    if employee_name == "":
        st.error("Enter Employee ID first")

    else:

        data = {
            "employee_name": employee_name,
            "employee_id": employee_id,
            "employee_category": employee_category,
            "basic_salary": basic_salary,
            "da_salary": da_salary,
            "basic_current": basic_current,
            "da_current": da_current,
            "earned_leave": earned_leave,
            "bonus": bonus,
            "nh_h": nh_h,
            "pf": pf,
            "esic_calc": esic_calc,
            "pt": pt,
            "canteen": canteen,
            "gross": gross,
            "total_deduction": total_deduction,
            "net_salary": net_salary,
            "uan": uan,
            "esic": esic,
            "wc_no": wc_no,
            "bank_name": bank_name,
            "bank_account": bank_account,
            "doj": doj,
            "pay_period": pay_period,
            "paid_days": paid_days,
            "pay_date": pay_date.strftime("%d/%m/%Y"),
            "extra_earnings": st.session_state.extra_earnings,
            "extra_deductions": st.session_state.extra_deductions,
        }

        file = create_pdf(data)

        with open(file, "rb") as f:
            st.download_button("Download Payslip", f, file_name=file)
# --------------------------------------------------
# BULK PAYSLIP GENERATOR
# --------------------------------------------------
st.divider()
st.subheader("Bulk Payslip Generator")

uploaded_excel = st.file_uploader("Upload Excel", type=["xlsx"])

if uploaded_excel:

    df = pd.read_excel(uploaded_excel)
    st.dataframe(df)

    if st.button("Generate Bulk Payslips"):

        employees_df = load_employees()
        zip_name = "all_payslips.zip"

        progress = st.progress(0)
        errors = []

        with zipfile.ZipFile(zip_name, "w") as zipf:

            for i, row in df.iterrows():
                pay_date_raw = row.get("pay_date", "")

                if isinstance(pay_date_raw, pd.Timestamp):
                    pay_date = pay_date_raw.strftime("%d/%m/%Y")
                else:
                    pay_date = str(pay_date_raw)
                progress.progress((i+1)/len(df))

                try:
                    emp_id = str(row["emp_id"])
                    pay_period_raw = row["pay_month"]

                    if isinstance(pay_period_raw, pd.Timestamp):
                        pay_period = pay_period_raw.strftime("%B %Y")
                    else:
                        pay_period = str(pay_period_raw)

                    emp = employees_df[
                        employees_df["emp_id"].astype(str) == emp_id
                    ]

                    if emp.empty:
                        errors.append(f"{emp_id} not found")
                        continue

                    emp = emp.iloc[0]

                    # ---------------- BASE CALCULATION ----------------
                    total_days = float(row["total_days"])
                    paid_days = float(row["paid_days"])

                    basic_current = r2((emp["basic"]/total_days)*paid_days)
                    da_current = r2((emp["da"]/total_days)*paid_days)

                    earned_leave = r2((basic_current+da_current)*0.0481)

                    # BONUS
                    bonus = r2((basic_current+da_current)*0.0833)

                    # NH&H from Excel (editable)
                    if "nh_h" in row and not pd.isna(row["nh_h"]):
                        nh_h = float(row["nh_h"])
                    else:
                        nh_h = r2((basic_current+da_current)*0.0288)

                   # ---------- PF CALCULATION ----------
                   # ---------- PF CALCULATION (FIXED) ----------
                    default_pf = r2((basic_current + da_current) * 0.12)

                    # Skilled & Manager → manual PF from Excel
                    if emp["category"] in ["Skilled", "Manager"]:
                        
                        if "pf" in row and not pd.isna(row["pf"]):
                            pf = float(row["pf"])   # ✅ take from Excel
                        else:
                            pf = default_pf         # fallback if missing

                    # Others → auto PF
                    else:
                        pf = default_pf

                    if emp["category"] == "Manager":
                        esic_calc = 0
                    else:
                        esic_calc = r2((basic_current+da_current)*0.0075)

                    # PT Logic
                    if emp["category"] == "Manager":
                        pt = 200
                    elif emp["category"] in ["Skilled", "Semi-Skilled"]:
                        pt = 150
                    else:
                        pt = 0

                    canteen = r2(paid_days * 10)

                    # ---------------- EXTRA COLUMNS ----------------
                    extra_earnings = 0
                    extra_deductions = 0

                    earning_breakup = []
                    deduction_breakup = []

                    for col in row.index:

                        col_name = col.lower()

                        if col_name in ["emp_id", "lop_days", "pay_month", "total_days", "paid_days", "pay_date", "pf","nh_h"]:
                            continue

                        value = row[col]

                        if pd.isna(value) or value == "":
                            continue

                        value = float(value)

                        # Prefix method
                        if col.startswith("E_"):
                            extra_earnings += value
                            earning_breakup.append((col.replace("E_", ""), value))
                            continue

                        if col.startswith("D_"):
                            extra_deductions += value
                            deduction_breakup.append((col.replace("D_", ""), value))
                            continue

                        # Keyword detection
                        if any(k in col_name for k in ["bonus","allowance","ot","incentive"]):
                            extra_earnings += value
                            earning_breakup.append((col, value))
                            continue

                        if any(k in col_name for k in ["loan","advance","fine","deduction"]):
                            extra_deductions += value
                            deduction_breakup.append((col, value))
                            continue

                        # fallback
                        if value >= 0:
                            extra_earnings += value
                            earning_breakup.append((col, value))
                        else:
                            extra_deductions += abs(value)
                            deduction_breakup.append((col, abs(value)))

                    # ---------------- FINAL ----------------
                    gross = basic_current + da_current + earned_leave + bonus + nh_h + extra_earnings

                    total_deduction = pf + esic_calc + pt + canteen + extra_deductions

                    net_salary = gross - total_deduction

                    # 🔥 IMPORTANT FIX
                    extra_earnings_list = [
                        {"name": n, "amount": v} for n, v in earning_breakup
                    ]

                    extra_deductions_list = [
                        {"name": n, "amount": v} for n, v in deduction_breakup
                    ]

                    # ---------------- DATA ----------------
                    data = {
                        "employee_name": emp["name"],
                        "employee_id": emp_id,
                        "employee_category": emp["category"],
                        "basic_salary": emp["basic"],
                        "da_salary": emp["da"],
                        "basic_current": basic_current,
                        "da_current": da_current,
                        "earned_leave": earned_leave,
                        "bonus": bonus,
                        "nh_h": nh_h,
                        "pf": pf,
                        "esic_calc": esic_calc,
                        "pt": pt,
                        "canteen": canteen,
                        "gross": gross,
                        "total_deduction": total_deduction,
                        "net_salary": net_salary,
                        "uan": emp["uan"],
                        "esic": emp["esic"],
                        "wc_no": emp["wc_no"],
                        "bank_name": emp["bank"],
                        "bank_account": emp["account"],
                        "doj": emp["doj"],
                        "pay_period": pay_period,
                        "paid_days": paid_days,
                         "pay_date": pay_date,
                        
                        "extra_earnings": extra_earnings_list,
                        "extra_deductions": extra_deductions_list,
                    }

                    file = create_pdf(data)

                    if os.path.exists(file):
                        zipf.write(file, arcname=os.path.basename(file))
                        os.remove(file)

                except Exception as e:
                    errors.append(f"{emp_id} → {str(e)}")

        with open(zip_name, "rb") as f:
            st.success("Bulk Payslips Generated")
            st.download_button("Download ZIP", f, file_name=zip_name)

        if errors:
            st.warning("Errors:")
            for err in errors:
                st.write(err)