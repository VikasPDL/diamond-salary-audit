"""Diamond Salary Audit - upload the 4 monthly files, download the reconciliation workbook."""
import hmac
import io
from pathlib import Path

import pandas as pd
import streamlit as st

import engine

APP_DIR = Path(__file__).parent
DEFAULT_MASTER = APP_DIR / "Employee_Master.xlsx"

st.set_page_config(page_title="Diamond Salary Audit", page_icon="💎", layout="wide")


def check_password():
    """Block the app until the password from st.secrets['APP_PASSWORD'] is entered."""
    try:
        expected = st.secrets["APP_PASSWORD"]
    except Exception:
        st.error("App password is not configured. Add APP_PASSWORD to the Streamlit secrets.")
        st.stop()
    if st.session_state.get("authed"):
        return
    st.title("💎 Diamond Salary Audit")
    pwd = st.text_input("Password", type="password")
    if pwd:
        if hmac.compare_digest(pwd, str(expected)):
            st.session_state["authed"] = True
            st.rerun()
        st.error("Wrong password")
    st.stop()


check_password()
st.title("💎 Diamond Salary Audit")
st.caption("Upload the month's 4 files → get the audit Excel (DAP earning vs bank transfer, attendance check, name matching).")

# ---------------------------------------------------------------- sidebar
with st.sidebar:
    st.header("Settings")
    tolerance = st.number_input("'Matched' tolerance (Rs)", min_value=0, value=500, step=100,
                                help="If |DAP earning − Bank transfer| is within this amount the remark says Matched.")
    st.caption("Remarks: <Department> – Matched / Paid Less than DAP / Paid High than DAP / No Production in DAP / Staff. "
               "If salary also went to another name's bank account, the name, A/c no. and amount are added.")
    st.markdown("---")
    st.markdown(
        "**Employee Master columns:** Emp Code, Employee Name, Designation, Department, Sheet Name, "
        "Salary Master, Name in Bank Paid file, Name in Employee Production file.\n\n"
        "To fix a name that doesn't match, fill the *Name in …* column in the master and upload it again."
    )

# ---------------------------------------------------------------- uploads
master_label = "0. Employee Master  (Employee Master - for App Upload.xlsx)"
if DEFAULT_MASTER.exists():
    master_label += " – optional, built-in master is used if empty"
master_up = st.file_uploader(master_label, type=["xlsx"])
c1, c2 = st.columns(2)
with c1:
    pdf_file = st.file_uploader("1. Attendance PDF  (e.g. DIAMOND AUG-26.pdf)", type=["pdf"])
    prod_file = st.file_uploader("2. DAP Worker Production  (e.g. Worker_Production_Aug_2026.csv)", type=["csv", "xlsx"])
with c2:
    bank_file = st.file_uploader("3. Bank Transfer  (e.g. AUGUST,2026 SALARY 2ND FLOOR TRANSFER.xlsx)", type=["xlsx"])
    sal_file = st.file_uploader("4. Salary Sheet  (e.g. DIAMOND AUGUST,2026 – Salary sheet.xlsx)", type=["xlsx"])

ready = all([pdf_file, prod_file, bank_file, sal_file]) and (master_up or DEFAULT_MASTER.exists())
if st.button("▶ Generate Audit Report", type="primary", disabled=not ready):
    try:
        with st.spinner("Reading files… the attendance PDF takes about a minute"):
            master_src = master_up if master_up is not None else str(DEFAULT_MASTER)
            if prod_file.name.lower().endswith(".xlsx"):
                prod_src = io.StringIO(pd.read_excel(prod_file).to_csv(index=False))
            else:
                prod_src = prod_file
            people, unmatched, meta, xlsx = engine.run(master_src, sal_file, bank_file, prod_src, pdf_file, tolerance)
        st.session_state["result"] = (people, unmatched, meta, xlsx)
    except Exception as exc:  # show the problem to the user instead of a stack trace
        st.session_state.pop("result", None)
        st.error(f"Could not process the files: {exc}")
        st.exception(exc)
elif not ready:
    st.info("Upload the Employee Master and the 4 monthly files to enable the button."
            if not DEFAULT_MASTER.exists() else "Upload all 4 files to enable the button.")

# ---------------------------------------------------------------- results
if "result" in st.session_state:
    people, unmatched, meta, xlsx = st.session_state["result"]
    st.success(f"Report ready for **{meta['period']}** – {len(people)} employees.")
    st.download_button("⬇ Download Excel report", data=xlsx, type="primary",
                       file_name=f"Diamond Salary Audit {meta['period']}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    dap_total = sum(p["dap"] for p in people if p["dap"] != engine.NO_PROD)
    bank_total = sum(p["bank_amt"] for p in people)
    base = [p["remark_base"] for p in people]
    m = st.columns(6)
    m[0].metric("Employees", len(people))
    m[1].metric("Earn Salary – DAP", f"₹{dap_total:,.0f}")
    m[2].metric("Bank transfer", f"₹{bank_total:,.0f}")
    m[3].metric("Paid High than DAP", sum(1 for b in base if b.endswith("Paid High than DAP")))
    m[4].metric("Paid to other person's A/c", sum(1 for p in people if p["other_bank"]))
    m[5].metric("Unmatched names", sum(len(v) for v in unmatched.values()))

    main = pd.DataFrame([{
        "Emp Code": p["code"], "Employee Name": p["name"], "Designation": p["designation"],
        "Salary Master": p["salary_master"], "Earn Salary - DAP": p["dap"], "Pcs Done": p["pcs"],
        "Cts Done": p["cts"], "Diff - Attdance": p["att_diff"],
        "Effiency - Bank trans - Salry DAP": p["eff"], "Remarks": p["remarks"]} for p in people])
    main["Earn Salary - DAP"] = main["Earn Salary - DAP"].astype(str)

    t1, t2, t3, t4, t5 = st.tabs(["Main report", "Remarks", "Attendance", "Name matching", "Unmatched names"])
    with t1:
        q = st.text_input("Search name / remark")
        view = main[main.apply(lambda r: q.upper() in f"{r['Employee Name']} {r['Remarks']}".upper(), axis=1)] if q else main
        st.dataframe(view, use_container_width=True, hide_index=True, height=520)
    with t2:
        rem = pd.DataFrame({"Department": [str(p["department"] or p["sheet"] or "Unassigned") for p in people],
                            "Remarks": base})
        st.subheader("Count by remark")
        cnt = rem["Remarks"].value_counts().sort_index()
        cnt.loc["Grand Total"] = cnt.sum()
        st.dataframe(cnt.rename("Count"), use_container_width=True)
        st.subheader("Department × Remarks")
        rem["Outcome"] = rem["Remarks"].map(lambda r: next((o for o in engine.OUTCOMES if r.endswith(o)), "Other"))
        pv = pd.crosstab(rem["Department"], rem["Outcome"], margins=True, margins_name="Total")
        st.dataframe(pv.reindex(columns=[c for c in engine.OUTCOMES + ["Other", "Total"] if c in pv.columns]),
                     use_container_width=True)
        ob = [{"Employee": p["name"], "Emp Code": p["code"], "Paid to": b["name"],
               "Account No": ", ".join(b["accounts"]), "Amount": b["amount"],
               "Paid in own name too?": "Yes" if p["bank"] else "No", "Net Salary": round(p["net_sal"]),
               "Linked by": b["basis"]} for p in people for b in p["other_bank"]]
        st.subheader(f"Salary transferred to other person's bank account ({len(ob)})")
        st.dataframe(pd.DataFrame(ob), use_container_width=True, hide_index=True)
        sa = [{"Employee": p["name"], "Emp Code": p["code"], "Account No": a, "Also used for": ", ".join(o)}
              for p in people for a, o in p["same_acc"]]
        if sa:
            st.subheader(f"Same account number used for different names ({len(sa)})")
            st.dataframe(pd.DataFrame(sa), use_container_width=True, hide_index=True)
    with t3:
        att = pd.DataFrame([{
            "Emp Code": p["code"], "Employee Name": p["name"],
            "Present": p["pdf"]["present"] if p["pdf"] else None,
            "Absent": p["pdf"]["absent"] if p["pdf"] else None,
            "Paid Days (PDF)": p["att_pdf"], "Paid Days (Salary Sheet)": p["att_sal"], "Diff": p["att_diff"],
        } for p in people])
        only = st.checkbox("Show only mismatches", value=True)
        st.dataframe(att[att["Diff"] != 0] if only else att, use_container_width=True, hide_index=True, height=480)
    with t4:
        nm = pd.DataFrame([{
            "Emp Code": p["code"], "Master Name": p["name"], "Salary Sheet": p["sal"]["name"] if p["sal"] else "",
            "Production": p["prod"]["name"] if p["prod"] else "", "Bank": "; ".join(([p["bank"]["name"]] if p["bank"] else []) + [b["name"] + " (other name)" for b in p["other_bank"]]),
            "Attendance PDF": p["pdf"]["name"] if p["pdf"] else "", "PDF match": p.get("pdf_how", ""),
        } for p in people])
        st.dataframe(nm, use_container_width=True, hide_index=True, height=480)
    with t5:
        st.subheader("Production names not linked")
        st.dataframe(pd.DataFrame(unmatched["prod"]).drop(columns=["key", "rows"], errors="ignore"),
                     use_container_width=True, hide_index=True)
        st.subheader("Bank names not linked")
        st.dataframe(pd.DataFrame(unmatched["bank"]), use_container_width=True, hide_index=True)
        st.subheader("Attendance PDF employees not linked")
        st.dataframe(pd.DataFrame([{"Emp Code": e["code"], "Name": e["name"], "Department": e["department"],
                                    "Paid Days": e["paid"]} for e in unmatched["pdf"]]),
                     use_container_width=True, hide_index=True)
