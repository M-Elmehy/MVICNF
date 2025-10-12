import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, PatternFill
from pathlib import Path
import io

def generate_excel(devices, blocks, rows_per_block, node_seq, rules):
    param_template = [
        "MCM.CONFIG.Port1MasterCmd[{i}].Enable",
        "MCM.CONFIG.Port1MasterCmd[{i}].IntAddress",
        "MCM.CONFIG.Port1MasterCmd[{i}].PollInt",
        "MCM.CONFIG.Port1MasterCmd[{i}].Count",
        "MCM.CONFIG.Port1MasterCmd[{i}].Swap",
        "MCM.CONFIG.Port1MasterCmd[{i}].Node",
        "MCM.CONFIG.Port1MasterCmd[{i}].Func",
        "MCM.CONFIG.Port1MasterCmd[{i}].DevAddress",
    ]

    count_map = rules["count_map"]
    devaddr_map = rules["devaddr_map"]
    func_value = rules["func"]
    enable_value = rules["enable"]
    int_offset = rules["int_offset"]

    rows = []
    global_block_idx = 0
    prev_intaddr = None
    prev_count = None
    prev_device = None

    for device_no in range(1, devices + 1):
        node_no = node_seq[device_no - 1]
        for block_no in range(1, blocks + 1):
            idx = global_block_idx
            for p in param_template:
                param = p.format(i=idx)
                base = param.split('.')[-1]
                cfg = ""
                if block_no in count_map or block_no in devaddr_map:
                    if base == "Enable":
                        cfg = enable_value
                    elif base == "Count":
                        cfg = count_map.get(block_no, "")
                    elif base == "IntAddress":
                        if prev_intaddr is None:
                            cfg = 0
                        else:
                            addition = prev_intaddr + (prev_count if prev_count else 0)
                            if prev_device and prev_device != device_no:
                                addition += int_offset
                            cfg = addition
                    elif base == "Node":
                        cfg = node_no
                    elif base == "Func":
                        cfg = func_value
                    elif base == "DevAddress":
                        cfg = devaddr_map.get(block_no, "")
                rows.append({
                    "Device No.": device_no,
                    "Block No.": block_no,
                    "Node No.": node_no,
                    "Parameter": param,
                    "ConfigValue": cfg
                })
            if block_no in count_map:
                prev_count = count_map[block_no]
                last_8 = rows[-8:]
                assigned_intaddr = last_8[1]['ConfigValue']
                prev_intaddr = assigned_intaddr
                prev_device = device_no
            global_block_idx += 1

    df = pd.DataFrame(rows)
    # Save to BytesIO instead of file
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Port1", index=False)
        pd.DataFrame(columns=df.columns).to_excel(writer, sheet_name="Port2", index=False)

    buffer.seek(0)
    return buffer


st.title("⚙️ MasterCmd Excel Generator")

devices = st.number_input("Number of Devices", min_value=1, value=26)
blocks = st.number_input("Blocks per Device", min_value=1, value=6)
rows_per_block = st.number_input("Rows per Block", min_value=1, value=8)
node_str = st.text_input("Node Numbers (comma-separated)", "26,27,28,29,30,31,32,33,34,35,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76")

st.markdown("### Rule Configuration")
enable_value = st.number_input(".Enable Value", value=1)
func_value = st.number_input(".Func Value", value=3)
int_offset = st.number_input("IntAddress Offset (new device)", value=10)

st.markdown("#### Block Mapping")
count_map = {}
devaddr_map = {}
for b in range(1, blocks + 1):
    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        st.text(f"Block {b}")
    with col2:
        cval = st.text_input(f".Count Value {b}", value=str(1 if b in [1, 3, 4] else 5), key=f"count{b}")
        if cval:
            count_map[b] = int(cval)
    with col3:
        dval = st.text_input(f".DevAddress Value {b}", value=str({1:3, 2:101, 3:116, 4:142}.get(b, "")), key=f"devaddr{b}")
        if dval:
            devaddr_map[b] = int(dval)

node_seq = [int(x.strip()) for x in node_str.split(",")]
rules = {"count_map": count_map, "devaddr_map": devaddr_map, "enable": enable_value, "func": func_value, "int_offset": int_offset}

if st.button("Generate Excel File"):
    if len(node_seq) != devices:
        st.error("❌ Number of node numbers must match number of devices.")
    else:
        buffer = generate_excel(devices, blocks, rows_per_block, node_seq, rules)
        st.success("✅ Excel file generated successfully!")
        st.download_button(
            label="📥 Download Excel",
            data=buffer,
            file_name="MasterCmd_Sequence.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
