import streamlit as st
import pandas as pd
from openpyxl import Workbook
from io import BytesIO

st.title("MasterCMD Generator")

st.markdown("### Enter Device-Node Mapping")
device_node_input = st.text_area("Device-Node Mapping (e.g., Dev1:Node1,Node2)", height=150)

st.markdown("### Configuration Parameters")
disable_nodes = st.text_input("Disable Nodes (comma-separated)", "Node3,Node4")
count_value = st.text_input("Count Value", "10")
dev_address = st.text_input("DevAddress Value", "0x01")

if st.button("Generate Excel"):
    wb = Workbook()
    ws = wb.active
    ws.title = "MasterCMD"

    ws.append(["Device", "Node", "Disable", "Count", "DevAddress"])

    for line in device_node_input.strip().splitlines():
        if ':' in line:
            device, nodes = line.split(':')
            for node in nodes.split(','):
                disable = "Yes" if node.strip() in disable_nodes.split(',') else "No"
                ws.append([device.strip(), node.strip(), disable, count_value, dev_address])

    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="Download Excel File",
        data=output,
        file_name="MasterCMD.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
