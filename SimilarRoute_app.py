import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder
from io import BytesIO
import os

# ================= PAGE CONFIG =================
st.set_page_config(
    page_title="ALPHA Similar Route Finder",
    layout="wide"
)

st.title("🏭 ALPHA MEASUREMENT SOLUTIONS - Similar Route Finder")
st.markdown("### Advanced Manufacturing Similarity Analytics System")

# ================= FILE SETTINGS =================
import os

BASE_DIR = os.path.dirname(__file__)  # path of the script
excel_file = os.path.join(BASE_DIR, "Route File Lean.xlsx")  # relative path to Excel in repo

# ================= LOAD DATA =================
@st.cache_data(show_spinner=False)
def load_data(path):
    try:
        apnrn = pd.read_excel(path, sheet_name="apnrn", usecols=["partno","routeno"])
        rodetail = pd.read_excel(path, sheet_name="rodetail", usecols=["routeno","laborgrade","cycletime"])
        immaster = pd.read_excel(path, sheet_name="immaster",
                                 usecols=["item","descrip","prodclas","misc02","misc05","misc10","misccode"])

        # ---- CLEAN DATA ----
        apnrn['partno'] = apnrn['partno'].astype(str).str.strip()
        immaster['item'] = immaster['item'].astype(str).str.strip()
        rodetail['cycletime'] = pd.to_numeric(rodetail['cycletime'], errors='coerce').fillna(0)
        rodetail['laborgrade'] = rodetail['laborgrade'].astype(str).str.upper().str.strip()

        # ---- TOTAL ROUTE TIME ----
        route_times = rodetail.groupby('routeno')['cycletime'].sum().reset_index(name='TotalCycleTime')
        for grade in ['FA1','CA1','ER1']:
            df = rodetail[rodetail['laborgrade']==grade] \
                    .groupby('routeno')['cycletime'] \
                    .sum() \
                    .reset_index(name=f'{grade}_Time')
            route_times = route_times.merge(df, on='routeno', how='left')

        route_times.fillna(0, inplace=True)
        route_times['FA_CA_ER_Total'] = route_times[['FA1_Time','CA1_Time','ER1_Time']].sum(axis=1)

        # ---- MERGE ----
        data = apnrn.merge(route_times, on='routeno', how='left')
        immaster = immaster.rename(columns={
            "item":"PartNumber",
            "descrip":"Description",
            "prodclas":"ProdClas",
            "misc02":"Misc02",
            "misc05":"Misc05",
            "misc10":"Misc10",
            "misccode":"MiscCode"
        })
        data = data.merge(immaster, left_on='partno', right_on='PartNumber', how='left')
        data.rename(columns={"partno":"PartNumber","routeno":"RouteNo"}, inplace=True)
        data = data.loc[:,~data.columns.duplicated()].copy()
        return data

    except FileNotFoundError:
        st.error(f"Excel file not found at: {path}")
        return pd.DataFrame()  # prevent crash
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return pd.DataFrame()


# ================= EXCEL REFRESH =================
if st.button("🔄 Refresh Excel File"):
    st.cache_data.clear()
    st.success("Excel file reloaded successfully!")

data = load_data(excel_file)

# ================= SIDEBAR FILTERS =================
st.sidebar.header("🔎 Filters Panel")

part_number = st.sidebar.selectbox(
    "Select Target Part",
    sorted(data['PartNumber'].astype(str).dropna().unique())
)

tolerance = st.sidebar.slider("Tolerance (%)", 1, 50, 10)/100

match_mode = st.sidebar.radio(
    "Matching Mode",
    [
        "Labor Content (FA+CA+ER)",
        "Total Route Time",
        "Weighted Labor Similarity",
        "FA1 Only",
        "CA1 Only",
        "ER1 Only"
    ]
)

prodclass_filter = st.sidebar.selectbox("Filter ProdClas", ["All"] + sorted(data['ProdClas'].dropna().astype(str).unique()))
misc02_filter = st.sidebar.selectbox("Filter Misc02", ["All"] + sorted(data['Misc02'].dropna().astype(str).unique()))
misccode_filter = st.sidebar.selectbox("Filter MiscCode", ["All"] + sorted(data['MiscCode'].dropna().astype(str).unique()))

# ================= TARGET =================
target = data[data['PartNumber']==part_number]
if target.empty:
    st.error("Selected part not found.")
    st.stop()

# Determine compare column
if match_mode == "Total Route Time":
    compare_column = "TotalCycleTime"
    target_time = float(target['TotalCycleTime'].values[0])
elif match_mode=="FA1 Only":
    compare_column = "FA1_Time"
    target_time = float(target['FA1_Time'].values[0])
elif match_mode=="CA1 Only":
    compare_column = "CA1_Time"
    target_time = float(target['CA1_Time'].values[0])
elif match_mode=="ER1 Only":
    compare_column = "ER1_Time"
    target_time = float(target['ER1_Time'].values[0])
else:
    compare_column = "FA_CA_ER_Total"
    target_time = float(target['FA_CA_ER_Total'].values[0])

# ================= FILTERING =================
filtered = data.copy()
if prodclass_filter!="All":
    filtered = filtered[filtered['ProdClas'].astype(str)==prodclass_filter]
if misc02_filter!="All":
    filtered = filtered[filtered['Misc02'].astype(str)==misc02_filter]
if misccode_filter!="All":
    filtered = filtered[filtered['MiscCode'].astype(str)==misccode_filter]

# ---- Tolerance Filtering ----
low, high = target_time*(1-tolerance), target_time*(1+tolerance)
filtered = filtered[(filtered[compare_column]>=low)&(filtered[compare_column]<=high)].copy()

# ================= SIMILARITY CALCULATION =================
if match_mode == "Weighted Labor Similarity":
    target_mix = [
        target['FA1_Time'].values[0],
        target['CA1_Time'].values[0],
        target['ER1_Time'].values[0]
    ]
    total_target = sum(target_mix)
    if total_target == 0:
        filtered['Similarity_Score (%)'] = 0
    else:
        target_mix = [x/total_target for x in target_mix]
        def similarity(row):
            total = row['FA1_Time'] + row['CA1_Time'] + row['ER1_Time']
            if total==0:
                return 0
            mix = [row['FA1_Time']/total, row['CA1_Time']/total, row['ER1_Time']/total]
            diff = sum(abs(a-b) for a,b in zip(target_mix,mix))
            return round((1-diff)*100,2)
        filtered['Similarity_Score (%)'] = filtered.apply(similarity, axis=1)
        filtered.sort_values(by='Similarity_Score (%)', ascending=False, inplace=True)
else:
    filtered['Percent_Difference'] = ((filtered[compare_column]-target_time)/target_time)*100
    filtered.sort_values(by='Percent_Difference', key=lambda x: x.abs(), inplace=True)

# ================= COLUMN SELECTION =================
st.sidebar.subheader("📊 Display Columns")
base_columns = ["PartNumber","Description","RouteNo","TotalCycleTime",
                "FA1_Time","CA1_Time","ER1_Time","FA_CA_ER_Total",
                "Misc02","Misc05","Misc10"]
if match_mode == "Weighted Labor Similarity":
    base_columns.append("Similarity_Score (%)")
else:
    base_columns.append("Percent_Difference")

selected_cols = st.sidebar.multiselect("Select Columns", base_columns, default=base_columns)
display_df = filtered[selected_cols]

# ================= KPI METRICS =================
st.subheader("📊 Performance Summary")
col1,col2,col3,col4 = st.columns(4)
col1.metric("Target Time", round(target_time,2))
col2.metric("Matching Parts", len(display_df))
col3.metric("Min Time", round(display_df[compare_column].min(),2))
col4.metric("Max Time", round(display_df[compare_column].max(),2))

st.divider()

# ================= AGGRID TABLE =================
gb = GridOptionsBuilder.from_dataframe(display_df)
gb.configure_default_column(filterable=True, sortable=True, resizable=True)
for col in display_df.select_dtypes(include=["float64","int64"]).columns:
    gb.configure_column(col, type=["numericColumn"])

AgGrid(display_df, gridOptions=gb.build(), height=500, fit_columns_on_grid_load=True)

# ================= COPY & DOWNLOAD =================
st.subheader("📋 Copy / Export")
if st.button("Generate Part List"):
    st.text_area("Copy Below:", "\n".join(display_df['PartNumber'].astype(str)), height=200)

buffer = BytesIO()
display_df.to_excel(buffer, index=False, engine='openpyxl')
buffer.seek(0)
st.download_button("⬇ Download Results as Excel", buffer, "Similar_Parts_Output.xlsx")
