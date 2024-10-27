from geopy.distance import geodesic
import streamlit as st
import pandas as pd
import googlemaps
import os

if "excel_path" not in st.session_state:
    st.session_state.excel_path = ""
if "map_path" not in st.session_state:
    st.session_state.map_path = ""


#住所を渡して緯度と経度を受け取る
def get_lat_lng(address_input:str, gmaps):
    try:
        geocode_result = gmaps.geocode(address_input)
        location = geocode_result[0]["geometry"]["location"]
        lat = location["lat"]
        lng = location["lng"]
        return (lat, lng)
    except:
        return (None, None)

st.title("Get Distance")
st.write("エクセルファイルを入れて作成してください。")
with st.form("get_distance", clear_on_submit=True):
    title = st.text_input("title")
    api_key = st.text_input("your api_key")
    point_loc = st.text_input("address of a point")
    insert_address = st.text_input("県や市を追加してください。")
    insert_check = st.checkbox("県や市を差し込む")
    file = st.file_uploader("excel file", type="xlsx")
    excel_submit = st.form_submit_button("マップを作成する")

    if excel_submit:
        gmaps = googlemaps.Client(key=api_key)
        df = pd.read_excel(file, header=0, usecols=[0,1])
        df = df.dropna()
        point_lat, point_lng = get_lat_lng(point_loc, gmaps)
        point = (point_lat, point_lng)

        for data in df.itertuples():
            if insert_check:
                address = f"{insert_address}{data.住所}"
            else:
                address = data.住所
            lat, lng = get_lat_lng(address, gmaps)
            df.at[data.Index, "緯度"] = lat
            df.at[data.Index, "経度"] = lng
            if lat == None or lng == None:
                continue
            try:
                data_loc = (lat, lng)
                distance_m = geodesic(point, data_loc).m
                distance_m = f"{round(distance_m)}m"
                df.at[data.Index, "distance"] = distance_m
            except:
                df.at[data.Index, "distance"] = "エラー"
        
        excel_path = f"{title}.xlsx"
        df.to_excel(excel_path)
        
        st.session_state.excel_path = excel_path        


if st.session_state.excel_path:  
    with open(st.session_state.excel_path, "rb") as f:
        st.download_button(
            label="EXCELログダウンロード",
            data=f,
            file_name=os.path.basename(st.session_state.excel_path),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

