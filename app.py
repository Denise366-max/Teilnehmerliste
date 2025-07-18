import streamlit as st
import requests
import pandas as pd
from io import BytesIO
import re

API_TOKEN = "dfaec32a80975c26802ef8fcf68bf4cc72990046"
BASE_URL = "https://api.pipedrive.com/v1"

# Mapping von Custom Field API-Key → Teilnehmer-Namen
CUSTOM_FIELD_MAPPING = {
    "b94970cec0683f8f5cadf4d3fab7079744ac28bb": "Teilnehmer 1",
    "fd69701c7c53f9854ad1df204cd81da18111a072": "Teilnehmer 2",
    "179be66b310cfbf4d1767437563e6fc902714c9a": "Teilnehmer 3",
    "6f9c097dc7bbc95739bebb8c48568b51c32aff26": "Teilnehmer 4",
    "01377c7c43ae80f19389c65d42efb810b6e8550a": "Teilnehmer 5",
    "d2fbe7e0ce8cf09700ea21d889c2a916e2ddf998": "Teilnehmer 6",
    "f23da18af8d33cea710381dc6153fc561cc851f8": "Teilnehmer 7",
    "bcfe6354dd76859d01e22c7ff091f5e2c072acda": "Teilnehmer 8",
    "8c1e8e89c4434d868075f21b98367b9cd9a2261c": "Teilnehmer 9",
    "9cd06a5ed78b5f958f8a7c44d8d4d721cfe406c6": "Teilnehmer 10",
    "de83542f0b2e38add0642101f11318ff095ddd21": "Teilnehmer 11",
    "4d64992cc7cae75827de3500849d8845ab48d0e3": "Teilnehmer 12",
    "309dd96701de50c73cf0cb3f2ba468bec57fa7aa": "Teilnehmer 13",
    "8d4f95f9410329cc17d4ad055c17310d718f6159": "Teilnehmer 14",
    "0c5c0e0865e3ae8d982f7cb30aa3c37eff13184b": "Teilnehmer 15",
    "f9be7b6ab264301204974afd9d8602352e02ce28": "Teilnehmer 16",
    "3d0a1e3f721f0abbc50843202476318ce4b3258e": "Teilnehmer 17",
    "214f05f31ed7d6bc7fac9dd51ef893c3a151462b": "Teilnehmer 18",
}

def get_deal_data(deal_id):
    url = f"{BASE_URL}/deals/{deal_id}?api_token={API_TOKEN}"
    response = requests.get(url)
    if response.status_code != 200:
        st.error(f"Fehler beim Abrufen des Deals: {response.status_code}")
        return None
    return response.json()

def extract_custom_participants(deal_json):
    if not deal_json or not deal_json.get("success"):
        return []

    deal_data = deal_json["data"]
    participants = []

    for field_key, label in CUSTOM_FIELD_MAPPING.items():
        value = deal_data.get(field_key)
        if not value:
            continue

        name = ""
        email = ""

        if isinstance(value, dict):
            name = value.get("name", "")
            emails = value.get("email", [])
            if isinstance(emails, list):
                for e in emails:
                    if e.get("value"):
                        email = e["value"]
                        break
            elif isinstance(emails, str):
                email = emails
        elif isinstance(value, str):
            name = value

        participants.append({
            "Teilnehmer": label,
            "Name": name,
            "E-Mail": email
        })

    return participants

st.title("Nur Custom Teilnehmerfelder aus Pipedrive Deal")

deal_id = st.text_input("Bitte Deal-ID eingeben:")

if st.button("Teilnehmer laden und Excel erstellen") and deal_id:
    deal_json = get_deal_data(deal_id)
    participants = extract_custom_participants(deal_json)

    if not participants:
        st.error("Keine Teilnehmer gefunden.")
    else:
        df = pd.DataFrame(participants)
        st.dataframe(df)

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False)
        buffer.seek(0)

        deal_title = deal_json["data"].get("title", "deal")
        filename = re.sub(r"[^a-zA-Z0-9_-]", "_", deal_title) + "_teilnehmer.xlsx"

        st.download_button(
            label="Excel herunterladen",
            data=buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )