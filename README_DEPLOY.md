# Deploy to Streamlit Community Cloud

This folder contains two Streamlit apps:

- `capos_app.py` for `สรุปปริมาณงาน รายได้ จากระบบ New CAPOS รายเดือน ป.xlsx`
- `app.py` for `กรองจาก New CA 6 ปณ.xlsx`

## Deploy steps

1. Create a GitHub repository.
2. Upload/push this `Digital twin` folder to that repository.
3. In Streamlit Community Cloud, choose **New app**.
4. Select the GitHub repository and branch.
5. Set **Main file path** to:

```text
capos_app.py
```

6. Deploy.

Do not use `run_capos_app.py` on Streamlit Community Cloud. That file is only a local launcher for Windows machines where `streamlit.exe` is blocked.

If the app file is inside a subfolder in your GitHub repo, use the relative path instead, for example:

```text
Digital twin/capos_app.py
```

## Required files

- `capos_app.py`
- `requirements.txt`
- `สรุปปริมาณงาน รายได้ จากระบบ New CAPOS รายเดือน ป.xlsx`

Power BI files in `PowerBI_CAPOS/` are not required for the Streamlit app.

If Streamlit Cloud shows `Import openpyxl failed`, make sure `requirements.txt` is in the root of the GitHub repository selected in Streamlit Cloud, not only inside a subfolder. The file must include:

```text
openpyxl>=3.1.2
```
