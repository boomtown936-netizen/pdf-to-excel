from fastapi import FastAPI, File, UploadFile, Form
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import pathlib
import io
import pandas as pd

# PDF parsing imports
import camelot
import pdfplumber

app = FastAPI()

# ✅ Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ✅ Serve static files (optional)
app.mount("/static", StaticFiles(directory="app/static"), name="static")

# ✅ Serve the frontend at "/"
@app.get("/", response_class=HTMLResponse)
async def serve_home():
    html_path = pathlib.Path("app/templates/index.html")
    return HTMLResponse(html_path.read_text())


# ✅ Conversion endpoint
@app.post("/convert")
async def convert_pdf(
    file: UploadFile = File(...),
    pages: str = Form("all"),
    flavor: str = Form("auto"),
):
    # Read PDF
    content = await file.read()

    # Try parsing with Camelot
    tables = []
    try:
        if flavor in ("auto", "camelot-lattice", "camelot-stream"):
            fl = "lattice" if flavor == "camelot-lattice" else "stream"
            tables = camelot.read_pdf(io.BytesIO(content), pages=pages, flavor=fl)
    except Exception:
        pass

    # Fallback: pdfplumber
    if not tables or len(tables) == 0 or flavor == "pdfplumber":
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables():
                    df = pd.DataFrame(table[1:], columns=table[0])
                    tables.append(df)

    # Convert tables to Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if isinstance(tables, list) and tables and isinstance(tables[0], pd.DataFrame):
            for i, df in enumerate(tables, start=1):
                df.to_excel(writer, index=False, sheet_name=f"Table_{i}")
        elif hasattr(tables, "n") and tables.n > 0:
            for i, t in enumerate(tables, start=1):
                t.df.to_excel(writer, index=False, sheet_name=f"Table_{i}")
        else:
            pd.DataFrame([["No tables found"]]).to_excel(writer, index=False, sheet_name="Sheet1")

    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={file.filename}.xlsx"},
    )
