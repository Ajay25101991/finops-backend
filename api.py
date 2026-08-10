import os
import io
import json
import tempfile
import importlib.util
from importlib.machinery import SourceFileLoader

import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openai import OpenAI

from data_cleaner import clean

# ── Load FS engine ────────────────────────────────────────────────────────────
SCRIPT_PATH = os.path.join(os.path.dirname(__file__), "TB to Financial Statements 1 click.Py")
_loader = SourceFileLoader("fs_engine", SCRIPT_PATH)
spec    = importlib.util.spec_from_loader("fs_engine", _loader)
engine  = importlib.util.module_from_spec(spec)
_loader.exec_module(engine)

# ── OpenAI client ─────────────────────────────────────────────────────────────
openai_client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY", ""))

# ── FastAPI app ───────────────────────────────────────────────────────────────
app = FastAPI(title="FinOps Report Generator")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# ── Health ────────────────────────────────────────────────────────────────────
@app.get("/health")
def health():
    return {"status": "ok"}


# ── Financial Statements generator ───────────────────────────────────────────
@app.post("/generate")
async def generate_report(
    tb:      UploadFile = File(...),
    mapping: UploadFile = File(...),
    company: str = Form(default="DemoCo Pvt. Ltd."),
    period:  str = Form(default="Year Ended 31 December 2025"),
):
    with tempfile.TemporaryDirectory() as tmpdir:
        tb_path  = os.path.join(tmpdir, "tb.xlsx")
        map_path = os.path.join(tmpdir, "mapping.xlsx")
        out_path = os.path.join(tmpdir, "Financial_Statements.xlsx")

        with open(tb_path,  "wb") as f: f.write(await tb.read())
        with open(map_path, "wb") as f: f.write(await mapping.read())

        engine.generate(tb_path, map_path, out_path, company=company, period=period)

        return FileResponse(
            out_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="Financial_Statements.xlsx",
            headers={"Access-Control-Expose-Headers": "Content-Disposition"},
        )


# ── Data Cleaner ──────────────────────────────────────────────────────────────
@app.post("/clean")
async def clean_data(
    file:    UploadFile = File(...),
    company: str = Form(default=""),
):
    # 1. Read uploaded file
    raw = await file.read()
    ext = file.filename.split(".")[-1].lower()

    try:
        if ext == "csv":
            df = pd.read_csv(io.BytesIO(raw))
        else:
            df = pd.read_excel(io.BytesIO(raw))
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": f"Could not read file: {str(e)}"})

    # 2. Run cleaning engine
    result = clean(df)

    issues    = result["issues"]
    df_clean  = result["df_clean"]

    # 3. Ask OpenAI to explain issues and write CFO summary
    ai_summary    = ""
    ai_per_issue  = []

    if openai_client.api_key and issues:
        # Per-issue plain English explanation
        issues_text = "\n".join(
            f"- [{i['severity']}] {i['type']}: {i['detail']}" for i in issues
        )
        prompt = f"""You are a CFO-level financial data quality analyst.

A financial data file{f' for {company}' if company else ''} was uploaded with {result['total_rows']} rows.
The following data quality issues were detected:

{issues_text}

For each issue, write ONE concise sentence (max 20 words) explaining the business risk to a CFO.
Then write a 2-sentence overall executive summary.

Respond in JSON:
{{
  "per_issue": ["explanation for issue 1", "explanation for issue 2", ...],
  "executive_summary": "2-sentence summary here."
}}"""

        try:
            response = openai_client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3,
                max_tokens=600,
            )
            raw_json = response.choices[0].message.content.strip()
            # strip markdown fences if present
            if raw_json.startswith("```"):
                raw_json = raw_json.split("```")[1]
                if raw_json.startswith("json"):
                    raw_json = raw_json[4:]
            parsed        = json.loads(raw_json)
            ai_per_issue  = parsed.get("per_issue", [])
            ai_summary    = parsed.get("executive_summary", "")
        except Exception:
            ai_summary   = "AI explanation unavailable — check OpenAI API key."
            ai_per_issue = ["" for _ in issues]

    # Attach AI explanation to each issue
    for i, issue in enumerate(issues):
        issue["ai_explanation"] = ai_per_issue[i] if i < len(ai_per_issue) else ""

    # 4. Save clean file to temp, stream back
    with tempfile.TemporaryDirectory() as tmpdir:
        clean_path = os.path.join(tmpdir, "Cleaned_Data.xlsx")
        df_clean.to_excel(clean_path, index=False)

        # Read bytes to return alongside JSON (as multipart would be complex)
        # Instead: return JSON with issues + base64 clean file
        import base64
        with open(clean_path, "rb") as f:
            clean_b64 = base64.b64encode(f.read()).decode()

    return JSONResponse(content={
        "total_rows":     result["total_rows"],
        "clean_rows":     result["clean_rows"],
        "critical_count": result["critical_count"],
        "warning_count":  result["warning_count"],
        "info_count":     result["info_count"],
        "total_issues":   result["total_issues"],
        "is_balanced":    result["is_balanced"],
        "issues":         issues,
        "ai_summary":     ai_summary,
        "clean_file_b64": clean_b64,
    })


# ── Run ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=5000)
