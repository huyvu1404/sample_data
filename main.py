
from io import BytesIO
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv
from get_sample_data import get_sample_data, read_excel
from calculate_sample_size import get_sample_size
import os
load_dotenv()

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/sample")
async def upload_file(file: UploadFile = File(...)):
    try:
        df, sentiment_rates = read_excel(file.file)
        if df.empty:
            return {"error": "No data found in the provided Excel files."}
        total_entries = len(df)
        if total_entries == 0:
            return {"error": "No valid entries found."}
        sample_size = get_sample_size(N=total_entries)
        sampled_df = get_sample_data(df, sample_size, sentiment_rates)
        if sampled_df is None:
            return {"error": "Sampling failed."}
        buf = BytesIO()
        sampled_df.to_excel(buf, index=False)
        buf.seek(0)
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=sampled_data.xlsx"},
        )
    except Exception as e:
        return {"error": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host=os.getenv("HOST", "0.0.0.0"), port=int(os.getenv("PORT", 8000)))