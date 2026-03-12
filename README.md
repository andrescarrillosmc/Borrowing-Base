# Henry — Financials Processing Platform

A locally-hosted web application for automating portfolio company financial statement processing. Built for a private equity firm to streamline the extraction, classification, and review of financial data from Excel workbooks.

## Tech Stack

- **Frontend**: React + TypeScript + Vite + Tailwind CSS
- **Backend**: Python FastAPI + SQLite
- **AI**: Anthropic Claude API
- **Excel Processing**: openpyxl + SheetJS (client-side rendering)

## Prerequisites

- Node.js 18+
- Python 3.11+

## Setup

### 1. Clone and navigate
```bash
cd financial-platform
```

### 2. Backend setup
```bash
cd backend
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt

# Create .env from template
cp .env.example .env
# Edit .env and add your ANTHROPIC_API_KEY
```

### 3. Frontend setup
```bash
cd frontend
npm install
```

### 4. Add prompt files
Place the four AI prompt markdown files in `backend/prompts/`:
- `layer1_income_statement.md`
- `layer1_balance_sheet.md`
- `layer2_income_statement.md`
- `layer2_balance_sheet.md`

### 5. Start the application

**Backend** (from `backend/` directory):
```bash
uvicorn app.main:app --reload --port 8000
```

**Frontend** (from `frontend/` directory):
```bash
npm run dev
```

The app will be available at `http://localhost:5173`.

## Project Structure

```
financial-platform/
├── frontend/          # React + TypeScript + Vite frontend
├── backend/           # Python FastAPI backend
│   ├── app/           # Application code
│   ├── prompts/       # AI prompt templates (loaded at runtime)
│   ├── templates/     # Firm's standardized output template
│   ├── uploads/       # Temporary uploaded file storage
│   └── processed/     # Processed PDFs and CSVs
└── README.md
```

## Wizard Flow

1. **Step 1 - Upload & Extract**: Upload Excel workbook → Layer 1 AI extraction → Analyst review
2. **Step 2 - Classify & Review**: Layer 2 AI classification into standardized template → Analyst corrections
3. **Step 3 - Finalize**: Review final output → Save to database

## Environment Variables

| Variable | Description | Default |
|---|---|---|
| `ANTHROPIC_API_KEY` | Anthropic API key | Required |
| `LAYER1_MODEL` | Claude model for Layer 1 extraction | `claude-sonnet-4-6` |
| `LAYER2_MODEL` | Claude model for Layer 2 classification | `claude-opus-4-6` |
| `LAYER_A_MODEL` | Claude model for Layer A instruction rewriting | `claude-sonnet-4-6` |
| `LAYER_B_MODEL` | Claude model for Layer B markdown integration | `claude-opus-4-6` |
