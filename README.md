# Power.io

Power.io takes a CSV or Excel file and turns it into a chart inside your existing PowerPoint deck, styled to match the rest of the presentation.

**Live demo:** https://powerio.vercel.app/

## What it was built for

My cousin at Harvard was getting tired of building and formatting PowerPoint charts all the time for his business consulting organization. I wanted to make it as easy as possible where the user only gives the basics (chart type, which slide, and where the data comes from) and gets back a finished deck.

## What I built

I built and orchestrated the main logic for what charts get built, when, and how they get placed and styled. The backend reads the data, builds each chart with python-pptx, then has an LLM decide which theme fits the chart based on the rest of the presentation.

I designed the frontend in Vercel v0 and handed off the wiring between it and the backend to a code agent.

## The hardest part

Getting the charts to cooperate with the theme decisions I was hardcoding. Fixed rules kept clashing with decks that didn't fit them. It only worked after I added an LLM judge on the fly which allows the backend to measure the deck's fonts, colors and slide types, and the model picks the chart's styling from that summary.

## How to use it

1. Upload your .pptx file.
2. For each chart, pick a chart type, a slide number and a data file (CSV or Excel).
3. Download the updated deck.

Supported chart types: column, bar, line and area (each with stacked and 100% stacked versions), pie, donut and scatter.

## Stack

- Backend: Python, FastAPI, pandas, python-pptx, Claude API. Deployed on Render.
- Frontend: Next.js. Deployed on Vercel.

## Run it locally

Backend:

```
cd backend
pip install -r requirements.txt
export ANTHROPIC_API_KEY=your_key
uvicorn main:app --reload
```

`CLAUDE_MODEL` is optional and defaults to `claude-sonnet-5-5`.

Frontend:

```
cd frontend
echo "NEXT_PUBLIC_BACKEND_URL=http://localhost:8000" > .env.local
pnpm install
pnpm dev
```

## API

`POST /process` takes a multipart form with:

- `pptxFile`: the deck
- `dataFiles`: one data file per chart, in the same order as the instructions
- `instructions`: a JSON list like `[{"chart_type": "line", "slide_number": 3}]`

It returns the edited .pptx.

## Project structure

```
backend/    FastAPI service (main.py routes, utils.py chart and theme logic)
frontend/   Next.js app
```
