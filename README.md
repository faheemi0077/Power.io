# PowerIO – AI Powered PowerPoint Chart Generator

PowerIO is an API that automatically generates charts in PowerPoint slides from raw data files.

The system accepts a PowerPoint template and one or more data files (CSV or Excel), interprets the data using a language model, and inserts properly formatted charts into the specified slides.

The goal is to automate the process of turning raw data into presentation-ready charts.

## Features

- PowerPoint chart generation via API
- Support for CSV and Excel data files
- Automatic data interpretation using an LLM
- Chart insertion into existing slide templates
- Chart styling that adapts to the presentation theme
- REST API built with FastAPI

## How It Works

1. A PowerPoint template is uploaded.
2. Data files are uploaded (CSV or Excel).
3. Instructions specify:
   - chart type
   - slide number
4. The system loads the data into pandas DataFrames.
5. The data is interpreted and formatted for charting.
6. A chart is generated and inserted into the PowerPoint slide.
7. The modified PowerPoint file is returned to the user.

## API Endpoint

POST `/process`

Request includes:

- PowerPoint template file
- One or more data files
- Chart instructions
- API key authentication

Response:

- A generated PowerPoint file with charts inserted.

## Technologies Used

- Python
- FastAPI
- pandas
- python-pptx
- Mistral AI API
