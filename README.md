# Composite Analyzer

A Streamlit application for composite drillhole data using Datamine StudioRM.

## Features

- Upload and process Datamine format files
- Create composites with custom size
- Analyze length and grade distributions
- Compare statistics between raw and composite samples
- Interactive visualizations

## Requirements

- Python 3.8+
- Streamlit
- Pandas
- Matplotlib
- Datamine StudioRM (for full functionality)

## Installation

1. Clone this repository:
```bash
git clone Composite-Analyzer
cd Composite-Drillholes-Dashboard
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Local Development:
```bash
streamlit run app.py
```

2. Demo Mode (without Datamine StudioRM):
```bash
set DEMO_MODE=true
streamlit run app.py
```

## Deployment

This application can be deployed to Streamlit Cloud:

1. Push your code to GitHub
2. Visit [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repository
4. Set environment variable `DEMO_MODE=true`
5. Deploy!
