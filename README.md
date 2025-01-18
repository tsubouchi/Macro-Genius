# Macro-Genius

An AI-powered macro generation service that converts natural language inputs into Excel-compatible macros.

## Features

- Natural language to Excel macro conversion using OpenAI GPT-4
- Template-based macro generation
- Macro categorization and management
- Excel macro export functionality
- Category-based macro organization
- Version control for macros
- Macro sharing capabilities

## Technical Stack

- FastAPI backend
- OpenAI integration for AI-powered macro generation
- PostgreSQL database
- Python macro generation
- Excel macro export functionality

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Set up environment variables:
- Create a `.env` file with the following variables:
  - `OPENAI_API_KEY`: Your OpenAI API key
  - `DATABASE_URL`: PostgreSQL database URL

3. Run the application:
```bash
uvicorn main:app --host 0.0.0.0 --port 8000
```

## Usage

1. Access the web interface
2. Enter your macro requirements in natural language
3. The AI will generate an appropriate Excel macro
4. Download the generated macro in Excel format

## Note

Make sure to keep your API keys and sensitive information secure and never commit them to version control.
