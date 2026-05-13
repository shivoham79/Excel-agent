# Excel Agent

A conversational AI agent that reads and writes Excel files using natural language. Built with **LangChain**, **LangGraph**, and **OpenAI GPT-4o**.

## Features

- List all available Excel files
- List worksheets in any file
- Read any cell range from any file
- Update/overwrite cells
- Append new rows
- Understands natural language (e.g., "Show me January orders above ₹500")

## Project Structure

```
Excel-agent/
├── .env.example          # Environment variable template
├── .env                  # Your actual config (not committed)
├── .gitignore
├── requirements.txt      # Python dependencies
├── graph_client.py       # Excel I/O layer (openpyxl)
├── tools.py              # LangChain tools (actions the agent can take)
├── agent.py              # LangGraph ReAct agent definition
├── main.py               # Entry point - interactive chat loop
└── data/                 # Excel files directory
    ├── invoice_jan_2026.xlsx
    ├── invoice_feb_2026.xlsx
    ├── invoice_mar_2026.xlsx
    └── invoice_apr_2026.xlsx
```

## Prerequisites

- Python 3.10+
- OpenAI API key (get one at https://platform.openai.com/api-keys)

## Setup & Run

### 1. Clone the repo

```bash
git clone <your-repo-url>
cd Excel-agent
```

### 2. Create virtual environment

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS/Linux
source .venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Configure environment

```bash
copy .env.example .env
```

Edit `.env` and add your OpenAI API key:

```
OPENAI_API_KEY=sk-your-key-here
EXCEL_DIR=./data
```

### 5. Run the agent

```bash
python main.py
```

You'll see:

```
Excel Agent ready. Type 'quit' to exit.

You: 
```

## Example Conversations

```
You: What files are available?
Agent: Available files: invoice_jan_2026.xlsx, invoice_feb_2026.xlsx, invoice_mar_2026.xlsx, invoice_apr_2026.xlsx

You: Show me the first 5 orders from January invoice
Agent: Here are the first 5 orders from invoice_jan_2026.xlsx...

You: Update the price of order ORD-202601-0003 to 450.00
Agent: Successfully updated the price in invoice_jan_2026.xlsx → Invoices!E4

You: Add a new order to the March invoice: ORD-202603-0031, Webcam HD, 2026-03-28, 2026-04-02, 59.99
Agent: Appended 1 row to invoice_mar_2026.xlsx → Invoices

You: quit
Bye!
```

## How It Works

```
User (natural language)
    → main.py (chat loop)
        → agent.py (GPT-4o decides which tool to call)
            → tools.py (executes the chosen tool)
                → graph_client.py (reads/writes .xlsx via openpyxl)
                    → data/*.xlsx files
```

1. **User** types a request in plain English
2. **LangGraph ReAct Agent** (GPT-4o) reasons about what to do and picks a tool
3. **Tool** executes the Excel operation via openpyxl
4. **Agent** observes the result and responds to the user

## Available Tools

| Tool | Description |
|------|-------------|
| `list_files` | List all .xlsx files in the data directory |
| `list_sheets` | List worksheet names in a specific file |
| `read_excel` | Read a cell range (e.g., A1:E10) |
| `write_excel` | Overwrite a cell range with new values |
| `append_rows` | Add new rows at the bottom of a sheet |

## Using Your Own Excel Files

Drop any `.xlsx` file into the `data/` folder and the agent can immediately access it. Just ask:

```
You: What files are available?
```

## Dependencies

- `langchain` — Tool definitions and LLM abstractions
- `langgraph` — ReAct agent orchestration
- `langchain-openai` — OpenAI ChatGPT integration
- `openpyxl` — Read/write Excel .xlsx files
- `python-dotenv` — Load environment variables from .env
