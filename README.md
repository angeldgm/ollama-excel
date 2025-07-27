# ollama-excel
Autocompletion for Excel using Ollama (local LLM)

## Setup
- Download & install Ollama from https://ollama.com/
- Run in terminal `ollama run gemma3` (or use a different model if you edit the script)
- In Excel, open a spreadsheet and press `Alt+F11`
- Paste the content from `ollamaExcel.txt`, from this repository, and save changes

## Usage example
- Enter in formula input: `=OllamaExcel("What is the capital of this country?", A1)`
- First parameter is your question for the LLM
- Second parameter is an optional cell to append to prompt

![ollama-excel](ollama-excel.gif)
