# ollama-excel
Autocompletion for Excel using Ollama (local LLM)

# Setup
- Download Ollama from https://ollama.com/
- Install Ollama
- Run in terminal `ollama run gemma3`
- In Excel, open a spreadsheet and press `Alt+F11`
- Paste the content from `ollamaExcel.txt`, from this repository
- Click the save button

# Usage example
- Enter in formula input: `=OllamaExcel("What is the capital of this country?", A1)`
- First parameter is your question for the LLM
- Second parameter is an optional cell to append to prompt

![ollama-excel](ollama-excel.gif)
