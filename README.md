# ollama-excel
Data autocompletion for Excel using Ollama (local LLM). Unlimited, free and fully private

## Setup
- Download & install Ollama from https://ollama.com/
- Run in terminal `ollama run gemma3` (or use a different model if you edit the script)
- In Excel, open a spreadsheet and press `Alt+F11`
- Paste the content from `ollamaExcel.txt`, from this repository, and save changes

## Plugin
- If you want to make the script persist when you open a new Excel file, save the script as an Add-In, then enable it.

## Usage example
- Enter in formula input: `=OllamaExcel("Who is the author of this book?", A1)`
- First parameter is your question for the LLM
- Second parameter is optional and appends a cell content to prompt

![ollama-excel](ollama-excel.gif)
