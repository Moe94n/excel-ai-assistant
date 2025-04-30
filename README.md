# Excel AI Assistant

A powerful desktop application that transforms your Excel and CSV data using AI, with support for both OpenAI's cloud API and local Ollama models.

![Excel AI Assistant Main Window](docs/screenshots/main_window.png)

## Overview

Excel AI Assistant is a Python-based desktop application that helps you apply intelligent transformations to your spreadsheet data. It seamlessly connects with OpenAI's powerful language models or your local Ollama open-source models to provide AI-driven data manipulation, cleaning, and analysis capabilities.

## Key Features

- **Load Excel and CSV files**: Easily open and manipulate your spreadsheet data
- **Intelligent Transformations**: Apply AI-powered transformations to cells and ranges
- **Dual AI Backend Options**:
  - **OpenAI API**: Use cloud-based GPT models for high-quality results
  - **Ollama Integration**: Process data locally with open-source models for privacy and cost savings
- **Customizable Templates**: Save your favorite prompts for quick reuse
- **Batch Processing**: Apply transformations to multiple cells efficiently
- **Interactive Data View**: View, edit, and copy cell data with an intuitive interface
- **Dark/Light Mode**: Choose your preferred theme or match your system theme

## Screenshots

| Main Interface | OpenAI Settings |
|:---:|:---:|
| ![Main Interface](docs/screenshots/main_interface.png) | ![OpenAI Settings](docs/screenshots/openai_settings.png) |

| Prompt Templates | Ollama Integration |
|:---:|:---:|
| ![Prompt Templates](docs/screenshots/prompt_templates.png) | ![Ollama Integration](docs/screenshots/ollama_settings.png) |

## Installation

### Prerequisites

- Python 3.13 or higher
- `uv` package manager (installation instructions below)
- For OpenAI API: An OpenAI API key
- For Ollama (optional): [Ollama](https://ollama.ai) installed with at least one model

### Installing uv

`uv` is a fast Python package installer and resolver built in Rust:

**macOS/Linux**:
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

**Windows**:
```powershell
iwr -useb https://astral.sh/uv/install.ps1 | iex
```

### Step-by-Step Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/georgekhananaev/excel-ai-assistant.git
   cd excel-ai-assistant
   ```

2. Create a virtual environment and install dependencies:
   ```bash
   # Create and activate a virtual environment
   uv venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate

   # Install dependencies using uv sync
   uv sync
   ```

3. Set up your OpenAI API key (optional):
   - Create a `.env` file in the project root directory
   - Add your API key: `OPENAI_API_KEY=your_api_key_here`

4. Set up Ollama (optional for local processing):
   - Install Ollama from [ollama.ai/download](https://ollama.ai/download)
   - Start the Ollama server: `ollama serve`
   - Pull a model: `ollama pull llama3.2` (or any other model you prefer) see here: [Ollama GitHub](https://github.com/ollama/ollama)

## Usage

### Starting the Application

Run the application using uv:

```bash
uv run main.py
```

Or if you've activated the virtual environment:
```bash
python main.py
```

### Basic Workflow

1. **Open a File**: Click "Open" in the toolbar or File menu to load an Excel or CSV file
2. **Select API**: Choose between OpenAI or Ollama for processing
3. **Set Range**: Specify the row range and columns to process
4. **Configure Prompt**: Enter instructions or select a template from the dropdown
5. **Process Data**: Click "Run on Selected Range" to apply the transformation


### Using Templates

The application comes with many built-in prompt templates for common transformations:

- Text formatting (capitalization, case conversion)
- Data formatting (dates, phone numbers, currency)
- Content transformation (summarization, grammar fixing)
- Code and markup handling (JSON, XML, CSV formatting)
- Multi-language translation
- Data cleanup operations

You can also create, save, and manage your own custom templates.

### Using Local Models with Ollama

For private data processing or offline use, Excel AI Assistant supports local models through Ollama:

1. In the toolbar, select "Ollama" from the API dropdown
2. Make sure Ollama is running (`ollama serve` in your terminal)
3. Use the "Test API" button to verify the connection
4. Select your preferred model from the dropdown
5. Process your data with complete privacy - nothing leaves your computer!

## Configuration

Access settings through Edit > Preferences to configure:

- API settings (OpenAI key, model, Ollama URL)
- Interface appearance (theme, fonts, table display)
- Processing options (batch size, temperature, tokens)
- Advanced settings (logging, performance)


The `design/` folder contains React mockups of the application's interface. These mockups serve two purposes:

1. **Visual Documentation**: They provide visual references for the application's UI and functionality
2. **Web Application Foundation**: They can be used as a starting point if you want to create a web version of this application

## Troubleshooting

### Connection Issues
- For OpenAI: Verify your API key is correct and check your internet connection
- For Ollama: Ensure Ollama is running (`ollama serve` in terminal) and check the URL (default: http://localhost:11434)

### Installation Issues
- If you see an error about "Multiple top-level packages discovered", use `uv sync` instead of `uv pip install -e .`

### Performance Tips
- For large files, increase the batch size in preferences
- If using Ollama, smaller models are faster but may be less accurate
- Enable multi-threading in Advanced settings for better performance

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.


## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Built with Python, Tkinter, and Pandas
- Uses OpenAI's API for cloud processing
- Integrates with Ollama for local AI processing
- Inspired by the need to bring AI capabilities to everyday data tasks

## Quick Links

- [Mockups Design Gallery](docs/mockups_gallery.md)
- [Ollama Setup](docs/ollama_setup.md)