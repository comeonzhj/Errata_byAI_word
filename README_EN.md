# AI Errata Tool for Word

English | [简体中文](README.md)

An AI-powered Word document proofreading tool that automatically checks and annotates typos while providing correction suggestions.

Created by Trae,prompt @ [How-to-prompt](How-to-prompt.md)

## Features

- 🔍 Automatic typo detection in Word documents
- 📝 Intelligent correction suggestions
- 🎯 Highlighted areas requiring modification
- 📊 Real-time progress display
- 📋 Error log export support
- ⚙️ Flexible API configuration

## Requirements

- Python 3.6+
- macOS (currently macOS only)
- Python packages:
  - PyQt6 >= 6.4.0
  - python-docx >= 0.8.11
  - requests >= 2.28.0
  - json5 >= 0.9.5
  - pyinstaller >= 5.13.0 (for packaging only)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/Errata_byAI_word.git
cd Errata_byAI_word
```

2. Install dependencies:
```bash
pip3 install -r requirements.txt
```

## Usage

### Running from Source

```bash
python3 main.py
```

### Building the Application

```bash
pyinstaller main.spec
```
The packaged application will be generated in the `dist` directory.

### Configuration

1. Click the "Model Configuration" button
2. Fill in the following information:
   - API URL: Your API endpoint
   - API Key: Your API key
   - Model Name: The model to use
   - System Prompt: System instructions (default provided)

### Workflow

1. Launch the application
2. Complete API configuration
3. Click "Open Document" to select a Word file
4. Click "Start Proofreading"
5. Wait for processing to complete
6. Review the new document with correction suggestions
7. Optional: Export error log

## Important Notes

- API keys and sensitive information are stored in `.errata_word/config.json` in your home directory
- Supported document format: .docx
- Ensure stable network connection when processing large documents

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Feel free to submit Issues and Pull Requests.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## Contact

For questions or suggestions, please open an Issue.
