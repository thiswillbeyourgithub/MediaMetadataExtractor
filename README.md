# MediaMetadataParser

![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)
![License](https://img.shields.io/badge/license-GPLv3-green.svg)

MediaMetadataParser is a powerful tool for extracting and organizing metadata from media files. It supports various media formats and provides detailed information in both Excel and JSON formats.

## Features

- Extracts comprehensive metadata including:
  - Duration
  - Resolution (width x height)
  - FPS (frames per second)
  - Codec information
  - File size
  - Creation date (if available in EXIF)
- Supports multiple media formats:
  - Video: .mp4, .avi, .mkv, .mov
  - Audio: .mp3, .wav, .flac, .m4a, .aac
- Recursively scans directories
- Excludes hidden files (those starting with '.')
- Provides:
  - Total number of media files
  - Total size in GB
  - Detailed metadata for each file
  - Results saved to Excel file
  - Optional JSON output
  - Remembers last used directory
  - Progress tracking and cancellation

## Installation

1. Clone the repository:
```bash
git clone https://github.com/thiswillbeyourgithub/MediaMetadataParser.git
cd MediaMetadataParser
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Run the script to launch the GUI application:
```bash
python MediaMetadataParser.py
```

1. Select a folder containing media files
2. Choose an output location
3. Click 'Start Processing' to begin metadata extraction

The application will:
- Scan the selected directory
- Extract metadata from all supported media files
- Save results to an Excel file
- Optionally save results in JSON format

## Requirements

- Python 3.8+
- Required packages:
  - moviepy
  - openpyxl
  - tkinter

## Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository
2. Create a new branch (`git checkout -b feature/YourFeatureName`)
3. Commit your changes (`git commit -m 'Add some feature'`)
4. Push to the branch (`git push origin feature/YourFeatureName`)
5. Create a new Pull Request

## License

This project is licensed under the GPLv3 License - see the [LICENSE](LICENSE) file for details.

## Support

If you find this project useful, please consider starring the repository ⭐

For issues or feature requests, please open an issue on GitHub.

## Screenshots

![Main Window](screenshots/main_window.png)
*Main application window*

![Processing](screenshots/processing.png)
*Metadata extraction in progress*

![Results](screenshots/results.png)
*Example Excel output*
