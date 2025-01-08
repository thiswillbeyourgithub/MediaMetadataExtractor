# MediaMetadataExtractor

> Une version française de ce fichier README est disponible : [README_fr.md](README_fr.md)  
> A French version of this README is available: [README_fr.md](README_fr.md)

![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)
![License](https://img.shields.io/badge/license-GPLv3-green.svg)

## Motivation

This project was created to help my significant other who works in video production and needed a quicker way to access metadata from large collections of dailies, rush footage, and other video files. The existing tools were either too slow, didn't provide the right information, or required manual processing of each file. MediaMetadataExtractor was designed to:

- Quickly scan entire directories of media files
- Extract all relevant technical metadata in one go
- Provide an organized, searchable output
- Handle the specific needs of video production workflows

MediaMetadataExtractor is a powerful tool for extracting and organizing metadata from media files. It supports various media formats and provides detailed information in both Excel and JSON formats.

## Features

- Extracts comprehensive metadata including:
  - Duration (in seconds and HH:MM:SS format)
  - Resolution (width x height)
  - FPS (frames per second)
  - Codec information
  - Pixel format
  - Bit depth
  - Rotation
  - Bitrate
  - Color space
  - File size (in bytes and MB)
  - Modification date
  - Additional technical metadata from file headers
- Supports multiple media formats:
  - Video: .mp4, .avi, .mkv, .mov
  - Audio: .mp3, .wav, .flac, .m4a, .aac
  - Supported formats: .mp3, .mp4, .avi, .mkv, .mov, .wav, .flac, .m4a, .aac
- Recursively scans directories
- Excludes hidden files (those starting with '.')
- Provides:
  - Total number of media files
  - Total size in GB
  - Detailed metadata for each file
  - Results saved to Excel file with auto-adjusted column widths
  - Optional JSON output with proper type conversion
  - Remembers last used directory via temporary file
  - Progress tracking with percentage complete
  - Cancellation support
  - Error handling for problematic files
  - GUI with:
    - Expandable documentation
    - GitHub link
    - Style enhancements
    - Input validation

## Installation

1. Clone the repository:
```bash
git clone https://github.com/thiswillbeyourgithub/MediaMetadataExtractor.git
cd MediaMetadataExtractor
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Run the script to launch the GUI application:
```bash
python MediaMetadataExtractor.py
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
  - moviepy (for media metadata extraction)
  - openpyxl (for Excel file creation)
  - tkinter (for GUI interface)
  - json (for optional JSON output)

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

## Example Output

The application generates a detailed Excel spreadsheet with metadata for each media file, including:
- File name and path
- File size in bytes and MB
- Modification date
- Duration in seconds and HH:MM:SS format
- Resolution (for video files)
- FPS (for video files)
- Codec information
- Additional technical details
