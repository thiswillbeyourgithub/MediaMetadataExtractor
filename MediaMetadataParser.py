#!/usr/bin/env python3
"""Media Metadata Extractor

This script extracts metadata from media files including:
- Duration
- Resolution (width x height)
- FPS (frames per second)
- Codec information
- File size
- Creation date (if available in EXIF)

The results are saved to an Excel file with each file's metadata in a row.

Dependencies:
- moviepy: For basic media metadata
- openpyxl: For Excel file creation
- exifread: For EXIF data extraction (optional)
"""

import os
import json
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import warnings
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font
from moviepy.video.io.VideoFileClip import VideoFileClip
from mutagen import File
from mutagen.mp4 import MP4
from mutagen.mp3 import MP3
from mutagen.flac import FLAC

MEDIA_EXTENSIONS = {'.mp3', '.mp4', '.avi', '.mkv', '.mov', '.wav', '.flac', '.m4a', '.aac'}

def get_media_metadata(file_path: Path) -> Dict[str, str]:
    """Extract metadata from a media file.
    
    Args:
        file_path: Path to the media file
        
    Returns:
        Dictionary containing extracted metadata
    """
    metadata = {
        'filename': file_path.name,
        'path': str(file_path),
        'size_B': file_path.stat().st_size,
        'size_MB': f"{file_path.stat().st_size / (1024 * 1024):.2f}",
        'modified_unix': file_path.stat().st_mtime,
        'modified_date': datetime.fromtimestamp(file_path.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
    }
    
    try:
        # Suppress warnings during metadata extraction
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            with VideoFileClip(str(file_path)) as clip:
                hours = clip.duration//60//60
                minutes = clip.duration//60 - hours * 60
                seconds = clip.duration - hours * 60 * 60 - minutes * 60
                metadata.update({
                    'duration_seconds': f"{clip.duration:.2f}",
                    'duration': f"{hours}H{minutes}::{seconds:.2f}",
                    'resolution': f"{clip.size[0]}x{clip.size[1]}",
                    'fps': f"{clip.fps:.2f}" if hasattr(clip, 'fps') else 'N/A',
                    'codec': clip.reader.codec if hasattr(clip.reader, 'codec') else 'N/A',
                    'pixel_format': clip.reader.pixel_format if hasattr(clip.reader, 'pixel_format') else 'N/A',
                    'depth': clip.reader.depth if hasattr(clip.reader, 'depth') else 'N/A',
                    'rotation': clip.reader.rotation if hasattr(clip.reader, 'rotation') else 'N/A',
                    'bitrate': clip.reader.bitrate if hasattr(clip.reader, 'bitrate') else 'N/A',
                    'extra_infos': clip.reader.infos if hasattr(clip.reader, 'infos') else 'N/A',
                })
    except Exception as e:
        metadata['error'] = str(e)
    
    # Extract additional metadata using mutagen
    try:
        audio = File(str(file_path))
        if audio:
            metadata.update({
                'title': audio.get('title', ['N/A'])[0],
                'artist': audio.get('artist', ['N/A'])[0],
                'album': audio.get('album', ['N/A'])[0],
                'genre': audio.get('genre', ['N/A'])[0],
                'track_number': audio.get('tracknumber', ['N/A'])[0],
                'year': audio.get('date', ['N/A'])[0],
                'bitrate': getattr(audio.info, 'bitrate', 'N/A'),
                'sample_rate': getattr(audio.info, 'sample_rate', 'N/A'),
                'channels': getattr(audio.info, 'channels', 'N/A'),
            })
            
            # Handle MP4 specific tags
            if file_path.suffix.lower() in {'.mp4', '.m4a'}:
                mp4_tags = audio.tags
                if mp4_tags:
                    metadata.update({
                        'composer': mp4_tags.get('\xa9wrt', ['N/A'])[0],
                        'album_artist': mp4_tags.get('aART', ['N/A'])[0],
                        'grouping': mp4_tags.get('\xa9grp', ['N/A'])[0],
                        'lyrics': mp4_tags.get('\xa9lyr', ['N/A'])[0],
                    })
    except Exception as e:
        metadata['audio_metadata_error'] = str(e)
    
    return metadata

def save_to_excel(data: List[Dict[str, str]], output_path: Path) -> None:
    """Save metadata to an Excel file.
    
    Args:
        data: List of metadata dictionaries
        output_path: Path to save the Excel file
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Media Metadata"
    
    # Create header row
    headers = [
        'Filename', 'Path', 'Size', 'Modified Date', 
        'Duration', 'Resolution', 'FPS', 'Codec',
        'Title', 'Artist', 'Album', 'Genre', 'Track Number',
        'Year', 'Bitrate', 'Sample Rate', 'Channels',
        'Composer', 'Album Artist', 'Grouping', 'Lyrics'
    ]
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = Font(bold=True)
    
    # Add data rows
    for row_num, item in enumerate(data, 2):
        ws.cell(row=row_num, column=1, value=item['filename'])
        ws.cell(row=row_num, column=2, value=item['path'])
        ws.cell(row=row_num, column=3, value=item['size'])
        ws.cell(row=row_num, column=4, value=item['modified'])
        ws.cell(row=row_num, column=5, value=item.get('duration', 'N/A'))
        ws.cell(row=row_num, column=6, value=item.get('resolution', 'N/A'))
        ws.cell(row=row_num, column=7, value=item.get('fps', 'N/A'))
        ws.cell(row=row_num, column=8, value=item.get('codec', 'N/A'))
        
        # Audio metadata
        ws.cell(row=row_num, column=9, value=item.get('title', 'N/A'))
        ws.cell(row=row_num, column=10, value=item.get('artist', 'N/A'))
        ws.cell(row=row_num, column=11, value=item.get('album', 'N/A'))
        ws.cell(row=row_num, column=12, value=item.get('genre', 'N/A'))
        ws.cell(row=row_num, column=13, value=item.get('track_number', 'N/A'))
        ws.cell(row=row_num, column=14, value=item.get('year', 'N/A'))
        ws.cell(row=row_num, column=15, value=item.get('bitrate', 'N/A'))
        ws.cell(row=row_num, column=16, value=item.get('sample_rate', 'N/A'))
        ws.cell(row=row_num, column=17, value=item.get('channels', 'N/A'))
        ws.cell(row=row_num, column=18, value=item.get('composer', 'N/A'))
        ws.cell(row=row_num, column=19, value=item.get('album_artist', 'N/A'))
        ws.cell(row=row_num, column=20, value=item.get('grouping', 'N/A'))
        ws.cell(row=row_num, column=21, value=item.get('lyrics', 'N/A'))
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column[0].column_letter].width = adjusted_width
    
    wb.save(output_path)

def process_directory(directory: Path, output_file: Path) -> None:
    """Process all media files in a directory and save metadata to Excel.
    
    Args:
        directory: Path to directory containing media files
        output_file: Path to save the Excel file
    """
    if not directory.exists():
        raise FileNotFoundError(f"Directory not found: {directory}")
    
    media_files = [
        f for f in directory.rglob('*') 
        if f.suffix.lower() in MEDIA_EXTENSIONS and not f.name.startswith('.')
    ]
    
    if not media_files:
        raise ValueError(f"No supported media files found in {directory}")
    
    print(f"Found {len(media_files)} media files. Processing...")
    
    metadata_list = []
    for file in media_files:
        print(f"Processing: {file.name}")
        metadata = get_media_metadata(file)
        metadata_list.append(metadata)
    
    print(f"Saving results to {output_file}")
    save_to_excel(metadata_list, output_file)
    print("Processing complete!")

if __name__ == '__main__':
    import argparse
    
    parser = argparse.ArgumentParser(description="Extract media metadata and save to Excel")
    parser.add_argument('directory', type=str, help="Directory containing media files")
    parser.add_argument('-o', '--output', type=str, default="media_metadata.xlsx",
                       help="Output Excel file name (default: media_metadata.xlsx)")
    
    args = parser.parse_args()
    
    try:
        directory = Path(args.directory)
        output_file = Path(args.output)
        process_directory(directory, output_file)
    except Exception as e:
        print(f"Error: {str(e)}")
