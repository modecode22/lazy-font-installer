import os
import shutil
import subprocess
from pathlib import Path
import ctypes
import sys
import argparse
import zipfile
import tempfile

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def find_font_files(directory):
    font_extensions = ['.ttf', '.otf', '.woff', '.woff2']
    font_files = []
    
    for root, dirs, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            if file.lower().endswith('.zip'):
                font_files.extend(find_fonts_in_zip(file_path))
            elif any(file.lower().endswith(ext) for ext in font_extensions):
                font_files.append(file_path)
    
    return font_files

def find_fonts_in_zip(zip_path):
    font_extensions = ['.ttf', '.otf', '.woff', '.woff2']
    font_files = []
    
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        for file in zip_ref.namelist():
            if any(file.lower().endswith(ext) for ext in font_extensions):
                # Extract font file to a temporary directory
                with tempfile.TemporaryDirectory() as tmpdirname:
                    zip_ref.extract(file, tmpdirname)
                    font_files.append(os.path.join(tmpdirname, file))
    
    return font_files

def install_font(font_path, dry_run=False):
    if os.name == 'nt':  # Windows
        font_dir = Path(os.environ['WINDIR']) / 'Fonts'
        
        try:
            font_name = os.path.basename(font_path)
            if dry_run:
                print(f"Would install: {font_name}")
                return True
            
            # Copy the font file to the Fonts directory
            shutil.copy(font_path, font_dir)
            
            # Register the font
            subprocess.run(['reg', 'add', r'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts', '/v', font_name, '/t', 'REG_SZ', '/d', font_name, '/f'], check=True)
            
            print(f"Installed: {font_name}")
            return True
        except PermissionError:
            print(f"Permission denied: {font_name}. Try running the script as administrator.")
        except subprocess.CalledProcessError:
            print(f"Failed to register font: {font_name}")
        except Exception as e:
            print(f"Error installing {font_name}: {str(e)}")
    elif os.name == 'posix':  # macOS or Linux
        if os.uname().sysname == 'Darwin':  # macOS
            font_dir = Path.home() / 'Library' / 'Fonts'
        else:  # Linux
            font_dir = Path.home() / '.local' / 'share' / 'fonts'
        
        try:
            font_name = os.path.basename(font_path)
            if dry_run:
                print(f"Would install: {font_name}")
                return True
            
            # Create the font directory if it doesn't exist
            font_dir.mkdir(parents=True, exist_ok=True)
            
            # Copy the font file to the appropriate directory
            shutil.copy(font_path, font_dir)
            
            print(f"Installed: {font_name}")
            return True
        except Exception as e:
            print(f"Error installing {font_name}: {str(e)}")
    else:
        print(f"Unsupported operating system: {os.name}")
    
    return False

def main():
    parser = argparse.ArgumentParser(description="Font Installer CLI Tool")
    parser.add_argument('-d', '--directory', type=str, default=os.getcwd(),
                        help='Directory to search for font files (default: current directory)')
    parser.add_argument('--dry-run', action='store_true',
                        help='Perform a dry run without actually installing fonts')
    parser.add_argument('-v', '--verbose', action='store_true',
                        help='Enable verbose output')
    args = parser.parse_args()

    if os.name == 'nt' and not is_admin() and not args.dry_run:
        print("This script requires administrator privileges. Please run it as an administrator.")
        print("Right-click on the script and select 'Run as administrator'.")
        return

    font_files = find_font_files(args.directory)
    
    if not font_files:
        print(f"No font files found in the specified directory: {args.directory}")
        return
    
    print(f"Found {len(font_files)} font files.")
    
    installed_count = 0
    skipped_count = 0
    
    for font_file in font_files:
        if args.verbose:
            print(f"Processing: {font_file}")
        if install_font(font_file, args.dry_run):
            installed_count += 1
        else:
            skipped_count += 1
    
    print(f"\nInstallation {'simulation ' if args.dry_run else ''}complete.")
    print(f"{'Would install' if args.dry_run else 'Installed'}: {installed_count}")
    print(f"Skipped/Failed: {skipped_count}")

if __name__ == "__main__":
    main()