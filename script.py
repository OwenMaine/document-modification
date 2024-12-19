#!/usr/bin/env python3

import sys
import os
from docx import Document
from pathlib import Path
import datetime

def replace_heading_text(folder_path, replacements):
    # Create a counter for modified files
    modified_count = 0
    
    # Print current time and user information
    print(f"Starting process at: {datetime.datetime.now()}")
    print(f"Processing files in: {folder_path}")
    print("-" * 50)
    
    # Ensure the folder path exists
    if not os.path.exists(folder_path):
        print(f"Error: The folder '{folder_path}' does not exist.")
        print("\nPlease make sure you have created the folder in:")
        print("/Users/user/Downloads/")
        return 0
    
    # Get all .docx files in the folder
    docx_files = list(Path(folder_path).glob('*.docx'))
    
    if not docx_files:
        print("No .docx files found in the specified folder.")
        return 0
    
    print(f"Found {len(docx_files)} .docx files:")
    for i, file in enumerate(docx_files, 1):
        print(f"{i}. {file.name}")
    print("-" * 50)
    
    for file_path in docx_files:
        try:
            print(f"\nProcessing: {file_path.name}")
            doc = Document(file_path)
            modified = False
            
            for paragraph in doc.paragraphs:
                original_text = paragraph.text
                new_text = original_text
                
                # Apply all replacements
                for old_text, new_text_replacement in replacements.items():
                    if old_text in new_text:
                        new_text = new_text.replace(old_text, new_text_replacement)
                
                # If any replacement was made
                if new_text != original_text:
                    print(f"Found target text in {file_path.name}")
                    paragraph.text = new_text
                    modified = True
            
            if modified:
                # Create backup filename
                backup_path = file_path.with_name(f"{file_path.stem}_backup{file_path.suffix}")
                print(f"Creating backup: {backup_path.name}")
                
                # Create backup of original file
                if not backup_path.exists():
                    os.rename(file_path, backup_path)
                
                # Save the modified document
                doc.save(file_path)
                modified_count += 1
                print(f"✓ Successfully modified: {file_path.name}")
            else:
                print(f"No changes needed in: {file_path.name}")
            
        except Exception as e:
            print(f"Error processing {file_path.name}: {str(e)}")
    
    return modified_count

def process_folders(folders):
    total_modified = 0
    for folder in folders:
        print(f"\nProcessing folder: {folder}")
        print("=" * 50)
        
        modified = replace_heading_text(
            folder,
            {
                "MEMORANDUM OF CUSTOMARY SALE OF LAND": "MEMORANDUM OF UNDERSTANDING",
                "Memorandum of Sale": "MEMORANDUM OF UNDERSTANDING",
                # Add more replacements here if needed
            }
        )
        total_modified += modified
        print(f"Completed processing folder: {folder}")
        print("=" * 50)
    
    return total_modified

if __name__ == "__main__":
    # Set the correct paths for your system
    folders_to_process = [
        "/Users/user/Downloads/manch"
    ]
    
    print("Script starting...")
    print(f"Current Date and Time (UTC): {datetime.datetime.utcnow()}")
    print(f"Current User's Login: {os.getlogin()}")
    print("\nWill process the following folders:")
    for folder in folders_to_process:
        print(f"- {folder}")
    print("-" * 50)
    
    total_modified = process_folders(folders_to_process)
    
    print("\nFinal Summary:")
    print("-" * 50)
    print(f"Total files modified across all folders: {total_modified}")
    print("Backup files were created for all modified documents with '_backup' suffix")
    print("\nScript completed successfully!")