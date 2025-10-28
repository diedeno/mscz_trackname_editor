#!/usr/bin/python3

"""
MuseScore Track Name Editor
--------------------------
Standalone utility to modify trackName in MuseScore .mscz files.
This helps fix concatenation issues by standardizing part names across files and allows changing the ytrack in the mixer.

Copyright (c) 2025 Diego Denolf (graffesmusic)

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program. If not, see <https://www.gnu.org/licenses/>.
"""



import zipfile
import xml.etree.ElementTree as ET
import tempfile
import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Application Info
VERSION = "1.0"
LAST_MODIFIED = "2025-10-28"
LICENSE = "GPLv3"

class PartNameEditor:
    def __init__(self, root):
        self.root = root
        self.root.title(f"MuseScore Track Name Editor {VERSION}" )
        self.root.minsize(600, 500)
        
        # Main container
        main_frame = tk.Frame(root, padx=10, pady=10)
        main_frame.pack(fill="both", expand=True)
        
        # File selection
        file_frame = tk.LabelFrame(main_frame, text="MuseScore File", padx=10, pady=5)
        file_frame.pack(fill="x", pady=5)
        
        file_btn_frame = tk.Frame(file_frame)
        file_btn_frame.pack(fill="x", pady=5)
        
        tk.Button(file_btn_frame, text="Open .mscz File", command=self.open_file, 
                 width=15).pack(side="left", padx=5)
        tk.Button(file_btn_frame, text="Reload", command=self.load_parts,
                 width=10).pack(side="left", padx=5)
        
        self.file_label = tk.Label(file_frame, text="No file selected", fg="gray", wraplength=500)
        self.file_label.pack(pady=5)
        
        # Parts list
        parts_frame = tk.LabelFrame(main_frame, text="Parts - Double-click to edit", padx=10, pady=5)
        parts_frame.pack(fill="both", expand=True, pady=5)
        
        # Treeview for parts with scrollbar
        tree_frame = tk.Frame(parts_frame)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        columns = ("Part ID", "Current Name", "New Name")
        self.parts_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=8)
        
        for col in columns:
            self.parts_tree.heading(col, text=col)
            self.parts_tree.column(col, width=150)
        
        # Add scrollbar
        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.parts_tree.yview)
        self.parts_tree.configure(yscrollcommand=tree_scroll.set)
        
        self.parts_tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")
        
        # Bind double-click to edit
        self.parts_tree.bind("<Double-1>", self.on_double_click)
        
        # Quick actions frame
        actions_frame = tk.Frame(parts_frame)
        actions_frame.pack(fill="x", pady=5)
        
        tk.Label(actions_frame, text="Quick actions:").pack(side="left", padx=5)
        tk.Button(actions_frame, text="Add numbers (Violin → Violin 1)", 
                 command=self.add_numbers).pack(side="left", padx=2)
        tk.Button(actions_frame, text="Reset all", 
                 command=self.reset_all).pack(side="left", padx=2)
        
        # Button frame at bottom
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)
        
        # Left side: Save button
        tk.Button(button_frame, text="Save Changes to File", command=self.save_file, 
                 bg="#4CAF50", fg="white", font=("Arial", 10, "bold"),
                 padx=20, pady=5).pack(side="left", padx=5)
        
        # Right side: Exit button
        tk.Button(button_frame, text="Exit", command=self.exit_app,
                 bg="#f44336", fg="white", font=("Arial", 10, "bold"),
                 padx=20, pady=5).pack(side="right", padx=5)
        
        tk.Label(main_frame, text="A backup will be created automatically when saving", 
                fg="gray", font=("Arial", 8)).pack(pady=2)
        
        self.current_file = None
        self.parts_data = []
        self.original_parts_data = []
        
        
        footer = ttk.Label(root, text=f"Version: {VERSION} | Last Modified: {LAST_MODIFIED} | License: {LICENSE}")
        footer.pack(side="bottom", pady=5)


    def exit_app(self):
        """Exit the application with confirmation if there are unsaved changes"""
        # Check if there are unsaved changes
        changes_made = any(
            part['new_name'] != part['current_name'] 
            for part in self.parts_data
        )
        
        if changes_made:
            response = messagebox.askyesnocancel(
                "Unsaved Changes", 
                "You have unsaved changes. Do you want to save before exiting?\n\n"
                "Yes: Save and exit\n"
                "No: Exit without saving\n"
                "Cancel: Return to application"
            )
            
            if response is None:  # Cancel
                return
            elif response:  # Yes - Save and exit
                self.save_file()
                # Don't exit immediately after save - let user see the success message
                return
        
        self.root.quit()
        
    def open_file(self):
        filename = filedialog.askopenfilename(
            title="Select MuseScore file",
            filetypes=[("MuseScore files", "*.mscz"), ("All files", "*.*")]
        )
        if filename:
            self.current_file = filename
            self.file_label.config(text=os.path.basename(filename))
            self.load_parts()
            
    def load_parts(self):
        """Extract part names from the .mscz file"""
        if not self.current_file:
            messagebox.showwarning("Warning", "No file selected")
            return
            
        try:
            with zipfile.ZipFile(self.current_file, 'r') as mscz:
                # Find the .mscx file inside
                mscx_files = [f for f in mscz.namelist() if f.endswith('.mscx')]
                if not mscx_files:
                    messagebox.showerror("Error", "No .mscx file found in the .mscz archive")
                    return
                    
                with mscz.open(mscx_files[0]) as mscx_file:
                    tree = ET.parse(mscx_file)
                    root = tree.getroot()
                    
                    # Clear existing data
                    for item in self.parts_tree.get_children():
                        self.parts_tree.delete(item)
                    self.parts_data = []
                    self.original_parts_data = []
                    
                    # Find all parts
                    for part in root.findall('.//Part'):
                        part_id = part.get('id', 'Unknown')
                        track_name_elem = part.find('trackName')
                        track_name = track_name_elem.text if track_name_elem is not None else "Unknown"
                        
                        # Store part data
                        part_data = {
                            'element': part,
                            'id': part_id,
                            'current_name': track_name,
                            'new_name': track_name
                        }
                        self.parts_data.append(part_data)
                        self.original_parts_data.append(part_data.copy())
                        
                        # Add to treeview
                        self.parts_tree.insert("", "end", values=(
                            part_id, track_name, track_name
                        ))
                        
            messagebox.showinfo("Success", f"Loaded {len(self.parts_data)} parts from file")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{str(e)}")
            
    def on_double_click(self, event):
        """Handle double-click to edit part name"""
        item = self.parts_tree.selection()
        if item:
            self.edit_part_name(item[0])
            
    def edit_part_name(self, item):
        """Edit the name of a specific part"""
        current_values = self.parts_tree.item(item, 'values')
        if not current_values:
            return
            
        # Create edit dialog
        edit_dialog = tk.Toplevel(self.root)
        edit_dialog.title("Edit Part Name")
        edit_dialog.geometry("300x150")
        edit_dialog.transient(self.root)
        
        # Make dialog modal without grab_set
        edit_dialog.focus_set()
        
        tk.Label(edit_dialog, text=f"Editing Part {current_values[0]}", 
                font=("Arial", 10, "bold")).pack(pady=10)
        
        tk.Label(edit_dialog, text="Current name:").pack()
        tk.Label(edit_dialog, text=current_values[1], fg="blue").pack()
        
        tk.Label(edit_dialog, text="New name:").pack(pady=(10, 0))
        new_name_var = tk.StringVar(value=current_values[2])
        name_entry = tk.Entry(edit_dialog, textvariable=new_name_var, width=30)
        name_entry.pack(pady=5)
        name_entry.select_range(0, tk.END)
        name_entry.focus()
        
        def apply_change():
            new_name = new_name_var.get().strip()
            if new_name:
                # Update treeview
                self.parts_tree.set(item, "New Name", new_name)
                # Update data
                index = self.parts_tree.index(item)
                self.parts_data[index]['new_name'] = new_name
            edit_dialog.destroy()
            
        def cancel_edit():
            edit_dialog.destroy()
            
        btn_frame = tk.Frame(edit_dialog)
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="Apply", command=apply_change, 
                 bg="#4CAF50", fg="white").pack(side="left", padx=5)
        tk.Button(btn_frame, text="Cancel", command=cancel_edit).pack(side="left", padx=5)
        
        # Bind Enter key to apply
        edit_dialog.bind('<Return>', lambda e: apply_change())

    def add_numbers(self):
        """Add numbers to parts with the same name"""
        name_count = {}
        name_first_occurrence = {}
        changes_made = False
        
        # First pass: count occurrences and track first occurrence
        for i, part_data in enumerate(self.parts_data):
            current_name = part_data['current_name']
            if current_name in name_count:
                name_count[current_name] += 1
            else:
                name_count[current_name] = 1
                name_first_occurrence[current_name] = i
        
        # Second pass: apply numbering
        for i, part_data in enumerate(self.parts_data):
            current_name = part_data['current_name']
            if name_count[current_name] > 1:  # Only number if there are duplicates
                # Find position in the sequence
                same_name_parts = [j for j, p in enumerate(self.parts_data) 
                                 if p['current_name'] == current_name]
                position = same_name_parts.index(i) + 1  # 1-based numbering
                
                new_name = f"{current_name} {position}"
                
                # Update data
                self.parts_data[i]['new_name'] = new_name
                # Update treeview
                item = self.parts_tree.get_children()[i]
                self.parts_tree.set(item, "New Name", new_name)
                changes_made = True
                    
        if changes_made:
            messagebox.showinfo("Success", "Added numbers to duplicate part names")
        else:
            messagebox.showinfo("Info", "No duplicate part names found")
            
    def reset_all(self):
        """Reset all names to original values"""
        if not messagebox.askyesno("Confirm", "Reset all part names to their original values?"):
            return
            
        for i, original_data in enumerate(self.original_parts_data):
            self.parts_data[i]['new_name'] = original_data['current_name']
            item = self.parts_tree.get_children()[i]
            self.parts_tree.set(item, "New Name", original_data['current_name'])
            
    def save_file(self):
        """Save changes back to the .mscz file"""
        if not self.current_file:
            messagebox.showwarning("Warning", "No file selected")
            return
            
        # Check if any changes were made
        changes_made = any(
            part['new_name'] != part['current_name'] 
            for part in self.parts_data
        )
        
        if not changes_made:
            messagebox.showinfo("Info", "No changes to save")
            return
            
        try:
            # Count changes BEFORE saving (so we can show the correct count)
            num_changes = sum(1 for p in self.parts_data if p['new_name'] != p['current_name'])
            
            # Create backup filename
            backup_file = self.current_file.replace('.mscz', '_backup.mscz')
            counter = 1
            while os.path.exists(backup_file):
                backup_file = self.current_file.replace('.mscz', f'_backup_{counter}.mscz')
                counter += 1
            
            # Create backup first
            shutil.copy2(self.current_file, backup_file)
            
            # Create temporary working directory
            temp_dir = tempfile.mkdtemp()
            
            try:
                with zipfile.ZipFile(self.current_file, 'r') as original:
                    with zipfile.ZipFile(os.path.join(temp_dir, "modified.mscz"), 'w') as modified:
                        # Copy all files, modifying the .mscx
                        for item in original.infolist():
                            if item.filename.endswith('.mscx'):
                                # Modify the .mscx file
                                with original.open(item.filename) as mscx_file:
                                    tree = ET.parse(mscx_file)
                                    root = tree.getroot()
                                    
                                    # Update part names
                                    for part in root.findall('.//Part'):
                                        part_id = part.get('id')
                                        for part_data in self.parts_data:
                                            if part_data['id'] == part_id and part_data['new_name'] != part_data['current_name']:
                                                track_name_elem = part.find('trackName')
                                                if track_name_elem is not None:
                                                    track_name_elem.text = part_data['new_name']
                                                    print(f"Changed part {part_id}: '{part_data['current_name']}' → '{part_data['new_name']}'")
                                    
                                    # Write modified content
                                    modified.writestr(item.filename, ET.tostring(root, encoding='unicode'))
                            else:
                                # Copy other files as-is
                                modified.writestr(item, original.read(item.filename))
                
                # Replace original file with modified one
                os.remove(self.current_file)
                shutil.move(os.path.join(temp_dir, "modified.mscz"), self.current_file)
                
                # Update current names to reflect the changes (AFTER showing the message)
                for part_data in self.parts_data:
                    part_data['current_name'] = part_data['new_name']
                
                # Also update the "Current Name" column in the treeview
                for i, part_data in enumerate(self.parts_data):
                    item = self.parts_tree.get_children()[i]
                    self.parts_tree.set(item, "Current Name", part_data['current_name'])
                
                messagebox.showinfo("Success", 
                    f"File saved successfully!\n\n"
                    f"Backup created as:\n{os.path.basename(backup_file)}\n\n"
                    f"Changes made: {num_changes} parts modified")
                
            finally:
                # Cleanup temp directory
                shutil.rmtree(temp_dir)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file:\n{str(e)}")

def main():
    root = tk.Tk()
    app = PartNameEditor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
