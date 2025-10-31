#!/usr/bin/python3

"""
MuseScore trackName Editor
--------------------------
Standalone utility to modify trackName in MuseScore .mscz files.
Usage:
-fix concatenation issues by standardizing part names across files
-change the track name in the mixer
-change the track name in midi export

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
import customtkinter as ctk
from tkinter import messagebox
import tkinter as tk
import threading  

# Try to import tkinterdnd2 for better Windows drag/drop
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    TKINTERDND2_AVAILABLE = True
except ImportError:
    TKINTERDND2_AVAILABLE = False

# Application Info
VERSION = "1.1"
LAST_MODIFIED = "2025-10-31"
LICENSE = "GPLv3"
AUTHOR = "Diego Denolf (graffesmusic)"


import zipfile
import xml.etree.ElementTree as ET
import tempfile
import os
import shutil
import customtkinter as ctk
from tkinter import messagebox
import tkinter as tk  
import threading

# Try to import tkinterdnd2 for better Windows drag/drop
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    TKINTERDND2_AVAILABLE = True
except ImportError:
    TKINTERDND2_AVAILABLE = False

# Application Info
VERSION = "1.1"
LAST_MODIFIED = "2025-10-31"
LICENSE = "GPLv3"
AUTHOR = "Diego Denolf (graffesmusic)"

class PartNameEditor:
    def __init__(self, root):
        self.root = root  # Use the provided root window
        
        self.root.title(f"MuseScore Track Name Editor {VERSION}")
        self.root.geometry("750x650")
        self.root.minsize(750, 650)
        
        # Store current theme settings
        self.current_theme = "blue"
        self.current_mode = "Light"
        
        # Configure CustomTkinter appearance
        ctk.set_appearance_mode(self.current_mode)
        ctk.set_default_color_theme(self.current_theme)
        
        # Store file data separately from UI
        self.current_file = None
        self.parts_data = []
        self.original_parts_data = []
        
        # Create the UI
        self.create_ui()
        
        # Enable drag and drop
        self.setup_drag_drop()

    def setup_drag_drop(self):
        """Set up drag and drop functionality with tkinterdnd2"""
        try:
            if TKINTERDND2_AVAILABLE:
                # Use tkinterdnd2 for reliable drag/drop
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind('<<Drop>>', self.on_drop)
                # Also register the main frame
                self.main_frame.drop_target_register(DND_FILES)
                self.main_frame.dnd_bind('<<Drop>>', self.on_drop)
                #print("Drag/drop enabled with tkinterdnd2")
            else:
                # Fallback to basic tkinter DND (less reliable)
                self.root.drop_target_register(tk.DND_FILES)
                self.root.dnd_bind('<<Drop>>', self.on_drop)
                #print("Drag/drop enabled with basic tkinter DND")
            
            # Update file label to show drag/drop is available
            if hasattr(self, 'file_label'):
                self.file_label.configure(text="No file selected (drag .mscz file here)")
                    
        except Exception as e:
            # Don't print the error - it's expected on some systems
            pass  # Silent fail - drag/drop is a bonus feature

    def on_drop(self, event):
        """Handle file drop event"""
        try:
            # Get the dropped files
            files = event.data.split()
            mscz_found = False
            
            for file_path in files:
                # Clean the file path
                clean_path = file_path.strip()
                if clean_path.startswith('{') and clean_path.endswith('}'):
                    clean_path = clean_path[1:-1]
                
                # Check if it's a .mscz file that exists
                if clean_path.lower().endswith('.mscz') and os.path.exists(clean_path):
                    self.root.after(0, lambda path=clean_path: self.load_file_from_path(path))
                    mscz_found = True
                    break  # Use first valid .mscz file
            
            if not mscz_found:
                self.root.after(0, lambda: messagebox.showwarning("Invalid File", "No valid .mscz file found. Please drop a MuseScore file."))
                
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Drag/Drop Error", f"Error processing dropped file: {str(e)}"))
            
            
    def load_file_from_path(self, file_path):
        """Load a file from the given path with safety checks"""
        # Check if we have unsaved changes in current file
        if self.current_file and self.has_unsaved_changes():
            response = messagebox.askyesnocancel(
                "Unsaved Changes",
                f"You have unsaved changes in '{os.path.basename(self.current_file)}'.\n\n"
                "Do you want to save before opening the new file?\n\n"
                "Yes: Save and open new file\n"
                "No: Discard changes and open new file\n" 
                "Cancel: Keep current file"
            )
            
            if response is None:  # Cancel
                return
            elif response:  # Yes - Save first
                self.save_file()
                # Don't proceed if save was cancelled or failed
                if self.has_unsaved_changes():
                    return
        
        # Now safe to load the new file
        self.current_file = file_path
        self.file_label.configure(text=os.path.basename(file_path))
        self.load_parts()

    def has_unsaved_changes(self):
        """Check if there are unsaved changes in the current file"""
        return any(
            part['new_name'] != part['current_name'] 
            for part in self.parts_data
        ) 
        

    def create_ui(self):
        """Create or recreate the entire UI with current theme settings"""
        # Clear existing widgets if recreating
        if hasattr(self, 'main_frame'):
            self.main_frame.destroy()
        
        # Main container
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # File selection
        file_frame = ctk.CTkFrame(self.main_frame)
        file_frame.pack(fill="x", pady=5)
        
        file_label = ctk.CTkLabel(file_frame, text="MuseScore File", 
                                 font=ctk.CTkFont(weight="bold"))
        file_label.pack(anchor="w", pady=(5, 0))
        
        file_btn_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        file_btn_frame.pack(fill="x", pady=5)
        
        ctk.CTkButton(file_btn_frame, text="Open .mscz File", 
                     command=self.open_file, width=140).pack(side="left", padx=5)
        ctk.CTkButton(file_btn_frame, text="Reload", 
                     command=self.load_parts, width=80).pack(side="left", padx=5)
        
        #self.file_label = ctk.CTkLabel(file_frame, text="No file selected", 
        #                              text_color="gray", wraplength=500)
        #self.file_label.pack(pady=5)
        
        # File label - set initial text based on whether we have a file loaded
        if self.current_file:
            file_display_text = os.path.basename(self.current_file)
        else:
            import platform
            if platform.system() == "Linux":
                file_display_text = "No file selected (drag a .mscz file here might not work on your system)"
            else:
                file_display_text = "No file selected (or drag a .mscz file here)"

        
        self.file_label = ctk.CTkLabel(file_frame, text=file_display_text, 
                                      text_color="gray", wraplength=500)
        self.file_label.pack(pady=5)
        
        # Parts list
        parts_frame = ctk.CTkFrame(self.main_frame)
        parts_frame.pack(fill="both", expand=True, pady=5)
        
        parts_label = ctk.CTkLabel(parts_frame, text="Parts - Double-click to edit", 
                                  font=ctk.CTkFont(weight="bold"))
        parts_label.pack(anchor="w", pady=(5, 0))
        
        # Treeview for parts with scrollbar
        tree_frame = ctk.CTkFrame(parts_frame)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Create a frame for the treeview and scrollbar
        tree_container = ctk.CTkFrame(tree_frame)
        tree_container.pack(fill="both", expand=True)
        
        # Use tkinter Treeview (CustomTkinter doesn't have its own treeview yet)
        columns = ("Part ID", "Current Name", "New Name")
        self.parts_tree = tk.ttk.Treeview(tree_container, columns=columns, show="headings", height=12)
        
        # Configure treeview style to match current theme
        self.update_treeview_style()
        
        for col in columns:
            self.parts_tree.heading(col, text=col)
            self.parts_tree.column(col, width=180)
        
        # Add scrollbar
        tree_scroll = ctk.CTkScrollbar(tree_container, orientation="vertical", command=self.parts_tree.yview)
        self.parts_tree.configure(yscrollcommand=tree_scroll.set)
        
        self.parts_tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")
        
        # Bind double-click to edit
        self.parts_tree.bind("<Double-1>", self.on_double_click)
        
        # Quick actions frame
        actions_frame = ctk.CTkFrame(parts_frame, fg_color="transparent")
        actions_frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(actions_frame, text="Quick actions:").pack(side="left", padx=5)
        ctk.CTkButton(actions_frame, text="Add numbers (Violin → Violin 1)", 
                     command=self.add_numbers, width=180).pack(side="left", padx=2)
        ctk.CTkButton(actions_frame, text="Reset all", 
                     command=self.reset_all, width=80).pack(side="left", padx=2)
        
        # Theme selector frame
        theme_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        theme_frame.pack(fill="x", pady=10)
        
        # Color theme selector
        theme_left_frame = ctk.CTkFrame(theme_frame, fg_color="transparent")
        theme_left_frame.pack(side="left")
        
        ctk.CTkLabel(theme_left_frame, text="Color Theme:", 
                    font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
        
        self.theme_var = ctk.StringVar(value=self.current_theme)
        theme_combo = ctk.CTkComboBox(theme_left_frame, 
                                     values=["blue", "green", "dark-blue"],
                                     variable=self.theme_var,
                                     command=self.change_theme,
                                     width=120)
        theme_combo.pack(side="left", padx=5)
        
        # Appearance mode selector
        theme_right_frame = ctk.CTkFrame(theme_frame, fg_color="transparent")
        theme_right_frame.pack(side="right")
        
        ctk.CTkLabel(theme_right_frame, text="Appearance Mode:", 
                    font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
        
        self.mode_var = ctk.StringVar(value=self.current_mode)
        mode_combo = ctk.CTkComboBox(theme_right_frame,
                                    values=["Light", "Dark"],
                                    variable=self.mode_var,
                                    command=self.change_appearance_mode,
                                    width=120)
        mode_combo.pack(side="left", padx=5)
        
        # Button frame at bottom
        button_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        button_frame.pack(fill="x", pady=10)

        # Left side: Save button
        ctk.CTkButton(button_frame, text="Save Changes to File", command=self.save_file,
                     font=ctk.CTkFont(weight="bold"), width=180, height=35).pack(side="left", padx=5)

        # Right side: About and Exit buttons
        right_buttons_frame = ctk.CTkFrame(button_frame, fg_color="transparent")
        right_buttons_frame.pack(side="right")

        ctk.CTkButton(right_buttons_frame, text="About", command=self.show_about,
                     font=ctk.CTkFont(weight="bold"), width=80, height=35).pack(side="left", padx=5)

        ctk.CTkButton(right_buttons_frame, text="Exit", command=self.exit_app,
                     font=ctk.CTkFont(weight="bold"), 
                     fg_color="#d9534f", hover_color="#c9302c",
                     width=80, height=35).pack(side="left", padx=5)
        
        # Reload parts data if we had a file loaded
        if self.current_file and self.parts_data:
            self.reload_parts_tree()

    def show_about(self):
        """Show About dialog with application information"""
        about_text = f"""
Features:
• Modify track names in MuseScore .mscz files to
    • Fix concatenation issues by standardizing part names
    • Change track names in the mixer
    • Change track names in MIDI export
• Automatic backup creation

--------------------------------------------------------------------------------------------------
Published under the GNU General Public License version 3
--------------------------------------------------------------------------------------------------

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

        # Create about dialog
        about_dialog = ctk.CTkToplevel(self.root)
        about_dialog.title("About MuseScore trackName Editor")
        about_dialog.geometry("600x500")
        about_dialog.transient(self.root)
        about_dialog.resizable(False, False)
        
        # Use focus_set instead of grab_set for CTkToplevel
        about_dialog.focus_set()
        about_dialog.lift()
        
        # Title
        title_label = ctk.CTkLabel(about_dialog, 
                                  text="MuseScore trackName Editor",
                                  font=ctk.CTkFont(size=18, weight="bold"))
        title_label.pack(pady=(20, 10))
        
        # Version info
        version_label = ctk.CTkLabel(about_dialog,
                                    text=f"Version {VERSION} • {LAST_MODIFIED}",
                                    font=ctk.CTkFont(size=14))
        version_label.pack(pady=(0, 10))
        
        # Copyright
        copyright_label = ctk.CTkLabel(about_dialog,
                                      text=f"Copyright (c) 2025 {AUTHOR}",
                                      font=ctk.CTkFont(size=12))
        copyright_label.pack(pady=(0, 20))
        
        # Scrollable text area for the rest of the content
        text_frame = ctk.CTkFrame(about_dialog)
        text_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Text widget with scrollbar
        text_widget = tk.Text(text_frame, wrap="word", width=70, height=15,
                             bg="#f5f5f5" if self.current_mode == "Light" else "#2b2b2b",
                             fg="#000000" if self.current_mode == "Light" else "#ffffff",
                             font=("Arial", 10),
                             relief="flat", padx=10, pady=10)
        
        text_scrollbar = ctk.CTkScrollbar(text_frame, orientation="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=text_scrollbar.set)
        
        text_widget.pack(side="left", fill="both", expand=True)
        text_scrollbar.pack(side="right", fill="y")
        
        # Insert about text
        text_widget.insert("1.0", about_text)
        text_widget.configure(state="disabled")  # Make read-only
        
        # Close button
        close_button = ctk.CTkButton(about_dialog, text="Close", 
                                    command=about_dialog.destroy,
                                    width=100, height=35)
        close_button.pack(pady=20)
        
        # Bind Enter key to close
        about_dialog.bind('<Return>', lambda e: about_dialog.destroy())
        about_dialog.bind('<Escape>', lambda e: about_dialog.destroy())

    def reload_parts_tree(self):
        """Reload parts data into the treeview after UI recreation"""
        # Clear existing data
        for item in self.parts_tree.get_children():
            self.parts_tree.delete(item)
        
        # Reload parts data
        for part_data in self.parts_data:
            self.parts_tree.insert("", "end", values=(
                part_data['id'], 
                part_data['current_name'], 
                part_data['new_name']
            ))

    def change_theme(self, choice):
        """Change color theme by recreating the UI"""
        if choice != self.current_theme:
            self.current_theme = choice
            ctk.set_default_color_theme(choice)
            # Recreate UI with new theme
            self.create_ui()
            messagebox.showinfo("Theme Changed", 
                               f"Color theme changed to {choice}.")

    def change_appearance_mode(self, choice):
        """Change appearance mode (light/dark)"""
        if choice != self.current_mode:
            self.current_mode = choice
            ctk.set_appearance_mode(choice)
            # Update treeview and recreate UI
            self.update_treeview_style()
            self.create_ui()
    
    def update_treeview_style(self):
        """Update treeview colors to match current appearance mode"""
        if hasattr(self, 'parts_tree'):
            style = tk.ttk.Style()
            if self.current_mode == "Dark":
                style.configure("Treeview", 
                               background="#2b2b2b",
                               foreground="white",
                               fieldbackground="#2b2b2b")
                style.configure("Treeview.Heading",
                               background="#3b3b3b",
                               foreground="white")
            else:
                style.configure("Treeview", 
                               background="#ffffff",
                               foreground="black",
                               fieldbackground="#ffffff")
                style.configure("Treeview.Heading",
                               background="#e0e0e0",
                               foreground="black")

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
        """Open file with safety check for unsaved changes"""
        # Check for unsaved changes first
        if self.current_file and self.has_unsaved_changes():
            response = messagebox.askyesnocancel(
                "Unsaved Changes",
                f"You have unsaved changes in '{os.path.basename(self.current_file)}'.\n\n"
                "Do you want to save before opening a new file?\n\n"
                "Yes: Save and open new file\n"
                "No: Discard changes and open new file\n" 
                "Cancel: Keep current file"
            )
            
            if response is None:  # Cancel
                return
            elif response:  # Yes - Save first
                self.save_file()
                # Don't proceed if save was cancelled or failed
                if self.has_unsaved_changes():
                    return
        
        # Now safe to open file dialog
        filename = tk.filedialog.askopenfilename(
            title="Select MuseScore file",
            filetypes=[("MuseScore files", "*.mscz"), ("All files", "*.*")]
        )
        if filename:
            self.current_file = filename
            self.file_label.configure(text=os.path.basename(filename))
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
        current_values = self.parts_tree.item(item, 'values')
        if not current_values:
            return
            
        # Create edit dialog using CustomTkinter
        edit_dialog = ctk.CTkToplevel(self.root)
        edit_dialog.title("Edit Part Name")
        edit_dialog.geometry("400x300")
        edit_dialog.transient(self.root)
        
        # Use focus_set instead of grab_set for CTkToplevel
        edit_dialog.focus_set()
        edit_dialog.lift()  # Bring to front
        
        ctk.CTkLabel(edit_dialog, text=f"Editing Part {current_values[0]}", 
                    font=ctk.CTkFont(weight="bold", size=14)).pack(pady=15)
        
        ctk.CTkLabel(edit_dialog, text="Current name:").pack()
        ctk.CTkLabel(edit_dialog, text=current_values[1], 
                    text_color="#1f6aa5").pack()
        
        ctk.CTkLabel(edit_dialog, text="New name:").pack(pady=(15, 0))
        name_entry = ctk.CTkEntry(edit_dialog, width=300, placeholder_text="Enter new part name")
        name_entry.pack(pady=10)
        name_entry.insert(0, current_values[2])
        name_entry.select_range(0, tk.END)
        name_entry.focus()
        
        def apply_change():
            new_name = name_entry.get().strip()
            if new_name:
                # Update treeview
                self.parts_tree.set(item, "New Name", new_name)
                # Update data
                index = self.parts_tree.index(item)
                self.parts_data[index]['new_name'] = new_name
            edit_dialog.destroy()
            
        def cancel_edit():
            edit_dialog.destroy()
            
        btn_frame = ctk.CTkFrame(edit_dialog, fg_color="transparent")
        btn_frame.pack(pady=15)
        
        ctk.CTkButton(btn_frame, text="Apply", command=apply_change, 
                     width=80).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Cancel", command=cancel_edit,
                     fg_color="gray", hover_color="#5a6268",
                     width=80).pack(side="left", padx=10)
        
        # Bind Enter key to apply
        edit_dialog.bind('<Return>', lambda e: apply_change())
        
        # Make dialog modal by waiting for it to close
        self.root.wait_window(edit_dialog)

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
                    # Use the same compression as original files for consistent file sizes
                    with zipfile.ZipFile(os.path.join(temp_dir, "modified.mscz"), 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=6) as modified:
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
                                                    #print(f"Changed part {part_id}: '{part_data['current_name']}' → '{part_data['new_name']}'")
                                    
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
    # Create the appropriate root window
    if TKINTERDND2_AVAILABLE:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    else:
        root = ctk.CTk()
        
    app = PartNameEditor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
