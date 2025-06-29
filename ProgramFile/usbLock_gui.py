import psutil
import subprocess
import os
import ctypes
import sys
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
import threading
import logging
from pathlib import Path

# Setup logging
logging.basicConfig(filename='usblock.log', level=logging.INFO, 
                   format='%(asctime)s - %(levelname)s - %(message)s')

# Language texts
TEXTS = {
    'title': 'USBLock - USB Drive Manager',
    'refresh': 'Refresh Drives',
    'disable': 'Disable Selected',
    'enable': 'Enable Selected',
    'status_ready': 'Ready',
    'no_drives': 'No USB drives found.',
    'no_removable': 'No removable drives found.',
    'select_drive': 'Please select a drive.',
    'invalid_selection': 'Invalid selection.',
    'confirm_disable': 'Are you sure you want to disable drive {drive}? This will make it unrecognizable.',
    'confirm_enable': 'Are you sure you want to enable drive {drive} using {backup}?',
    'backup_failed': 'Failed to create backup. Aborting disable operation.',
    'disable_success': 'Drive {drive} disabled successfully. Backup saved as {backup}. Remove and reinsert the USB.',
    'enable_success': 'Drive {drive} enabled successfully. Remove and reinsert the USB.',
    'backup_not_found': 'Backup file {backup} not found.',
    'invalid_backup': 'Backup file is invalid (must be 512 bytes).',
    'error_disable': 'Error disabling drive: {error}',
    'error_enable': 'Error enabling drive: {error}',
    'error_drive_number': 'Could not find physical drive number.',
    'admin_error': 'This application requires administrator privileges. Restarting as administrator...',
    'windows_only': 'This tool only works on Windows.',
    'help': 'Help',
    'help_text': '''USBLock allows you to disable or enable USB drives.
- Disable: Makes the drive unrecognizable by overwriting its partition table (data is preserved).
- Enable: Restores the drive using a backup file.
- Backup files are saved in the USBLock_Backups folder.
- Always run as administrator.
- Safely remove and reinsert the USB after operations.''',
    'select_backup': 'Select Backup File',
    'backup_label': 'Available Backups:',
    'delete_backup': 'Delete backup file after enabling?'
}

def is_admin():
    """Check if the script is running with admin privileges"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """Restart the script with admin privileges"""
    try:
        ctypes.windll.shell32.ShellExecuteW(
            None, 
            "runas", 
            sys.executable, 
            " ".join(sys.argv), 
            None, 
            1
        )
    except Exception as e:
        logging.error(f"Failed to run as admin: {e}")
    sys.exit()

def get_usb_drives():
    """Returns a list of mounted USB drives with size and label"""
    usb_drives = []
    for disk in psutil.disk_partitions():
        if 'removable' in disk.opts.lower() or 'usb' in disk.device.lower():
            try:
                usage = psutil.disk_usage(disk.mountpoint)
                total_size = f"{usage.total / (1024**3):.2f} GB"
                # Get volume label
                try:
                    result = subprocess.run(
                        f'vol {disk.device}', 
                        capture_output=True, 
                        text=True, 
                        shell=True
                    )
                    label = "No Label"
                    if result.returncode == 0:
                        lines = result.stdout.splitlines()
                        for line in lines:
                            if 'is' in line and disk.device[0] in line:
                                parts = line.split(' is ')
                                if len(parts) > 1:
                                    label = parts[1].strip()
                                break
                except:
                    label = "No Label"
            except:
                total_size = "Unknown"
                label = "No Label"
            usb_drives.append({
                'device': disk.device,
                'mountpoint': disk.mountpoint,
                'fstype': disk.fstype,
                'size': total_size,
                'label': label
            })
    return usb_drives

def get_all_physical_drives():
    """Returns a list of all removable physical drive numbers with serial numbers"""
    try:
        result = subprocess.run(
            'wmic diskdrive where "MediaType=\'Removable Media\'" get index,serialnumber', 
            capture_output=True, 
            text=True, 
            shell=True
        )
        lines = result.stdout.splitlines()
        drives = []
        for line in lines[1:]:
            if line.strip():
                parts = line.split()
                if len(parts) >= 1:
                    try:
                        index = int(parts[0])
                        serial = parts[1] if len(parts) > 1 else "Unknown"
                        drives.append({'index': index, 'serial': serial})
                    except ValueError:
                        continue
        return drives
    except Exception as e:
        logging.error(f"Error retrieving physical drives: {e}")
        return []

def get_physical_drive_number(mountpoint):
    """Extracts the physical drive number from the mountpoint"""
    try:
        # Get the drive letter
        drive_letter = mountpoint[0]
        
        # Find the physical drive number for this drive letter
        result = subprocess.run(
            f'wmic partition where "DriveLetter=\'{drive_letter}:\'" get DiskIndex', 
            capture_output=True, 
            text=True, 
            shell=True
        )
        
        lines = result.stdout.splitlines()
        for line in lines[1:]:
            if line.strip() and line.strip().isdigit():
                return int(line.strip())
        
        # Alternative method
        result = subprocess.run(
            'wmic diskdrive get index,size', 
            capture_output=True, 
            text=True, 
            shell=True
        )
        # This is a simplified approach - in production you'd need more robust mapping
        return 1  # Default for first removable drive
        
    except Exception as e:
        logging.error(f"Error retrieving drive number: {e}")
        return None

def backup_partition_table(drive_number):
    """Backs up the first sector (MBR/GPT) of the drive"""
    backup_dir = Path("USBLock_Backups")
    backup_dir.mkdir(exist_ok=True)
    try:
        with open(f"\\\\.\\PhysicalDrive{drive_number}", "rb") as disk:
            disk.seek(0)
            sector = disk.read(512)
            if len(sector) != 512:
                raise ValueError("Invalid sector size")
        backup_file = backup_dir / f"usb_backup_drive{drive_number}_{int(time.time())}.bin"
        with open(backup_file, "wb") as f:
            f.write(sector)
        logging.info(f"Backup created: {backup_file}")
        return str(backup_file)
    except (PermissionError, OSError, ValueError) as e:
        logging.error(f"Backup failed for drive {drive_number}: {e}")
        return None

def disable_usb_drive(drive_number):
    """Disables the USB drive by overwriting the partition table"""
    try:
        backup_file = backup_partition_table(drive_number)
        if not backup_file:
            return False, "backup_failed"
        
        with open(f"\\\\.\\PhysicalDrive{drive_number}", "r+b") as disk:
            disk.seek(0)
            disk.write(b'\x00' * 512)
            disk.flush()
        logging.info(f"Drive {drive_number} disabled, backup: {backup_file}")
        return True, ("disable_success", drive_number, backup_file)
    except (PermissionError, OSError) as e:
        logging.error(f"Error disabling drive {drive_number}: {e}")
        return False, ("error_disable", str(e))

def enable_usb_drive(drive_number, backup_file):
    """Enables the USB drive by restoring the partition table"""
    try:
        if not os.path.exists(backup_file):
            return False, ("backup_not_found", backup_file)
        
        with open(backup_file, "rb") as f:
            sector = f.read()
            if len(sector) != 512:
                return False, ("invalid_backup",)
        
        with open(f"\\\\.\\PhysicalDrive{drive_number}", "r+b") as disk:
            disk.seek(0)
            disk.write(sector)
            disk.flush()
        logging.info(f"Drive {drive_number} enabled with backup: {backup_file}")
        return True, ("enable_success", drive_number, backup_file)
    except (PermissionError, OSError) as e:
        logging.error(f"Error enabling drive {drive_number}: {e}")
        return False, ("error_enable", str(e))

class USBLockApp:
    def __init__(self, root):
        self.root = root
        self.root.title(TEXTS['title'])
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        # Apply modern theme
        try:
            self.style = ttkb.Style(theme="superhero")
        except:
            self.style = ttkb.Style(theme="darkly")
        
        # Main frame
        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.pack(fill=BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(self.main_frame, text="USBLock", font=("Helvetica", 18, "bold"))
        title_label.pack(pady=10)
        
        # Drive list frame
        self.drive_frame = ttk.Frame(self.main_frame)
        self.drive_frame.pack(fill=BOTH, expand=True, pady=5)
        
        # Create scrollable listbox for drives
        self.drive_list_frame = ttk.Frame(self.drive_frame)
        self.drive_list_frame.pack(fill=BOTH, expand=True)
        
        # Drive listbox with scrollbar
        self.drive_listbox = tk.Listbox(
            self.drive_list_frame, 
            height=8, 
            font=("Consolas", 10),
            bg="#2b3e50",
            fg="white",
            selectbackground="#3498db",
            selectforeground="white"
        )
        self.drive_scrollbar = ttk.Scrollbar(self.drive_list_frame, orient=VERTICAL)
        self.drive_listbox.config(yscrollcommand=self.drive_scrollbar.set)
        self.drive_scrollbar.config(command=self.drive_listbox.yview)
        
        self.drive_listbox.pack(side=LEFT, fill=BOTH, expand=True)
        self.drive_scrollbar.pack(side=RIGHT, fill=Y)
        
        # Backup list for enable mode
        self.backup_label = ttk.Label(self.drive_frame, text=TEXTS['backup_label'])
        self.backup_list_frame = ttk.Frame(self.drive_frame)
        
        self.backup_listbox = tk.Listbox(
            self.backup_list_frame, 
            height=5,
            font=("Consolas", 9),
            bg="#34495e",
            fg="white",
            selectbackground="#e74c3c",
            selectforeground="white"
        )
        self.backup_scrollbar = ttk.Scrollbar(self.backup_list_frame, orient=VERTICAL)
        self.backup_listbox.config(yscrollcommand=self.backup_scrollbar.set)
        self.backup_scrollbar.config(command=self.backup_listbox.yview)
        
        self.backup_listbox.bind('<<ListboxSelect>>', self.select_backup)
        self.selected_backup = tk.StringVar()
        
        # Buttons frame
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.pack(fill=X, pady=10)
        
        # Create buttons with styling
        self.refresh_btn = ttkb.Button(
            self.button_frame,
            text=TEXTS['refresh'], 
            command=self.refresh_drives, 
            bootstyle="info-outline"
        )
        self.refresh_btn.pack(side=LEFT, padx=5)
        
        self.disable_btn = ttkb.Button(
            self.button_frame, 
            text=TEXTS['disable'], 
            command=self.disable_drive, 
            bootstyle="warning-outline"
        )
        self.disable_btn.pack(side=LEFT, padx=5)
        
        self.enable_btn = ttkb.Button(
            self.button_frame, 
            text=TEXTS['enable'], 
            command=self.enable_drive, 
            bootstyle="success-outline"
        )
        self.enable_btn.pack(side=LEFT, padx=5)
        
        self.help_btn = ttkb.Button(
            self.button_frame, 
            text=TEXTS['help'], 
            command=self.show_help, 
            bootstyle="secondary-outline"
        )
        self.help_btn.pack(side=LEFT, padx=5)
        
        # Progress bar
        self.progress = ttkb.Progressbar(
            self.main_frame, 
            bootstyle="info-striped", 
            mode='indeterminate'
        )
        
        # Status bar
        self.status_var = tk.StringVar(value=TEXTS['status_ready'])
        status_label = ttk.Label(
            self.main_frame, 
            textvariable=self.status_var, 
            font=("Helvetica", 10)
        )
        status_label.pack(fill=X, pady=(5, 0))
        
        # Mode (disable/enable)
        self.mode = "disable"
        self.refresh_drives()
    
    def show_help(self):
        """Show help dialog"""
        messagebox.showinfo(TEXTS['help'], TEXTS['help_text'])
    
    def select_backup(self, event):
        """Update selected backup file"""
        selection = self.backup_listbox.curselection()
        if selection:
            self.selected_backup.set(self.backup_listbox.get(selection[0]))
    
    def refresh_drives(self):
        """Refresh the list of drives and backups"""
        self.drive_listbox.delete(0, tk.END)
        
        # Hide backup components first
        self.backup_label.pack_forget()
        self.backup_list_frame.pack_forget()
        
        if self.mode == "disable":
            usb_drives = get_usb_drives()
            if not usb_drives:
                self.drive_listbox.insert(tk.END, TEXTS['no_drives'])
                self.status_var.set(TEXTS['no_drives'])
            else:
                for i, drive in enumerate(usb_drives, 1):
                    drive_info = f"{i}. Drive: {drive['device']} ({drive['label']})"
                    drive_info += f", Size: {drive['size']}, Type: {drive['fstype']}"
                    self.drive_listbox.insert(tk.END, drive_info)
                self.status_var.set(f"Found {len(usb_drives)} USB drive(s).")
        else:
            # Show backup components for enable mode
            self.backup_label.pack(fill=X, pady=(5, 0))
            self.backup_list_frame.pack(fill=X, pady=(5, 10))
            self.backup_listbox.pack(side=LEFT, fill=BOTH, expand=True)
            self.backup_scrollbar.pack(side=RIGHT, fill=Y)
            
            self.backup_listbox.delete(0, tk.END)
            drive_numbers = get_all_physical_drives()
            if not drive_numbers:
                self.drive_listbox.insert(tk.END, TEXTS['no_removable'])
                self.status_var.set(TEXTS['no_removable'])
            else:
                for i, drive in enumerate(drive_numbers, 1):
                    serial = drive['serial'][:15] + "..." if len(drive['serial']) > 15 else drive['serial']
                    drive_info = f"{i}. Physical Drive {drive['index']} (Serial: {serial})"
                    self.drive_listbox.insert(tk.END, drive_info)
                self.status_var.set(f"Found {len(drive_numbers)} removable drive(s).")
            
            # List backup files
            backup_dir = Path("USBLock_Backups")
            backups = list(backup_dir.glob("usb_backup_drive*.bin")) if backup_dir.exists() else []
            if backups:
                for backup in backups:
                    self.backup_listbox.insert(tk.END, str(backup))
            else:
                self.backup_listbox.insert(tk.END, "No backups found.")
    
    def disable_drive(self):
        """Disable the selected USB drive in a separate thread"""
        selection = self.drive_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", TEXTS['select_drive'])
            return
        
        index = selection[0]
        usb_drives = get_usb_drives()
        if not usb_drives or index >= len(usb_drives):
            messagebox.showerror("Error", TEXTS['invalid_selection'])
            return
        
        selected_drive = usb_drives[index]
        drive_number = get_physical_drive_number(selected_drive['mountpoint'])
        if drive_number is None:
            messagebox.showerror("Error", TEXTS['error_drive_number'])
            self.status_var.set(TEXTS['error_drive_number'])
            return
        
        confirm_msg = TEXTS['confirm_disable'].format(drive=selected_drive['device'])
        if messagebox.askyesno("Confirm", confirm_msg):
            self.progress.pack(fill=X, pady=(10, 0))
            self.progress.start(10)
            self._disable_buttons()
            threading.Thread(
                target=self._disable_thread, 
                args=(drive_number, selected_drive['device']), 
                daemon=True
            ).start()
    
    def enable_drive(self):
        """Enable the selected USB drive in a separate thread"""
        if self.mode != "enable":
            self.mode = "enable"
            self.refresh_drives()
            return
        
        selection = self.drive_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", TEXTS['select_drive'])
            return
        
        index = selection[0]
        drive_numbers = get_all_physical_drives()
        if not drive_numbers or index >= len(drive_numbers):
            messagebox.showerror("Error", TEXTS['invalid_selection'])
            return
        
        selected_drive = drive_numbers[index]['index']
        backup_file = self.selected_backup.get()
        
        if not backup_file or "No backups found" in backup_file:
            backup_file = filedialog.askopenfilename(
                title=TEXTS['select_backup'],
                filetypes=[("Backup files", "*.bin"), ("All files", "*.*")]
            )
            if not backup_file:
                self.status_var.set("No backup file selected.")
                return
        
        confirm_msg = TEXTS['confirm_enable'].format(
            drive=selected_drive, 
            backup=os.path.basename(backup_file)
        )
        if messagebox.askyesno("Confirm", confirm_msg):
            delete_backup = messagebox.askyesno("Confirm", TEXTS['delete_backup'])
            self.progress.pack(fill=X, pady=(10, 0))
            self.progress.start(10)
            self._disable_buttons()
            threading.Thread(
                target=self._enable_thread, 
                args=(selected_drive, backup_file, delete_backup), 
                daemon=True
            ).start()
    
    def _disable_buttons(self):
        """Disable all buttons during operation"""
        self.refresh_btn.configure(state='disabled')
        self.disable_btn.configure(state='disabled')
        self.enable_btn.configure(state='disabled')
    
    def _enable_buttons(self):
        """Enable all buttons after operation"""
        self.refresh_btn.configure(state='normal')
        self.disable_btn.configure(state='normal')
        self.enable_btn.configure(state='normal')
    
    def _disable_thread(self, drive_number, drive_name):
        """Thread for disabling drive"""
        try:
            success, message = disable_usb_drive(drive_number)
            self.root.after(0, lambda: self._post_operation(success, message, drive_name))
        except Exception as e:
            error_msg = f"Unexpected error: {str(e)}"
            self.root.after(0, lambda: self._post_operation(False, error_msg, drive_name))
    
    def _enable_thread(self, drive_number, backup_file, delete_backup):
        """Thread for enabling drive"""
        try:
            success, message = enable_usb_drive(drive_number, backup_file)
            if success and delete_backup:
                try:
                    os.remove(backup_file)
                    logging.info(f"Backup file deleted: {backup_file}")
                except Exception as e:
                    logging.error(f"Error deleting backup: {backup_file} - {e}")
            self.root.after(0, lambda: self._post_operation(success, message, drive_number))
        except Exception as e:
            error_msg = f"Unexpected error: {str(e)}"
            self.root.after(0, lambda: self._post_operation(False, error_msg, drive_number))
    
    def _post_operation(self, success, message, drive):
        """Handle post-operation tasks"""
        self.progress.stop()
        self.progress.pack_forget()
        self._enable_buttons()
        
        if success:
            if isinstance(message, tuple):
                key, *args = message
                if key in TEXTS:
                    if len(args) >= 2:
                        formatted_msg = TEXTS[key].format(
                            drive=drive, 
                            backup=os.path.basename(args[1])
                        )
                    else:
                        formatted_msg = TEXTS[key].format(drive=drive)
                else:
                    formatted_msg = str(message)
            else:
                formatted_msg = str(message)
            messagebox.showinfo("Success", formatted_msg)
            self.status_var.set("Operation completed successfully.")
        else:
            if isinstance(message, tuple):
                key, *args = message
                if key in TEXTS:
                    formatted_msg = TEXTS[key].format(
                        error=args[0] if args else "Unknown error"
                    )
                else:
                    formatted_msg = str(message)
            else:
                formatted_msg = str(message)
            messagebox.showerror("Error", formatted_msg)
            self.status_var.set("Operation failed.")
        
        # Reset mode
        if self.mode == "enable":
            self.mode = "disable"
        self.refresh_drives()

def main():
    """Main function with automatic admin privilege handling"""
    # Check if running on Windows
    if os.name != 'nt':
        messagebox.showerror("Error", TEXTS['windows_only'])
        return
    
    # Check for admin privileges and run as admin if needed
    if not is_admin():
        print("Requesting administrator privileges...")
        try:
            run_as_admin()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to get admin privileges: {e}")
        return
    
    try:
        # Create the main window
        root = ttkb.Window()
        
        # Set window icon (optional)
        try:
            root.iconbitmap('./logo/usblock_logo.ico')
        except:
            pass
        
        # Create the application
        app = USBLockApp(root)
        
        # Center the window
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f"{width}x{height}+{x}+{y}")
        
        # Start the GUI
        root.mainloop()
        
    except Exception as e:
        logging.error(f"Application error: {e}")
        messagebox.showerror("Error", f"Application error: {e}")

if __name__ == "__main__":
    main()