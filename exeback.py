# -*- coding: utf-8 -*-
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import fileinput
import subprocess
import paramiko
import win32serviceutil
import win32service
import win32event
import servicemanager
import time
import shutil
import configparser

class SFTPClient:
    def __init__(self):
        self.sftp = None
        self.transport = None

    def connect(self, username: str, password: str, host: str, port: int = 22):
        try:
            self.transport = paramiko.Transport((host.encode("utf-8"), port))
            self.transport.connect(username=username.encode("utf-8"), password=password.encode("utf-8"))
            self.sftp = paramiko.SFTPClient.from_transport(self.transport)
        except Exception as e:
            raise e

    def disconnect(self):
        try:
            if self.sftp:
                self.sftp.close()
            if self.transport:
                self.transport.close()
        except Exception as e:
            raise e

    def upload_file(self, local_path: str, remote_path: str):
        try:
            if self.sftp:
                self.sftp.put(local_path.encode("utf-8"), remote_path.encode("utf-8"))
        except Exception as e:
            raise e

class SFTPService(win32serviceutil.ServiceFramework):
    _svc_name_ = "SFTPUploadService"
    _svc_display_name_ = "SFTP Upload Service"
    _svc_description_ = "A service that uploads files to an SFTP server periodically."
    _svc_deps_ = ["Tcpip", "Dnscache"]


    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.running = True
        self.client = SFTPClient()
        self.config = self.read_config()
        
        self.local_folder = self.config.get('local_folder', '')
        self.remote_folder = self.config.get('remote_folder', '/')
        self.host = self.config.get('host', '')
        self.port = int(self.config.get('port', 22))
        self.username = self.config.get('username', '')
        self.password = self.config.get('password', '')
        
        self.uploaded_files = set()

    @staticmethod
    def read_config():
        config = configparser.ConfigParser()
        config_path = os.path.join(os.path.dirname(__file__), 'sftp_config.ini').encode("utf-8")
        config.read(config_path)
        return dict(config['Settings']) if config.has_section('Settings') else {}

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)
        self.running = False

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                               servicemanager.PYS_SERVICE_STARTED,
                               (self._svc_name_, ''))
        
        while self.running:
            try:
                if not self.client.sftp:
                    self.client.connect(self.username, self.password, self.host, self.port)

                upload_files_from_folder(self.local_folder, self.remote_folder, self.client, self.uploaded_files)
                
                move_uploaded_files(self.local_folder)

                win32event.WaitForSingleObject(self.stop_event, 60000)  

            except Exception as e:
                print(f"Error in service: {e}")
                time.sleep(10)  

        self.client.disconnect()

def upload_files_from_folder(local_folder: str, remote_folder: str, sftp_client: SFTPClient, uploaded_files: set):
    for filename in os.listdir(local_folder):
        local_file_path = os.path.join(local_folder, filename)
        
        if os.path.isfile(local_file_path) and filename not in uploaded_files:
            remote_file_path = os.path.join(remote_folder, filename)
            sftp_client.upload_file(local_file_path, remote_file_path)
            uploaded_files.add(filename)  

def move_uploaded_files(local_folder: str):
    uploaded_dir = os.path.join(local_folder, "uploaded")
    os.makedirs(uploaded_dir, exist_ok=True)
    
    for filename in os.listdir(local_folder):
        local_file_path = os.path.join(local_folder, filename)
        
        if os.path.isfile(local_file_path):  
            moved_file_path = os.path.join(uploaded_dir, filename)
            shutil.move(local_file_path, moved_file_path)


class SFTPServiceConfigurator:
    def __init__(self, master):
        self.master = master
        master.title("SFTP Service Configurator")
        master.geometry("500x400")

        self.config = configparser.ConfigParser()
        self.config_path = 'sftp_config.ini'
        self.load_config()

        self.status_label = tk.Label(master, text="Status: Not Installed", font=("Arial", 12), fg="red")
        self.status_label.pack(pady=(10, 10))

        server_frame = tk.LabelFrame(master, text="SFTP Server Configuration")
        server_frame.pack(padx=20, pady=10, fill='x')

        tk.Label(server_frame, text="Host:").pack()
        self.host_entry = tk.Entry(server_frame, width=40)
        self.host_entry.pack()
        self.host_entry.insert(0, self.config.get('Settings', 'host', fallback=''))

        tk.Label(server_frame, text="Port:").pack()
        self.port_entry = tk.Entry(server_frame, width=40)
        self.port_entry.pack()
        self.port_entry.insert(0, self.config.get('Settings', 'port', fallback='22'))

        tk.Label(server_frame, text="Username:").pack()
        self.username_entry = tk.Entry(server_frame, width=40)
        self.username_entry.pack()
        self.username_entry.insert(0, self.config.get('Settings', 'username', fallback=''))

        tk.Label(server_frame, text="Password:").pack()
        self.password_entry = tk.Entry(server_frame, show="*", width=40)
        self.password_entry.pack()
        self.password_entry.insert(0, self.config.get('Settings', 'password', fallback=''))

        tk.Label(master, text="Local Folder Path:").pack()
        self.folder_frame = tk.Frame(master)
        self.folder_frame.pack(pady=(5, 10), padx=20, fill='x')
        
        self.folder_path = tk.StringVar()
        self.folder_path.set(self.config.get('Settings', 'local_folder', fallback=''))
        self.folder_entry = tk.Entry(self.folder_frame, textvariable=self.folder_path, width=40)
        self.folder_entry.pack(side=tk.LEFT, expand=True, fill='x', padx=(0, 10))
        
        browse_btn = tk.Button(self.folder_frame, text="Browse", command=self.browse_folder)
        browse_btn.pack(side=tk.RIGHT)

        tk.Label(master, text="Remote Folder Path:").pack()
        self.remote_folder_entry = tk.Entry(master, width=40)
        self.remote_folder_entry.pack()
        self.remote_folder_entry.insert(0, self.config.get('Settings', 'remote_folder', fallback='/'))

        self.button_frame = tk.Frame(master)
        self.button_frame.pack(pady=(10, 20))

        self.install_btn = tk.Button(self.button_frame, text="Install Service", command=self.install_service)
        self.install_btn.pack(side=tk.LEFT, padx=(0, 20))

        self.stop_btn = tk.Button(self.button_frame, text="Stop Service", command=self.stop_service)
        self.stop_btn.pack(side=tk.LEFT)

    def load_config(self):
        if not os.path.exists(self.config_path):
            self.config['Settings'] = {}
        else:
            self.config.read(self.config_path)

    def save_config(self):
        if not self.config.has_section('Settings'):
            self.config.add_section('Settings')

        self.config.set('Settings', 'host', self.host_entry.get())
        self.config.set('Settings', 'port', self.port_entry.get())
        self.config.set('Settings', 'username', self.username_entry.get())
        self.config.set('Settings', 'password', self.password_entry.get())
        self.config.set('Settings', 'local_folder', self.folder_path.get())
        self.config.set('Settings', 'remote_folder', self.remote_folder_entry.get())

        with open(self.config_path, 'w') as configfile:
            self.config.write(configfile)

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)

    def install_service(self):
        # Validate inputs
        if not all([self.host_entry.get(), self.port_entry.get(), 
                    self.username_entry.get(), self.password_entry.get(), 
                    self.folder_path.get()]):
            messagebox.showerror("Error", "Please fill all configuration fields")
            return

        try:
            self.save_config()
            if hasattr(sys, 'frozen'):  # Check if running as a frozen executable
                service_executable = sys.executable
            else:
                service_executable = sys.argv[0]

            install_result = subprocess.run(
                [service_executable, 'install'], 
                capture_output=True, 
                text=True
            )

            if install_result.returncode == 0:
                messagebox.showinfo("Success", "SFTP Upload Service installed")
                self.status_label.config(text="Status: Installed", fg="green")
                
                start_result = subprocess.run(
                    [service_executable, sys.argv[0], 'start'], 
                    capture_output=True, 
                    text=True
                )
                
                if start_result.returncode == 0:
                    messagebox.showinfo("Success", "SFTP Upload Service started")
                    self.status_label.config(text="Status: Running", fg="green")
                else:
                    messagebox.showwarning("Start Warning", f"Service installed but failed to start: {start_result.stderr}")
                    self.status_label.config(text=f"Status: Installed but not running\nError: {start_result.stderr}", fg="orange")
            else:
                messagebox.showerror("Installation Error", install_result.stderr)
                self.status_label.config(text=f"Status: Installation failed\nError: {install_result.stderr}", fg="red")

        except Exception as e:
            messagebox.showerror("Installation Error", str(e))
            self.status_label.config(text=f"Status: Installation error\nError: {str(e)}", fg="red")

    def stop_service(self):
        try:
            esecutproc = sys.executable
            result = subprocess.run(
                [esecutproc, sys.argv[0], 'stop'], 
                capture_output=True, 
                text=True
            )

            if result.returncode == 0:
                messagebox.showinfo("Success", "SFTP Upload Service stopped")
                self.status_label.config(text="Status: Not Running", fg="red")
            else:
                messagebox.showerror("Stop Error", result.stderr)
                self.status_label.config(text=f"Status: Error stopping service\nError: {result.stderr}", fg="orange")

        except Exception as e:
            messagebox.showerror("Stop Error", str(e))
            self.status_label.config(text=f"Status: Stop error\nError: {str(e)}", fg="orange")

def main():
    if len(sys.argv) > 1:
        win32serviceutil.HandleCommandLine(SFTPService)
    else:
        root = tk.Tk()
        app = SFTPServiceConfigurator(root)
        root.mainloop()

if __name__ == "__main__":
    main()