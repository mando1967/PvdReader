import win32com.client
import os
import sys
from pathlib import Path
import pythoncom
import time
import inspect

class PolytecFileReader:
    def __init__(self):
        self.file_access = None
        # Initialize COM for the thread
        pythoncom.CoInitialize()
        self.initialize_file_access()
    
    def initialize_file_access(self):
        """Initialize the PolyFile control and explore its attributes"""
        try:
            # Create an instance of the PolyFile control
            self.file_access = win32com.client.Dispatch("PolyFile.PolyFile")
            print("Successfully initialized PolyFile control")
            
            # Explore and print all attributes
            print("\nExploring PolyFile object attributes:")
            print("=====================================")
            
            # Get all attributes
            attributes = dir(self.file_access)
            
            # Categorize attributes
            methods = []
            properties = []
            other = []
            
            for attr in attributes:
                if attr.startswith('_'):
                    continue
                    
                try:
                    # Try to get the attribute
                    attr_value = getattr(self.file_access, attr)
                    
                    # Check if it's callable (method)
                    if callable(attr_value):
                        methods.append(attr)
                    else:
                        properties.append(attr)
                except:
                    other.append(attr)
            
            # Print methods
            print("\nMethods:")
            print("--------")
            for method in sorted(methods):
                try:
                    # Try to get method signature if possible
                    attr_value = getattr(self.file_access, method)
                    print(f"  {method}()")
                except:
                    print(f"  {method}()")
            
            # Print properties
            print("\nProperties:")
            print("-----------")
            for prop in sorted(properties):
                try:
                    value = getattr(self.file_access, prop)
                    print(f"  {prop} = {value}")
                except Exception as e:
                    print(f"  {prop} (unable to read value: {str(e)})")
            
            # Print other attributes
            if other:
                print("\nOther Attributes:")
                print("----------------")
                for attr in sorted(other):
                    print(f"  {attr}")
            
            return True
            
        except Exception as e:
            print(f"Error initializing PolyFile: {str(e)}")
            return False
    
    def read_pvd_file(self, file_path):
        """
        Read a PVD file and return its contents
        
        Args:
            file_path (str): Path to the PVD file
            
        Returns:
            dict: Dictionary containing the file data and metadata
        """
        file_path = Path(file_path)
        if not file_path.exists():
            print(f"Error: File {file_path} does not exist")
            return None
        
        if not self.file_access:
            if not self.initialize_file_access():
                return None
        
        try:
            print(f"\nAttempting to load file: {file_path}")
            
            # Try to open the file
            try:
                self.file_access.OpenFile(str(file_path))
                print("Successfully opened file")
                
                # After opening the file, explore the object again to see any new attributes
                print("\nExploring PolyFile object attributes after file open:")
                print("================================================")
                self.initialize_file_access()  # This will re-print all attributes
                
            except Exception as e:
                print(f"Error opening file: {str(e)}")
                return None
            
            return True
                
        except Exception as e:
            print(f"Error reading PVD file: {str(e)}")
            return None
    
    def close(self):
        """Clean up COM objects"""
        if self.file_access:
            try:
                self.file_access.CloseFile()
            except:
                pass
            self.file_access = None
            # Uninitialize COM
            pythoncom.CoUninitialize()

def main():
    # Create file reader instance
    reader = PolytecFileReader()
    
    # Get file path from command line or use default
    if len(sys.argv) > 1:
        test_file = sys.argv[1]
        
        # Read the file
        reader.read_pvd_file(test_file)
    else:
        print("No file specified - showing available attributes of uninitialized PolyFile object")
    
    # Clean up
    reader.close()
        
if __name__ == "__main__":
    main()