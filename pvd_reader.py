import win32com.client
import os
import sys
from pathlib import Path
import pythoncom
import time

class PolytecFileReader:
    def __init__(self):
        self.file_access = None
        # Initialize COM for the thread
        pythoncom.CoInitialize()
        self.initialize_file_access()
    
    def initialize_file_access(self):
        """Initialize the Polytec PolyFileAccess control"""
        try:
            # Try to create an instance of the Polytec PolyFileAccess control
            self.file_access = win32com.client.Dispatch("PolyFileAccess.PolyFileAccess")
            print("Successfully initialized Polytec PolyFileAccess control")
            
            # Try to get available methods and properties
            print("\nExploring available methods and properties:")
            for item in dir(self.file_access):
                if not item.startswith('_'):  # Skip internal attributes
                    print(f"Found: {item}")
            
            return True
        except Exception as e:
            print(f"Error initializing Polytec PolyFileAccess: {str(e)}")
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
            # Try different methods to load the file
            try:
                success = self.file_access.OpenFile(str(file_path))
                print("Used OpenFile method")
            except:
                try:
                    success = self.file_access.Open(str(file_path))
                    print("Used Open method")
                except Exception as e:
                    print(f"All file loading methods failed: {str(e)}")
                    return None
            
            # Get file information
            data = self._extract_file_data()
            return data
                
        except Exception as e:
            print(f"Error reading PVD file: {str(e)}")
            return None
    
    def _extract_file_data(self):
        """Extract data from the currently loaded file"""
        data = {}
        
        try:
            # Common properties to try for PolyFileAccess
            properties_to_try = [
                # File information
                'FileVersion',
                'FileType',
                'FileComment',
                # Measurement settings
                'SamplingFrequency',
                'NumberOfChannels',
                'NumberOfSamples',
                'MeasurementComment',
                'MeasurementDirection',
                'SignalType',
                # Channel information
                'ChannelComment',
                'ChannelQuantity',
                'ChannelUnit',
                # Signal properties
                'SignalDomain',
                'SignalUnit',
                'SignalQuantity'
            ]
            
            print("\nFile Information:")
            for prop in properties_to_try:
                try:
                    value = getattr(self.file_access, prop)
                    data[prop] = value
                    print(f"{prop}: {value}")
                except:
                    # Try alternative getter methods
                    try:
                        getter = f"Get{prop}"
                        value = getattr(self.file_access, getter)()
                        data[prop] = value
                        print(f"{prop} (via {getter}): {value}")
                    except:
                        try:
                            getter = f"get_{prop}"
                            value = getattr(self.file_access, getter)()
                            data[prop] = value
                            print(f"{prop} (via {getter}): {value}")
                        except:
                            print(f"Could not retrieve property: {prop}")
            
            # Try to get signal data
            try:
                # Get number of channels and samples
                num_channels = data.get('NumberOfChannels', 1)
                num_samples = data.get('NumberOfSamples', 0)
                
                if num_samples > 0:
                    print("\nAttempting to read signal data...")
                    
                    # Try different methods to get signal data
                    try:
                        # Try to get data for each channel
                        for channel in range(num_channels):
                            channel_data = self.file_access.GetSignalData(channel)
                            data[f'Channel_{channel}_Data'] = channel_data
                            print(f"Successfully read data for channel {channel}")
                    except Exception as e:
                        print(f"Error reading signal data: {str(e)}")
            except Exception as e:
                print(f"Error accessing signal data: {str(e)}")
            
        except Exception as e:
            print(f"Error extracting file data: {str(e)}")
        
        return data
    
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
    else:
        print("Please provide a PVD file path as command line argument")
        return
    
    # Read the file
    data = reader.read_pvd_file(test_file)
    
    # Clean up
    reader.close()
    
    if data:
        print("\nSuccessfully read PVD file")
        # You can process data further here
        
if __name__ == "__main__":
    main()