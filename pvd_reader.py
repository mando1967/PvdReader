import win32com.client
import os
import sys
from pathlib import Path
import pythoncom
import time

class PolytecViewer:
    def __init__(self):
        self.viewer = None
        # Initialize COM for the thread
        pythoncom.CoInitialize()
        self.initialize_viewer()
    
    def initialize_viewer(self):
        """Initialize the Polytec ScanViewer ActiveX control"""
        try:
            # Try to create an instance of the Polytec ScanViewer ActiveX control
            self.viewer = win32com.client.Dispatch("Polytec.ScanViewer.ScanViewerActiveXLib.ScanViewer")
            print("Successfully initialized Polytec ScanViewer ActiveX control")
            
            # Try to get available methods and properties
            print("\nExploring available methods and properties:")
            for item in dir(self.viewer):
                if not item.startswith('_'):  # Skip internal attributes
                    print(f"Found: {item}")
            
            return True
        except Exception as e:
            print(f"Error initializing Polytec ScanViewer: {str(e)}")
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
        
        if not self.viewer:
            if not self.initialize_viewer():
                return None
        
        try:
            print(f"\nAttempting to load file: {file_path}")
            # Try different methods to load the file
            try:
                success = self.viewer.LoadFile(str(file_path))
                print("Used LoadFile method")
            except:
                try:
                    success = self.viewer.Open(str(file_path))
                    print("Used Open method")
                except:
                    try:
                        success = self.viewer.OpenFile(str(file_path))
                        print("Used OpenFile method")
                    except Exception as e:
                        print(f"All file loading methods failed: {str(e)}")
                        return None
            
            if success:
                print(f"Successfully loaded file: {file_path}")
                # Give the viewer some time to load the file
                time.sleep(1)
                
                # Get file information
                data = self._extract_file_data()
                return data
            else:
                print(f"Failed to load file: {file_path}")
                return None
                
        except Exception as e:
            print(f"Error reading PVD file: {str(e)}")
            return None
    
    def _extract_file_data(self):
        """Extract data from the currently loaded file"""
        data = {}
        
        try:
            # Common properties to try
            properties_to_try = [
                # File metadata
                'FileName', 'FilePath', 'FileType',
                # Measurement settings
                'SampleRate', 'NumberOfSamples', 'NumberOfChannels',
                'MeasurementMode', 'MeasurementType',
                # Data properties
                'Channels', 'DataLength', 'Domain',
                'XAxisLabel', 'YAxisLabel', 'ZAxisLabel',
                'XAxisUnit', 'YAxisUnit', 'ZAxisUnit',
                # Specific data arrays
                'TimeData', 'FrequencyData', 'AmplitudeData',
                'XData', 'YData', 'ZData',
                # Additional properties
                'Version', 'Description', 'Comment'
            ]
            
            print("\nFile Information:")
            for prop in properties_to_try:
                try:
                    value = getattr(self.viewer, prop)
                    data[prop] = value
                    print(f"{prop}: {value}")
                except:
                    # Try alternative getter methods
                    try:
                        getter = f"Get{prop}"
                        value = getattr(self.viewer, getter)()
                        data[prop] = value
                        print(f"{prop} (via {getter}): {value}")
                    except:
                        try:
                            getter = f"get_{prop}"
                            value = getattr(self.viewer, getter)()
                            data[prop] = value
                            print(f"{prop} (via {getter}): {value}")
                        except:
                            print(f"Could not retrieve property: {prop}")
            
            # Try to get any available methods that might provide data
            print("\nTrying additional methods:")
            methods_to_try = [
                'GetData', 'GetAllData', 'GetChannelData',
                'GetTimeData', 'GetFrequencyData',
                'GetXData', 'GetYData', 'GetZData',
                'GetMetadata', 'GetProperties'
            ]
            
            for method in methods_to_try:
                try:
                    func = getattr(self.viewer, method)
                    result = func()
                    data[method] = result
                    print(f"Successfully called {method}")
                except:
                    print(f"Method not available: {method}")
            
        except Exception as e:
            print(f"Error extracting file data: {str(e)}")
        
        return data
    
    def close(self):
        """Clean up COM objects"""
        if self.viewer:
            try:
                self.viewer.Close()
            except:
                pass
            self.viewer = None
            # Uninitialize COM
            pythoncom.CoUninitialize()

def main():
    # Create viewer instance
    viewer = PolytecViewer()
    
    # Get file path from command line or use default
    if len(sys.argv) > 1:
        test_file = sys.argv[1]
    else:
        print("Please provide a PVD file path as command line argument")
        return
    
    # Read the file
    data = viewer.read_pvd_file(test_file)
    
    # Clean up
    viewer.close()
    
    if data:
        print("\nSuccessfully read PVD file")
        # You can process data further here
        
if __name__ == "__main__":
    main()