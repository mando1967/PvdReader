# PVD Reader

A Python tool for reading and visualizing Polytec measurement files (PVD and SVD formats).

## File Formats

### PVD (Polytec Vibrometer Data)
- Binary file format used by Polytec's scanning vibrometer systems
- Contains time-domain vibration measurement data
- Supports multiple channels
- Includes metadata such as sample rate and measurement parameters

### SVD (Scanning Vibrometer Data)
- Binary file format for scanning vibrometer measurements
- Contains frequency-domain data
- Includes spatial scanning information
- Supports multiple measurement points and frequencies

## Setup

1. Create Conda environment:
```bash
conda create -n pvdreader python=3.10
conda activate pvdreader
```

2. Install dependencies:
```bash
conda install numpy pandas matplotlib scipy
pip install struct logging
```

## Usage

The tool can be used from the command line:
```bash
python pvdreader.py <filename>
```

Or imported as a module:
```python
from pvdreader import PVDReader, SVDReader

# For PVD files
reader = PVDReader("measurement.pvd")
reader.read_file()
reader.plot_data(channel=0)  # Plot first channel

# For SVD files
reader = SVDReader("scan.svd")
reader.read_file()
reader.plot_data()
```

## Features
- Reads both PVD and SVD file formats
- Supports multiple data channels
- Automatic file format detection
- Time-domain and frequency-domain plotting
- Error handling and logging
- Command-line interface

## Note
This is an initial implementation and may need adjustments based on the actual file format specifications from Polytec.
