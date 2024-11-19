# PVD Reader

A Python tool for reading and processing Polytec PVD (Polytec Vibrometer Data) files using the Polytec ScanViewer ActiveX interface.

## Prerequisites

- Python 3.9+
- Polytec Scan Viewer 2.9 installed
- Conda environment manager

## Installation

1. Create and activate the conda environment:
```bash
conda create -n pvdreader python=3.9
conda activate pvdreader
```

2. Install required packages:
```bash
conda install -c conda-forge pywin32
```

## Usage

Run the script with a PVD file path:
```bash
python pvd_reader.py path/to/your/file.pvd
```

The script will:
1. Initialize the Polytec ScanViewer ActiveX control
2. Load the specified PVD file
3. Extract available data and metadata
4. Display the extracted information

## Features

- Automatic COM interface initialization and cleanup
- Multiple file loading method attempts
- Comprehensive property and method exploration
- Flexible data extraction with multiple getter method attempts
- Clean error handling and reporting

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License

[MIT](https://choosealicense.com/licenses/mit/)
