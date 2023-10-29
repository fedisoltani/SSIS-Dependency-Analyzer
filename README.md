<div align="center">
  <h1>SSIS Dependency Analyzer</h1>
  <p>
    <strong>Analyze SQL Server Integration Services (SSIS) packages and their dependencies.</strong>
  </p>
</div>

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)
- [Sample Output](#sample-output)
- [Authors](#authors)
- [Contributing](#contributing)
- [Feedback](#feedback)

## Overview

**SSIS Dependency Analyzer** is a Python script that simplifies the analysis of SQL Server Integration Services (SSIS) packages. It extracts valuable information about connection managers, related tables, and procedures used in your SSIS packages, making it easier to understand and document your data integration workflows. Whether you are managing a complex ETL process or want to gain insights into your SSIS packages, this script has you covered.

## Features

- :floppy_disk: **Connection Manager Data Extraction:** Extracts connection manager data from SSIS package files.
- :mag: **Dependency Identification:** Identifies related tables and procedures utilized within SSIS packages.
- :file_folder: **Output in Multiple Formats:** Generates a JSON file containing connection manager information and creates an Excel file with a detailed summary of SSIS package dependencies.

## Installation

Getting started with **SSIS Dependency Analyzer** is a straightforward process:

1. **Clone** this repository to your local machine:

   ```shell
   git clone https://github.com/fedisoltani/SSIS-Dependency-Analyzer.git
   ```

2. Navigate to the project directory:

   ```shell
   cd SSIS-Dependency-Analyzer
   ```

3. Install the necessary Python packages using `pip`:

   ```shell
   pip install pandas openpyxl
   ```

4. **Execute** the script:

   ```shell
   python main.py
   ```

   The script also requires the following libraries, which are part of the standard Python library and are included by default:

   - `os` (Operating System Interface): Used for interacting with the operating system.
   - `re` (Regular Expressions): A library for working with regular expressions.
   - `xml.etree.ElementTree`: A library for parsing XML files.
   - `pandas` (Python Data Analysis Library): A library for data manipulation and analysis.
   - `json` (JavaScript Object Notation): A library for handling JSON data.

## Usage

Analyzing your SSIS packages is just a few steps away:

1. Run the script by executing the `main.py` file:

   ```shell
   python main.py
   ```

2. Follow the prompts to specify the paths to your SSIS package files and the desired output file paths for both JSON and Excel files.

3. The script will process your SSIS packages and generate JSON and Excel files that provide you with connection manager data and a comprehensive summary of SSIS package dependencies.

## Configuration

Customizing the script to meet your specific requirements is possible. You can modify the `main.py` file to streamline your workflow further.

## Sample Output

The script generates two essential output files:

- `cnx_mngrs.json`: A JSON file containing detailed connection manager information.
- `dependencies.xlsx`: An Excel file summarizing SSIS package dependencies, including related connections, tables, and procedures.

## Authors

- **Fedi Soltani** - [GitHub](https://github.com/fedisoltani)

## Contributing

Contributions and suggestions to enhance this script are always welcome. Please visit the [GitHub repository](https://github.com/fedisoltani/SSIS-Dependency-Analyzer) to report issues or contribute to the project's development.

## Feedback

We encourage feedback and suggestions to improve this script. If you encounter any issues or have ideas for enhancements, please don't hesitate to [create an issue](https://github.com/fedisoltani/SSIS-Dependency-Analyzer/issues).

## Enjoy Analyzing Your SSIS Packages with SSIS Dependency Analyzer!

This script is provided as open-source software, and anyone is welcome to use it. Happy analyzing!
