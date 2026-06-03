 # ReportBom 🛠️📊

    [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
    [![Platform](https://img.shields.io/badge/Platform-Windows-blue.svg)]()
    [![Framework](https://img.shields.io/badge/.NET-C%23-purple.svg)]()

    **ReportBom** is an open-source utility designed to automate the extraction, traversal, and generation of structured Bills of Materials (BOM) directly from Solid Edge.

    By interfacing with CAD application programming interfaces, ReportBom recursively analyzes assembly structures, extracts component metadata, and outputs standardized, production-ready reports for  manufacturing, procurement, and enterprise lifecycle systems.

    ---

    ## ✨ Key Features

    * **Recursive Assembly Traversal:** Deep traversal of CAD assembly structures.
    * **Metadata & Property Extraction:** Extraction of custom properties, physical properties (such as mass and volume), and document attributes.
    * **Flexible Export Pipelines:** Generates clean structured data output ready for downstream enterprise and planning integrations.
    * **Component Grouping:** Automatic identification of unique parts and quantity summation across deeply nested sub-assemblies.

    ---

    ## 🛠️ Tech Stack & Prerequisites

    * **Language:** C#
    * **Framework:** .NET Core / .NET Framework (configured for CAD API interoperability)

    ---

    ## 🚀 Getting Started

    ### Installation
    Clone the repository:
    ```bash
    git clone https://github.com/dexhek/ReportBom.git
    cd ReportBom

  ### Build

  Build the project using standard build utilities:

    dotnet build
  ──────
  ## 🤝 Contributing

  Contributions are welcome!

  ## 📄 License

  This project is licensed under the MIT License - see the LICENSE file for details.
