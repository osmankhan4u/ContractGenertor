# ContractGeneratorBlazor

A Blazor Server application for automated contract generation from Excel data and Word templates.

## Features

- **Contract Type Management**
  - Add, edit, and delete contract types via the web UI
  - Contract types and placeholders stored in `contracts.json`
  - Upload Word template for each contract type

- **Contract Generation**
  - Upload Excel file with contract data
  - Select contract type to generate contracts
  - Placeholders in Word template are replaced with Excel data
  - Generates individual Word documents for each row
  - All generated documents are packaged into a downloadable ZIP file

- **Frontend (Blazor UI)**
  - Easy-to-use interface for uploading Excel files and selecting contract types
  - Progress indicator during contract generation
  - Download link for generated contracts
  - Separate page for contract type management

- **Backend (C# Services)**
  - Uses DocumentFormat.OpenXml for Word document manipulation
  - Uses ClosedXML for Excel file reading
  - All contract logic handled in `ContractService`

- **Configuration**
  - Contract templates and placeholder mappings managed in `contracts.json`
  - Template files stored in `Contracts/` folder

## How to Use

1. Start the application (`dotnet run`).
2. Go to the Contract Generator page.
3. Upload an Excel file and select a contract type.
4. Click Generate to create contracts.
5. Download the ZIP file containing all generated Word documents.
6. Manage contract types and templates from the Contract Type Management page.

## Project Structure

```
ContractGeneratorBlazor/
├── Models/
│   └── ContractConfig.cs
├── Pages/
│   ├── ContractGenerator.razor
│   └── AddContractType.razor
├── Services/
│   ├── ContractService.cs
│   └── DocumentGenerator.cs
├── Contracts/
│   ├── contracts.json
│   └── [template files]
├── appsettings.json
├── Program.cs
└── README.md
```

## Technologies Used
- Blazor Server (.NET 7)
- DocumentFormat.OpenXml
- ClosedXML

## License
This project uses only free and open-source libraries.
