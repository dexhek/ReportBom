# BOM Generator - .NET 8

#### 1. **File di Progetto (.csproj)**
- Convertito da formato legacy a **SDK-style**
- Target framework: `net8.0-windows`
- Abilitato `UseWindowsForms` e `ImplicitUsings`
- Gestione automatica dei pacchetti NuGet

### Come Compilare

```bash
# Debug
dotnet build

# Release
dotnet build --configuration Release

# Eseguire
dotnet run
```

### Struttura File

```
ReportBom/
├── Program.cs          # Logica principale e classi
├── BomItem.cs          # Modello dati BOM
├── ReportBom.csproj    # File progetto .NET 8
└── Properties/
    └── AssemblyInfo.cs # Informazioni assembly
```
