![GitHub](https://img.shields.io/github/license/kolenichsj/stego-csharp)

# stego-csharp

This library allows you to use [steganography](https://wikipedia.org/wiki/Steganography) to insert data into an overt file such as HTML or DOCX.

## Build
```
dotnet build
```
```
dotnet run --project ./hips
```

### Tests
```
dotnet test --collect:"XPlat Code Coverage"
```

## Usage:
```
hips [options] [command]
```  

#### Options:

  --version         Show version information
  
  -?, -h, --help    Show help and usage information

#### Commands:
- insert: Insert covert data into overt file
- generate: Create overt file with covert data inserted
- extract: Extract covert data from overt file