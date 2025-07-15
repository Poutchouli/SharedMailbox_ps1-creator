# Exchange Mailbox Permissions Script Generator

A Streamlit web application that generates PowerShell scripts for managing Exchange mailbox permissions based on CSV input.

## Features

- Upload CSV files with mailbox permission requirements
- Review and select specific actions to include in the script
- Generate PowerShell scripts for adding/removing mailbox permissions
- Download generated scripts as `.ps1` files

## CSV Format

The CSV file should use semicolon (`;`) as delimiter and contain the following columns:
- `Identity`: Mailbox identity (email address or name)
- `User`: User/trustee to grant or remove permissions
- `Action`: Either "Add" or "Remove"

Example:
```csv
Identity;User;Action
shared.mailbox@company.com;john.doe@company.com;Add
sales@company.com;jane.smith@company.com;Remove
```

## Running with Docker

### Using Docker Compose (Recommended)

1. Clone or download this repository
2. Navigate to the project directory
3. Run the application:
   ```bash
   docker-compose up -d
   ```
4. Access the application at http://localhost:17654

### Using Docker directly

1. Build the image:
   ```bash
   docker build -t mailbox-script-generator .
   ```

2. Run the container:
   ```bash
   docker run -p 17654:17654 mailbox-script-generator
   ```

3. Access the application at http://localhost:17654

## Running Locally

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Run the application:
   ```bash
   streamlit run app.py
   ```

## Generated PowerShell Script

The generated script will:
1. Connect to Exchange Online
2. Process each selected action (Add/Remove permissions)
3. Grant or remove both FullAccess and SendAs permissions
4. Disconnect from Exchange Online

## Requirements

- Python 3.11+
- Streamlit
- Pandas
- Docker (for containerized deployment)

## License

This project is open source and available under the MIT License.
