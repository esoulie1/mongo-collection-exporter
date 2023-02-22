# Mongo Collection Exporter

This Node.js script exports all collections of a MongoDB database to a single Excel (.xlsx) file.

## Usage

1. Clone this repository.
2. Install the required packages using `npm install`.
3. Create a `.env` file in the project directory with the following environment variables:
```
MONGODB_URI=<your MongoDB URI>
DB_NAME=<your database name>
```
4. Run the script using `npm start`. The exported Excel file will be saved to the project directory as `output.xlsx`.

## Requirements

- Node.js (version 18)
- A MongoDB database
- A MongoDB URI (in the format `mongodb://<username>:<password>@<hostname>:<port>`)
- The `xlsx`, `mongodb` and `dotenv` packages (installed automatically by running `npm install`)
