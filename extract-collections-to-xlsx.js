require('dotenv').config();

const { MongoClient, ObjectId } = require('mongodb');
const XLSX = require('xlsx');
const fs = require('fs');

const {
  MONGODB_URI: url,
  DB_NAME: dbName,
} = process.env;

const fileName = 'output.xlsx';

if (!url || !dbName || !fileName) {
  throw new Error('Missing environment variables');
}

const client = new MongoClient(url);

async function main() {
  await client.connect();

  const db = client.db(dbName);

  const collections = await db.listCollections().toArray();

  const workbook = XLSX.utils.book_new();

  for await (const collection of collections) {
    const worksheetName = collection.name;

    const col = await db.collection(worksheetName).find().toArray();

    const colWithId = col.map((c) => {
      const keys = Object.keys(c);

      return keys.reduce((acc, key) => {
        if (ObjectId.isValid(c[key]) && Object.prototype.hasOwnProperty.call(c, key)) {
          acc[key] = c[key].toString();
        } else {
           if (Array.isArray(c[key])) {
             acc[key] = c[key].map((ck) => ObjectId.isValid(ck) ? ck.toString() : ck).join(', ');
           } else {
             acc[key] = c[key];
           }
        }

        return acc;
      }, {});
    });

    const worksheet = XLSX.utils.json_to_sheet(colWithId);

    XLSX.utils.book_append_sheet(workbook, worksheet, worksheetName);

    const isLast = collections.findIndex((co => co.name === worksheetName)) === (collections.length - 1);

    if (isLast) {
      // Save the workbook to a file when all collections have been processed
      const buffer = XLSX.write(workbook, {
        bookType: 'xlsx',
        type: 'buffer'
      });

      fs.writeFileSync(fileName, buffer);

      console.log('Excel file saved to ' + fileName);

      return 'done'
    }
  }
}

main()
  .then(console.log)
  .catch(console.error)
  .finally(() => client.close());
