'use strict';

const fs = require('fs');
const path = require('path');
const db = require('../db');
const mongoStore = require('../mongo-store');

async function main() {
  if (!mongoStore.configured()) {
    throw new Error('Set MONGODB_URI before running the workbook migration.');
  }

  const snapshot = await db.readWorkbookSnapshot();
  const total = snapshot.packages.length + snapshot.requests.length + snapshot.numbers.length + snapshot.files.length;
  if (!total) {
    throw new Error(`No workbook records were found at ${db.DB_PATH}. Set DATABASE_PATH if the workbook is elsewhere.`);
  }

  const fileBuffers = new Map();
  let missingFiles = 0;
  for (const file of snapshot.files) {
    const storedPath = file.storagePath ? path.join(db.SHARED_DIR, file.storagePath) : '';
    if (storedPath && fs.existsSync(storedPath)) {
      fileBuffers.set(file.fileId, await fs.promises.readFile(storedPath));
    } else if (file.fileId) {
      missingFiles++;
    }
  }

  await mongoStore.importSnapshot(snapshot, fileBuffers);
  console.log(`Imported workbook records into MongoDB Atlas: ${snapshot.packages.length} plans, ${snapshot.requests.length} requests, ${snapshot.numbers.length} numbers, ${snapshot.files.length} file records.`);
  if (missingFiles) console.warn(`${missingFiles} logo file(s) were not found on disk; their metadata was imported but their content was not.`);
}

main().catch(error => {
  console.error('Workbook migration failed:', error);
  process.exitCode = 1;
});
