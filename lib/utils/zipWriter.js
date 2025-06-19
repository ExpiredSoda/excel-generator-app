// utils/zipWriter.js
// Minimal ZIP writer for Excel file export

// Helper: UTF-8 encoding for string to bytes
function utf8ToBytes(str) {
  const encoder = new TextEncoder();
  return Array.from(encoder.encode(str));
}

function stringToBytes(str) {
  return utf8ToBytes(str); // Use proper UTF-8 encoding instead of charCodeAt
}

function toBytesLE(num, len) {
  const arr = [];
  for (let i = 0; i < len; i++) {
    arr.push(num & 0xff);
    num >>= 8;
  }
  return arr;
}

export function createZip(files) {
  let offset = 0;
  let allData = [];
  let centralDir = [];
  
  files.forEach(file => {
    // Use UTF-8 encoding for both filename and content
    const fileBytes = stringToBytes(file.content);
    const fileLen = fileBytes.length;
    const filenameBytes = stringToBytes(file.name);
    
    const localHeader = [
      0x50,0x4b,0x03,0x04, // Local file header signature
      0x14,0x00, // Version needed to extract
      0x00,0x00, // General purpose bit flag
      0x00,0x00, // Compression method (0 = no compression)
      0x00,0x00,0x00,0x00, // Last mod file time and date
      0x00,0x00,0x00,0x00, // CRC-32
      ...toBytesLE(fileLen, 4), // Compressed size
      ...toBytesLE(fileLen, 4), // Uncompressed size
      ...toBytesLE(filenameBytes.length, 2), // File name length
      0x00,0x00 // Extra field length
    ];
    
    const local = [...localHeader, ...filenameBytes, ...fileBytes];
    allData.push(...local);
    
    const central = [
      0x50,0x4b,0x01,0x02, // Central directory file header signature
      0x14,0x00, // Version made by
      0x14,0x00, // Version needed to extract
      0x00,0x00, // General purpose bit flag
      0x00,0x00, // Compression method
      0x00,0x00,0x00,0x00, // Last mod file time and date
      0x00,0x00,0x00,0x00, // CRC-32
      ...toBytesLE(fileLen, 4), // Compressed size
      ...toBytesLE(fileLen, 4), // Uncompressed size
      ...toBytesLE(filenameBytes.length, 2), // File name length
      0x00,0x00, // Extra field length
      0x00,0x00, // File comment length
      0x00,0x00, // Disk number start
      0x00,0x00, // Internal file attributes
      0x00,0x00,0x00,0x00, // External file attributes
      ...toBytesLE(offset, 4) // Relative offset of local header
    ];
    
    centralDir.push(...central, ...filenameBytes);
    offset += local.length;
  });

  const centralDirLen = centralDir.length;
  const allDataLen = allData.length;
  
  const endCentral = [
    0x50,0x4b,0x05,0x06, // End of central directory signature
    0x00,0x00, // Number of this disk
    0x00,0x00, // Disk where central directory starts
    ...toBytesLE(files.length, 2), // Number of central directory records on this disk
    ...toBytesLE(files.length, 2), // Total number of central directory records
    ...toBytesLE(centralDirLen, 4), // Size of central directory
    ...toBytesLE(allDataLen, 4), // Offset of start of central directory
    0x00,0x00 // ZIP file comment length
  ];
  
  return new Uint8Array([...allData, ...centralDir, ...endCentral]);
}

// createZip and ZIP utilities in zipWriter.js match the logic in script.js. No changes needed.
