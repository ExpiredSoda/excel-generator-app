// excel/utils/zipUtils.js
// Shared ZIP utility for creating Excel files

// Debug logs removed

export function createZip(files) {
  // Debug logs removed
  // Create a simple ZIP file structure
  // This is a minimal ZIP implementation for Excel files
  
  const centralDir = [];
  let offset = 0;
  let zipData = new Uint8Array(0);
  
  // Add each file to the ZIP
  files.forEach((file, index) => {
    const content = new TextEncoder().encode(file.content);
    const filename = file.name;
    
    // Local file header
    const localHeader = new Uint8Array(30 + filename.length);
    const view = new DataView(localHeader.buffer);
    
    view.setUint32(0, 0x04034b50, true); // Local file header signature
    view.setUint16(4, 20, true); // Version needed to extract
    view.setUint16(6, 0, true); // General purpose bit flag
    view.setUint16(8, 0, true); // Compression method (stored)
    view.setUint16(10, 0, true); // Last mod file time
    view.setUint16(12, 0, true); // Last mod file date
    view.setUint32(14, crc32(content), true); // CRC-32
    view.setUint32(18, content.length, true); // Compressed size
    view.setUint32(22, content.length, true); // Uncompressed size
    view.setUint16(26, filename.length, true); // File name length
    view.setUint16(28, 0, true); // Extra field length
    
    // Add filename
    const filenameBytes = new TextEncoder().encode(filename);
    localHeader.set(filenameBytes, 30);
    
    // Store central directory info
    centralDir.push({
      filename: filename,
      filenameLength: filename.length,
      offset: offset,
      crc32: crc32(content),
      compressedSize: content.length,
      uncompressedSize: content.length
    });
    
    // Append to ZIP data
    const newZipData = new Uint8Array(zipData.length + localHeader.length + content.length);
    newZipData.set(zipData);
    newZipData.set(localHeader, zipData.length);
    newZipData.set(content, zipData.length + localHeader.length);
    
    offset += localHeader.length + content.length;
    zipData = newZipData;
  });
  
  // Create central directory
  const centralDirStart = zipData.length;
  let centralDirData = new Uint8Array(0);
  
  centralDir.forEach(file => {
    const centralDirEntry = new Uint8Array(46 + file.filenameLength);
    const view = new DataView(centralDirEntry.buffer);
    
    view.setUint32(0, 0x02014b50, true); // Central directory file header signature
    view.setUint16(4, 20, true); // Version made by
    view.setUint16(6, 20, true); // Version needed to extract
    view.setUint16(8, 0, true); // General purpose bit flag
    view.setUint16(10, 0, true); // Compression method
    view.setUint16(12, 0, true); // Last mod file time
    view.setUint16(14, 0, true); // Last mod file date
    view.setUint32(16, file.crc32, true); // CRC-32
    view.setUint32(20, file.compressedSize, true); // Compressed size
    view.setUint32(24, file.uncompressedSize, true); // Uncompressed size
    view.setUint16(28, file.filenameLength, true); // File name length
    view.setUint16(30, 0, true); // Extra field length
    view.setUint16(32, 0, true); // File comment length
    view.setUint16(34, 0, true); // Disk number start
    view.setUint16(36, 0, true); // Internal file attributes
    view.setUint32(38, 0, true); // External file attributes
    view.setUint32(42, file.offset, true); // Relative offset of local header
    
    // Add filename
    const filenameBytes = new TextEncoder().encode(file.filename);
    centralDirEntry.set(filenameBytes, 46);
    
    // Append to central directory
    const newCentralDirData = new Uint8Array(centralDirData.length + centralDirEntry.length);
    newCentralDirData.set(centralDirData);
    newCentralDirData.set(centralDirEntry, centralDirData.length);
    centralDirData = newCentralDirData;
  });
  
  // End of central directory record
  const endOfCentralDir = new Uint8Array(22);
  const endView = new DataView(endOfCentralDir.buffer);
  
  endView.setUint32(0, 0x06054b50, true); // End of central dir signature
  endView.setUint16(4, 0, true); // Number of this disk
  endView.setUint16(6, 0, true); // Number of the disk with the start of the central directory
  endView.setUint16(8, centralDir.length, true); // Total number of entries in the central directory on this disk
  endView.setUint16(10, centralDir.length, true); // Total number of entries in the central directory
  endView.setUint32(12, centralDirData.length, true); // Size of central directory
  endView.setUint32(16, centralDirStart, true); // Offset of start of central directory
  endView.setUint16(20, 0, true); // .ZIP file comment length
  
  // Combine everything
  const finalZip = new Uint8Array(zipData.length + centralDirData.length + endOfCentralDir.length);
  finalZip.set(zipData);
  finalZip.set(centralDirData, zipData.length);
  finalZip.set(endOfCentralDir, zipData.length + centralDirData.length);
  
  return finalZip.buffer;
}

// Simple CRC32 implementation
function crc32(data) {
  const table = makeCRCTable();
  let crc = 0 ^ (-1);
  
  for (let i = 0; i < data.length; i++) {
    crc = (crc >>> 8) ^ table[(crc ^ data[i]) & 0xFF];
  }
  
  return (crc ^ (-1)) >>> 0;
}

function makeCRCTable() {
  let c;
  const crcTable = [];
  for (let n = 0; n < 256; n++) {
    c = n;
    for (let k = 0; k < 8; k++) {
      c = ((c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1));
    }
    crcTable[n] = c;
  }
  return crcTable;
}