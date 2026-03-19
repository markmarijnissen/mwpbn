import JSZip from 'jszip'

/**
 * Scans a ZIP file for .xlsx files and extracts them.
 * * @param {File} zipFile - The zip file from the file input
 * @returns {Promise<Array<{filename: string, content: ArrayBuffer}>>}
 */
export async function unzip(zipFile) {
  const jszip = new JSZip()
  const zip = await jszip.loadAsync(zipFile)
  
  const extractedFiles = []
  const processingPromises = []

  // Iterate over each file inside the zip
  zip.forEach((relativePath, zipEntry) => {
    // Check if it's an .xlsx file and not a directory
    const isXlsx = relativePath.toLowerCase().endsWith('.xlsx')
    
    // Ignore macOS hidden metadata files
    const isHidden = relativePath.includes('__MACOSX') || relativePath.split('/').pop().startsWith('._')

    if (!zipEntry.dir && isXlsx && !isHidden) {
      // Extract the content as an ArrayBuffer
      const processFile = zipEntry.async('arraybuffer').then((content) => {
        extractedFiles.push({
          filename: zipEntry.name,
          content: content 
        })
      })
      
      processingPromises.push(processFile)
    }
  })

  // Wait for all files to finish extracting
  await Promise.all(processingPromises)
  
  return extractedFiles
}