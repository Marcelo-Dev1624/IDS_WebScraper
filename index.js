const XLSX = require("xlsx");  // Importing xlsx library for Excel file operations
const fs = require("fs");  // Importing fs (file system) module for file operations
const path = require("path");  // Importing path module for handling file paths
const puppeteer = require("puppeteer");  // Importing puppeteer for web scraping
const Tesseract = require("tesseract.js");  // Importing Tesseract.js for OCR (Optical Character Recognition)
const sharp = require("sharp");  // Importing sharp for image processing

// Function to read and normalize data from an Excel file
// Function to read and normalize data from an Excel file
function readExcelFile(filePath) {
  if (!fs.existsSync(filePath)) {  // Check if file exists at the given path
    throw new Error(`File not found: ${filePath}`);
  }
  
  const workbook = XLSX.readFile(filePath);  // Read the Excel file
  const sheetName = workbook.SheetNames[0];  // Get the first sheet name
  const worksheet = workbook.Sheets[sheetName];  // Get the worksheet by its name
  let data = XLSX.utils.sheet_to_json(worksheet);  // Convert the worksheet to JSON array

  // Normalize data to ensure all values are strings
  data = data.map(row => {
    return {
        Aduana: row['Aduana'] ? row['Aduana'].toString() : '',  // Convert Aduana to string or empty string if undefined
        'Año': row['Año'] ? row['Año'].toString() : '',  // Convert Año to string or empty string if undefined
        Numero: row['Numero'] ? row['Numero'].toString() : '',  // Convert Numero to string or empty string if undefined
        Item: row['Item'] ? row['Item'].toString() : '',  // Convert Item to string or empty string if undefined
        Fecha: row['Fecha'] ? row['Fecha'].toString() : '',  // Convert Fecha to string or empty string if undefined
        Aduana_1: row['Aduana_1'] ? row['Aduana_1'].toString() : '',  // Convert Aduana_1 to string or empty string if undefined
        Regimen: row['Regimen'] ? row['Regimen'].toString() : '',  // Convert Regimen to string or empty string if undefined
        Modalidad: row['Modalidad'] ? row['Modalidad'].toString() : '',  // Convert Modalidad to string or empty string if undefined
        Importador: row['Importador'] ? row['Importador'].toString() : '',  // Convert Importador to string or empty string if undefined
        Marca: row['Marca'] ? row['Marca'].toString() : '',  // Convert Marca to string or empty string if undefined
        Modelo: row['Modelo'] ? row['Modelo'].toString() : '',  // Convert Modelo to string or empty string if undefined
        Factura: row['Factura'] ? row['Factura'].toString() : '',  // Convert Factura to string or empty string if undefined
        'Código SAC': row['Código SAC'] ? row['Código SAC'].toString() : '',  // Convert Código SAC to string or empty string if undefined
        'Vía Transporte': row['Vía Transporte'] ? row['Vía Transporte'].toString() : '',  // Convert Vía Transporte to string or empty string if undefined
        'País de Origen': row['País de Origen'] ? row['País de Origen'].toString() : '',  // Convert País de Origen to string or empty string if undefined
        'Pais de Procedencia': row['Pais de Procedencia'] ? row['Pais de Procedencia'].toString() : '',  // Convert Pais de Procedencia to string or empty string if undefined
        'Pais de Adquisición': row['Pais de Adquisición'] ? row['Pais de Adquisición'].toString() : '',  // Convert Pais de Adquisición to string or empty string if undefined
        'Cantidad Comercial': row['Cantidad Comercial'] ? row['Cantidad Comercial'].toString() : '',  // Convert Cantidad Comercial to string or empty string if undefined
        'Unidad de Medida': row['Unidad de Medida'] ? row['Unidad de Medida'].toString() : '',  // Convert Unidad de Medida to string or empty string if undefined
        Bultos: row['Bultos'] ? row['Bultos'].toString() : '',  // Convert Bultos to string or empty string if undefined
        'U$S FOB': row['U$S FOB'] ? row['U$S FOB'].toString() : '',  // Convert U$S FOB to string or empty string if undefined
        'U$S FOB, Unit.': row['U$S FOB, Unit.'] ? row['U$S FOB, Unit.'].toString() : '',  // Convert U$S FOB, Unit. to string or empty string if undefined
        'U$S Flete': row['U$S Flete'] ? row['U$S Flete'].toString() : '',  // Convert U$S Flete to string or empty string if undefined
        'U$S Seguro': row['U$S Seguro'] ? row['U$S Seguro'].toString() : '',  // Convert U$S Seguro to string or empty string if undefined
        'U$S CIF': row['U$S CIF'] ? row['U$S CIF'].toString() : '',  // Convert U$S CIF to string or empty string if undefined
        'U$S Unitario': row['U$S Unitario'] ? row['U$S Unitario'].toString() : '',  // Convert U$S Unitario to string or empty string if undefined
        'KGS. Netos': row['KGS. Netos'] ? row['KGS. Netos'].toString() : '',  // Convert KGS. Netos to string or empty string if undefined
        'Kgs. Brutos': row['Kgs. Brutos'] ? row['Kgs. Brutos'].toString() : '',  // Convert Kgs. Brutos to string or empty string if undefined
        'Descripción de Mercadería': row['Descripción de Mercadería'] ? row['Descripción de Mercadería'].toString() : '',  // Convert Descripción de Mercadería to string or empty string if undefined
        'Localización Actual': row['Localización Actual'] ? row['Localización Actual'].toString() : '',  // Convert Localización Actual to string or empty string if undefined
        'Localización Destino': row['Localización Destino'] ? row['Localización Destino'].toString() : ''  // Convert Localización Destino to string or empty string if undefined
    };
  });

  // Find the first record without 'Localización Actual' and 'Localización Destino'
  const incompleteRecord = data.find(row =>
    !row['Localización Actual'] || !row['Localización Destino']
  );

  if (incompleteRecord) {
    console.log("First incomplete record found:", incompleteRecord);
  } else {
    console.log("All records are complete.");
  }

  return data;
}


// Function to update a specific column in the Excel file
function updateExcelFile(filePath, updatedData) {
  if (!fs.existsSync(filePath)) {  // Check if file exists at the given path
    throw new Error(`File not found: ${filePath}`);
  }

  const workbook = XLSX.readFile(filePath);  // Read the Excel file
  const sheetName = workbook.SheetNames[0];  // Get the first sheet name
  const worksheet = workbook.Sheets[sheetName];  // Get the worksheet by its name
  const data = XLSX.utils.sheet_to_json(worksheet);  // Convert the worksheet to JSON array

  updatedData.forEach(updatedRow => {
    const index = data.findIndex(row => row["Numero"] === updatedRow["Numero"]);  // Find the index of the row to update
    if (index !== -1) {
      data[index] = { ...data[index], ...updatedRow };  // Merge updatedRow into data[index]
    }
  });

  const updatedWorksheet = XLSX.utils.json_to_sheet(data);  // Convert updated JSON data back to worksheet
  workbook.Sheets[sheetName] = updatedWorksheet;  // Update the worksheet in the workbook
  XLSX.writeFile(workbook, filePath);  // Write the updated workbook back to file
}

// Function to solve the captcha
async function solveCaptcha(page) {
  const captchaDimensions = await page.evaluate(() => {
    const captchaImageElement = document.querySelector("#captchaImage img");  // Get captcha image element
    if (!captchaImageElement) return null;
    const { x, y, width, height } = captchaImageElement.getBoundingClientRect();  // Get dimensions of captcha image
    return { x, y, width, height };
  });

  if (!captchaDimensions) {
    throw new Error("Captcha dimensions not found");
  }

  console.log("Captcha Dimensions: ", captchaDimensions);

  const captchaPath = path.join(__dirname, "/Screenshots/captcha.png");  // Path to save original captcha image
  await page.screenshot({
    path: captchaPath,
    clip: {
      x: captchaDimensions.x,
      y: captchaDimensions.y,
      width: captchaDimensions.width,
      height: captchaDimensions.height,
    },
  });

  console.log(`Captcha screenshot saved at ${captchaPath}`);

  const processedImagePath = path.join(__dirname, "/Screenshots/processed_captcha.png");  // Path to save processed captcha image
  await sharp(captchaPath)
    .grayscale()
    .normalize()
    .linear(1.2, 0) // Adjust contrast
    .modulate({ brightness: 1.2 }) // Adjust brightness
    .blur(1) // Apply blur to reduce noise
    .toFile(processedImagePath);

  console.log(`Processed Captcha image saved at ${processedImagePath}`);

  const captchaText = await new Promise((resolve, reject) => {
    Tesseract.recognize(processedImagePath, 'eng', {
      tessedit_char_whitelist: 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'  // Set whitelist characters for recognition
    })
      .then((result) => {
        resolve(result.data.text.trim());  // Resolve recognized text
      })
      .catch((err) => {
        reject(err);  // Reject on error
      });
  }).catch(err => {
    console.error("Error in Tesseract recognition:", err);
    return "";
  });

  console.log("Captcha Text: ", captchaText);

  if (!captchaText) {
    throw new Error("Captcha text could not be recognized");
  }

  await page.type("#_cfield", captchaText);  // Type recognized captcha text into input field
  console.log("Captcha Input: ", captchaText);

  const currentUrl = page.url();  // Get current page URL
  await page.click('input[name="DETALLE"]');  // Click on the input button

  try {
    // Wait for the URL to change or timeout after 5 seconds
    await page.waitForFunction(
      `window.location.href !== "${currentUrl}"`,
      { timeout: 5000 } // Reduced to 5 seconds
    );

    console.log("Successfully navigated to the next page");
    return true;
  } catch (error) {
    console.log("Navigation failed, refreshing page...");
    return false;
  }
}

// Function to scrape data from the website
async function scrapeData(page, record) {
  const colAduana = record["Aduana"];  // Get Aduana from record
  const colAno = record["Año"];  // Get Año from record
  const colNumero = record["Numero"];  // Get Numero from record

  await page.type('#vVCODI_ADUA', colAduana);  // Type Aduana into input field
  await page.type('#vVANO_PRE', colAno);  // Type Año into input field
  await page.type('#vVNUME_CORR', colNumero);  // Type Numero into input field

  console.log("Aduana: " + colAduana);
  console.log("Año: " + colAno);
  console.log("Numero: " + colNumero);
  

  let captchaSolved = false;
  while (!captchaSolved) {
    captchaSolved = await solveCaptcha(page);  // Solve captcha

    if (captchaSolved) {
      const localizacionActualElement = await page.evaluate(() => {
        const codAlma = document.querySelector('#span_CODI_ALMA');  // Get Localización Actual element
        const drSocial = document.querySelector('#span_vVDRSOCIAL');
        return codAlma && drSocial ? `${codAlma.textContent}-${drSocial.textContent}` : null;  // Concatenate Localización Actual
      });

      const localizacionDestinoElement = await page.evaluate(() => {
        const aduDest = document.querySelector('#span_vVCALMDEST');  // Get Localización Destino element
        const aduDsc = document.querySelector('#span_vVRGRSOC');
        return aduDest && aduDsc ? `${aduDest.textContent}-${aduDsc.textContent}` : null;  // Concatenate Localización Destino
      });

      if (localizacionActualElement && localizacionDestinoElement) {
        console.log(`Localización Actual: ${localizacionActualElement}`);
        console.log(`Localización Destino: ${localizacionDestinoElement}`);
        
        const updatedData = [
          {
            Numero: colNumero,
            'Localización Actual': localizacionActualElement,
            'Localización Destino': localizacionDestinoElement
          }
        ];
        
        updateExcelFile('Penta.xlsx', updatedData);  // Update Excel file with Localización Actual and Localización Destino
        return true;
      } else {
        console.log("No data found on the page.");
      }
    }

    // Refresh the page if captcha was incorrect or data was not found
    console.log("Refreshing page...");
    await page.reload({ waitUntil: ["networkidle0", "domcontentloaded"] });
    await page.waitForTimeout(5000); // Wait a bit before retrying
  }

  return false;
}

// Main function to control the workflow
async function main() {
  const filePath = path.join(__dirname, 'Penta.xlsx');  // Define Excel file path
  
  // Check if file exists before proceeding
  if (!fs.existsSync(filePath)) {
    console.error(`File not found: ${filePath}`);
    return;
  }

  const data = readExcelFile(filePath);  // Read data from Excel file

  const browser = await puppeteer.launch({ headless: true, defaultViewport: null });  // Launch Puppeteer browser
  const page = await browser.newPage();  // Create new page instance

  for (let i = 0; i < data.length; i++) {
    try {
      await page.goto('https://ticaconsultas.hacienda.go.cr/Tica/hcimppon.aspx');  // Navigate to the specified URL
      const success = await scrapeData(page, data[i]);  // Scrape data for each record in data array

      if (success) {
        console.log(`Record ${i + 1} processed successfully.`);
      } else {
        console.log(`Record ${i + 1} failed to process. Retrying...`);
        i--; // Decrement the counter to retry the current record
      }

    } catch (err) {
      console.error("Error during processing:", err);
      i--; // Decrement the counter to retry the current record
    }
  }

  await browser.close();  // Close the Puppeteer browser
  console.log("All data processed.");
}

// Execute the main function and handle errors
main().catch(err => console.error(err));
