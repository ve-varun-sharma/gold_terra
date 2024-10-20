import { verifyEmail } from "@devmehq/email-validator-js";
import * as xlsx from "xlsx";
import * as path from "path";
import {
  parsePhoneNumberFromString,
  type CountryCode,
  isValidPhoneNumber,
} from "libphonenumber-js";

import {
  GoogleGenerativeAI,
  HarmCategory,
  HarmBlockThreshold,
} from "@google/generative-ai";

async function validateEmail(email: string): Promise<boolean> {
  try {
    const { validFormat, validMx } = await verifyEmail({
      emailAddress: email,
      verifyMx: true,
    });

    console.log(validFormat, validMx);
    if (!validFormat || !validMx) {
      console.log("Invalid Email for: ", email);
      return false;
    }
    console.log("Valid email for : ", email);
    return true;
  } catch (error) {
    throw error;
  }
}
// Function to process the Excel file
async function processEmails(filePath: string) {
  const absolutePath = path.resolve(__dirname, filePath);

  // Read the Excel file
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Convert the worksheet to JSON
  const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

  // Iterate through the email column and validate emails
  for (let i = 1; i < data.length; i++) {
    // @ts-expect-error
    const email = data[i][6]; // Assuming emails are in the first column
    if (typeof email === "string") {
      const isValid = await validateEmail(email);
      // @ts-expect-error
      data[i][7] = isValid; // Write the result to the next column
    }
  }

  // Convert the JSON back to a worksheet
  // @ts-expect-error
  const newWorksheet = xlsx.utils.aoa_to_sheet(data);

  // Replace the old worksheet with the new one
  workbook.Sheets[sheetName] = newWorksheet;

  // Save the updated Excel file
  xlsx.writeFile(workbook, filePath);
}

// processExcelFile(fullMiningList)
//   .then(() => {
//     console.log("Email validation completed and results saved.");
//   })
//   .catch((error) => {
//     console.error("Error processing Excel file:", error);
//   });

// Function to clean phone numbers
function cleanPhoneNumber(phone: string): string {
  if (!phone) return "";
  const cleanedPhone = phone.replace(/\D/g, ""); // Remove all non-numeric characters
  return cleanedPhone;
}

// Function to validate phone numbers without specifying a country code
function validatePhoneNumber(phone: string): boolean {
  if (!phone) return false;

  // Try to parse the phone number with the default country code (US)
  let phoneNumber = parsePhoneNumberFromString(phone, "US");

  // If the phone number is not valid, try to parse it without specifying a country code
  if (!phoneNumber || !phoneNumber.isValid()) {
    phoneNumber = parsePhoneNumberFromString(phone);
  }

  // Check if the phone number is valid
  return phoneNumber ? phoneNumber.isValid() : false;
}
// New function to process the Excel file for phone numbers
async function processPhoneNumbers(filePath: string) {
  // Resolve the absolute path
  const absolutePath = path.resolve(__dirname, filePath);

  // Read the Excel file
  const workbook = xlsx.readFile(absolutePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Convert the worksheet to JSON
  const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

  // Iterate through the phone column and clean phone numbers
  for (let i = 1; i < data.length; i++) {
    // Convert phone number to string if it's a number

    // @ts-expect-error
    let phone = data[i][2]; // Assuming phone numbers are in the second column
    if (typeof phone === "number") {
      phone = phone.toString();
    }

    if (typeof phone === "string") {
      const cleanedPhone = cleanPhoneNumber(phone);
      const isValid = validatePhoneNumber(cleanedPhone);
      // @ts-expect-error
      data[i][3] = phone === cleanedPhone ? phone : cleanedPhone; // Copy the original if already cleaned, otherwise use cleaned
      // @ts-expect-error
      data[i][4] = isValid; // Write the validation result to the fifth column
    } else {
      // @ts-expect-error
      data[i][3] = ""; // Leave the cell blank if the phone number is not a string
      // @ts-expect-error
      data[i][4] = false; // Mark as invalid
    }
  }

  // Convert the JSON back to a worksheet
  // @ts-expect-error
  const newWorksheet = xlsx.utils.aoa_to_sheet(data);

  // Replace the old worksheet with the new one
  workbook.Sheets[sheetName] = newWorksheet;

  // Save the updated Excel file
  xlsx.writeFile(workbook, absolutePath);
}

const projectId = "experimentation-learning"; // Replace with your project ID
const location = "us-central1"; // Replace with your desired location
const modelId = "gemini-1.5-flash-8b"; // Replace with the desired model ID

// Create a new client
const apiKey = process.env.GEMINI_API_KEY;
const genAI = new GoogleGenerativeAI(apiKey as string);

const model = genAI.getGenerativeModel({
  model: modelId,
});

// Define the prompt template
const promptTemplate = `
You are a country code expert. 
Please analyze the following data:
**State:** {state} 
**Country:** {country}

Based on the provided information, determine the **Alpha-2 Country Code** for the given state or country.
State or Country may be missing, and if it is, use the non-Unknown State or use the non-Unknown country depending which is available for your analysis.

If the state is unknown, use the country information to determine the Alpha-2 Country Code.
If the country is unknown, use the state information to determine the Alpha-2 Country Code.
If both state and  country are unknown, return "Unknown".

**Important Note:** The country code must be in Alpha-2 format, which is a two-letter abbreviation. 
Only return the Alpha-2 format code.
`;

// Function to process data and generate Alpha-2 Country Codes
async function generateAlpha2Codes(
  state: string,
  country: string
): Promise<string> {
  // Construct the prompt with the provided state and country
  const prompt = promptTemplate
    .replace("{state}", state || "Unknown")
    .replace("{country}", country || "Unknown");

  try {
    const result = await model.generateContent([prompt]);

    // Extract the generated Alpha-2 Country Code from the response
    const alpha2Code = result.response.text();
    console.log(alpha2Code);

    return alpha2Code;
  } catch (error) {
    console.error("Error generating Alpha-2 Country Code:", error);
    return "Unknown";
  }
}

// Function to process the Excel file for state and country analysis
async function processStateCountry(filePath: string) {
  // Resolve the absolute path
  const absolutePath = path.resolve(__dirname, filePath);

  // Read the Excel file
  const workbook = xlsx.readFile(absolutePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Convert the worksheet to JSON
  const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

  // Iterate through the state and country columns
  for (let i = 1; i < data.length; i++) {
    // @ts-expect-error
    const state = data[i][5]; // Assuming states are in the first column
    // @ts-expect-error
    const country = data[i][6]; // Assuming countries are in the second column

    if (typeof state === "string" || typeof country === "string") {
      const countryCode = await generateAlpha2Codes(state, country);
      // @ts-expect-error
      data[i][7] = countryCode || ""; // Write the cleaned country code to the third column
    } else {
      // @ts-expect-error
      data[i][7] = ""; // Leave the cell blank if both state and country are not strings
    }
  }

  // Convert the JSON back to a worksheet
  // @ts-expect-error
  const newWorksheet = xlsx.utils.aoa_to_sheet(data);

  // Replace the old worksheet with the new one
  workbook.Sheets[sheetName] = newWorksheet;

  // Save the updated Excel file
  xlsx.writeFile(workbook, absolutePath);
}

// Function to count rows with both specified columns having the value true
async function countRowsWithBothColumnsTrue(
  filePath: string,
  col1Index: number,
  col2Index: number
): Promise<number> {
  // Resolve the absolute path
  const absolutePath = path.resolve(__dirname, filePath);

  // Read the Excel file
  const workbook = xlsx.readFile(absolutePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Convert the worksheet to JSON
  const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

  // Initialize the counter
  let count = 0;

  // Iterate through the rows and count the rows with both columns true
  for (let i = 1; i < data.length; i++) {
    // @ts-expect-error
    const col1Value = data[i][col1Index];
    // @ts-expect-error
    const col2Value = data[i][col2Index];

    if (col1Value === true && col2Value === true) {
      count++;
    }
  }

  console.log(count);
  return count;
}

// Example usage

// Main function to call both email validation and phone number cleaning
async function main() {
  const testRunFilePath =
    "/Users/ve/fractional_cto_projects/gold_terra_mining/data/test_run.xlsx";
  const fullMiningList =
    "/Users/ve/fractional_cto_projects/gold_terra_mining/data/gold_terra_mining_list.xlsx";

  const col1Index = 3; // Index of the first column to check
  const col2Index = 9; // Index of the second column to check

  try {
    // await processEmails(testRunFilePath);
    // await processPhoneNumbers(fullMiningList);
    // await processStateCountry(fullMiningList);
    const count = await countRowsWithBothColumnsTrue(
      fullMiningList,
      col1Index,
      col2Index
    );

    console.log(
      "Email validation and phone number cleaning completed and results saved."
    );
  } catch (error) {
    console.error("Error processing Excel file:", error);
  }
}

// Example usage
main();
