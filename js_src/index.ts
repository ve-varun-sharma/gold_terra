import { verifyEmail } from "@devmehq/email-validator-js";
import * as xlsx from "xlsx";
import * as path from "path";

async function validateEmail(email: string): Promise<boolean> {
  try {
    const { validFormat, validSmtp, validMx } = await verifyEmail({
      emailAddress: email,
      verifyMx: true,
      verifySmtp: true,
      timeout: 3000,
    });

    console.log(validFormat, validMx, validSmtp);
    if (!validFormat || !validMx || !validSmtp) {
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
async function processExcelFile(filePath: string) {
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

// Example usage
processExcelFile(
  "/Users/ve/fractional_cto_projects/gold_terra_mining/gold_terra_mining_list.xlsx"
)
  .then(() => {
    console.log("Email validation completed and results saved.");
  })
  .catch((error) => {
    console.error("Error processing Excel file:", error);
  });
