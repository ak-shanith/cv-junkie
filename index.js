const fs = require("fs");
const path = require("path");
const pdfParse = require("pdf-parse");
const xlsx = require("xlsx");
const OpenAI = require("openai");
require("dotenv").config();

// OpenAI API configuration
const openaiApiKey = process.env.OPENAI_API_KEY;
const projectId = process.env.OPENAI_PROJECT_ID;
const orgId = process.env.OPENAI_ORG_ID;

// Paths and file names
const folderPath = "./cvs"; // Change to your folder path
const outputFilePath = "./out/result.xlsx";
const jobDescriptionPath = "./jd.txt"; // Path to your job description file

const openai = new OpenAI({
  apiKey: openaiApiKey,
  organization: orgId,
  project: projectId,
});

// Function to read PDFs and extract candidates' information
async function readPDFs(folderPath, jobDescription) {
  const files = fs.readdirSync(folderPath);
  const pdfFiles = files.filter(
    (file) => path.extname(file).toLowerCase() === ".pdf"
  );
  const candidates = [];

  for (const file of pdfFiles) {
    const filePath = path.join(folderPath, file);
    const dataBuffer = fs.readFileSync(filePath);
    const data = await pdfParse(dataBuffer);
    const text = data.text;

    // Send text to OpenAI API for processing
    const candidate = await extractInfoWithGPT(text, jobDescription);
    candidates.push(candidate);
  }

  // Convert 'Score' to number and sort the array
  candidates.forEach((item) => {
    item.Score = Number(item.Score);
  });
  candidates.sort((a, b) => b.Score - a.Score);

  return candidates;
}

// Function to extract information using GPT
async function extractInfoWithGPT(candidateCV, jobDescription) {
  const prompt = `
    Here is a job description and a CV for analysis. Please follow the instructions to extract and rate the candidate's suitability.

    Extract the following information from the given CV text based on the job description:

    Job Description:
    ${jobDescription}

    CV Text:
    ${candidateCV}

    Extracted Information:
    - Name:
    - Email:
    - Phone:
    - Skills:
    - Experience:
    - Education:

    Rate this candidate on a scale of 1 to 5 based on how suitable the candidate for the job.

    Also give a overall suitability score based on the job description.
    Calculate the suitability score for the candidate based on their alignment with the job requirements.
      
    Consider the following criteria:
    - Skills Match (40%)
    - Experience (50%)
    - Education (10%)
      
    Scoring:
    - Assign points for each criterion (out of 100) and calculate the total suitability score.

    Output should be in JSON format and should only contain the following keys: 
    [Name, Email, Phone, Skills, Experience, Education, Rating, Reason, Score].

    The values for the keys should be in string format.
    Clean and trim the output as needed.
    `;
  try {
    const response = await openai.chat.completions.create({
      model: "gpt-4o-mini",
      messages: [
        {
          role: "system",
          content:
            "You are an expert in job matching and CV analysis. Your task is to extract and analyze information from CVs based on given job descriptions.",
        },
        {
          role: "user",
          content: prompt,
        },
      ],
      max_tokens: 4096,
      temperature: 0.7,
    });
    const extractedText = response.choices[0].message.content;
    const extractedInfo = parseExtractedText(extractedText);
    console.log(extractedInfo);
    return extractedInfo;
  } catch (error) {
    console.error("Error extracting information:", error);
    return {};
  }
}

// Function to parse the extracted information
function parseExtractedText(extractedText) {
  const v1 = extractedText.replace("```json", "");
  const v2 = v1.replace("```", "");
  const info = JSON.parse(v2);
  return info;
}

// Function to write data to Excel
function writeToExcel(data, outputFilePath) {
  const workbook = xlsx.utils.book_new();
  const worksheet = xlsx.utils.json_to_sheet(data);
  xlsx.utils.book_append_sheet(workbook, worksheet, "CVs");
  xlsx.writeFile(workbook, outputFilePath);
}

async function main() {
  const jd = fs.readFileSync(jobDescriptionPath, "utf-8");
  const candidates = await readPDFs(folderPath, jd);
  writeToExcel(candidates, outputFilePath);

  console.log(`Extracted data has been saved to ${outputFilePath}`);
}

main().catch(console.error);
