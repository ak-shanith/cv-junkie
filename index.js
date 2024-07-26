const fs = require("fs");
const path = require("path");
const pdfParse = require("pdf-parse");
const xlsx = require("xlsx");
const OpenAI = require("openai");

// OpenAI API configuration
const openaiApiKey = "";
const projectId = "";
const orgId = "";

const openai = new OpenAI({
  apiKey: openaiApiKey,
  organization: orgId,
  project: projectId,
});

// Function to read PDFs and extract text
async function readPDFs(folderPath, jobDescription) {
  const files = fs.readdirSync(folderPath);
  const pdfFiles = files.filter(
    (file) => path.extname(file).toLowerCase() === ".pdf"
  );
  const extractedData = [];

  for (const file of pdfFiles) {
    const filePath = path.join(folderPath, file);
    const dataBuffer = fs.readFileSync(filePath);
    const data = await pdfParse(dataBuffer);
    const text = data.text;

    // Send text to OpenAI API for processing
    const extractedInfo = await extractInfoWithGPT(text, jobDescription);
    extractedData.push(extractedInfo);
  }

  return extractedData;
}

// Function to extract information using GPT
async function extractInfoWithGPT(text, jobDescription) {
  const prompt = `
Extract the following information from the given CV text based on the job description:

Job Description:
${jobDescription}

CV Text:
${text}

Extracted Information:
- Name:
- Email:
- Phone:
- Skills:
- Experience:
- Education:

Rate this candidate on a scale of 1 to 5 based on how suitable the candidate for the job.

Output should be in JSON format.
Output JSON should only contain following Keys: [Name, Email, Phone, Skills, Experience, Education, Candidate Rating, Reason for rating]
And the value for above keys should be in string format.
Clean and trip output as needed.
`;
  try {
    const response = await openai.chat.completions.create({
      model: "gpt-4o",
      messages: [
        {
          role: "user",
          content: prompt,
        },
      ],
      max_tokens: 3000,
      temperature: 0.7,
    });
    console.log(response.choices[0].message.content);
    const extractedText = response.choices[0].message.content;
    const extractedInfo = parseExtractedText(extractedText);
    return extractedInfo;
  } catch (error) {
    console.error("Error extracting information:", error);
    return {};
  }
}

// Analyze CVs with the job description
async function analyzeCVs(jobDescription, cvList) {
  const prompt = `
  Analyze the following CVs in the context of the job description.
  
  Job Description:
  ${jobDescription}
  
  CVs:
  ${JSON.stringify(cvList, null, 2)}

  Give a suitability score for each CV based on the job description.
  Calculate the suitability score for each candidate based on their alignment with the job requirements for Senior DevOps Engineer role.
  
  Consider the following criteria:
  Skills Match (40%)
  Experience (50%)
  Education (10%)
  
  Scoring
  Assign points for each criterion (out of 100) and calculate the total suitability score.

  Output should be in JSON format. Sort the candidates based on their suitability score.
`;

  try {
    const response = await openai.chat.completions.create({
      model: "gpt-4o",
      messages: [
        {
          role: "system",
          content: "You are an expert in job matching and CV analysis.",
        },
        { role: "user", content: prompt },
      ],
      max_tokens: 3000, // Adjust based on the expected length of response
      temperature: 0.7,
    });

    return response.choices[0].message.content;
  } catch (error) {
    console.error("Error analyzing CVs:", error);
    throw error;
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

// Function to read the job description from a text file
function readJobDescription(filePath) {
  return fs.readFileSync(filePath, "utf-8");
}

async function main() {
  const folderPath = "./cvs"; // Change to your folder path
  const outputFilePath = "./result.xlsx";
  const jobDescriptionPath = "./jd.txt"; // Path to your job description file

  const jobDescription = readJobDescription(jobDescriptionPath);
  const extractedData = await readPDFs(folderPath, jobDescription);

  writeToExcel(extractedData, outputFilePath);
  console.log(`Extracted data has been written to ${outputFilePath}`);

  const analysisResult = await analyzeCVs(jobDescription, extractedData);
  console.log("Analysis Result:", analysisResult);
}

main().catch(console.error);
