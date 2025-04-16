const fs = require("fs");
const path = require("path");
const { exec } = require("child_process");

// Define paths
const schemaFolder = path.join(__dirname, "teams");
const outputFolder = path.join(__dirname, "src", "teams");

// Ensure output folder exists
if (!fs.existsSync(outputFolder)) {
  fs.mkdirSync(outputFolder, { recursive: true });
}

// Function to process schema files
function generateTypeScriptFiles() {
  fs.readdir(schemaFolder, (err, files) => {
    if (err) {
      console.error("Error reading schema folder:", err);
      return;
    }

    files.forEach((version) => {
      const convertedVersion = version.replace(".", "d");
      const schemaDir = path.join(schemaFolder, version);
      const schemaFile = "MicrosoftTeams.schema.json";
      const outputFileName = `MicrosoftTeams.${convertedVersion}.ts`;
      const outputPath = path.join(outputFolder, outputFileName);
      const command = `cd ${schemaDir} && npx quicktype -s schema ${schemaFile} -o ${outputPath}`;
      exec(command, (error, stdout, stderr) => {
        if (error) {
          console.error(`Error generating TypeScript for ${version}:`, error);
          return;
        }
        if (stderr) {
          console.error(`stderr for ${version}:`, stderr);
        }
        console.log(`Generated TypeScript file: ${outputPath}`);
      });
    });
  });
}

// Run the function
generateTypeScriptFiles();
