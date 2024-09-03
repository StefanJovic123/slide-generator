const fs = require("fs");
const { google } = require("googleapis");
const { authenticate } = require("@google-cloud/local-auth");
const path = require("path");

async function uploadToGoogleSlides() {
  const auth = await authenticate({
    keyfilePath: path.join(__dirname, "credentials.json"),
    scopes: ["https://www.googleapis.com/auth/presentations"],
  });

  const slides = google.slides({ version: "v1", auth });

  const presentation = fs.readFileSync("output/presentation.pptx");

  const fileMetadata = {
    name: "Generated Presentation",
    mimeType: "application/vnd.google-apps.presentation",
  };

  const media = {
    mimeType:
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    body: presentation,
  };

  const drive = google.drive({ version: "v3", auth });
  const response = await drive.files.create({
    resource: fileMetadata,
    media: media,
    fields: "id",
  });

  console.log(
    `Google Slides file created: https://docs.google.com/presentation/d/${response.data.id}/edit`
  );
}

uploadToGoogleSlides().catch(console.error);
