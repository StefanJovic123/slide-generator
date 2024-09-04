const fs = require('fs');
const { google } = require('googleapis');
const readline = require('readline');
const path = require('path');
const TOKEN_PATH = 'token.json';
const credentials = require('./credentials.json');

authorize(credentials.web, createPresentation);

// Authorize with OAuth2
function authorize(credentials, callback) {
  console.log(credentials); 
  const { client_secret, client_id, redirect_uris } = credentials;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getAccessToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
  });
}

function getAccessToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: ['https://www.googleapis.com/auth/presentations'],
  });
  console.log('Authorize this app by visiting this url:', authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question('Enter the code from that page here: ', (code) => {
    rl.close();
    oAuth2Client.getToken((err, token) => {
      if (err) return console.error('Error retrieving access token', err);
      oAuth2Client.setCredentials(token);
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) return console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}

// Main function to create the presentation
async function createPresentation(auth) {
  const slidesService = google.slides({ version: 'v1', auth });

  // Create a new presentation from a template
  const templatePresentationId = '1uFhAjxnZvl13AWgtjJwPS0GO_M9U3Ajfh-9nveiFzbQ'; // Set the ID of your Google Slides template
  const newPresentation = await slidesService.presentations.create({
    title: 'Generated Presentation',
  });
  const presentationId = newPresentation.data.presentationId;

  // Copy slides from the template
  const templateSlides = await slidesService.presentations.get({ presentationId: templatePresentationId });
  const slideIds = templateSlides.data.slides.map(slide => slide.objectId);
  
  // Copy the slides from the template into the new presentation
  await slidesService.presentations.batchUpdate({
    presentationId: presentationId,
    resource: {
      requests: slideIds.map(slideId => ({
        duplicateObject: {
          objectId: slideId,
          objectIds: {},
        },
      })),
    },
  });

  console.log(`Created new presentation with ID: ${presentationId}`);

  // Update Slide 1: Insert an Image
  await insertImage(auth, presentationId, slideIds[0], 'images/your_image.png');

  // Update Slide 2: Add Title and Table
  await replaceTextPlaceholders(auth, presentationId, { title: 'Sales Data Table' }, slideIds[1]);
  await addTable(auth, presentationId, slideIds[1]);

  // Update Slide 3: Add Bar Chart and Pie Chart
  await insertCharts(auth, presentationId, slideIds[2]);

  console.log('Presentation updated successfully!');
}

// Function to insert an image in slide 1
async function insertImage(auth, presentationId, slideId, imagePath) {
  const slidesService = google.slides({ version: 'v1', auth });

  const imageFile = fs.readFileSync(imagePath);
  const imageBase64 = Buffer.from(imageFile).toString('base64');

  const requests = [
    {
      createImage: {
        url: `data:image/jpeg;base64,${imageBase64}`,
        elementProperties: {
          pageObjectId: slideId,
          size: {
            height: { magnitude: 200, unit: 'PT' },
            width: { magnitude: 300, unit: 'PT' },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 100,
            translateY: 100,
            unit: 'PT',
          },
        },
      },
    },
  ];

  await slidesService.presentations.batchUpdate({
    presentationId,
    resource: { requests },
  });

  console.log('Image inserted into slide 1');
}

// Function to replace text placeholders
async function replaceTextPlaceholders(auth, presentationId, placeholders, slideId) {
  const slidesService = google.slides({ version: 'v1', auth });

  const requests = Object.keys(placeholders).map((placeholder) => ({
    replaceAllText: {
      containsText: {
        text: `{{${placeholder}}}`,
        matchCase: true,
      },
      replaceText: placeholders[placeholder],
    },
  }));

  await slidesService.presentations.batchUpdate({
    presentationId,
    resource: { requests },
  });

  console.log('Text placeholders replaced');
}

// Function to add a table in slide 2
async function addTable(auth, presentationId, slideId) {
  const slidesService = google.slides({ version: 'v1', auth });

  const tableData = [
    ['Product', 'Q1', 'Q2', 'Q3', 'Q4'],
    ['Apples', '10', '15', '20', '25'],
    ['Oranges', '20', '25', '30', '35'],
    ['Bananas', '30', '35', '40', '45'],
  ];

  const requests = [
    {
      createTable: {
        elementProperties: {
          pageObjectId: slideId,
          size: {
            height: { magnitude: 150, unit: 'PT' },
            width: { magnitude: 400, unit: 'PT' },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 50,
            translateY: 150,
            unit: 'PT',
          },
        },
        rows: tableData.length,
        columns: tableData[0].length,
      },
    },
    ...tableData.flatMap((row, rowIndex) =>
      row.map((cell, columnIndex) => ({
        insertText: {
          objectId: slideId, // Assuming the table object has the same ID as the slide
          cellLocation: { rowIndex, columnIndex },
          text: cell,
        },
      }))
    ),
  ];

  await slidesService.presentations.batchUpdate({
    presentationId,
    resource: { requests },
  });

  console.log('Table inserted into slide 2');
}

// Function to add a bar and pie chart in slide 3
async function insertCharts(auth, presentationId, slideId) {
  const slidesService = google.slides({ version: 'v1', auth });

  const chartRequests = [
    {
      createSheetsChart: {
        spreadsheetId: '1UYhtEzCtbPtGyU_MF1RXWinNKiFZmt_dmYgf6BKKR6I', // This must be a Google Sheets document with your data
        chartId: 8766122, // This is the chart ID in Google Sheets
        linkingMode: 'LINKED',
        elementProperties: {
          pageObjectId: slideId,
          size: {
            height: { magnitude: 200, unit: 'PT' },
            width: { magnitude: 300, unit: 'PT' },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 50,
            translateY: 100,
            unit: 'PT',
          },
        },
      },
    },
    {
      createSheetsChart: {
        spreadsheetId: '1UYhtEzCtbPtGyU_MF1RXWinNKiFZmt_dmYgf6BKKR6I',
        chartId: 76128633, // Another chart ID for the pie chart
        linkingMode: 'LINKED',
        elementProperties: {
          pageObjectId: slideId,
          size: {
            height: { magnitude: 200, unit: 'PT' },
            width: { magnitude: 300, unit: 'PT' },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 400,
            translateY: 100,
            unit: 'PT',
          },
        },
      },
    },
  ];

  await slidesService.presentations.batchUpdate({
    presentationId,
    resource: { requests: chartRequests },
  });

  console.log('Charts inserted into slide 3');
}