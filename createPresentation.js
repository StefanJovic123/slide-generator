const {
  Automizer,
  modify,
  ModifyTextHelper,
  ModifyShapeHelper,
  ModifyImageHelper,
  CmToDxa,
} = require("pptx-automizer");
const path = require("path");
const fs = require("fs").promises;

const templatePath = path.join(__dirname, "templates", "branded_template.pptx");
const outputDir = path.join(__dirname, "output");
const outputPath = path.join(outputDir, "presentation.pptx");

async function createPresentation() {
  // Ensure the output directory exists
  await fs.mkdir(outputDir, { recursive: true });

  // Initialize Automizer with preferences
  const automizer = new Automizer({
    templateDir: `templates`,
    outputDir: `output`,
    useCreationIds: false,
    autoImportSlideMasters: true,
    removeExistingSlides: true,
    cleanup: false,
    compression: 0,
  });

  // Load the root template and additional templates
  let pres = automizer
    .loadRoot(templatePath)
    .load("SlideWithShapes.pptx", "shapes")
    .load("SlideWithGraph.pptx", "graph")
    .load("SlideWithImages.pptx", "images");

  // Create a new slide with text and table
  await pres.addSlide("shapes", 1, (slide) => {
    slide.addElement("shapes", 1, "TextBox", [
      ModifyTextHelper.setText("{{Introduction}}"),
      ModifyShapeHelper.setPosition({
        x: CmToDxa(1),
        y: CmToDxa(0.5),
        w: CmToDxa(10),
        h: CmToDxa(2),
      }),
    ]);

    slide.addElement("shapes", 1, "Table", [
      modify.setTable({
        body: [
          ["Header 1", "Header 2", "Header 3"],
          ["Row 1", "Data 1", "Data 2"],
          ["Row 2", "Data 1", "Data 2"],
        ],
      }),
      ModifyShapeHelper.setPosition({
        x: CmToDxa(0.5),
        y: CmToDxa(3),
        w: CmToDxa(15),
        h: CmToDxa(5),
      }),
    ]);
  });

  // Create a new slide with bar and pie chart
  await pres.addSlide("graph", 1, (slide) => {
    slide.addElement("graph", 1, "ChartTitle", [
      ModifyTextHelper.setText("{{Charts}}"),
      ModifyShapeHelper.setPosition({
        x: CmToDxa(1),
        y: CmToDxa(0.5),
        w: CmToDxa(10),
        h: CmToDxa(2),
      }),
    ]);

    slide.addElement("graph", 1, "BarChart", [
      modify.setChartData({
        series: [
          { label: "Category 1", values: [30, 40, 50] },
          { label: "Category 2", values: [20, 50, 60] },
        ],
      }),
      ModifyShapeHelper.setPosition({
        x: CmToDxa(0.5),
        y: CmToDxa(3),
        w: CmToDxa(10),
        h: CmToDxa(5),
      }),
    ]);

    slide.addElement("graph", 1, "PieChart", [
      modify.setChartData({
        series: [
          { label: "Slice 1", values: [30] },
          { label: "Slice 2", values: [70] },
        ],
      }),
      ModifyShapeHelper.setPosition({
        x: CmToDxa(11),
        y: CmToDxa(3),
        w: CmToDxa(10),
        h: CmToDxa(5),
      }),
    ]);
  });

  // Create a new slide with an image
  await pres.addSlide("images", 1, (slide) => {
    slide.addElement("images", 1, "TextBox", [
      ModifyTextHelper.setText("{{Image}}"),
      ModifyShapeHelper.setPosition({
        x: CmToDxa(1),
        y: CmToDxa(0.5),
        w: CmToDxa(10),
        h: CmToDxa(2),
      }),
    ]);

    slide.addElement("images", 1, "ImagePlaceholder", [
      ModifyImageHelper.setRelationTarget(
        path.join(__dirname, "images", "your_image.png")
      ),
      ModifyShapeHelper.setPosition({
        x: CmToDxa(1),
        y: CmToDxa(3),
        w: CmToDxa(15),
        h: CmToDxa(10),
      }),
    ]);
  });

  // Replace placeholders with actual values
  await pres.modifyElement("shapes", 1, "TextBox", [
    modify.replaceText([
      { replace: "Introduction", by: { text: "Welcome to the Presentation" } },
    ]),
  ]);

  await pres.modifyElement("graph", 1, "ChartTitle", [
    modify.replaceText([{ replace: "Charts", by: { text: "Sales Data" } }]),
  ]);

  await pres.modifyElement("images", 1, "TextBox", [
    modify.replaceText([{ replace: "Image", by: { text: "Company Logo" } }]),
  ]);

  // Save the presentation
  await pres.write(outputPath);
  console.log(`Presentation saved to ${outputPath}`);
}

createPresentation().catch((err) => console.error(err));
