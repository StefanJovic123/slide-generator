const PptxGenJS = require("nodejs-pptx");

async function createPptxWithNodejsPptx() {
  const pptx = new PptxGenJS();

  // Load branded styles from a branded template
  await pptx.load("templates/branded_template.pptx");
  console.log("Branded template loaded successfully.");

  // Now create the new slides using the loaded branded styles
  // Slide 1: Text and Table
  let slide1 = pptx.addSlide();
  slide1.addText("Introduction", {
    x: 1,
    y: 0.5,
    fontSize: 32,
    color: "000000",
  });

  let tableData = [
    [
      { text: "Header 1", options: { fill: "FF0000" } },
      { text: "Header 2", options: { fill: "00FF00" } },
      { text: "Header 3", options: { fill: "0000FF" } },
    ],
    ["Row 1", "Data 1", "Data 2"],
    ["Row 2", "Data 1", "Data 2"],
  ];
  slide1.addTable(tableData, { x: 1, y: 1.5, w: 8 });

  // Slide 2: Charts
  let slide2 = pptx.addSlide();
  slide2.addText("Charts", { x: 1, y: 0.5, fontSize: 32, color: "000000" });

  let barChartData = [
    {
      name: "Category 1",
      labels: ["2010", "2020", "2030"],
      values: [30, 40, 50],
    },
    {
      name: "Category 2",
      labels: ["2010", "2020", "2030"],
      values: [20, 50, 60],
    },
  ];
  slide2.addChart(pptx.ChartType.bar, barChartData, {
    x: 0.5,
    y: 1.5,
    w: 4,
    h: 3,
  });

  let pieChartData = [
    { name: "Slice 1", labels: ["A"], values: [30] },
    { name: "Slice 2", labels: ["B"], values: [70] },
  ];
  slide2.addChart(pptx.ChartType.pie, pieChartData, {
    x: 5,
    y: 1.5,
    w: 4,
    h: 3,
  });

  // Slide 3: Image
  let slide3 = pptx.addSlide();
  slide3.addText("Image", { x: 1, y: 0.5, fontSize: 32, color: "000000" });
  slide3.addImage({ path: "images/your_image.png", x: 1, y: 1.5, w: 6, h: 4 });

  await pptx.writeFile({ fileName: "output/presentation_nodejs_pptx.pptx" });
  console.log("Presentation created with nodejs-pptx");
}

createPptxWithNodejsPptx().catch(console.error);
