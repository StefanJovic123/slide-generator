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

const templatePath = path.join(__dirname, "pptx-templates", "RootTemplateWithImages.pptx");
const outputDir = path.join(__dirname, "output");
const outputPath = path.join(outputDir, "presentation.pptx");

const bulletPoints = ['first line', 'second line', 'third line'].join(`
  `);

async function createPresentation() {
  // Ensure the output directory exists
  await fs.mkdir(outputDir, { recursive: true });

  const data1 = {
    body: [
      { label: 'item test r1', values: ['Header 1', 'header 2', 'header 3', 'header 4', 'header 4'] },
      { label: 'item test r1', values: ['test1', 10, 16, 12, 11] },
      { label: 'item test r2', values: ['test2', 12, 18, 15, 12] },
      { label: 'item test r3', values: ['test3', 14, 12, 11, 14] },
    ],
  };


  // Initialize Automizer with preferences
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/output`,
    mediaDir: `${__dirname}/images`,
    removeExistingSlides: true,
    useCreationIds: true,
    cleanup: true,
  });

  // Log the template directory to verify the path
  console.log(`Template directory: ${path.join(__dirname, 'templates')}`);


  // Load the root template and additional templates
  let pres = automizer
    .loadRoot(templatePath)
    .loadMedia([`test.png`])
    .load(`RootTemplateWithImages.pptx`, 'base')
    .load(`SlideWithTables.pptx`, 'tables')
    .load(`EmptySlide.pptx`, 'empty')
    .load(`SlideWithCharts.pptx`, 'charts')
    .load(`twoTextElementPres.pptx`)
    .load(`TextReplace.pptx`)
    .load(`SlideWithImages.pptx`, 'images')

    await pres.addSlide('base', 1, async (slide) => {
      slide.modifyElement('Titel 1', modify.setText('My Awesome title'))
      slide.modifyElement('Untertitel 2', modify.setText('Description'))
    })


    await pres
      .addSlide('charts', 2, async (slide) => {
      slide.useSlideLayout(2);
      slide.modifyElement('ColumnChart', [
        modify.setChartData({
          series: [
            {label: 'series 1'},
            {label: 'series 2'},
            {label: 'series 3'},
          ],
          categories: [
            {label: 'cat 2-1', values: [50, 50, 20]},
            {label: 'cat 2-2', values: [14, 50, 20]},
            {label: 'cat 2-3', values: [15, 50, 20]},
            {label: 'cat 2-4', values: [26, 50, 20]}
          ]
        }),
        ModifyShapeHelper.setPosition({
          x: CmToDxa((25.4 - 15) / 2 - 1), // Center horizontally and shift left
          y: CmToDxa(6), // Position 6 cm from the top
          w: CmToDxa(7.5), // Element width
          h: CmToDxa(10), // Element height
        })
      ]);

      slide.modifyElement('PieChart', [
        modify.setChartData({
          series: [{ label: 'Pie Chart title' }],
          categories: [
            { label: 'cat 1-1', values: [50] },
            { label: 'cat 1-2', values: [14] },
          ],
        }),
        ModifyShapeHelper.setPosition({
          x: CmToDxa((25.4 - 15) / 2 + 11), // Center horizontally and shift right
          y: CmToDxa(6), // Position 6 cm from the top
          w: CmToDxa(7.5), // Element width
          h: CmToDxa(10), // Element height
        })
      ]);
    })

    await pres
    .addSlide('tables', 3, async (slide) => {
      slide.useSlideLayout(3);
      slide.removeElement('TableWithEmptyCells')
      slide.removeElement('TableWithFormattedCells')
      // slide.removeElement('Titel 2')

      slide.addElement('tables', 3, 'Titel 2', [
        modify.setText('Table title'),
        ModifyShapeHelper.setPosition({
          x: CmToDxa((25.4 - 15) / 2), // Center horizontally (slide width - element width) / 2
          y: CmToDxa(4), // Position 4 cm from the top
          w: CmToDxa(15), // Element width
          h: CmToDxa(2), // Element height
        })
      ]);

      slide.modifyElement('EmptyTable', [
        modify.setTable(data1),
        ModifyShapeHelper.setPosition({
          x: CmToDxa((25.4 - 15) / 2), // Center horizontally (slide width - element width) / 2
          y: CmToDxa(6), // Position 4 cm from the top
          w: CmToDxa(15), // Element width
          h: CmToDxa(5), // Element height
        }),
        // modify.dump
      ]);
    })
  

  // const data2 = {
  //   body: [
  //     {
  //       values: ['test1', 10, 16, 12, 11],
  //       styles: [
  //         {
  //           color: {
  //             type: 'srgbClr',
  //             value: '00FF00',
  //           },
  //           background: {
  //             type: 'srgbClr',
  //             value: 'CCCCCC',
  //           },
  //           isItalics: true,
  //           isBold: true,
  //           size: 1200,
  //         },
  //       ],
  //     },
  //     {
  //       values: ['test2', 12, 18, 15, 12],
  //       styles: [
  //         null,
  //         null,
  //         null,
  //         null,
  //         {
  //           // If you want to style a cell border, you
  //           // need to style adjacent borders as well:
  //           border: [
  //             {
  //               // This is required to complete top border
  //               // of adjacent cell in row below:
  //               tag: 'lnB',
  //               type: 'solid',
  //               weight: 5000,
  //               color: {
  //                 type: 'srgbClr',
  //                 value: '00FF00',
  //               },
  //             },
  //           ],
  //         },
  //       ],
  //     },
  //     {
  //       values: ['test3', 14, 12, 11, 14],
  //       styles: [
  //         null,
  //         null,
  //         null,
  //         {
  //           border: [
  //             {
  //               tag: 'lnR',
  //               type: 'solid',
  //               weight: 5000,
  //               color: {
  //                 type: 'srgbClr',
  //                 value: '00FF00',
  //               },
  //             },
  //           ],
  //         },
  //         {
  //           color: {
  //             type: 'srgbClr',
  //             value: 'FF0000',
  //           },
  //           background: {
  //             type: 'srgbClr',
  //             value: 'ffffff',
  //           },
  //           isItalics: true,
  //           isBold: true,
  //           size: 600,
  //           border: [
  //             {
  //               // This will only work in case you style
  //               // adjacent cell in row above with 'lnB':
  //               tag: 'lnT',
  //               type: 'solid',
  //               weight: 5000,
  //               color: {
  //                 type: 'srgbClr',
  //                 value: '00FF00',
  //               },
  //             },
  //             {
  //               tag: 'lnB',
  //               type: 'solid',
  //               weight: 5000,
  //               color: {
  //                 type: 'srgbClr',
  //                 value: '00FF00',
  //               },
  //             },
  //             {
  //               tag: 'lnL',
  //               type: 'solid',
  //               weight: 5000,
  //               color: {
  //                 type: 'srgbClr',
  //                 value: '00FF00',
  //               },
  //             },
  //             {
  //               tag: 'lnR',
  //               type: 'solid',
  //               weight: 5000,
  //               color: {
  //                 type: 'srgbClr',
  //                 value: '00FF00',
  //               },
  //             },
  //           ],
  //         },
  //       ],
  //     },
  //   ],
  // };

  await pres
  .addSlide('images', 1, async (slide) => {
    slide.useSlideLayout(2);
    const elements = await slide.getAllElements();
    elements.forEach(element => {
      slide.removeElement({ name: element.name })
    })

    slide.addElement('images', 1, elements[0].name);
  })

  // Save the presentation
  await pres.write("presentation.pptx");
  console.log(`Presentation saved to ${outputPath}`);
}

createPresentation().catch((err) => console.error(err));
