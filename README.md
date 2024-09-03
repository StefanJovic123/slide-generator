# Slide Generator

Slide Generator is a Proof of Concept (PoC) project for generating slides using branded templates. It leverages `pptx-automizer` and `nodejs-pptx` for creating PowerPoint presentations and `googleapis` for uploading them to Google Slides.

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
  - [Create PPTX](#create-pptx)
  - [Upload to Google Slides](#upload-to-google-slides)
  - [Create PPTX with nodejs-pptx](#create-pptx-with-nodejs-pptx)
- [Testing](#testing)
- [Dependencies](#dependencies)
- [Contributing](#contributing)
- [License](#license)

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/slide-generator.git
    cd slide-generator
    ```

2. Install the dependencies:
    ```sh
    npm install
    ```

3. Ensure you have a `credentials.json` file for Google API authentication in the root directory.

## Usage

### Create PPTX

To generate a PowerPoint presentation using `pptx-automizer`, run:
    ```sh
    npm run create-pptx
    ```

### Upload to Google Slides

To upload the generated PowerPoint presentation to Google Slides, run:
    ```sh
    npm run upload-to-google-slides
    ```

### Create PPTX with nodejs-pptx

To generate a PowerPoint presentation using `nodejs-pptx`, run:
    ```sh
    npm run create-pptx-nodejs
    ```

## Testing

To run tests, use:
    ```sh
    npm test
    ```

## Dependencies

- `pptx-automizer`
- `nodejs-pptx`
- `googleapis`
