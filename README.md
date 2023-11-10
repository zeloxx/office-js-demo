# Demo Excel Add-In Guide

## Notes
- The backend is only boilerplate and doesn't do anything for frontend yet
- Most of the add-in logic can be found here: [App.tsx](https://github.com/zeloxx/office-js-demo/blob/main/yo/demo-example-1/src/taskpane/components/App.tsx)

## Setup and Run Instructions

### Initial Setup

1. Navigate to the project directory:

   ```sh
   cd yo/demo-example-1
   ```

2. Install the necessary dependencies:

   ```sh
   npm install
   ```

3. Build the CSS:
   ```sh
   npm run build:css
   ```

### Development

- To update styles in real-time during development, run the following command:
  ```sh
  npm run watch:css
  ```

### Running the App Locally

- To start the application:
  ```sh
  npm run start
  ```
  This command will open the Excel application and load the add-in for you to view.

## Styling

- TailwindCSS is used for styling due to familiarity with the framework, ensuring a streamlined development process.

## To-Do List

- [ ] Refactor and organize logic into reusable React hooks.
- [ ] Break down the user interface into smaller React components, such as buttons, inputs, and typography elements.
- [ ] Implement functionality to send JSON data extracted from CSV files to the backend for processing.
- [ ] Upon receiving the processed response, generate a new Excel sheet populated with data in the desired format.

## Preview

For a quick look at the current UI, see the screenshot below:

![Excel Add-In Screenshot](https://i.gyazo.com/1e4d593a606ff4692ea3667c50bb2a01.png)
