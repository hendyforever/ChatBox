# ChatBox Plugin

## Overview
ChatBox is an Office Add-in that provides a task pane for enhanced user interaction. It allows users to upload files, interact with AI, and manage chat history seamlessly within Microsoft Office applications.

## Project Structure
```
ChatBox
├── src
│   ├── taskpane
│   │   ├── taskpane.html      # HTML structure for the task pane
│   │   └── taskpane.js        # JavaScript functionality for the task pane
│   ├── manifest.xml           # Manifest file for the Office Add-in
│   └── assets
│       └── style.css          # CSS styles for the task pane
├── package.json                # npm configuration file
└── README.md                   # Project documentation
```

## Installation
1. Clone the repository:
   ```
   git clone <repository-url>
   cd ChatBox
   ```

2. Install dependencies:
   ```
   npm install
   ```

## Usage
1. Open the project in your preferred code editor.
2. Build the project using:
   ```
   npm run build
   ```

3. Load the add-in in your Office application:
   - For Excel or Word, go to `Insert` > `My Add-ins` > `Manage My Add-ins` > `Upload My Add-in` and select the `manifest.xml` file.

## Features
- **File Upload**: Users can upload Word documents and other file types for processing.
- **Chat History**: The task pane maintains a history of user interactions.
- **AI Integration**: The add-in integrates with an AI service to provide responses based on user input and uploaded content.

## Contributing
Contributions are welcome! Please submit a pull request or open an issue for any enhancements or bug fixes.

## License
This project is licensed under the MIT License. See the LICENSE file for details.