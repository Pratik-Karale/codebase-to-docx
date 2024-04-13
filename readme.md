# Codebase to docx file convertor

This is a command-line tool for converting code snippets into a formatted DOCX document. It allows you to specify an input directory containing code files (JavaScript, Python, Java, C++, C#, Ruby), ignores specific paths, and generates a DOCX document with basic styles.

This project was made for my college project submission where we were required to print all the code files in our respective projects. 
This was also inspired by other projects as well but doesn't rely on libreoffice and is much faster ⚡

## Prerequisites

Make sure you have [Node.js](https://nodejs.org/) installed on your machine.

## Installation

1. Clone this repository:

   ```bash
   git clone https://github.com/Pratik-Karale/codebase-to-docx.git
   ```

2. Navigate to the project directory:

   ```bash
   cd code-conversion-tool
   ```

3. Install dependencies:

   ```bash
   npm install
   ```

## Usage

1. Run the tool:

   ```bash
   npm start
   ```

2. Follow the prompts to provide the path to the input directory, specify paths to ignore (if any), and choose whether to open the converted file.

3. Once the conversion is completed, the output.docx file will be generated in the project directory.

4. Edit template.docx in the main project folder to edit styles of how the code snippets should look and run `npm start` again to generate the output file.

5. There might be a few spacing issues in the generated docx so you could resolve it by manually editing it.

## Features

- Supports various programming languages including JavaScript, Python, Java, C++, C#, and Ruby.
- Allows customization of input directory and ignored paths.
- Option to open the converted file in the default software.

## Contributing

Contributions are welcome! Feel free to open issues or pull requests for any improvements or new features you'd like to see.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
