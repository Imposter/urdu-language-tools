# Urdu Language Tools for Microsoft Word

This project is a VSTO (Visual Studio Tools for Office) plugin for Microsoft Word that provides tools for formatting Urdu poetry, including Ghazals, Nazams, and Nasars.

## Features

- Format Ghazals with specific styles and lines per verse.
- Add formatted Ghazals to the table of contents.
- Remove multiple spaces from selected text.
- Paste and format Ghazal text from the clipboard.

![Controls](docs/images/001_controls.png)
![Pasting](docs/images/002_pasting_ghazal.gif)

## Building the Project

To build the project, follow these steps:

1. Clone the repository:
    ```sh
    git clone https://github.com/imposter/UrduLanguageTools.git
    cd UrduLanguageTools
    ```

2. Open the solution file `UrduLanguageTools.sln` in JetBrains Rider or Visual Studio.

3. Restore the NuGet packages:
    ```sh
    dotnet restore
    ```

4. Build the solution:
    ```sh
    dotnet build
    ```

5. Run the project to start debugging the VSTO plugin in Microsoft Word.

## Contributing

We welcome contributions to improve the Urdu Language Tools plugin. To contribute, follow these steps:

1. Fork the repository.

2. Create a new branch for your feature or bugfix:
    ```sh
    git checkout -b feature-or-bugfix-name
    ```

3. Make your changes and commit them with a descriptive message:
    ```sh
    git commit -am "Add new feature or fix bug"
    ```

4. Push your changes to your forked repository:
    ```sh
    git push origin feature-or-bugfix-name
    ```

5. Create a pull request to the main repository.

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.